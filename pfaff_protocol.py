"""
PFAFF protocol handler
Handles communication protocol with the sewing machine
"""

import os
import struct
import logging
import time

logger = logging.getLogger(__name__)


class PFAFFProtocol:
    """PFAFF sewing machine protocol handler"""
    
    # Control characters
    CTRL_ETX = 0x03 # End of Text
    CTRL_EOT = 0x04 # End of Transmission
    CTRL_ENQ = 0x05 # Enquiry
    CTRL_ACK = 0x06 # Acknowledge
    CTRL_BEL = 0x07 # Bell
    CTRL_NAK = 0x15 # Negative Acknowledge
    CTRL_ETB = 0x17 # End of Transmission Block

    # Protocol constants
    COMMAND_READ_PMEMORY = 0x01
    COMMAND_WRITE_PMEMORY = 0x02
    COMMAND_READ_MMEMORY = 0x03
    COMMAND_WRITE_MMEMORY = 0x04
    COMMAND_STATUS = 0x05
    COMMAND_EXECUTE = 0x06
    
    # Response codes
    RESPONSE_OK = 0x00
    RESPONSE_ERROR = 0x01
    RESPONSE_BUSY = 0x02

    # Text command strings (received without the terminating CTRL_ETX)
    CMD_LIST_PMEMORY = "PI"
    CMD_DELETE_PMEMORY_PREFIX = "PL"   # followed by 2 hex-ASCII chars = slot number
    CMD_WRITE_PMEMORY_PREFIX = "PN"   # followed by 11 hex-ASCII chars: slot(2)+size(6)+CTRL_ETB+checksum(2)
    CMD_READ_PMEMORY_PREFIX  = "RM"   # followed by 5 chars: 06(2)+slot(2)+type(1)

    # Internal state machine states
    _STATE_IDLE = 0
    _STATE_WRITE_DATA = 1         # collecting chunk data bytes
    _STATE_WRITE_CHECKSUM = 2     # collecting 2-byte hex checksum after chunk terminator
    _STATE_WAIT_ACK = 3           # waiting for CTRL_ACK from host after PI response
    _STATE_READ_WAIT_ACK = 4      # waiting for CTRL_ACK after a read chunk
    _STATE_READ_WAIT_EOT = 5      # waiting for CTRL_EOT after final-chunk ACK
    _STATE_WRITE_HEADER = 6       # collecting header bytes after PN init ACK

    # Read chunk size (max ASCII chars per chunk = 2 * raw bytes)
    READ_CHUNK_SIZE = 250

    # Bell identification strings per model
    MODEL_BELL_STRINGS = {
        "PFAFF Creative 7570":   "Copyright 1992 - 97       G.M. PFAFF AG Creative 7570B    Vers. 2.1.",
        "PFAFF Creative 7550":   "Copyright 1992 - 97       G.M. PFAFF AG Creative 7550 CD  Vers. 1.0.",
        "PFAFF Creative 1475 CD": "Copyright 1992 - 97       G.M. PFAFF AG Creative 1475 CD  Vers. 1.0.",
    }

    # Bell command debounce time (seconds)
    BELL_DEBOUNCE_SECONDS = 2.0
    
    def __init__(self, machine_state=None, on_pmemory_changed=None):
        self.machine_state = machine_state
        self.on_pmemory_changed = on_pmemory_changed  # Optional callback: called when P-Memory is modified
        self._model_name = "PFAFF Creative 7570"
        self.cmd_buffer = bytearray()  # Accumulates bytes for text commands
        self.last_bell_time = 0  # Timestamp of last bell command processed

        # Write P-Memory state machine
        self._state = self._STATE_IDLE
        self._write_slot_id = None
        self._write_stitch_type = None
        self._write_expected_size = None
        self._write_data_accumulated = bytearray()
        self._write_chunk_buffer = bytearray()
        self._write_checksum_chars = bytearray()
        self._write_header_buffer = bytearray()  # raw bytes of the header message
        self._write_header = bytearray()         # stored header data (without checksum)

        # 1475CD list experiment counter (each call adds one extra '0' to ascii_data)
        self._1475cd_list_call_count = 0

        # Read P-Memory state machine
        self._read_data = bytearray()   # full hex-ASCII encoded slot data to send
        self._read_offset = 0
        self._read_last_chunk_sent = False

    def process_incoming(self, data: bytes) -> bytes:
        """
        Process bytes received from serial port.
        Uses a state machine to handle both idle text commands and multi-chunk
        P-Memory write transfers.
        Returns any response bytes that should be sent back.
        """
        response = bytearray()
        for byte in data:

            if self._state == self._STATE_IDLE:
                if byte == self.CTRL_EOT:
                    if self.cmd_buffer:
                        logger.debug(f"Discarding partial buffer on CTRL_EOT: {bytes(self.cmd_buffer)!r}")
                        self.cmd_buffer.clear()
                    self._state = self._STATE_IDLE
                    logger.info("End of transmission")
                elif byte == self.CTRL_BEL:
                    if self.cmd_buffer:
                        logger.debug(f"Discarding partial buffer on CTRL_BEL: {bytes(self.cmd_buffer)!r}")
                        self.cmd_buffer.clear()
                    response.extend(self.handle_bell_command())
                elif byte == self.CTRL_ETX:
                    cmd_str = self.cmd_buffer.decode('ascii', errors='replace')
                    self.cmd_buffer.clear()
                    response.extend(self._dispatch_text_command(cmd_str))
                else:
                    self.cmd_buffer.append(byte)

            elif self._state == self._STATE_WAIT_ACK:
                if byte == self.CTRL_EOT:
                    logger.info("EOT received while waiting for ACK - resetting to idle")
                    self._state = self._STATE_IDLE
                elif byte == self.CTRL_ACK:
                    logger.debug("ACK received from host")
                    self._state = self._STATE_IDLE
                elif byte == self.CTRL_BEL:
                    logger.warning("CTRL_BEL received while waiting for ACK - aborting wait")
                    self._state = self._STATE_IDLE
                    response.extend(self.handle_bell_command())
                else:
                    logger.warning(f"Unexpected byte 0x{byte:02X} while waiting for ACK - ignored")

            elif self._state == self._STATE_READ_WAIT_ACK:
                if byte == self.CTRL_EOT:
                    logger.info("EOT received while waiting for read ACK - resetting to idle")
                    self._abort_read()
                elif byte == self.CTRL_ACK:
                    if self._read_last_chunk_sent:
                        self._state = self._STATE_READ_WAIT_EOT
                    else:
                        response.extend(self._send_next_read_chunk())
                elif byte == self.CTRL_NAK:
                    logger.warning("Read P-Memory: NAK received, aborting")
                    self._abort_read()
                else:
                    logger.warning(f"Read P-Memory: unexpected byte 0x{byte:02X} while waiting for ACK")

            elif self._state == self._STATE_READ_WAIT_EOT:
                if byte == self.CTRL_EOT:
                    logger.info("Read P-Memory: transfer complete (EOT received) - resetting to idle")
                    self._abort_read()
                elif byte == self.CTRL_NAK:
                    logger.warning("Read P-Memory: NAK received after last chunk, aborting")
                    self._abort_read()
                else:
                    logger.warning(f"Read P-Memory: unexpected byte 0x{byte:02X} while waiting for EOT")

            elif self._state == self._STATE_WRITE_DATA:
                if byte == self.CTRL_EOT:
                    logger.info("EOT received during write data - resetting to idle")
                    self._abort_write()
                elif byte == self.CTRL_BEL:
                    logger.warning("Write P-Memory: aborted by CTRL_BEL")
                    self._abort_write()
                    response.extend(self.handle_bell_command())
                elif byte == self.CTRL_ETX:
                    # CTRL_ETX after a completed chunk checksum signals end of all chunks
                    if self._write_chunk_buffer:
                        logger.warning("Write P-Memory: unexpected CTRL_ETX mid-chunk - aborting")
                        self._abort_write()
                    else:
                        response.extend(self._commit_write())
                elif byte == self.CTRL_ETB:
                    self._write_checksum_chars = bytearray()
                    self._state = self._STATE_WRITE_CHECKSUM
                else:
                    self._write_chunk_buffer.append(byte)

            elif self._state == self._STATE_WRITE_CHECKSUM:
                if byte == self.CTRL_EOT:
                    logger.info("EOT received during write checksum - resetting to idle")
                    self._abort_write()
                else:
                    self._write_checksum_chars.append(byte)
                    if len(self._write_checksum_chars) == 2:
                        response.extend(self._process_write_chunk())

            elif self._state == self._STATE_WRITE_HEADER:
                if byte == self.CTRL_EOT:
                    logger.info("CTRL_EOT received during write header - resetting to idle")
                    self._abort_write()
                elif byte == self.CTRL_BEL:
                    logger.warning("Write P-Memory: header aborted by CTRL_BEL")
                    self._abort_write()
                    response.extend(self.handle_bell_command())
                elif byte == self.CTRL_ETX:
                    response.extend(self._process_write_header())
                else:
                    self._write_header_buffer.append(byte)

        return bytes(response)

    def _dispatch_text_command(self, cmd: str) -> bytes:
        """Dispatch a complete text command (stripped of its CTRL_ETX terminator)."""
        logger.debug(f"Text command received: {cmd!r}")
        if cmd == self.CMD_LIST_PMEMORY:
            return self.handle_list_pmemory()
        if cmd.startswith(self.CMD_DELETE_PMEMORY_PREFIX) and len(cmd) == 4:
            return self.handle_delete_pmemory(cmd[2:])
        if cmd.startswith(self.CMD_WRITE_PMEMORY_PREFIX) and (len(cmd) == 13 or len(cmd) == 21): # 13 chars for 75xx, 21 chars for 1475 CD
            return self.handle_write_pmemory_init(cmd[2:])
        if cmd.startswith(self.CMD_READ_PMEMORY_PREFIX) and len(cmd) == 7:
            return self.handle_read_pmemory_init(cmd[2:])
        logger.unknown_cmd(f"Unknown text command: {cmd!r}")
        return b""

    def create_read_pmemory_command(self, slot_id: int) -> bytes:
        """
        Create command to read P-Memory slot
        
        Args:
            slot_id: Memory slot ID (0-31)
        
        Returns:
            Command bytes
        """
        if not (0 <= slot_id <= 31):
            raise ValueError(f"Invalid slot ID: {slot_id}")
        
        cmd = bytearray()
        cmd.append(self.COMMAND_READ_PMEMORY)
        cmd.append(slot_id)
        checksum = self._calculate_checksum(cmd)
        cmd.append(checksum)
        return bytes(cmd)
    
    def create_write_pmemory_command(self, slot_id: int, data: bytes, pattern_type: str = "9mm") -> bytes:
        """
        Create command to write P-Memory slot
        
        Args:
            slot_id: Memory slot ID (0-31)
            data: Data to write
            pattern_type: Pattern type (9mm, MAXI, etc)
        
        Returns:
            Command bytes
        """
        if not (0 <= slot_id <= 31):
            raise ValueError(f"Invalid slot ID: {slot_id}")
        
        if len(data) > 255:
            raise ValueError(f"Data too large: {len(data)} bytes")
        
        cmd = bytearray()
        cmd.append(self.COMMAND_WRITE_PMEMORY)
        cmd.append(slot_id)
        cmd.append(len(data))
        cmd.extend(data)
        checksum = self._calculate_checksum(cmd)
        cmd.append(checksum)
        return bytes(cmd)
    
    def create_status_command(self) -> bytes:
        """
        Create command to request machine status
        
        Returns:
            Command bytes
        """
        cmd = bytearray()
        cmd.append(self.COMMAND_STATUS)
        checksum = self._calculate_checksum(cmd)
        cmd.append(checksum)
        return bytes(cmd)
    
    def configure_model(self, model_name: str):
        """Set the active machine model (affects bell response string)."""
        if model_name not in self.MODEL_BELL_STRINGS:
            raise ValueError(f"Unknown model: {model_name}")
        self._model_name = model_name
        logger.info(f"Protocol model set to: {model_name}")

    def handle_bell_command(self) -> bytes:
        """
        Handle CTRL_BEL (bell/identify) command with debouncing.
        
        Returns the identification string followed by CTRL_ETX if debounce period has passed,
        otherwise returns empty bytes (command ignored).
        
        Returns:
            Identification response or empty bytes if debounced
        """
        current_time = time.time()
        time_since_last_bell = current_time - self.last_bell_time
        
        # Check if we should ignore this bell (debouncing)
        if time_since_last_bell < self.BELL_DEBOUNCE_SECONDS:
            logger.debug(f"Bell command ignored (debounced). Last bell was {time_since_last_bell:.2f}s ago")
            return b""
        
        # Update last bell time
        self.last_bell_time = current_time
        
        # Build response: identification string + CTRL_ETX
        resp_ident = self.MODEL_BELL_STRINGS.get(self._model_name, self.MODEL_BELL_STRINGS["PFAFF Creative 7570"])

        response = bytearray()
        response.extend(resp_ident.encode('ascii'))
        response.append(self.CTRL_ETX)
        
        logger.info("Bell command received - sending identification string [%s]", resp_ident)
        return bytes(response)

    def handle_list_pmemory(self) -> bytes:
        """Handle 'PI' + CTRL_ETX (List P-Memory) command.

        Dispatches to the model-specific implementation.
        """
        if self._model_name == "PFAFF Creative 1475 CD":
            return self._handle_list_pmemory_1475cd()
        else:
            return self._handle_list_pmemory_75xx()

    def _handle_list_pmemory_75xx(self) -> bytes:
        """List P-Memory handler for PFAFF Creative 7570 and 7550.

        Format (all values hex-ASCII encoded, 2 chars per byte, uppercase):
          [num_slots: 2]  +  per slot: [type: 2][size_be: 4][00 00: 4]  +  [free: 4]
          + CTRL_ETB (raw byte)  +  [xor_checksum: 2]

        Type codes: 0x00 = 9mm or Empty, 0x01 = MAXI
        Size is number of data bytes stored in the slot (big-endian 2 bytes).
        Free memory is calculated as the difference between total P-Memory size and used bytes.
        Checksum is XOR of every ASCII byte that precedes CTRL_ETB.
        """

        logger.info("List P-Memory command received - sending response")

        slots = self.machine_state.p_memory_slots if self.machine_state else []

        ascii_data = bytearray()

        # Number of slots: 1 byte value → 2 ASCII hex chars
        ascii_data.extend(f"{len(slots):02X}".encode('ascii'))

        # Per-slot: type(1) + size_be(2) + reserved_zeros(2) = 5 bytes → 10 ASCII chars
        for slot in slots:
            type_byte = 0x01 if slot.pattern_type == "MAXI" else 0x00
            size = slot.get_size_bytes()
            ascii_data.extend(f"{type_byte:02X}".encode('ascii'))
            ascii_data.extend(f"{size:04X}".encode('ascii'))
            ascii_data.extend(b"0000")  # 2 reserved zero bytes

        # Free memory calculated as the difference between total P-Memory size and used bytes
        used_bytes = sum(slot.get_size_bytes() for slot in slots)
        free_bytes = self.machine_state.p_memory_total_size - used_bytes
        ascii_data.extend(f"{free_bytes:04X}".encode('ascii'))

        # Checksum over all ASCII data bytes (before CTRL_ETB)
        checksum = self._calculate_checksum(ascii_data)

        # Final response: ascii_data + CTRL_ETB (raw) + checksum (2 ASCII hex chars)
        response = bytearray(ascii_data)
        response.append(self.CTRL_ETB)
        response.extend(f"{checksum:02X}".encode('ascii'))

        self._state = self._STATE_WAIT_ACK
        return bytes(response)

    def _handle_list_pmemory_1475cd(self) -> bytes:
        """List P-Memory handler for PFAFF Creative 1475 CD.

        Format (all values hex-ASCII encoded, 2 chars per byte, uppercase):
          [num_slots: 2]  +  per slot: [type: 2][size_be: 4][00 00 00 00: 8]  +  [free: 4]
          + CTRL_ETB (raw byte)  +  [xor_checksum: 2]

        Type codes: 0x00 = 9mm or Empty, 0x01 = MAXI
        Size is number of data bytes stored in the slot (big-endian 2 bytes).
        Free memory is calculated as the difference between total P-Memory size and used bytes.
        Checksum is XOR of every ASCII byte that precedes CTRL_ETB.
        """
        logger.info("List P-Memory (1475 CD) command received - sending response")

        slots = self.machine_state.p_memory_slots if self.machine_state else []

        ascii_data = bytearray()

        # Number of slots: 1 byte value → 2 ASCII hex chars
        ascii_data.extend(f"{len(slots):02X}".encode('ascii'))

        # Per-slot: type(1) + size_be(2) + reserved_zeros(4) = 7 bytes → 14 ASCII chars
        for slot in slots:
            type_byte = 0x01 if slot.pattern_type == "MAXI" else 0x00
            size = slot.get_size_bytes()
            ascii_data.extend(f"{type_byte:02X}".encode('ascii'))
            ascii_data.extend(f"{size:04X}".encode('ascii'))
            ascii_data.extend(b"00000000")  # 4 reserved zero bytes

        # Free memory calculated as the difference between total P-Memory size and used bytes
        used_bytes = sum(slot.get_size_bytes() for slot in slots)
        free_bytes = self.machine_state.p_memory_total_size - used_bytes
        ascii_data.extend(f"{free_bytes:04X}".encode('ascii'))

        # Checksum over all ASCII data bytes (before CTRL_ETB)
        checksum = self._calculate_checksum(ascii_data)

        # Final response: ascii_data + CTRL_ETB (raw) + checksum (2 ASCII hex chars)
        response = bytearray(ascii_data)
        response.append(self.CTRL_ETB)
        response.extend(f"{checksum:02X}".encode('ascii'))

        self._state = self._STATE_WAIT_ACK
        return bytes(response)
    
    def handle_delete_pmemory(self, slot_hex: str) -> bytes:
        """Handle 'PL<XX>' + CTRL_ETX (Delete P-Memory) command.

        slot_hex is the 2-character hex-ASCII encoded slot number (e.g. '05' for slot 5).
        Clears the slot data and resets its type to Empty.
        Returns CTRL_ACK on success, CTRL_NAK on invalid slot or decode error.
        """
        try:
            slot_id = int(slot_hex, 16)
        except ValueError:
            logger.warning(f"Delete P-Memory: invalid slot hex {slot_hex!r}")
            return bytes([self.CTRL_NAK])

        if self.machine_state is None:
            logger.warning("Delete P-Memory: no machine state available")
            return bytes([self.CTRL_NAK])

        try:
            slot = self.machine_state.get_p_memory_slot(slot_id)
        except IndexError:
            logger.warning(f"Delete P-Memory: slot {slot_id} out of range")
            return bytes([self.CTRL_NAK])

        slot.clear()
        logger.info(f"Delete P-Memory: slot {slot_id} cleared")
        if self.on_pmemory_changed:
            self.on_pmemory_changed()
        return bytes([self.CTRL_ACK])

    def _abort_write(self):
        """Abort an in-progress P-Memory write and return to idle state."""
        self._state = self._STATE_IDLE
        self._write_slot_id = None
        self._write_stitch_type = None
        self._write_expected_size = None
        self._write_data_accumulated = bytearray()
        self._write_chunk_buffer = bytearray()
        self._write_checksum_chars = bytearray()
        self._write_header_buffer = bytearray()
        self._write_header = bytearray()


    def handle_write_pmemory_init(self, params: str) -> bytes:
        """Handle 'PN<11 hex chars>' + CTRL_ETX (Write P-Memory header) command.

        Dispatches to the model-specific implementation.
        """
        if self._model_name == "PFAFF Creative 1475 CD":
            return self._handle_write_pmemory_init_1475cd(params)
        else:
            return self._handle_write_pmemory_init_75xx(params)
        
    def _handle_write_pmemory_init_75xx(self, params: str) -> bytes:
        """Handle 'PN<11 hex chars>' + CTRL_ETX (Write P-Memory header) command for PFAFF Creative 7570 and 7550.

        params is 11 chars: slot(2) + stitch_type(2) + size(4) + CTRL_ETB + checksum(2)
        Checksum is verified over 'PN' + slot + size (10 ASCII bytes).
        On success: transitions to STATE_WRITE_HEADER and returns CTRL_ACK.
        On error: returns CTRL_NAK.
        """
        slot_hex     = params[0:2]
        type_hex     = params[2:4]
        size_hex     = params[4:8]
        ctrl         = params[8]  # should be CTRL_ETB
        checksum_hex = params[9:11]

        try:
            slot_id           = int(slot_hex, 16)
            stitch_type       = int(type_hex, 16)
            expected_size     = int(size_hex, 16)
            ctrl_byte         = ord(ctrl)
            received_checksum = int(checksum_hex, 16)
        except ValueError:
            logger.warning(f"Write P-Memory: invalid params {params!r}")
            return bytes([self.CTRL_NAK])
        
        if ctrl_byte != self.CTRL_ETB:
            logger.warning(f"Write P-Memory: expected CTRL_ETB at position 8, got {ctrl!r}")
            return bytes([self.CTRL_NAK])

        # Verify checksum over the full command prefix (PN + slot + type + size)
        cmd_bytes  = ("PN" + slot_hex + type_hex + size_hex).encode('ascii')
        calculated = self._calculate_checksum(cmd_bytes)
        if calculated != received_checksum:
            logger.warning(
                f"Write P-Memory: checksum mismatch "
                f"(received 0x{received_checksum:02X}, calculated 0x{calculated:02X})"
            )
            return bytes([self.CTRL_NAK])

        if self.machine_state is None:
            logger.warning("Write P-Memory: no machine state available")
            return bytes([self.CTRL_NAK])

        try:
            self.machine_state.get_p_memory_slot(slot_id)
        except IndexError:
            logger.warning(f"Write P-Memory: slot {slot_id} out of range")
            return bytes([self.CTRL_NAK])

        if stitch_type == 0x00:
            pattern_type_str = "9mm"
        elif stitch_type == 0x01:
            pattern_type_str = "MAXI"
        else:
            logger.warning(f"Write P-Memory: unknown stitch type {stitch_type:#02X}")
            return bytes([self.CTRL_NAK])

        # Valid header - set up write state
        self._write_slot_id           = slot_id
        self._write_stitch_type       = pattern_type_str
        self._write_expected_size     = expected_size
        self._write_data_accumulated  = bytearray()
        self._write_chunk_buffer      = bytearray()
        self._write_is_last_chunk     = False
        self._write_checksum_chars    = bytearray()
        self._state = self._STATE_WRITE_HEADER

        logger.info(f"Write P-Memory: slot {slot_id}, expecting {expected_size} bytes pattern of type: {pattern_type_str} - requesting header")
        return bytes([self.CTRL_ENQ])  # Using ENQ to indicate ready for data (ACK is used for chunk ACKs)

    def _handle_write_pmemory_init_1475cd(self, params: str) -> bytes:
        """Handle 'PN<19 hex chars>' + CTRL_ETX (Write P-Memory header) command for PFAFF Creative 1475CD.

        params is 19 chars: slot(2) + stitch_type(2) + size(4) + header (8) + CTRL_ETB + checksum(2)
        Checksum is verified over 'PN' + slot + size (10 ASCII bytes).
        On success: transitions to STATE_WRITE_DATA and returns CTRL_ACK.
        On error: returns CTRL_NAK.
        """
        slot_hex     = params[0:2]
        type_hex     = params[2:4]
        size_hex     = params[4:8]
        header       = params[8:16]
        ctrl         = params[16]  # should be CTRL_ETB
        checksum_hex = params[17:19]

        try:
            slot_id           = int(slot_hex, 16)
            stitch_type       = int(type_hex, 16)
            expected_size     = int(size_hex, 16)
            ctrl_byte         = ord(ctrl)
            received_checksum = int(checksum_hex, 16)
        except ValueError:
            logger.warning(f"Write P-Memory: invalid params {params!r}")
            return bytes([self.CTRL_NAK])
        
        if ctrl_byte != self.CTRL_ETB:
            logger.warning(f"Write P-Memory: expected CTRL_ETB at position 16, got {ctrl!r}")
            return bytes([self.CTRL_NAK])

        # Verify checksum over the full command prefix (PN + slot + type + size + header)
        cmd_bytes  = ("PN" + slot_hex + type_hex + size_hex + header).encode('ascii')
        calculated = self._calculate_checksum(cmd_bytes)
        if calculated != received_checksum:
            logger.warning(
                f"Write P-Memory: checksum mismatch "
                f"(received 0x{received_checksum:02X}, calculated 0x{calculated:02X})"
            )
            return bytes([self.CTRL_NAK])

        if self.machine_state is None:
            logger.warning("Write P-Memory: no machine state available")
            return bytes([self.CTRL_NAK])

        try:
            self.machine_state.get_p_memory_slot(slot_id)
        except IndexError:
            logger.warning(f"Write P-Memory: slot {slot_id} out of range")
            return bytes([self.CTRL_NAK])

        if stitch_type == 0x00:
            pattern_type_str = "9mm"
        elif stitch_type == 0x01:
            pattern_type_str = "MAXI"
        else:
            logger.warning(f"Write P-Memory: unknown stitch type {stitch_type:#02X}")
            return bytes([self.CTRL_NAK])

        # Valid header - set up write state
        self._write_slot_id           = slot_id
        self._write_stitch_type       = pattern_type_str
        self._write_expected_size     = expected_size
        self._write_header            = bytearray(header.encode('ascii'))
        self._write_data_accumulated  = bytearray()
        self._write_chunk_buffer      = bytearray()
        self._write_is_last_chunk     = False
        self._write_checksum_chars    = bytearray()
        self._state = self._STATE_WRITE_DATA

        logger.info(f"Write P-Memory: slot {slot_id}, expecting {expected_size} bytes pattern of type: {pattern_type_str} - requesting data")
        return bytes([self.CTRL_ENQ])  # Using ENQ to indicate ready for data (ACK is used for chunk ACKs)

    def _process_write_header(self) -> bytes:
        """Verify checksum of the write header, store it, and transition to data collection.

        Header format: <header_data> + <2-char hex checksum> + CTRL_ETX
        Checksum is calculated over all header_data bytes.
        On success: stores raw header_data, sends ACK, transitions to _STATE_WRITE_DATA.
        On error: sends NAK and aborts.
        """
        buf = self._write_header_buffer
        if len(buf) < 2:
            logger.warning("Write P-Memory: header too short")
            self._abort_write()
            return bytes([self.CTRL_NAK])

        header_data    = buf[:-3]  # all bytes before the last 3 (CTRL_ETB + checksum)
        ctrl           = buf[-3]   # should be CTRL_ETB
        checksum_chars = buf[-2:]

        try:
            received_checksum = int(checksum_chars.decode('ascii'), 16)
        except (ValueError, UnicodeDecodeError):
            logger.warning(
                f"Write P-Memory: invalid header checksum bytes {bytes(checksum_chars)!r}"
            )
            self._abort_write()
            return bytes([self.CTRL_NAK])

        calculated = self._calculate_checksum(header_data)
        if calculated != received_checksum:
            logger.warning(
                f"Write P-Memory: header checksum mismatch "
                f"(received 0x{received_checksum:02X}, calculated 0x{calculated:02X})"
            )
            self._abort_write()
            return bytes([self.CTRL_NAK])


        if ctrl != self.CTRL_ETB:
            logger.warning(f"Write P-Memory: expected CTRL_ETB at position 8, got {ctrl!r}")
            return bytes([self.CTRL_NAK])
        
        self._write_header = bytearray(header_data)
        self._write_header_buffer = bytearray()
        logger.info(
            f"Write P-Memory: header received ({len(self._write_header)} bytes) - requesting chunks"
        )
        header_ascii = self._write_header.decode('ascii', errors='replace')
        # Separate with spaces every 2 chars for readability
        header_ascii_spaced = ' '.join(header_ascii[i:i+2] for i in range(0, len(header_ascii), 2))
        logger.debug(f"Header data (ASCII): {header_ascii_spaced!r}")
        def _try_hex(s):
            try:
                return str(int(s, 16))
            except ValueError:
                return '??'
        header_decimal = ' '.join(
            _try_hex(header_ascii[i:i+2]) for i in range(0, len(header_ascii) - 1, 2)
        )
        logger.debug(f"Header data (decimal): {header_decimal}")


        self._state = self._STATE_WRITE_DATA
        return bytes([self.CTRL_ENQ])  # Using ENQ to indicate ready for data (ACK is used for chunk ACKs)

    def _process_write_chunk(self) -> bytes:
        """Verify checksum of the current chunk, accumulate its data, and ACK or NAK."""
        try:
            received_checksum = int(self._write_checksum_chars.decode('ascii'), 16)
        except (ValueError, UnicodeDecodeError):
            logger.warning(
                f"Write P-Memory: invalid chunk checksum bytes {bytes(self._write_checksum_chars)!r}"
            )
            self._abort_write()
            return bytes([self.CTRL_NAK])

        calculated = self._calculate_checksum(self._write_chunk_buffer)
        if calculated != received_checksum:
            logger.warning(
                f"Write P-Memory: chunk checksum mismatch "
                f"(received 0x{received_checksum:02X}, calculated 0x{calculated:02X})"
            )
            self._abort_write()
            return bytes([self.CTRL_NAK])

        # Checksum OK - accumulate chunk and await more data or CTRL_ETX
        self._write_data_accumulated.extend(self._write_chunk_buffer)
        self._write_chunk_buffer = bytearray()
        logger.debug(
            f"Write P-Memory: chunk OK ({len(self._write_data_accumulated)} bytes accumulated so far)"
        )
        self._state = self._STATE_WRITE_DATA
        return bytes([self.CTRL_ACK])

    def _commit_write(self) -> bytes:
        """Commit all accumulated data to the target P-Memory slot and return to idle."""
        if self.machine_state is None:
            logger.warning("Write P-Memory: no machine state on commit")
            self._abort_write()
            return bytes([])

        slot = self.machine_state.get_p_memory_slot(self._write_slot_id)

        if self._write_stitch_type in ("9mm", "MAXI"):
            slot.set_slot_data(
                self._write_stitch_type,
                self._write_header.decode('ascii', errors='replace'),
                self._write_data_accumulated.decode('ascii', errors='replace')
            )
            logger.info(
                f"Write P-Memory: slot {self._write_slot_id} written with a {self._write_stitch_type} stitch ({slot.get_size_bytes()} bytes)"
            )
        else:
            slot.set_slot_data(
                self._write_stitch_type,
                self._write_header.decode('ascii', errors='replace'),
                self._write_data_accumulated.decode('ascii', errors='replace')
            )
            logger.warning(f"Write P-Memory: unknown stitch type {self._write_stitch_type} on commit - storing raw data only ({slot.get_size_bytes()} bytes)")

        self._append_write_log(slot)
        self._abort_write()  # resets state to IDLE
        if self.on_pmemory_changed:
            self.on_pmemory_changed()
        return bytes([])

    def _append_write_log(self, slot) -> None:
        """Append header bytes and pattern statistics (in hex) to pmem_write_head_log.txt."""
        log_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "pmem_write_head_log.txt")
        try:
            with open(log_path, 'a', encoding='ascii') as f:
                import datetime
                ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                f.write(f"--- {ts}  slot={slot.slot_id} ---\n")

                # Header: ascii bytes are pairs of hex chars; convert each pair to int → print as hex
                header_ascii = slot.header_raw
                header_hex = ' '.join(
                    f"{int(header_ascii[i:i+2], 16):02X}"
                    if all(c in '0123456789ABCDEFabcdef' for c in header_ascii[i:i+2])
                    else '??'
                    for i in range(0, len(header_ascii) - 1, 2)
                )
                f.write(f"HEADER: {header_hex}\n")

                if slot.pattern_type in ("9mm", "MAXI"):
                    stats = slot.get_pattern_stats()
                    if stats["n"] >= 1:
                        f.write(
                            f"STATISTICS: n={stats['n']:02X}, n_bytes={slot.get_size_bytes():02X}, checksum={stats['checksum']:02X}, "
                            f"x_min={stats['x_min']:02X} x_max={stats['x_max']:02X}, span_x={stats['span_x']:02X}, "
                            f"dx_max={stats['dx_max']:02X}, dx_min={stats['dx_min'] & 0xFF:02X}, dx_min_abs={stats['dx_min_abs']:02X}, "
                            f"y_min={stats['y_min']:02X} y_max={stats['y_max']:02X}, span_y={stats['span_y']:02X}, "
                            f"d0y_max={stats['d0y_max']:02X}, d0y_min={stats['d0y_min'] & 0xFF:02X}, d0y_min_abs={stats['d0y_min_abs']:02X}, "
                            f"p0_x={stats['p0_x']:02X}, p0_y={stats['p0_y']:02X}, pn_x={stats['pn_x']:02X}, pn_y={stats['pn_y']:02X}\n\n"
                        )
                else:
                    f.write(
                        f"STATISTICS: unavailable for slot type {slot.pattern_type}\n\n"
                    )

        except OSError as e:
            logger.warning(f"Write log: could not append to {log_path}: {e}")

    def _abort_read(self):
        """Abort an in-progress P-Memory read and return to idle state."""
        self._state = self._STATE_IDLE
        self._read_data = bytearray()
        self._read_offset = 0
        self._read_last_chunk_sent = False

    def handle_read_pmemory_init(self, params: str) -> bytes:
        """Handle 'RM<5 chars>' + CTRL_ETX (Read P-Memory) command.

        params is 5 chars: fixed_06(2) + slot_hex(2) + pattern_type(1)
        Responds with slot data in hex-ASCII chunks of up to READ_CHUNK_SIZE chars each.
        Each chunk is followed by CTRL_ETB + 2-char hex checksum.
        The last chunk additionally gets CTRL_ETX appended after the checksum.
        Returns NAK if params are invalid, slot is out of range, or slot is empty.
        """
        fixed_byte = params[0:2]
        slot_hex   = params[2:4]
        pattern_type  = params[4]

        if pattern_type == "0":
            requested_pattern_type = "9mm"
        elif pattern_type == "1":
            requested_pattern_type = "MAXI" 
        else:
            logger.warning(f"Read P-Memory: unknown requested pattern type {pattern_type!r}")
            return bytes([self.CTRL_NAK])

        if fixed_byte.upper() != "06":
            logger.warning(f"Read P-Memory: unexpected fixed byte {fixed_byte!r} (expected '06')")
            return bytes([self.CTRL_NAK])

        try:
            slot_id = int(slot_hex, 16)
        except ValueError:
            logger.warning(f"Read P-Memory: invalid slot hex {slot_hex!r}")
            return bytes([self.CTRL_NAK])

        if self.machine_state is None:
            logger.warning("Read P-Memory: no machine state available")
            return bytes([self.CTRL_NAK])

        try:
            slot = self.machine_state.get_p_memory_slot(slot_id)
        except IndexError:
            logger.warning(f"Read P-Memory: slot {slot_id} out of range")
            return bytes([self.CTRL_NAK])

        if slot.pattern_type == "Empty" or not slot.pattern_xy:
            logger.warning(f"Read P-Memory: slot {slot_id} is empty")
            return bytes([self.CTRL_NAK])
        
        if requested_pattern_type != slot.pattern_type:
            logger.warning(f"Read P-Memory: unexpected requested pattern type (is: {slot.pattern_type}, requested: {requested_pattern_type})")
            return bytes([self.CTRL_NAK])

        if slot.pattern_type == "9mm":
            # Encode slot data as 3-digit x + 2-digit y, iterating flat list in steps of 2
            self._read_data = bytearray()
            pattern_xy = slot.pattern_xy
            for i in range(0, len(pattern_xy) - 1, 2):
                self._read_data.extend(f"{pattern_xy[i]:03d}{pattern_xy[i+1]:02d}".encode('ascii'))
            self._read_offset = 0
        elif slot.pattern_type == "MAXI":
            # Encode slot data as 3-digit x + 2-digit y + side transport with sign
            self._read_data = bytearray()
            pattern_bytes = slot.pattern_bytes
            for i in range(0, len(pattern_bytes) - 1, 3):
                x = pattern_bytes[i]
                y = pattern_bytes[i+1]
                side = pattern_bytes[i+2]
                self._read_data.extend(f"{x:03d}{y:02d}{side:+d}".encode('ascii'))
        else:
            logger.warning(f"Read P-Memory: unknown pattern type {slot.pattern_type} in slot {slot_id}")
            return bytes([self.CTRL_NAK])


        self._read_last_chunk_sent = False

        logger.info(
            f"Read P-Memory: slot {slot_id}, {len(slot.pattern_xy)} bytes "
            f"({len(self._read_data)} ASCII chars) - sending first chunk"
        )
        return self._send_next_read_chunk()

    def _send_next_read_chunk(self) -> bytes:
        """Build and return the next read chunk, advancing _read_offset.

        Chunk format: <hex-ASCII data> + CTRL_ETB + <2-char hex checksum>
        Last chunk:   <hex-ASCII data> + CTRL_ETB + <2-char hex checksum> + CTRL_ETX
        Sets state to _STATE_READ_WAIT_ACK before returning.
        """
        chunk_ascii = bytes(self._read_data[self._read_offset : self._read_offset + self.READ_CHUNK_SIZE])
        self._read_offset += len(chunk_ascii)
        is_last = self._read_offset >= len(self._read_data)

        checksum = self._calculate_checksum(chunk_ascii)
        response = bytearray(chunk_ascii)
        response.append(self.CTRL_ETB)
        response.extend(f"{checksum:02X}".encode('ascii'))
        if is_last:
            response.append(self.CTRL_ETX)
            self._read_last_chunk_sent = True

        logger.debug(
            f"Read P-Memory: chunk {len(chunk_ascii)} ASCII chars, "
            f"checksum 0x{checksum:02X}, last={is_last}"
        )
        self._state = self._STATE_READ_WAIT_ACK
        return bytes(response)

    def parse_response(self, data: bytes) -> dict:
        """
        Parse response from machine

        Args:
            data: Response bytes

        Returns:
            Dictionary with parsed response data
        """
        if len(data) < 2:
            return {"error": "Response too short"}

        response_code = data[0]

        if response_code == self.RESPONSE_OK:
            return {"status": "OK"}
        elif response_code == self.RESPONSE_ERROR:
            error_code = data[1] if len(data) > 1 else 0
            return {"status": "ERROR", "code": error_code}
        elif response_code == self.RESPONSE_BUSY:
            return {"status": "BUSY"}
        else:
            return {"status": "UNKNOWN", "code": response_code}
    
    @staticmethod
    def _calculate_checksum(data: bytes) -> int:
        # Sum all bytes modulo 256
        checksum = 0
        for byte in data:
            checksum = (checksum + byte) & 0xFF
        return checksum
