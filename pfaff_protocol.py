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
    CMD_LIST_CARD = "KI"
    CMD_WRITE_CARD = "KN"
    CMD_READ_CARD_PREVIEW = "KB"
    CMD_DELETE_CARD = "KL"

    # Internal state machine states
    _STATE_IDLE = 0
    _STATE_WRITE_DATA = 1         # collecting chunk data bytes
    _STATE_WRITE_CHECKSUM = 2     # collecting 2-byte hex checksum after chunk terminator
    _STATE_WAIT_ACK = 3           # waiting for CTRL_ACK from host after PI response
    _STATE_READ_WAIT_ACK = 4      # waiting for CTRL_ACK after a read chunk
    _STATE_WRITE_HEADER = 6       # collecting header bytes after PN init ACK
    _STATE_WRITE_CARD_HEADER = 7  # collecting 31 raw header bytes for KN command
    _STATE_WRITE_CARD_DATA = 8    # collecting card data chunks
    _STATE_READ_KB_PARAMS = 9     # collecting 8 raw parameter bytes for KB command
    _STATE_READ_KB_WAIT_ETX = 10  # after 8 KB params, waiting for CTRL_ETX terminator
    _STATE_READ_KB_WAIT_ACK = 11  # waiting for CTRL_ACK after a KB preview chunk
    _STATE_READ_KL_PARAMS = 12    # collecting 7 raw parameter bytes for KL command
    _STATE_READ_KL_WAIT_ETX = 13  # after 7 KL params, waiting for CTRL_ETX terminator
    _STATE_RAW_CMD_MNEMONIC = 14     # received CTRL_ETX in idle, expecting 2-byte raw command mnemonic
    _STATE_WRITE_CARD_WAIT_ETX = 15  # collected 30 KN header bytes, waiting for CTRL_ETX terminator

    # Card data chunk sub-states (used when _state == _STATE_WRITE_CARD_DATA)
    _CARD_CHUNK_WAIT_START  = 0  # waiting for CTRL_ENQ (or bare size byte)
    _CARD_CHUNK_WAIT_SIZE   = 1  # received CTRL_ENQ, waiting for payload size byte
    _CARD_CHUNK_PAYLOAD     = 2  # collecting payload bytes
    _CARD_CHUNK_SIZE_REPEAT = 3  # waiting for repeated size confirmation byte
    _CARD_CHUNK_ETB         = 4  # waiting for CTRL_ETB
    _CARD_CHUNK_CHECKSUM    = 5  # collecting 2-byte checksum

    # Read chunk size (max ASCII chars per chunk = 2 * raw bytes)
    READ_CHUNK_SIZE = 250

    # Bell identification strings per model
    MODEL_BELL_STRINGS = {
        "PFAFF Creative 7570":    "Copyright 1992 - 97       G.M. PFAFF AG Creative 7570B    Vers. 2.1", # From real machine
        "PFAFF Creative 7550":    "Copyright 1992,-93,-94    G.M. PFAFF AG Creative 7550 CD  Vers. 2.0", # From real machine
        "PFAFF Creative 1475 CD": "Copyright 1992,-93,-94    G.M. PFAFF AG Creative 1475 CD  Vers. 1.0", # Guess
    }

    # Bell command debounce time (seconds)
    BELL_DEBOUNCE_SECONDS = 2.0
    
    def __init__(self, machine_state=None, on_pmemory_changed=None, on_card_changed=None):
        self.machine_state = machine_state
        self.on_pmemory_changed = on_pmemory_changed  # Optional callback: called when P-Memory is modified
        self.on_card_changed = on_card_changed        # Optional callback: called when card memory is modified
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

        # Write Card state machine
        self._write_card_header_buffer = bytearray()
        self._write_card_header_raw = bytearray()
        self._write_card_stitch_type = None
        self._write_card_preview_size = 0
        self._write_card_pattern_size = 0
        self._write_card_filename_len = 0
        self._write_card_slot_id = None
        self._write_card_data_accumulated = bytearray()
        self._write_card_data_substate = self._CARD_CHUNK_WAIT_START
        self._write_card_chunk_has_enq = False
        self._write_card_chunk_size = 0
        self._write_card_chunk_size_repeat = 0
        self._write_card_chunk_buffer = bytearray()
        self._write_card_chunk_checksum_buf = bytearray()

        # Read Card Preview (KB) state machine
        self._kb_params_buffer = bytearray()
        self._kb_preview_data = bytearray()
        self._kb_preview_offset = 0

        # Delete Card Slot (KL) state machine
        self._kl_params_buffer = bytearray()

        # Raw command mnemonic collection (KN, KB, KL — preceded by CTRL_ETX)
        self._raw_cmd_mnemonic_buffer = bytearray()

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
                    if self.cmd_buffer:
                        cmd_str = self.cmd_buffer.decode('ascii', errors='replace')
                        self.cmd_buffer.clear()
                        response.extend(self._dispatch_text_command(cmd_str))
                    else:
                        # CTRL_ETX with empty buffer is the prefix for a raw command (KN, KB, KL)
                        self._raw_cmd_mnemonic_buffer = bytearray()
                        self._state = self._STATE_RAW_CMD_MNEMONIC
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
                        logger.info("Read P-Memory: transfer complete (ACK received) - resetting to idle")
                        self._abort_read()
                    else:
                        response.extend(self._send_next_read_chunk())
                elif byte == self.CTRL_NAK:
                    logger.warning("Read P-Memory: NAK received, aborting")
                    self._abort_read()
                else:
                    logger.warning(f"Read P-Memory: unexpected byte 0x{byte:02X} while waiting for ACK")

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

            elif self._state == self._STATE_WRITE_CARD_HEADER:
                # Collect raw bytes — any byte value is valid (including 0x03).
                # The header is exactly 30 bytes; the CTRL_ETX terminator follows separately.
                self._write_card_header_buffer.append(byte)
                if len(self._write_card_header_buffer) == 30:
                    self._state = self._STATE_WRITE_CARD_WAIT_ETX

            elif self._state == self._STATE_WRITE_CARD_DATA:
                response.extend(self._handle_card_data_byte(byte))

            elif self._state == self._STATE_READ_KB_PARAMS:
                # Accept any byte value (including control characters) as raw parameter
                self._kb_params_buffer.append(byte)
                if len(self._kb_params_buffer) == 8:
                    self._state = self._STATE_READ_KB_WAIT_ETX

            elif self._state == self._STATE_READ_KB_WAIT_ETX:
                if byte == self.CTRL_ETX:
                    response.extend(self._handle_read_card_preview())
                elif byte == self.CTRL_EOT:
                    logger.info("KB: EOT received waiting for ETX - resetting to idle")
                    self._state = self._STATE_IDLE
                    self._kb_params_buffer = bytearray()
                else:
                    logger.warning(f"KB: unexpected byte 0x{byte:02X} waiting for CTRL_ETX - ignored")

            elif self._state == self._STATE_READ_KB_WAIT_ACK:
                if byte == self.CTRL_EOT:
                    logger.info("KB preview: EOT received - aborting transfer")
                    self._abort_read_kb()
                elif byte == self.CTRL_ACK:
                    if self._kb_preview_offset >= len(self._kb_preview_data):
                        logger.info("KB preview: transfer complete, sending CTRL_ETX")
                        self._abort_read_kb()
                        response.append(self.CTRL_ETX)
                    else:
                        response.extend(self._send_next_kb_chunk())
                elif byte == self.CTRL_NAK:
                    logger.warning("KB preview: NAK received - aborting transfer")
                    self._abort_read_kb()
                else:
                    logger.warning(f"KB preview: unexpected byte 0x{byte:02X} waiting for ACK - ignored")

            elif self._state == self._STATE_READ_KL_PARAMS:
                # Accept any byte value (including control characters) as raw parameter
                self._kl_params_buffer.append(byte)
                if len(self._kl_params_buffer) == 7:
                    self._state = self._STATE_READ_KL_WAIT_ETX

            elif self._state == self._STATE_READ_KL_WAIT_ETX:
                if byte == self.CTRL_ETX:
                    response.extend(self._handle_delete_card_slot())
                elif byte == self.CTRL_EOT:
                    logger.info("KL: EOT received waiting for ETX - resetting to idle")
                    self._state = self._STATE_IDLE
                    self._kl_params_buffer = bytearray()
                else:
                    logger.warning(f"KL: unexpected byte 0x{byte:02X} waiting for CTRL_ETX - ignored")

            elif self._state == self._STATE_WRITE_CARD_WAIT_ETX:
                if byte == self.CTRL_ETX:
                    response.extend(self._process_write_card_header())
                elif byte == self.CTRL_EOT:
                    logger.info("KN: EOT received waiting for header CTRL_ETX - aborting")
                    self._abort_write_card()
                else:
                    logger.warning(f"KN: unexpected byte 0x{byte:02X} waiting for header CTRL_ETX - aborting")
                    self._abort_write_card()

            elif self._state == self._STATE_RAW_CMD_MNEMONIC:
                if byte == self.CTRL_EOT:
                    logger.info("Raw cmd: EOT received - resetting to idle")
                    self._state = self._STATE_IDLE
                    self._raw_cmd_mnemonic_buffer = bytearray()
                elif byte == self.CTRL_ETX:
                    # Another CTRL_ETX — treat as a repeated raw-command prefix
                    logger.debug("Raw cmd: CTRL_ETX received while collecting mnemonic - resetting")
                    self._raw_cmd_mnemonic_buffer = bytearray()
                elif byte == self.CTRL_BEL:
                    logger.warning("Raw cmd: CTRL_BEL received - aborting, handling bell")
                    self._state = self._STATE_IDLE
                    self._raw_cmd_mnemonic_buffer = bytearray()
                    response.extend(self.handle_bell_command())
                else:
                    self._raw_cmd_mnemonic_buffer.append(byte)
                    if len(self._raw_cmd_mnemonic_buffer) == 2:
                        mnemonic = bytes(self._raw_cmd_mnemonic_buffer).decode('ascii', errors='replace')
                        self._raw_cmd_mnemonic_buffer = bytearray()
                        if mnemonic == self.CMD_WRITE_CARD:
                            response.extend(self.handle_write_card_init())
                        elif mnemonic == self.CMD_READ_CARD_PREVIEW:
                            self._kb_params_buffer = bytearray()
                            self._state = self._STATE_READ_KB_PARAMS
                        elif mnemonic == self.CMD_DELETE_CARD:
                            self._kl_params_buffer = bytearray()
                            self._state = self._STATE_READ_KL_PARAMS
                        else:
                            logger.warning(f"Raw cmd: unknown mnemonic {mnemonic!r} - resetting to idle")
                            self._state = self._STATE_IDLE

        return bytes(response)

    def _dispatch_text_command(self, cmd: str) -> bytes:
        """Dispatch a complete text command (stripped of its CTRL_ETX terminator)."""
        logger.debug(f"Text command received: {cmd!r}")
        if cmd == self.CMD_LIST_PMEMORY:
            return self.handle_list_pmemory()
        if cmd == self.CMD_LIST_CARD:
            return self.handle_list_card()
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

    def handle_list_card(self) -> bytes:
        """Handle 'KI' + CTRL_ETX (List Memory Card content) command.

        Response format (raw bytes):
          06 00 00 10 02 18 01 00 <N9mm> 03 C8 <NEmbr>
          00 00 00 00 00 00 02 00 <NMaxi> 00 00 00 00 00 00 00 00 00 18
          + CTRL_ETB (raw byte) + checksum (2 raw bytes, 16-bit big-endian sum)

        <N9mm>  = number of 9mm patterns on the card
        <NEmbr> = number of Embroidery patterns on the card
        <NMaxi> = number of MAXI patterns on the card
        Checksum is the 16-bit sum of all data bytes before CTRL_ETB.
        """
        logger.info("List Card Memory command received - sending response")

        if self.machine_state is not None:
            n_9mm  = len(self.machine_state.card_9mm.slots)
            n_embr = len(self.machine_state.card_embroidery.slots)
            n_maxi = len(self.machine_state.card_maxi.slots)
        else:
            n_9mm = n_embr = n_maxi = 0

        raw_bytes = bytes([
            0x00, 0x00, 0x10, 0x02, 0x18, 0x01, 0x00,
            n_9mm & 0xFF,
            0x03, 0xC8,
            n_embr & 0xFF,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x02, 0x00,
            n_maxi & 0xFF,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x18,
        ])

        checksum = self._calculate_checksum(raw_bytes)

        response = bytearray()
        response.append(0x06)
        response.extend(raw_bytes)
        response.append(self.CTRL_ETB)
        response.extend(f"{checksum:02X}".encode('ascii'))

        self._state = self._STATE_IDLE
        return bytes(response)

    def handle_write_card_init(self) -> bytes:
        """Handle KN raw command (preceded by CTRL_ETX).

        Resets the card-write state and transitions to _STATE_WRITE_CARD_HEADER
        to collect 30 raw header bytes, followed by a CTRL_ETX terminator.
        Returns empty bytes (no immediate response — wait for header bytes).
        """
        logger.info("Write Card Memory: KN command received - collecting 30-byte header")
        self._write_card_header_buffer = bytearray()
        self._write_card_header_raw = bytearray()
        self._write_card_stitch_type = None
        self._write_card_preview_size = 0
        self._write_card_pattern_size = 0
        self._write_card_filename_len = 0
        self._write_card_slot_id = None
        self._state = self._STATE_WRITE_CARD_HEADER
        return b""

    def _process_write_card_header(self) -> bytes:
        """Process the 30-byte raw card-write header and respond with ACK + slot info.

        Header layout (raw bytes, 0-indexed):
          [6]     stitch type: 0x01=9mm, 0x02=MAXI, 0x03=Embroidery
          [24-25] preview image size in bytes (big-endian uint16)
          [26-27] pattern data size in bytes (big-endian uint16)
          [28]    filename field length including null-terminator

        Any byte value is valid; the header is delimited by byte count (30).
        The CTRL_ETX terminator is received separately in _STATE_WRITE_CARD_WAIT_ETX.

        Response on success: CTRL_ACK + 2 bytes
          9mm:        0xC0, slot_id
          MAXI:       0xD0, slot_id
          Embroidery: 0xC0, slot_id + 0xC8   (slot 0 → 0xC8, slot 1 → 0xC9, …)

        On error: CTRL_NAK, resets to idle.
        Transitions to _STATE_WRITE_CARD_DATA to await incoming data chunks.
        """
        buf = self._write_card_header_buffer

        stitch_byte = buf[6]
        if stitch_byte == 0x01:
            self._write_card_stitch_type = "9mm"
        elif stitch_byte == 0x02:
            self._write_card_stitch_type = "MAXI"
        elif stitch_byte == 0x03:
            self._write_card_stitch_type = "Embroidery"
        else:
            logger.warning(f"Write Card: unknown stitch type byte 0x{stitch_byte:02X} - aborting")
            self._abort_write_card()
            return bytes([self.CTRL_NAK])

        self._write_card_preview_size = (buf[24] << 8) | buf[25]
        self._write_card_pattern_size = (buf[27] << 8) | buf[28]
        self._write_card_filename_len = buf[29]
        self._write_card_header_raw = bytes(buf)

        if self.machine_state is None:
            logger.warning("Write Card: no machine state available - aborting")
            self._abort_write_card()
            return bytes([self.CTRL_NAK])

        if self._write_card_stitch_type == "9mm":
            space = self.machine_state.card_9mm
            slot_id = self._next_free_card_slot(space)
            response_code = 0xC0
            response_slot = slot_id
        elif self._write_card_stitch_type == "MAXI":
            space = self.machine_state.card_maxi
            slot_id = self._next_free_card_slot(space)
            response_code = 0xD0
            response_slot = slot_id
        else:  # Embroidery
            space = self.machine_state.card_embroidery
            slot_id = self._next_free_card_slot(space)
            response_code = 0xC0
            response_slot = slot_id + 0xC8

        self._write_card_slot_id = slot_id

        logger.info(
            f"Write Card: {self._write_card_stitch_type} pattern assigned to slot {slot_id}, "
            f"preview={self._write_card_preview_size} bytes, "
            f"pattern={self._write_card_pattern_size} bytes, "
            f"filename_len={self._write_card_filename_len} - ACK 0x{response_code:02X} 0x{response_slot:02X}"
        )
        self._state = self._STATE_WRITE_CARD_DATA
        return bytes([self.CTRL_ACK, response_code, response_slot])

    def _next_free_card_slot(self, space) -> int:
        """Return the lowest non-negative slot_id not currently occupied in the given card space."""
        used = set(space.slots.keys())
        i = 0
        while i in used:
            i += 1
        return i

    def _abort_write_card(self) -> None:
        """Abort an in-progress card write and return to idle state."""
        self._state = self._STATE_IDLE
        self._write_card_header_buffer = bytearray()
        self._write_card_header_raw = bytearray()
        self._write_card_stitch_type = None
        self._write_card_preview_size = 0
        self._write_card_pattern_size = 0
        self._write_card_filename_len = 0
        self._write_card_slot_id = None
        self._write_card_data_accumulated = bytearray()
        self._write_card_data_substate = self._CARD_CHUNK_WAIT_START
        self._write_card_chunk_has_enq = False
        self._write_card_chunk_size = 0
        self._write_card_chunk_size_repeat = 0
        self._write_card_chunk_buffer = bytearray()
        self._write_card_chunk_checksum_buf = bytearray()

    def _handle_card_data_byte(self, byte: int) -> bytes:
        """Route one incoming byte through the card data chunk sub-state machine."""
        sub = self._write_card_data_substate

        if sub == self._CARD_CHUNK_WAIT_START:
            if byte == self.CTRL_EOT:
                logger.info("Write Card: CTRL_EOT between chunks - aborting transfer")
                self._abort_write_card()
            elif byte == self.CTRL_ENQ:
                self._write_card_chunk_has_enq = True
                self._write_card_chunk_buffer = bytearray()
                self._write_card_chunk_checksum_buf = bytearray()
                self._write_card_data_substate = self._CARD_CHUNK_WAIT_SIZE
            else:
                # Chunk starts directly with size byte (missing CTRL_ENQ) — will be NAK'd
                self._write_card_chunk_has_enq = False
                self._write_card_chunk_size = byte
                self._write_card_chunk_buffer = bytearray()
                self._write_card_chunk_checksum_buf = bytearray()
                self._write_card_data_substate = (
                    self._CARD_CHUNK_SIZE_REPEAT if byte == 0 else self._CARD_CHUNK_PAYLOAD
                )
            return b""

        if sub == self._CARD_CHUNK_WAIT_SIZE:
            if byte == self.CTRL_EOT:
                logger.info("Write Card: CTRL_EOT during chunk size byte - aborting")
                self._abort_write_card()
                return b""
            self._write_card_chunk_size = byte
            self._write_card_data_substate = (
                self._CARD_CHUNK_SIZE_REPEAT if byte == 0 else self._CARD_CHUNK_PAYLOAD
            )
            return b""

        if sub == self._CARD_CHUNK_PAYLOAD:
            # Any byte value is valid payload — do NOT special-case control chars here.
            self._write_card_chunk_buffer.append(byte)
            if len(self._write_card_chunk_buffer) == self._write_card_chunk_size:
                self._write_card_data_substate = self._CARD_CHUNK_SIZE_REPEAT
            return b""

        if sub == self._CARD_CHUNK_SIZE_REPEAT:
            self._write_card_chunk_size_repeat = byte
            self._write_card_data_substate = self._CARD_CHUNK_ETB
            return b""

        if sub == self._CARD_CHUNK_ETB:
            if byte != self.CTRL_ETB:
                logger.warning(
                    f"Write Card: expected CTRL_ETB in chunk, got 0x{byte:02X} - NAK, reset chunk"
                )
                self._write_card_data_substate = self._CARD_CHUNK_WAIT_START
                return bytes([self.CTRL_NAK])
            self._write_card_data_substate = self._CARD_CHUNK_CHECKSUM
            return b""

        if sub == self._CARD_CHUNK_CHECKSUM:
            self._write_card_chunk_checksum_buf.append(byte)
            if len(self._write_card_chunk_checksum_buf) == 2:
                return self._process_card_chunk()

        return b""

    def _process_card_chunk(self) -> bytes:
        """Validate the completed chunk; ACK if valid, NAK and reset chunk if not.

        Validation rules (any failure → NAK, sub-state resets to WAIT_START for retransmit):
          1. Repeated size byte must match the original size byte.
          2. Checksum of payload bytes must match the received checksum.
          3. Chunk must have started with CTRL_ENQ; if not, all other checks may pass
             but the chunk is still NAK'd (spec: discard and wait for retransmit).

        On ACK: if total accumulated bytes reach the expected transfer size, the card
        slot is committed and the write-card state machine is reset to idle.
        """
        payload = self._write_card_chunk_buffer
        size    = self._write_card_chunk_size
        size_rep = self._write_card_chunk_size_repeat
        has_enq = self._write_card_chunk_has_enq
        cs_buf  = self._write_card_chunk_checksum_buf

        # Always reset sub-state — ready for next chunk or retransmit.
        self._write_card_data_substate = self._CARD_CHUNK_WAIT_START

        if self._write_card_chunk_size_repeat != size:
            logger.warning(
                f"Write Card: chunk size repeat mismatch "
                f"(hdr={size}, repeat={size_rep})"
            )
            return bytes([self.CTRL_NAK])


        if not has_enq:
            logger.warning("Write Card: chunk missing leading CTRL_ENQ")
            return bytes([self.CTRL_NAK])
        # concatenate size, payload, and size repeat for checksum calculation
        data_for_checksum = bytes([size]) + payload + bytes([size_rep])
        calculated = self._calculate_checksum(data_for_checksum)
        received_checksum = int(cs_buf, 16)
        if calculated != received_checksum:
            logger.warning(
                f"Write Card: chunk checksum mismatch "
                f"(received 0x{received_checksum:02X}, calculated 0x{calculated:02X})"
            )
            return bytes([self.CTRL_NAK])

        # Valid chunk — accumulate
        self._write_card_data_accumulated.extend(payload)
        total_expected = (
            self._write_card_filename_len
            + self._write_card_preview_size
            + self._write_card_pattern_size
        )
        total_received = len(self._write_card_data_accumulated)
        logger.debug(
            f"Write Card: chunk OK ({len(payload)} B), total {total_received}/{total_expected}"
        )

        if total_received >= total_expected:
            self._commit_write_card()

        return bytes([self.CTRL_ACK])

    def _commit_write_card(self) -> None:
        """Decode accumulated card payload and store it in the appropriate card space.

        Payload layout (lengths from the KN header):
          [0 : fn_len]                    → filename bytes (incl. null terminator)
          [fn_len : fn_len+prev_size]     → preview image (raw bytes)
          [fn_len+prev_size : ...+pat_sz] → stitch pattern (raw bytes)

        All three fields are stored as lowercase hex strings in the CardMemorySlot.
        Pattern parsing is deferred — the on-card binary stitch format is not yet known;
        pattern_xy / pattern_bytes will be empty until proper parsing is added.
        """
        import datetime
        from machine_state import CardMemorySlot

        data     = bytes(self._write_card_data_accumulated)
        fn_len   = self._write_card_filename_len
        prev_sz  = self._write_card_preview_size
        pat_sz   = self._write_card_pattern_size

        filename_bytes = data[:fn_len]
        preview_bytes  = data[fn_len : fn_len + prev_sz]
        pattern_bytes  = data[fn_len + prev_sz : fn_len + prev_sz + pat_sz]

        try:
            filename = filename_bytes.rstrip(b'\x00').decode('latin-1', errors='replace')
        except Exception:
            filename = ""

        slot = CardMemorySlot(
            slot_id      = self._write_card_slot_id,
            pattern_type = self._write_card_stitch_type,
            header_raw   = self._write_card_header_raw.hex(),
            preview_raw  = preview_bytes.hex(),
            pattern_raw  = pattern_bytes.hex(),
            filename     = filename,
        )
        slot.parse_pattern_data()  # no-op for hex-encoded raw binary; safe to call

        if self.machine_state is not None:
            if self._write_card_stitch_type == "9mm":
                self.machine_state.card_9mm.set_slot(slot)
            elif self._write_card_stitch_type == "MAXI":
                self.machine_state.card_maxi.set_slot(slot)
            else:
                self.machine_state.card_embroidery.set_slot(slot)

        logger.info(
            f"Write Card: committed {self._write_card_stitch_type} slot {self._write_card_slot_id} "
            f"(filename={filename!r}, preview={prev_sz} B, pattern={pat_sz} B)"
        )
        if self.on_card_changed:
            self.on_card_changed()
        self._abort_write_card()

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
        free_bytes = self.machine_state.p_memory_total_size - used_bytes + 1 # Unknown why but real machines report 1 byte more free than what the original SW shows
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
        free_bytes = self.machine_state.p_memory_total_size - used_bytes + 1 # Unknown why but real machines report 1 byte more free than what the original SW shows
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

    def _abort_read_kb(self) -> None:
        """Abort an in-progress KB card-preview read and return to idle state."""
        self._state = self._STATE_IDLE
        self._kb_params_buffer = bytearray()
        self._kb_preview_data = bytearray()
        self._kb_preview_offset = 0

    def _handle_read_card_preview(self) -> bytes:
        """Handle KB command: read preview image from a card memory slot.

        Parameter bytes layout (8 raw bytes):
          [0-3]  fixed: 00 00 10 02
          [4]    BANK: 0xC0 = 9mm/Embroidery, 0xD0 = MAXI
          [5]    SLOT: raw slot byte (Embroidery adds 0xC8 offset, so 0xC8=slot0, 0xC9=slot1)
          [6]    TYPE: 0x01=9mm, 0x02=MAXI, 0x03=Embroidery
          [7]    fixed: 00

        First response chunk:
          CTRL_ACK | 00 00 00 00 | <pattern_size 2 bytes BE> | <fn_len_byte>
          | <filename+NUL> | <chunk_size> | <preview_bytes> | <chunk_size> | CTRL_ETB
          | <checksum 2 bytes BE (sum of preview payload)>

        Subsequent chunks (sent after each ACK while preview remains):
          CTRL_ENQ | <chunk_size> | <preview_bytes> | <chunk_size> | CTRL_ETB
          | <checksum 2 bytes BE>

        After the final ACK (all preview bytes sent): respond with CTRL_ETX.
        """
        params = self._kb_params_buffer
        type_byte = params[6]
        slot_raw  = params[5]

        if type_byte == 0x01:
            stitch_type = "9mm"
            space = self.machine_state.card_9mm if self.machine_state else None
            slot_id = slot_raw
        elif type_byte == 0x02:
            stitch_type = "MAXI"
            space = self.machine_state.card_maxi if self.machine_state else None
            slot_id = slot_raw
        elif type_byte == 0x03:
            stitch_type = "Embroidery"
            space = self.machine_state.card_embroidery if self.machine_state else None
            slot_id = slot_raw - 0xC8
        else:
            logger.warning(f"KB: unknown type byte 0x{type_byte:02X}")
            self._state = self._STATE_IDLE
            return bytes([self.CTRL_NAK])

        if self.machine_state is None:
            logger.warning("KB: no machine state available")
            self._state = self._STATE_IDLE
            return bytes([self.CTRL_NAK])

        slot = space.get_slot(slot_id)
        if slot is None:
            logger.warning(f"KB: {stitch_type} slot {slot_id} not found")
            self._state = self._STATE_IDLE
            return bytes([self.CTRL_NAK])

        try:
            preview_bytes = bytes.fromhex(slot.preview_raw) if slot.preview_raw else b""
        except ValueError:
            logger.warning(f"KB: invalid preview_raw hex in {stitch_type} slot {slot_id} - using empty")
            preview_bytes = b""

        try:
            pattern_size = len(bytes.fromhex(slot.pattern_raw)) if slot.pattern_raw else 0
        except ValueError:
            pattern_size = 0

        filename_bytes = slot.filename.encode('ascii', errors='replace') + b'\x00'
        fn_len_byte = len(filename_bytes)  # length including null terminator

        self._kb_preview_data = bytearray(preview_bytes)
        self._kb_preview_offset = 0

        chunk_size = min(0x80, len(self._kb_preview_data))
        chunk_data = bytes(self._kb_preview_data[:chunk_size])
        self._kb_preview_offset = chunk_size

        response = bytearray()
        response.append(self.CTRL_ACK)
        response.extend([0x00, 0x00, 0x00, 0x00])
        response.append((pattern_size >> 8) & 0xFF)
        response.append(pattern_size & 0xFF)
        response.append(fn_len_byte)
        response.extend(filename_bytes)
        response.append(chunk_size)
        response.extend(chunk_data)
        response.append(chunk_size)
        checksum = self._calculate_checksum(response[1:])  # checksum over all bytes after CTRL_ACK
        response.append(self.CTRL_ETB)
        response.extend(f"{checksum:02X}".encode('ascii'))

        self._state = self._STATE_READ_KB_WAIT_ACK
        logger.info(
            f"KB: {stitch_type} slot {slot_id}, preview={len(self._kb_preview_data)} B, "
            f"pattern={pattern_size} B, filename={slot.filename!r} - sending first chunk ({chunk_size} B)"
        )
        return bytes(response)

    def _send_next_kb_chunk(self) -> bytes:
        """Build and return the next KB preview chunk (starts with CTRL_ENQ)."""
        remaining = len(self._kb_preview_data) - self._kb_preview_offset
        chunk_size = min(0x80, remaining)
        chunk_data = bytes(
            self._kb_preview_data[self._kb_preview_offset : self._kb_preview_offset + chunk_size]
        )
        self._kb_preview_offset += chunk_size

        response = bytearray()
        response.append(self.CTRL_ENQ)
        response.append(chunk_size)
        response.extend(chunk_data)
        response.append(chunk_size)
        checksum = self._calculate_checksum(response) # checksum over the full chunk message (including CTRL_ENQ and chunk size bytes)
        response.append(self.CTRL_ETB)
        response.extend(f"{checksum:02X}".encode('ascii'))

        self._state = self._STATE_READ_KB_WAIT_ACK
        logger.debug(
            f"KB preview: chunk {chunk_size} B, "
            f"offset {self._kb_preview_offset}/{len(self._kb_preview_data)}"
        )
        return bytes(response)

    def _handle_delete_card_slot(self) -> bytes:
        """Handle KL command: delete a slot from card memory.

        Parameter bytes layout (7 raw bytes):
          [0-3]  fixed: 00 00 10 02
          [4]    BANK: 0xC0 = 9mm/Embroidery, 0xD0 = MAXI
          [5]    SLOT: raw slot byte (Embroidery adds 0xC8 offset, so 0xC8=slot0, 0xC9=slot1)
          [6]    TYPE: 0x01=9mm, 0x02=MAXI, 0x03=Embroidery

        Response: CTRL_ACK on success, CTRL_NAK on error.
        """
        params = self._kl_params_buffer
        type_byte = params[6]
        slot_raw  = params[5]

        if type_byte == 0x01:
            stitch_type = "9mm"
            space = self.machine_state.card_9mm if self.machine_state else None
            slot_id = slot_raw
        elif type_byte == 0x02:
            stitch_type = "MAXI"
            space = self.machine_state.card_maxi if self.machine_state else None
            slot_id = slot_raw
        elif type_byte == 0x03:
            stitch_type = "Embroidery"
            space = self.machine_state.card_embroidery if self.machine_state else None
            slot_id = slot_raw - 0xC8
        else:
            logger.warning(f"KL: unknown type byte 0x{type_byte:02X} - NAK")
            self._state = self._STATE_IDLE
            return bytes([self.CTRL_NAK])

        if self.machine_state is None:
            logger.warning("KL: no machine state available - NAK")
            self._state = self._STATE_IDLE
            return bytes([self.CTRL_NAK])

        if space.get_slot(slot_id) is None:
            logger.warning(f"KL: {stitch_type} slot {slot_id} not found - NAK")
            self._state = self._STATE_IDLE
            return bytes([self.CTRL_NAK])

        space.delete_slot(slot_id)
        logger.info(f"KL: deleted {stitch_type} slot {slot_id}")

        if self.on_card_changed:
            self.on_card_changed()

        self._state = self._STATE_IDLE
        return bytes([self.CTRL_ACK])

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
