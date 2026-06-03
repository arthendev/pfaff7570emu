"""
Serial communication handler
"""

import threading
import time
from typing import Callable, Optional
import serial
from PyQt5.QtCore import QObject, pyqtSignal
import logging

logger = logging.getLogger(__name__)


class SerialHandler(QObject):
    """Handles serial communication with the sewing machine"""
    
    # Signals
    data_received = pyqtSignal(bytes)
    connection_changed = pyqtSignal(bool)
    error_occurred = pyqtSignal(str)
    
    def __init__(self):
        super().__init__()
        self.serial_port = None
        self.is_connected = False
        self.reader_thread = None
        self.running = False
    
    def connect(self, port: str, baudrate: int = 4800) -> bool:
        """
        Connect to serial port
        
        Args:
            port: COM port name (e.g., 'COM3')
            baudrate: Baud rate (default: 4800)
        
        Returns:
            True if connection successful, False otherwise
        """
        try:
            self.serial_port = serial.Serial(
                port=port,
                baudrate=baudrate,
                bytesize=serial.EIGHTBITS,
                parity=serial.PARITY_NONE,
                stopbits=serial.STOPBITS_ONE,
                timeout=1
            )
            
            self.is_connected = True
            self.running = True
            
            # Start reader thread
            self.reader_thread = threading.Thread(target=self._read_loop, daemon=True)
            self.reader_thread.start()
            
            self.connection_changed.emit(True)
            logger.info(f"Connected to {port} at {baudrate} baud")
            return True
        except Exception as e:
            error_msg = f"Failed to connect to {port}: {str(e)}"
            logger.error(error_msg)
            self.error_occurred.emit(error_msg)
            return False
    
    def disconnect(self):
        """Disconnect from serial port"""
        try:
            self.running = False
            if self.reader_thread and self.reader_thread.is_alive():
                self.reader_thread.join(timeout=2)
            
            if self.serial_port and self.serial_port.is_open:
                self.serial_port.close()
            
            self.is_connected = False
            self.connection_changed.emit(False)
            logger.info("Disconnected from serial port")
        except Exception as e:
            logger.error(f"Error disconnecting: {str(e)}")
    
    def send_data(self, data: bytes) -> bool:
        """
        Send data through serial port
        
        Args:
            data: Data to send
        
        Returns:
            True if successful, False otherwise
        """
        if not self.is_connected or not self.serial_port:
            logger.warning("Cannot send data: not connected")
            return False
        
        try:
            self.serial_port.write(data)
            # logger.debug(f"Sent {len(data)} bytes")
            return True
        except Exception as e:
            error_msg = f"Error sending data: {str(e)}"
            logger.error(error_msg)
            self.error_occurred.emit(error_msg)
            return False
    
    def _read_loop(self):
        """Read loop running in separate thread"""
        # Use a blocking read to avoid busy-waiting when there's no data.
        # `read(1)` will block up to the port timeout (set in `connect`) and
        # return empty bytes on timeout, keeping CPU usage low.
        while self.running and self.is_connected:
            try:
                # Block waiting for at least 1 byte (respects Serial.timeout)
                first = self.serial_port.read(1)
                if not first:
                    # timeout occurred, loop again (low CPU usage)
                    continue

                data = first

                # Read any additional bytes that have already arrived
                try:
                    remaining = self.serial_port.in_waiting
                except Exception:
                    remaining = 0

                if remaining:
                    data += self.serial_port.read(remaining)

                if data:
                    # logger.debug(f"Received {len(data)} bytes")
                    self.data_received.emit(data)

            except Exception as e:
                if self.running:
                    error_msg = f"Error reading from serial port: {str(e)}"
                    logger.error(error_msg)
                    self.error_occurred.emit(error_msg)
                    # Attempt a clean disconnect; break to end thread
                    try:
                        self.disconnect()
                    except Exception:
                        pass
                    break
