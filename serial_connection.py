"""
Serial connection dialog
"""

from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, 
                             QComboBox, QPushButton, QMessageBox)
from PyQt5.QtGui import QFont
from PyQt5.QtCore import Qt
import serial
import serial.tools.list_ports


class SerialConnectionDialog(QDialog):
    """Dialog for selecting serial connection parameters"""
    
    def __init__(self, parent=None, last_port=None, last_baudrate=None):
        super().__init__(parent)
        self.setWindowTitle("Open Serial Connection")
        self.setModal(True)
        self.selected_port = None
        self.selected_baudrate = None
        self._last_port = last_port
        self._last_baudrate = last_baudrate
        self.setup_ui()
    
    def setup_ui(self):
        """Setup dialog UI"""
        layout = QVBoxLayout()
        
        # Port selection
        port_layout = QHBoxLayout()
        port_label = QLabel("COM Port:")
        port_label.setMinimumWidth(100)
        self.port_combo = QComboBox()
        self.populate_ports()
        refresh_button = QPushButton("Refresh")
        refresh_button.clicked.connect(self.populate_ports)
        port_layout.addWidget(port_label)
        port_layout.addWidget(self.port_combo)
        port_layout.addWidget(refresh_button)
        layout.addLayout(port_layout)
        
        # Baud rate selection
        baudrate_layout = QHBoxLayout()
        baudrate_label = QLabel("Baud Rate:")
        baudrate_label.setMinimumWidth(100)
        self.baudrate_combo = QComboBox()
        self.baudrate_combo.addItem("4800", 4800)
        self.baudrate_combo.addItem("10472", 10472)
        if self._last_baudrate is not None:
            idx = self.baudrate_combo.findData(self._last_baudrate)
            if idx >= 0:
                self.baudrate_combo.setCurrentIndex(idx)
        baudrate_layout.addWidget(baudrate_label)
        baudrate_layout.addWidget(self.baudrate_combo)
        layout.addLayout(baudrate_layout)
        
        # Data bits, stop bits, parity (fixed as 8N1)
        settings_label = QLabel("Data: 8 bits, Parity: None, Stop bits: 1")
        settings_font = QFont()
        settings_font.setPointSize(9)
        settings_label.setFont(settings_font)
        layout.addWidget(settings_label)
        
        layout.addSpacing(20)
        
        # Buttons
        button_layout = QHBoxLayout()
        ok_button = QPushButton("OK")
        ok_button.clicked.connect(self.accept)
        cancel_button = QPushButton("Cancel")
        cancel_button.clicked.connect(self.reject)
        button_layout.addStretch()
        button_layout.addWidget(ok_button)
        button_layout.addWidget(cancel_button)
        layout.addLayout(button_layout)
        
        self.setLayout(layout)
        self.setMinimumWidth(400)
    
    def populate_ports(self):
        """Populate COM port list"""
        self.port_combo.clear()
        ports = serial.tools.list_ports.comports()
        
        if not ports:
            self.port_combo.addItem("No ports available", None)
        else:
            for port in ports:
                self.port_combo.addItem(
                    f"{port.device} - {port.description}", 
                    port.device
                )
        if self._last_port is not None:
            idx = self.port_combo.findData(self._last_port)
            if idx >= 0:
                self.port_combo.setCurrentIndex(idx)
    
    def get_selected_connection(self):
        """Get selected port and baud rate"""
        port = self.port_combo.currentData()
        baudrate = self.baudrate_combo.currentData()
        return port, baudrate
