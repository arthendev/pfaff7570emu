#!/usr/bin/env python3
"""
PFAFF Creative 75xx Emulator
Main application entry point
"""

import sys
import json
import logging
from pathlib import Path
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QTabWidget, QTextEdit, QFileDialog, 
                             QMessageBox, QMenu, QAction, QSplitter)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont

from machine_state import MachineState
from pmemory_tab import PMemoryTab
from mmemory_tab import MMemoryTab
from card_memory_tab import CardMemoryTab
from serial_connection import SerialConnectionDialog
from serial_handler import SerialHandler
from pfaff_protocol import PFAFFProtocol
from logger import setup_logger

logger = setup_logger(__name__)


class PfaffCreativeEmulator(QMainWindow):
    """Main application window for the sewing machine emulator"""
    
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PFAFF Creative 75xx Emulator")
        self.setGeometry(100, 100, 1500, 900)
        
        # Initialize machine state
        self.machine_state = MachineState()
        self.current_file = None
        
        # Initialize serial handler and protocol
        self.serial_handler = SerialHandler()
        self.serial_handler.data_received.connect(self.on_serial_data_received)
        self.serial_handler.error_occurred.connect(self.on_serial_error)
        self.protocol = PFAFFProtocol(self.machine_state, on_pmemory_changed=self._on_pmemory_changed)
        
        # Setup UI
        self.setup_ui()
        self.create_menu()
        
        logger.info("Application started")
    
    def setup_ui(self):
        """Setup main UI layout"""
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        layout = QVBoxLayout()
        
        # Create splitter for upper and lower sections
        splitter = QSplitter(Qt.Vertical)
        
        # Upper section: Tab widget
        self.tab_widget = QTabWidget()
        self.pmemory_tab = PMemoryTab(self.machine_state)
        self.mmemory_tab = MMemoryTab(self.machine_state)
        self.card_memory_tab = CardMemoryTab(self.machine_state)
        
        self.tab_widget.addTab(self.pmemory_tab, "P-Memory")
        self.tab_widget.addTab(self.mmemory_tab, "M-Memory")
        self.tab_widget.addTab(self.card_memory_tab, "Card Memory")
        
        splitter.addWidget(self.tab_widget)
        
        # Lower section: Console log
        self.console = QTextEdit()
        self.console.setReadOnly(True)
        font = QFont("Courier New", 9)
        self.console.setFont(font)
        self.console.setContextMenuPolicy(Qt.CustomContextMenu)
        self.console.customContextMenuRequested.connect(self._show_console_context_menu)
        splitter.addWidget(self.console)
        
        # Set splitter sizes (50% upper, 50% lower) - scales with window
        splitter.setSizes([450, 600])
        # splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 1)
        
        layout.addWidget(splitter)
        central_widget.setLayout(layout)
        
        # Setup logger output to console
        self.setup_console_logging()
    
    def setup_console_logging(self):
        """Redirect logger output to console widget and set up filterable Python console handler"""
        from logger import ConsoleHandler, FilteringStreamHandler, FilteringFileHandler
        formatter = logging.Formatter('%(levelname)s - %(name)s - %(message)s')

        # Qt console handler
        self.console_handler = ConsoleHandler(self.console)
        self.console_handler.setLevel(logging.DEBUG)

        # Filterable Python (stdout) handler — replaces per-module StreamHandlers
        self.python_console_handler = FilteringStreamHandler()
        self.python_console_handler.setLevel(logging.DEBUG)
        self.python_console_handler.setFormatter(formatter)

        self.file_handler = None  # created on demand when file logging is enabled

        root_logger = logging.getLogger()
        root_logger.setLevel(logging.DEBUG)
        root_logger.addHandler(self.console_handler)
        root_logger.addHandler(self.python_console_handler)

        # Remove StreamHandlers already added to named loggers by setup_logger
        # so output is not duplicated now that the root handler covers them.
        for log in list(logging.Logger.manager.loggerDict.values()):
            if isinstance(log, logging.Logger):
                for h in list(log.handlers):
                    if isinstance(h, logging.StreamHandler) and not isinstance(h, logging.FileHandler):
                        log.removeHandler(h)
    
    def create_menu(self):
        """Create application menu"""
        menubar = self.menuBar()
        
        # File menu
        file_menu = menubar.addMenu("File")
        
        new_action = QAction("New", self)
        new_action.triggered.connect(self.new_file)
        file_menu.addAction(new_action)
        
        open_action = QAction("Open", self)
        open_action.triggered.connect(self.open_file)
        file_menu.addAction(open_action)
        
        save_action = QAction("Save", self)
        save_action.triggered.connect(self.save_file)
        file_menu.addAction(save_action)
        
        save_as_action = QAction("Save As", self)
        save_as_action.triggered.connect(self.save_file_as)
        file_menu.addAction(save_as_action)
        
        file_menu.addSeparator()
        
        exit_action = QAction("Exit", self)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)
        
        # Log menu
        log_menu = menubar.addMenu("Log")

        log_window_submenu = QMenu("Log window", self)
        log_menu.addMenu(log_window_submenu)
        self._build_log_level_submenu(
            log_window_submenu, self._on_log_level_toggled
        )

        python_console_submenu = QMenu("Python console", self)
        log_menu.addMenu(python_console_submenu)
        self._build_log_level_submenu(
            python_console_submenu, self._on_python_log_level_toggled
        )

        log_menu.addSeparator()
        self._log_to_file_action = QAction("Log to file enabled", self)
        self._log_to_file_action.setCheckable(True)
        self._log_to_file_action.setChecked(False)
        self._log_to_file_action.toggled.connect(self._on_log_to_file_toggled)
        log_menu.addAction(self._log_to_file_action)

        log_to_file_submenu = QMenu("Log to file", self)
        log_menu.addMenu(log_to_file_submenu)
        self._build_log_level_submenu(
            log_to_file_submenu, self._on_file_log_level_toggled
        )

        # Connection menu
        connection_menu = menubar.addMenu("Connection")
        
        open_connection_action = QAction("Open Connection", self)
        open_connection_action.triggered.connect(self.open_serial_connection)
        connection_menu.addAction(open_connection_action)
        
        close_connection_action = QAction("Close Connection", self)
        close_connection_action.triggered.connect(self.close_serial_connection)
        connection_menu.addAction(close_connection_action)
    
    def new_file(self):
        """Create a new machine state file"""
        self.machine_state = MachineState()
        self.current_file = None
        self.pmemory_tab.update_ui(self.machine_state)
        logger.info("New file created")
        self.setWindowTitle("PFAFF 75xx Sewing Machine Emulator - [New]")
    
    def open_file(self):
        """Open a machine state file"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, 
            "Open Sewing Machine State",
            "",
            "JSON files (*.json);;All files (*.*)"
        )
        
        if file_path:
            try:
                self.machine_state.load_from_file(file_path)
                self.current_file = file_path
                self.pmemory_tab.update_ui(self.machine_state)
                self.mmemory_tab.update_ui(self.machine_state)
                self.card_memory_tab.update_ui(self.machine_state)
                logger.info(f"File opened: {file_path}")
                self.setWindowTitle(f"PFAFF 75xx Sewing Machine Emulator - {Path(file_path).name}")
            except Exception as e:
                logger.error(f"Failed to open file: {str(e)}")
                QMessageBox.critical(self, "Error", f"Failed to open file: {str(e)}")
    
    def save_file(self):
        """Save current machine state"""
        if self.current_file:
            try:
                self.machine_state.save_to_file(self.current_file)
                logger.info(f"File saved: {self.current_file}")
                self.setWindowTitle(f"PFAFF 75xx Sewing Machine Emulator - {Path(self.current_file).name}")
            except Exception as e:
                logger.error(f"Failed to save file: {str(e)}")
                QMessageBox.critical(self, "Error", f"Failed to save file: {str(e)}")
        else:
            self.save_file_as()
    
    def save_file_as(self):
        """Save machine state to a new file"""
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Sewing Machine State",
            "",
            "JSON files (*.json);;All files (*.*)"
        )
        
        if file_path:
            try:
                self.machine_state.save_to_file(file_path)
                self.current_file = file_path
                logger.info(f"File saved as: {file_path}")
                self.setWindowTitle(f"PFAFF 75xx Sewing Machine Emulator - {Path(file_path).name}")
            except Exception as e:
                logger.error(f"Failed to save file: {str(e)}")
                QMessageBox.critical(self, "Error", f"Failed to save file: {str(e)}")
    
    def open_serial_connection(self):
        """Open serial connection dialog"""
        dialog = SerialConnectionDialog(self)
        if dialog.exec_():
            port, baudrate = dialog.get_selected_connection()
            if port is None:
                logger.warning("No COM port available")
                QMessageBox.warning(self, "Connection", "No COM ports available")
                return
            
            logger.info(f"Opening serial connection: {port} at {baudrate} baud")
            if self.serial_handler.connect(port, baudrate):
                QMessageBox.information(
                    self, 
                    "Connection", 
                    f"Serial connection opened on {port} at {baudrate} baud"
                )
            else:
                QMessageBox.critical(
                    self, 
                    "Connection Error", 
                    f"Failed to open connection on {port}"
                )
    
    def close_serial_connection(self):
        """Close serial connection"""
        self.serial_handler.disconnect()
        logger.info("Serial connection closed")
        QMessageBox.information(self, "Connection", "Serial connection closed")
    
    def _build_log_level_submenu(self, menu: QMenu, slot):
        """Add Warning / Info / Debug checkable actions to *menu*, connected to *slot*."""
        for label, level in (("Warning", logging.WARNING), ("Info", logging.INFO), ("Debug", logging.DEBUG)):
            action = QAction(label, self)
            action.setCheckable(True)
            action.setChecked(True)
            action.toggled.connect(lambda checked, lvl=level: slot(lvl, checked))
            menu.addAction(action)

    def _on_log_level_toggled(self, level: int, checked: bool):
        """Show or hide a log level in the Qt console."""
        self.console_handler.set_level_visible(level, checked)

    def _on_python_log_level_toggled(self, level: int, checked: bool):
        """Show or hide a log level in the Python (stdout) console."""
        self.python_console_handler.set_level_visible(level, checked)

    def _on_log_to_file_toggled(self, checked: bool):
        """Enable or disable logging to file."""
        import datetime
        import os
        root_logger = logging.getLogger()
        if checked:
            if self.file_handler is None:
                logs_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logs")
                os.makedirs(logs_dir, exist_ok=True)
                ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                log_path = os.path.join(logs_dir, f"log_{ts}.txt")
                from logger import FilteringFileHandler
                formatter = logging.Formatter('%(asctime)s %(levelname)s - %(name)s - %(message)s')
                self.file_handler = FilteringFileHandler(log_path, encoding='utf-8')
                self.file_handler.setLevel(logging.DEBUG)
                self.file_handler.setFormatter(formatter)
                root_logger.addHandler(self.file_handler)
                logger.info(f"File logging started: {log_path}")
        else:
            if self.file_handler is not None:
                logger.info("File logging stopped")
                root_logger.removeHandler(self.file_handler)
                self.file_handler.close()
                self.file_handler = None

    def _on_file_log_level_toggled(self, level: int, checked: bool):
        """Show or hide a log level in the file log."""
        if self.file_handler is not None:
            self.file_handler.set_level_visible(level, checked)

    def _show_console_context_menu(self, pos):
        """Show right-click context menu on the log console."""
        menu = self.console.createStandardContextMenu()
        menu.addSeparator()
        clear_action = QAction("Clear", self)
        clear_action.triggered.connect(self.console.clear)
        menu.addAction(clear_action)
        menu.exec_(self.console.mapToGlobal(pos))

    def _on_pmemory_changed(self):
        """Refresh P-Memory tab after a delete or write operation"""
        self.pmemory_tab.update_ui(self.machine_state)

    def on_serial_data_received(self, data: bytes):
        """Handle received serial data - pass through protocol dispatcher"""
        logger.debug(f"Serial RX ({len(data)} bytes): {data.hex()}")
        response = self.protocol.process_incoming(data)
        if response:
            logger.debug(f"Serial TX ({len(response)} bytes): {response.hex()}")
            self.serial_handler.send_data(response)
    
    def on_serial_error(self, error_msg: str):
        """Handle serial communication errors"""
        logger.error(error_msg)
        QMessageBox.critical(self, "Serial Error", error_msg)
    
    def closeEvent(self, event):
        """Handle application close event"""
        if self.serial_handler.is_connected:
            self.serial_handler.disconnect()
        event.accept()


def main():
    """Application entry point"""
    app = QApplication(sys.argv)
    window = PfaffCreativeEmulator()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
