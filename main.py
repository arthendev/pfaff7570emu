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
                             QMessageBox, QMenu, QAction, QSplitter, QActionGroup)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont

from machine_state import MachineState
from pmemory_tab import PMemoryTab
from mmemory_tab import MMemoryTab
from card_memory_tab import CardMemoryTab
from serial_connection import SerialConnectionDialog
from serial_handler import SerialHandler
from pfaff_protocol import PFAFFProtocol
from preferences_dialog import PreferencesDialog
from slot_detail_window import SlotDetailWindow
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
        self._modified = False
        self._title_name = ""
        self._slot_detail_windows: dict = {}
        self._config = self._load_config()
        self._recent_files: list = self._config.get("recent_files", [])
        
        # Initialize serial handler and protocol
        self.serial_handler = SerialHandler()
        self.serial_handler.data_received.connect(self.on_serial_data_received)
        self.serial_handler.error_occurred.connect(self.on_serial_error)
        self.serial_handler.connection_changed.connect(self._on_connection_changed)
        self.protocol = PFAFFProtocol(self.machine_state, on_pmemory_changed=self._on_pmemory_changed)
        
        # Setup UI
        self.setup_ui()
        self.pmemory_tab.slot_clicked.connect(self._open_slot_detail)
        self.create_menu()

        logger.info("Application started")
        self._try_auto_connect()
        self._try_auto_open_state()
    
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
        
        open_action = QAction("Open...", self)
        open_action.triggered.connect(self.open_file)
        file_menu.addAction(open_action)

        self._recent_menu = QMenu("Open recent", self)
        file_menu.addMenu(self._recent_menu)
        self._rebuild_recent_menu()
        
        self._save_action = QAction("Save", self)
        self._save_action.setEnabled(False)
        self._save_action.triggered.connect(self.save_file)
        file_menu.addAction(self._save_action)
        
        save_as_action = QAction("Save As...", self)
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
        self._log_window_actions = self._build_log_level_submenu(
            log_window_submenu, self._on_log_level_toggled
        )

        python_console_submenu = QMenu("Python console", self)
        log_menu.addMenu(python_console_submenu)
        self._python_console_actions = self._build_log_level_submenu(
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
        self._log_to_file_level_actions = self._build_log_level_submenu(
            log_to_file_submenu, self._on_file_log_level_toggled
        )

        # Connection menu
        connection_menu = menubar.addMenu("Connection")
        
        open_connection_action = QAction("Open Connection", self)
        open_connection_action.triggered.connect(self.open_serial_connection)
        connection_menu.addAction(open_connection_action)
        
        close_connection_action = QAction("Close Connection", self)
        close_connection_action.triggered.connect(self.close_serial_connection)
        close_connection_action.setEnabled(False)
        self._close_connection_action = close_connection_action
        connection_menu.addAction(close_connection_action)

        # Machine menu
        machine_menu = menubar.addMenu("Machine")
        model_submenu = QMenu("Model", self)
        machine_menu.addMenu(model_submenu)

        model_group = QActionGroup(self)
        model_group.setExclusive(True)
        self._model_actions = {}
        for model_name in ("PFAFF Creative 7570", "PFAFF Creative 7550", "PFAFF Creative 1475 CD"):
            action = QAction(model_name, self)
            action.setCheckable(True)
            action.triggered.connect(lambda checked, m=model_name: self._on_model_selected(m))
            model_group.addAction(action)
            model_submenu.addAction(action)
            self._model_actions[model_name] = action
        self._model_actions["PFAFF Creative 7570"].setChecked(True)

        pmemory_submenu = QMenu("P-Memory", self)
        machine_menu.addMenu(pmemory_submenu)

        clear_all_action = QAction("Clear all", self)
        clear_all_action.triggered.connect(self._clear_all_pmemory)
        pmemory_submenu.addAction(clear_all_action)

        # Settings menu
        settings_menu = menubar.addMenu("Settings")
        preferences_action = QAction("Preferences", self)
        preferences_action.triggered.connect(self._open_preferences)
        settings_menu.addAction(preferences_action)

        # Help menu
        help_menu = menubar.addMenu("Help")
        about_action = QAction("About", self)
        about_action.triggered.connect(self._show_about)
        help_menu.addAction(about_action)

        self._apply_config_to_menu()

    # ------------------------------------------------------------------
    # Modified-state tracking
    # ------------------------------------------------------------------

    def _set_modified(self, modified: bool):
        """Update the dirty flag, Save action, and window title."""
        self._modified = modified
        self._save_action.setEnabled(modified)
        self._refresh_title()

    def _refresh_title(self):
        """Rebuild the window title from the current file name and dirty flag."""
        suffix = " *" if self._modified else ""
        if self._title_name:
            self.setWindowTitle(f"PFAFF Creative 75xx Emulator - {self._title_name}{suffix}")
        else:
            self.setWindowTitle(f"PFAFF Creative 75xx Emulator{suffix}")

    # ------------------------------------------------------------------
    # Config persistence
    # ------------------------------------------------------------------

    _CONFIG_FILE = "config.json"

    def _load_config(self) -> dict:
        """Load configuration from JSON file, merging missing keys with defaults."""
        default = {
            "recent_files": [],
            "log_window":      {"warning": True, "info": True, "debug": True},
            "python_console":  {"warning": True, "info": True, "debug": True},
            "log_to_file":     {"enabled": False, "warning": True, "info": True, "debug": True},
            "serial":          {"port": None, "baudrate": 4800},
            "machine":         {"model": "PFAFF Creative 7570"},
            "general":         {"auto_connect": False, "open_state_on_start": False},
        }
        base_dir = Path(sys.executable).parent if getattr(sys, 'frozen', False) else Path(__file__).parent
        config_path = base_dir / self._CONFIG_FILE
        if config_path.exists():
            try:
                with open(config_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                for key, val in default.items():
                    if key not in data:
                        data[key] = val
                    elif isinstance(val, dict):
                        for k, v in val.items():
                            data[key].setdefault(k, v)
                return data
            except Exception as e:
                import logging as _log
                _log.getLogger(__name__).warning(f"Failed to load config: {e}")
        return default

    def _save_config(self):
        """Persist current configuration to JSON file."""
        base_dir = Path(sys.executable).parent if getattr(sys, 'frozen', False) else Path(__file__).parent
        config_path = base_dir / self._CONFIG_FILE
        try:
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(self._config, f, indent=2)
        except Exception as e:
            logger.warning(f"Failed to save config: {e}")

    def _apply_config_to_menu(self):
        """Apply loaded configuration to menu checked states (called once after menu creation)."""
        for level, key in self._LEVEL_KEY.items():
            if not self._config["log_window"].get(key, True):
                self._log_window_actions[level].setChecked(False)
            if not self._config["python_console"].get(key, True):
                self._python_console_actions[level].setChecked(False)
        if self._config["log_to_file"].get("enabled", False):
            self._log_to_file_action.setChecked(True)
        for level, key in self._LEVEL_KEY.items():
            if not self._config["log_to_file"].get(key, True):
                self._log_to_file_level_actions[level].setChecked(False)
        saved_model = self._config.get("machine", {}).get("model", "PFAFF Creative 7570")
        if saved_model in self._model_actions:
            self._model_actions[saved_model].setChecked(True)
            self.machine_state.configure_model(saved_model)
            self.protocol.configure_model(saved_model)

    # ------------------------------------------------------------------
    # Recent-files helpers
    # ------------------------------------------------------------------

    def _sync_model_menu_to_state(self):
        """Update the Machine → Model menu checkmark to match machine_state.machine_model."""
        model = self.machine_state.machine_model
        if model and model in self._model_actions:
            self._model_actions[model].setChecked(True)
            self.protocol.configure_model(model)
            self._config.setdefault("machine", {})["model"] = model

    RECENT_MAX = 20

    def _add_to_recent(self, file_path: str):
        """Insert *file_path* at the top of the recent-files list and persist."""
        path = str(Path(file_path).resolve())
        if path in self._recent_files:
            self._recent_files.remove(path)
        self._recent_files.insert(0, path)
        self._recent_files = self._recent_files[:self.RECENT_MAX]
        self._config["recent_files"] = self._recent_files
        self._save_config()
        self._rebuild_recent_menu()

    def _rebuild_recent_menu(self):
        """Rebuild the Open recent submenu from the current list."""
        self._recent_menu.clear()
        for path in self._recent_files:
            action = QAction(path, self)
            action.triggered.connect(lambda checked, p=path: self._open_recent_file(p))
            self._recent_menu.addAction(action)
        if self._recent_files:
            self._recent_menu.addSeparator()
        clear_action = QAction("Clear list", self)
        clear_action.setEnabled(bool(self._recent_files))
        clear_action.triggered.connect(self._clear_recent_files)
        self._recent_menu.addAction(clear_action)
        self._recent_menu.setEnabled(True)

    def _clear_recent_files(self):
        """Clear the recent files list."""
        self._recent_files.clear()
        self._config["recent_files"] = self._recent_files
        self._save_config()
        self._rebuild_recent_menu()

    def _open_recent_file(self, file_path: str):
        """Open a file from the recent list."""
        try:
            self.machine_state.load_from_file(file_path)
            self.current_file = file_path
            self.pmemory_tab.update_ui(self.machine_state)
            self.mmemory_tab.update_ui(self.machine_state)
            self.card_memory_tab.update_ui(self.machine_state)
            self._sync_model_menu_to_state()
            logger.info(f"File opened: {file_path}")
            self._title_name = Path(file_path).name
            self._set_modified(False)
            self._add_to_recent(file_path)
        except Exception as e:
            logger.error(f"Failed to open file: {str(e)}")
            QMessageBox.critical(self, "Error", f"Failed to open file: {str(e)}")
            # Remove from recent list if the file is no longer accessible
            path = str(Path(file_path).resolve())
            if path in self._recent_files:
                self._recent_files.remove(path)
                self._config["recent_files"] = self._recent_files
                self._save_config()
                self._rebuild_recent_menu()

    def new_file(self):
        """Create a new machine state file"""
        for slot in self.machine_state.p_memory_slots:
            slot.clear()
        self.machine_state.m_memory = []
        self.machine_state.clear_card_memory()
        self.current_file = None
        self.pmemory_tab.update_ui(self.machine_state)
        self.mmemory_tab.update_ui(self.machine_state)
        self.card_memory_tab.update_ui(self.machine_state)
        logger.info("New file created")
        self._title_name = "new file"
        self._set_modified(False)
    
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
                self._sync_model_menu_to_state()
                logger.info(f"File opened: {file_path}")
                self._title_name = Path(file_path).name
                self._set_modified(False)
                self._add_to_recent(file_path)
            except Exception as e:
                logger.error(f"Failed to open file: {str(e)}")
                QMessageBox.critical(self, "Error", f"Failed to open file: {str(e)}")

    def save_file(self) -> bool:
        """Save current machine state. Returns True on success."""
        if self.current_file:
            try:
                self.machine_state.save_to_file(self.current_file)
                logger.info(f"File saved: {self.current_file}")
                self._set_modified(False)
                return True
            except Exception as e:
                logger.error(f"Failed to save file: {str(e)}")
                QMessageBox.critical(self, "Error", f"Failed to save file: {str(e)}")
                return False
        else:
            return self.save_file_as()

    def save_file_as(self) -> bool:
        """Save machine state to a new file. Returns True on success."""
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
                self._title_name = Path(file_path).name
                self._set_modified(False)
                self._add_to_recent(file_path)
                return True
            except Exception as e:
                logger.error(f"Failed to save file: {str(e)}")
                QMessageBox.critical(self, "Error", f"Failed to save file: {str(e)}")
                return False
        return False
    
    def open_serial_connection(self):
        """Open serial connection dialog"""
        serial_cfg = self._config.get("serial", {})
        dialog = SerialConnectionDialog(
            self,
            last_port=serial_cfg.get("port"),
            last_baudrate=serial_cfg.get("baudrate", 4800)
        )
        if dialog.exec_():
            port, baudrate = dialog.get_selected_connection()
            if port is None:
                logger.warning("No COM port available")
                QMessageBox.warning(self, "Connection", "No COM ports available")
                return
            
            if self.serial_handler.is_connected:
                logger.info("Closing existing connection before opening a new one")
                self.serial_handler.disconnect()

            logger.info(f"Opening serial connection: {port} at {baudrate} baud")
            if self.serial_handler.connect(port, baudrate):
                self._config["serial"] = {"port": port, "baudrate": baudrate}
                self._save_config()
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

    def _on_connection_changed(self, connected: bool):
        """Enable/disable Close Connection action based on connection state."""
        self._close_connection_action.setEnabled(connected)
    
    _LEVEL_KEY = {logging.WARNING: "warning", logging.INFO: "info", logging.DEBUG: "debug"}

    def _build_log_level_submenu(self, menu: QMenu, slot) -> dict:
        """Add Warning / Info / Debug checkable actions to *menu*, connected to *slot*.
        Returns a dict mapping logging level int to QAction."""
        actions = {}
        for label, level in (("Warning", logging.WARNING), ("Info", logging.INFO), ("Debug", logging.DEBUG)):
            action = QAction(label, self)
            action.setCheckable(True)
            action.setChecked(True)
            action.toggled.connect(lambda checked, lvl=level: slot(lvl, checked))
            menu.addAction(action)
            actions[level] = action
        return actions

    def _open_slot_detail(self, slot):
        """Open (or raise) a detail window for the given slot."""
        slot_id = slot.slot_id
        existing = self._slot_detail_windows.get(slot_id)
        if existing is not None:
            existing.raise_()
            existing.activateWindow()
            return

        def on_clear():
            self._set_modified(True)
            self.pmemory_tab.update_ui(self.machine_state)

        def on_navigate(old_id, new_id):
            if new_id in self._slot_detail_windows:
                self._slot_detail_windows[new_id].raise_()
                self._slot_detail_windows[new_id].activateWindow()
                return False
            w = self._slot_detail_windows.pop(old_id, None)
            if w:
                self._slot_detail_windows[new_id] = w
            return True

        win = SlotDetailWindow(self.machine_state.p_memory_slots, slot_id,
                               on_clear=on_clear, on_navigate=on_navigate,
                               machine_model=self.machine_state.machine_model, parent=self)
        win.destroyed.connect(lambda: [
            self._slot_detail_windows.pop(k, None)
            for k, v in list(self._slot_detail_windows.items()) if v is win
        ])
        self._slot_detail_windows[slot_id] = win
        win.show()

    def _show_about(self):
        """Show the About dialog."""
        QMessageBox.about(
            self,
            "About PFAFF Creative 75xx Emulator",
            "<h3>PFAFF Creative 75xx Emulator</h3>"
            "<p>An emulator for the PFAFF Creative 7570, 7550 and 1475 CD sewing machines, "
            "enabling experiments with communication over a serial interface.</p>"
            "<b>Project:</b> "
            '<a href="https://github.com/arthendev/pfaff7570emu">'
            "github.com/arthendev/pfaff7570emu</a>"
            "<p>© 2026 A. Frej (arthendev)</p>"
        )

    def _open_preferences(self):
        """Open the Preferences dialog and save any changes."""
        dlg = PreferencesDialog(self._config, parent=self)
        if dlg.exec_():
            self._save_config()

    def _try_auto_open_state(self):
        """Load the most recent machine state file at startup if configured."""
        if not self._config.get("general", {}).get("open_state_on_start", False):
            return
        if not self._recent_files:
            return
        file_path = self._recent_files[0]
        try:
            self.machine_state.load_from_file(file_path)
            self.current_file = file_path
            self.pmemory_tab.update_ui(self.machine_state)
            self.mmemory_tab.update_ui(self.machine_state)
            self.card_memory_tab.update_ui(self.machine_state)
            self._sync_model_menu_to_state()
            logger.info(f"Auto-loaded machine state: {file_path}")
            self._title_name = Path(file_path).name
            self._set_modified(False)
        except Exception as e:
            logger.warning(f"Auto-load machine state failed: {e}")

    def _try_auto_connect(self):
        """Attempt automatic serial connection at startup if configured."""
        if not self._config.get("general", {}).get("auto_connect", False):
            return
        serial_cfg = self._config.get("serial", {})
        port = serial_cfg.get("port")
        baudrate = serial_cfg.get("baudrate", 4800)
        if not port:
            return
        import serial.tools.list_ports
        available = [p.device for p in serial.tools.list_ports.comports()]
        if port not in available:
            logger.info(f"Auto-connect: port {port} not available, skipping")
            return
        logger.info(f"Auto-connect: opening {port} at {baudrate} baud")
        if not self.serial_handler.connect(port, baudrate):
            logger.warning(f"Auto-connect: failed to open {port}")

    def _on_model_selected(self, model_name: str):
        """Apply model configuration and persist the selected model."""
        self.machine_state.configure_model(model_name)
        self.protocol.configure_model(model_name)
        self.pmemory_tab.update_ui(self.machine_state)
        self._config.setdefault("machine", {})["model"] = model_name
        self._save_config()
        logger.info(f"Model set to: {model_name}")

    def _on_log_level_toggled(self, level: int, checked: bool):
        """Show or hide a log level in the Qt console."""
        self.console_handler.set_level_visible(level, checked)
        key = self._LEVEL_KEY.get(level)
        if key:
            self._config["log_window"][key] = checked
            self._save_config()

    def _on_python_log_level_toggled(self, level: int, checked: bool):
        """Show or hide a log level in the Python (stdout) console."""
        self.python_console_handler.set_level_visible(level, checked)
        key = self._LEVEL_KEY.get(level)
        if key:
            self._config["python_console"][key] = checked
            self._save_config()

    def _on_log_to_file_toggled(self, checked: bool):
        """Enable or disable logging to file."""
        import datetime
        import os
        self._config["log_to_file"]["enabled"] = checked
        self._save_config()
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
        key = self._LEVEL_KEY.get(level)
        if key:
            self._config["log_to_file"][key] = checked
            self._save_config()

    def _show_console_context_menu(self, pos):
        """Show right-click context menu on the log console."""
        menu = self.console.createStandardContextMenu()
        menu.addSeparator()
        clear_action = QAction("Clear", self)
        clear_action.triggered.connect(self.console.clear)
        menu.addAction(clear_action)
        menu.exec_(self.console.mapToGlobal(pos))

    def _clear_all_pmemory(self):
        """Clear all P-Memory slots, resetting them to Empty."""
        for slot in self.machine_state.p_memory_slots:
            slot.clear()
        self.pmemory_tab.update_ui(self.machine_state)
        logger.info("P-Memory: all slots cleared")
        self._set_modified(True)

    def _on_pmemory_changed(self):
        """Refresh P-Memory tab after a delete or write operation"""
        self.pmemory_tab.update_ui(self.machine_state)
        self._set_modified(True)
        for win in list(self._slot_detail_windows.values()):
            win._load_slot()

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
        if self._modified:
            reply = QMessageBox.question(
                self,
                "Unsaved Changes",
                "There are unsaved changes. Do you want to save them?",
                QMessageBox.Save | QMessageBox.Discard | QMessageBox.Cancel,
                QMessageBox.Save
            )
            if reply == QMessageBox.Save:
                if not self.save_file():
                    event.ignore()
                    return
            elif reply == QMessageBox.Cancel:
                event.ignore()
                return
        if self.serial_handler.is_connected:
            self.serial_handler.disconnect()
        self.console_handler.close()
        event.accept()


def main():
    """Application entry point"""
    app = QApplication(sys.argv)
    window = PfaffCreativeEmulator()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
