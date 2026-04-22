"""
Preferences dialog
"""

from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QTabWidget,
                             QWidget, QCheckBox, QPushButton, QLabel)
from PyQt5.QtCore import Qt


class PreferencesDialog(QDialog):
    """Application preferences dialog with tabbed layout."""

    def __init__(self, config: dict, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Preferences")
        self.setModal(True)
        self.setMinimumWidth(380)
        self._config = config
        self._setup_ui()
        self._load_values()

    def _setup_ui(self):
        layout = QVBoxLayout()

        self._tabs = QTabWidget()

        # --- General tab ---
        general_tab = QWidget()
        general_layout = QVBoxLayout()
        general_layout.setAlignment(Qt.AlignTop)

        self._auto_connect_cb = QCheckBox("Open connection at start")
        self._auto_connect_cb.setToolTip(
            "If the last used COM port is available at startup, "
            "open the serial connection automatically."
        )
        general_layout.addWidget(self._auto_connect_cb)

        self._open_state_cb = QCheckBox("Open machine state upon start")
        self._open_state_cb.setToolTip(
            "Automatically load the most recently opened machine state file at startup."
        )
        general_layout.addWidget(self._open_state_cb)
        general_layout.addStretch()
        general_tab.setLayout(general_layout)

        self._tabs.addTab(general_tab, "General")
        layout.addWidget(self._tabs)

        # OK / Cancel buttons
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        ok_btn = QPushButton("OK")
        ok_btn.setDefault(True)
        ok_btn.clicked.connect(self._accept)
        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        btn_layout.addWidget(ok_btn)
        btn_layout.addWidget(cancel_btn)
        layout.addLayout(btn_layout)

        self.setLayout(layout)

    def _load_values(self):
        general = self._config.get("general", {})
        self._auto_connect_cb.setChecked(general.get("auto_connect", False))
        self._open_state_cb.setChecked(general.get("open_state_on_start", False))

    def _accept(self):
        general = self._config.setdefault("general", {})
        general["auto_connect"] = self._auto_connect_cb.isChecked()
        general["open_state_on_start"] = self._open_state_cb.isChecked()
        self.accept()
