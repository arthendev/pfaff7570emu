"""
M-Memory Slot detail window - shows information about a single M-Memory slot.
"""

from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel,
                             QPushButton, QTextEdit, QGroupBox, QTabWidget, QWidget,
                             QSizePolicy, QShortcut, QCheckBox, QTableWidget,
                             QTableWidgetItem)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont, QKeySequence

from machine_state import MMemorySlot
from pmemory_tab import PatternPreview


class MMemSlotDetailWindow(QDialog):
    """Non-modal window showing detailed information about an M-Memory slot."""

    def __init__(self, slots: list, slot_id: int, on_navigate=None, parent=None):
        super().__init__(parent)
        self._slots = slots
        self.slot = slots[slot_id]
        self._on_navigate = on_navigate
        self.setWindowTitle(f"Slot M {slot_id} - Details")
        self.setWindowFlags(Qt.Window)
        self.setAttribute(Qt.WA_DeleteOnClose)
        self.setMinimumWidth(400)
        self.setMinimumHeight(400)
        self.resize(720, 800)
        self._setup_ui()
        self._load_slot()

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------

    def _setup_ui(self):
        layout = QVBoxLayout()
        layout.setSpacing(8)

        # Info row: Slot, Patterns No
        info_layout = QHBoxLayout()
        bold = QFont()
        bold.setBold(True)
        bold.setPointSize(10)

        self._prev_btn = QPushButton("◀")
        self._prev_btn.clicked.connect(lambda: self._navigate(-1))

        self._next_btn = QPushButton("▶")
        self._next_btn.clicked.connect(lambda: self._navigate(+1))

        QShortcut(QKeySequence("Right"), self).activated.connect(lambda: self._navigate(+1))
        QShortcut(QKeySequence("Left"), self).activated.connect(lambda: self._navigate(-1))

        self._slot_label = QLabel()
        self._slot_label.setFont(bold)
        info_layout.addWidget(self._slot_label)

        self._patterns_label = QLabel()
        self._patterns_label.setFont(bold)
        info_layout.addWidget(self._patterns_label)

        info_layout.addStretch()

        for btn in (self._prev_btn, self._next_btn):
            sp = btn.sizePolicy()
            sp.setHorizontalPolicy(QSizePolicy.Minimum)
            btn.setSizePolicy(sp)
            info_layout.addWidget(btn)

        layout.addLayout(info_layout)

        # Pattern preview
        preview_group = QGroupBox("Pattern Preview")
        preview_layout = QVBoxLayout()
        self._preview = PatternPreview([], "", show_points=True)
        self._preview.setFixedHeight(120)
        preview_layout.addWidget(self._preview)
        preview_group.setLayout(preview_layout)
        layout.addWidget(preview_group)

        # Logical split checkbox
        split_row = QHBoxLayout()
        self._logical_split_cb = QCheckBox("Logical split")
        self._logical_split_cb.stateChanged.connect(self._refresh_raw_display)
        split_row.addWidget(self._logical_split_cb)
        split_row.addStretch()
        layout.addLayout(split_row)

        # Tab widget
        tabs = QTabWidget()

        # Tab: Raw data
        raw_tab = QWidget()
        raw_layout = QVBoxLayout()

        header_label = QLabel("Header (raw)")
        header_font = QFont()
        header_font.setBold(True)
        header_label.setFont(header_font)
        raw_layout.addWidget(header_label)

        self._header_edit = QTextEdit()
        self._header_edit.setReadOnly(True)
        self._header_edit.setFont(QFont("Courier New", 9))
        self._header_edit.setFixedHeight(80)
        raw_layout.addWidget(self._header_edit)

        pattern_label = QLabel("Pattern (raw)")
        pattern_label.setFont(header_font)
        raw_layout.addWidget(pattern_label)

        self._pattern_edit = QTextEdit()
        self._pattern_edit.setReadOnly(True)
        self._pattern_edit.setFont(QFont("Courier New", 9))
        raw_layout.addWidget(self._pattern_edit)

        raw_tab.setLayout(raw_layout)
        tabs.addTab(raw_tab, "Raw data")

        # Tab: Pattern (pattern entries table)
        pattern_tab = QWidget()
        pattern_layout = QVBoxLayout()
        self._pattern_table = QTableWidget()
        self._pattern_table.setColumnCount(9)
        self._pattern_table.setHorizontalHeaderLabels(
            ["#", "Mirror", "Pattern", "Scale-L", "Scale-W", "Pat. group", "Pat. No", "Mirror-W", "Mirror-L"]
        )
        self._pattern_table.setSelectionBehavior(QTableWidget.SelectRows)
        self._pattern_table.setSelectionMode(QTableWidget.SingleSelection)
        self._pattern_table.verticalHeader().setVisible(False)
        self._pattern_table.setEditTriggers(QTableWidget.NoEditTriggers)
        pattern_layout.addWidget(self._pattern_table)
        pattern_tab.setLayout(pattern_layout)
        tabs.addTab(pattern_tab, "Pattern")

        layout.addWidget(tabs)

        # Buttons
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.close)
        btn_layout.addWidget(close_btn)
        layout.addLayout(btn_layout)

        self.setLayout(layout)

    # ------------------------------------------------------------------
    # Data helpers
    # ------------------------------------------------------------------

    def _navigate(self, delta: int):
        """Switch to an adjacent slot."""
        old_id = self.slot.slot_id
        new_id = old_id + delta
        if not (0 <= new_id < len(self._slots)):
            return
        if self._on_navigate and not self._on_navigate(old_id, new_id):
            return
        self.slot = self._slots[new_id]
        self.setWindowTitle(f"Slot M {new_id} - Details")
        self._load_slot()

    def _update_nav_buttons(self):
        idx = self.slot.slot_id
        self._prev_btn.setEnabled(idx > 0)
        self._next_btn.setEnabled(idx < len(self._slots) - 1)

    def _load_slot(self):
        """Populate all fields from the current slot data."""
        self._slot_label.setText(f"Slot:  M {self.slot.slot_id}")
        self._patterns_label.setText(f"    Patterns No:  {self.slot.get_size_patterns()}")
        self._preview.pattern_xy = list(self.slot.pattern_xy)
        self._preview.pattern_type = ""
        self._preview.update()
        self._refresh_raw_display()
        self._populate_pattern_table()
        self._update_nav_buttons()

    def _refresh_raw_display(self):
        """Update header and pattern raw text edits (respects logical split)."""
        header_text = self.slot.header_raw
        self._header_edit.setPlainText(header_text)
        raw_text = self.slot.sequence_raw
        if self._logical_split_cb.isChecked() and raw_text:
            lines = []
            for i in range(0, len(raw_text), 8):
                group = raw_text[i:i+8]
                if len(group) >= 8:
                    group = f"{group[:4]} {group[4:6]} {group[6:8]}"
                elif len(group) >= 6:
                    group = f"{group[:4]} {group[4:6]} {group[6:]}"
                elif len(group) >= 4:
                    group = f"{group[:4]} {group[4:]}"
                lines.append(group)
            raw_text = '\n'.join(lines)
        self._pattern_edit.setPlainText(raw_text)

    _PAT_GROUP_MAP = {
        0x0: "9mm",
        0x2: "MAXI",
        0x3: "Script",
        0x4: "Block",
        0x5: "Outline",
        0xA: "Cursive",
    }

    def _populate_pattern_table(self):
        """Fill the Pattern tab with the sequence of stitch pattern entries."""
        raw_text = self.slot.sequence_raw
        n = self.slot.get_size_patterns()
        mono = QFont("Courier New", 9)

        if n == 0:
            self._pattern_table.setRowCount(1)
            it = QTableWidgetItem("--")
            it.setFont(mono)
            self._pattern_table.setItem(0, 0, it)
            for col in range(1, 9):
                self._pattern_table.setItem(0, col, QTableWidgetItem(""))
            return

        self._pattern_table.setRowCount(n)
        for i in range(n):
            group = raw_text[i * 8 : i * 8 + 8]
            if len(group) < 8:
                break

            # Parse the 4 bytes from hex ASCII
            try:
                b0 = int(group[0:2], 16)
                b0h = int(group[0:1], 16) # mirror
                b0l = int(group[1:2], 16) # pattern group
                b1 = int(group[2:4], 16)  # pattern number
                b2 = int(group[4:6], 16)  # scale-L
                b3 = int(group[6:8], 16)  # scale-W
            except ValueError:
                continue

            # Col 0: row number
            idx_it = QTableWidgetItem(str(i + 1))
            idx_it.setFont(mono)

            # Col 1: Mirror – high nibble of the first byte
            mirror_it = QTableWidgetItem(f"{group[0:1]}")
            mirror_it.setFont(mono)

            # Col 2: Pattern – low nibble of first byte + full second byte
            pat_it = QTableWidgetItem(f"{group[1:2]} {group[2:4]}")
            pat_it.setFont(mono)

            # Col 3: Scale-L (second byte)
            scalel_it = QTableWidgetItem(f"{group[4:6]}")
            scalel_it.setFont(mono)

            # Col 4: Scale-W (third byte)
            scalew_it = QTableWidgetItem(f"{group[6:8]}")
            scalew_it.setFont(mono)

            # Col 5 : Pat. group (low nibble of byte 0 mapped to description)
            group_name = self._PAT_GROUP_MAP.get(b0l, f"0x{b0l:02X}")
            group_it = QTableWidgetItem(group_name)
            group_it.setFont(mono)

            # Col 6: Pat. No (second byte in decimal)
            pat_no_it = QTableWidgetItem(str(b1))
            pat_no_it.setFont(mono)

            # Col 7: Mirror-W – bit 3 of Mirror
            mirror_w = (b0h >> 3) & 1
            mirrorw_it = QTableWidgetItem(str(mirror_w))
            mirrorw_it.setFont(mono)

            # Col 8: Mirror-L – bit 1 of Mirror
            mirror_l = (b0h >> 1) & 1
            mirrorl_it = QTableWidgetItem(str(mirror_l))
            mirrorl_it.setFont(mono)

            self._pattern_table.setItem(i, 0, idx_it)
            self._pattern_table.setItem(i, 1, mirror_it)
            self._pattern_table.setItem(i, 2, pat_it)
            self._pattern_table.setItem(i, 3, scalel_it)
            self._pattern_table.setItem(i, 4, scalew_it)
            self._pattern_table.setItem(i, 5, group_it)
            self._pattern_table.setItem(i, 6, pat_no_it)
            self._pattern_table.setItem(i, 7, mirrorw_it)
            self._pattern_table.setItem(i, 8, mirrorl_it)

        try:
            self._pattern_table.resizeColumnsToContents()
        except Exception:
            pass
