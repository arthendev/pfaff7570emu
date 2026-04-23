"""
Slot detail window - shows full information about a single P-Memory slot.
"""

from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel,
                             QPushButton, QTextEdit, QGroupBox, QTabWidget, QWidget,
                             QScrollArea, QGridLayout, QSizePolicy, QSpacerItem,
                             QTableWidget, QTableWidgetItem, QMenu, QApplication)
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QFont, QColor

from machine_state import MemorySlot
from pmemory_tab import PatternPreview


class ClickableLabel(QLabel):
    """A QLabel that emits clicked(index) when clicked."""
    clicked = pyqtSignal(int)

    def __init__(self, text="", idx=None, parent=None):
        super().__init__(text, parent)
        self._idx = idx

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            try:
                self.clicked.emit(self._idx)
            except Exception:
                pass
        super().mousePressEvent(event)


class SlotDetailWindow(QDialog):
    """Non-modal window showing detailed information about a P-Memory slot."""

    def __init__(self, slots: list, slot_id: int, on_clear=None, on_navigate=None, parent=None):
        super().__init__(parent)
        self._slots = slots
        self.slot = slots[slot_id]
        self._on_clear_callback = on_clear
        self._on_navigate = on_navigate
        self.setWindowTitle(f"Slot P {slot_id} - Details")
        self.setWindowFlags(Qt.Window)
        self.setAttribute(Qt.WA_DeleteOnClose)
        self.setMinimumWidth(620)
        self.setMinimumHeight(800)
        self._setup_ui()
        self._load_slot()

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------

    def _setup_ui(self):
        layout = QVBoxLayout()
        layout.setSpacing(8)

        # Info row
        info_layout = QHBoxLayout()
        bold = QFont()
        bold.setBold(True)
        bold.setPointSize(10)

        self._prev_btn = QPushButton("◀")
        self._prev_btn.clicked.connect(lambda: self._navigate(-1))

        self._next_btn = QPushButton("▶")
        self._next_btn.clicked.connect(lambda: self._navigate(+1))

        self._slot_label = QLabel()
        self._type_label = QLabel()
        self._bytes_label = QLabel()
        self._stitches_label = QLabel()
        for lbl in (self._slot_label, self._type_label, self._bytes_label, self._stitches_label):
            lbl.setFont(bold)
            info_layout.addWidget(lbl)
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

        # Tabbed widget
        tabs = QTabWidget()

        # Tab 1: Raw data
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

        # Tab 2: Header
        header_tab = QWidget()
        header_layout = QVBoxLayout()

        h_header_label = QLabel("Header (raw)")
        h_header_font = QFont()
        h_header_font.setBold(True)
        h_header_label.setFont(h_header_font)
        header_layout.addWidget(h_header_label)

        self._header_edit_2 = QTextEdit()
        self._header_edit_2.setReadOnly(True)
        self._header_edit_2.setFont(QFont("Courier New", 9))
        self._header_edit_2.setFixedHeight(60)
        header_layout.addWidget(self._header_edit_2)

        header_layout.addWidget(QLabel("Byte Analysis:"))

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        self._header_grid_widget = QWidget()
        grid_widget = self._header_grid_widget
        self._header_grid = QGridLayout(grid_widget)
        self._header_grid.setHorizontalSpacing(2)
        self._header_grid.setVerticalSpacing(5)
        grid_widget.setContextMenuPolicy(Qt.CustomContextMenu)
        grid_widget.customContextMenuRequested.connect(self._on_header_grid_context_menu)
        scroll.setWidget(grid_widget)
        header_layout.addWidget(scroll)

        header_tab.setLayout(header_layout)
        tabs.addTab(header_tab, "Header")

        # Tab 3: Pattern (points table)
        pattern_tab = QWidget()
        pattern_layout = QVBoxLayout()
        pattern_layout.setAlignment(Qt.AlignTop)

        pattern_scroll = QScrollArea()
        pattern_scroll.setWidgetResizable(True)
        # Use a table for points: easier selection and built-in features
        table_container = QWidget()
        table_layout = QVBoxLayout(table_container)
        table_layout.setContentsMargins(0, 0, 0, 0)
        self._points_table = QTableWidget()
        self._points_table.setColumnCount(5)
        self._points_table.setHorizontalHeaderLabels(["#", "Hex (x,y)", "Dec (x,y)", "Hex diff(n, n-1)", "Dec diff(n, n-1)"])
        self._points_table.setSelectionBehavior(QTableWidget.SelectRows)
        self._points_table.setSelectionMode(QTableWidget.SingleSelection)
        self._points_table.verticalHeader().setVisible(False)
        self._points_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self._points_table.itemSelectionChanged.connect(self._on_point_table_selection_changed)
        self._points_table.setContextMenuPolicy(Qt.CustomContextMenu)
        self._points_table.customContextMenuRequested.connect(self._on_points_table_context_menu)
        table_layout.addWidget(self._points_table)
        pattern_scroll.setWidget(table_container)
        pattern_layout.addWidget(pattern_scroll)

        pattern_tab.setLayout(pattern_layout)
        tabs.addTab(pattern_tab, "Pattern")

        # Tab 4: Stats (statistics table)
        stats_tab = QWidget()
        stats_layout = QVBoxLayout()
        stats_layout.setAlignment(Qt.AlignTop)

        stats_scroll = QScrollArea()
        stats_scroll.setWidgetResizable(True)
        self._pattern_grid_widget = QWidget()
        pattern_grid_widget = self._pattern_grid_widget
        self._pattern_grid = QGridLayout(pattern_grid_widget)
        self._pattern_grid.setSpacing(6)
        pattern_grid_widget.setContextMenuPolicy(Qt.CustomContextMenu)
        pattern_grid_widget.customContextMenuRequested.connect(self._on_stats_grid_context_menu)
        stats_scroll.setWidget(pattern_grid_widget)
        stats_layout.addWidget(stats_scroll)

        stats_tab.setLayout(stats_layout)
        tabs.addTab(stats_tab, "Stats")

        layout.addWidget(tabs)

        # Buttons
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        self._clear_btn = QPushButton("Clear")
        self._clear_btn.clicked.connect(self._clear_slot)
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.close)
        btn_layout.addWidget(self._clear_btn)
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
        self.setWindowTitle(f"Slot P {new_id} - Details")
        self._load_slot()

    def _update_nav_buttons(self):
        idx = self.slot.slot_id
        self._prev_btn.setEnabled(idx > 0)
        self._next_btn.setEnabled(idx < len(self._slots) - 1)

    def _parse_pattern_stats(self):
        """Return pattern statistics computed from the slot's pattern data."""
        return self.slot.get_pattern_stats()

    def _get_header_bytes(self):
        """Extract header bytes from header_raw string."""
        bytes_list = []
        raw = self.slot.header_raw
        for i in range(0, len(raw), 2):
            try:
                byte_val = int(raw[i:i+2], 16)
                bytes_list.append(byte_val)
            except (ValueError, IndexError):
                bytes_list.append(None)
        return bytes_list

    def _populate_header_grid(self):
        """Populate the header byte analysis grid."""
        while self._header_grid.count():
            item = self._header_grid.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        header_bytes = self._get_header_bytes()
        stats = self._parse_pattern_stats()

        mapping = {
            0: ("y_min", "min(ys)"),
            1: ("y_max", "max(ys)"),
            2: ("dx_abs_max", "max(abs(dxs))"),
            5: (None, "Unknown"),
            7: ("d0x_min_abs", "abs(min(dxs)-xs[0]))"),
            9: ("pn_x", "xs[end]"),
            11: ("span_x", "max(xs) - min(xs)"),
            13: ("y_min_to_bound", "0x36 - min(ys)"),
            15: ("span_y", "max(ys) - min(ys)"),
        }

        row = 0
        for idx in range(max(16, len(header_bytes))):
            # Get header byte value
            h_byte = header_bytes[idx] if idx < len(header_bytes) else None
            h_text = f"0x{h_byte:02X} {h_byte:4d}" if h_byte is not None else "--"

            # Column 1: Label + hex value
            lbl = QLabel(f"H[{idx:2d}]  {h_text}")
            lbl.setFont(QFont("Courier New", 9))
            self._header_grid.addWidget(lbl, row, 0)

            # Column 2: Statistic & value - format: 0xNN  dec  var_name
            if idx in mapping:
                stat_key, stat_label = mapping[idx]
                lbl.setToolTip(stat_label)
                if stat_key is not None:
                    stat_val = stats.get(stat_key)
                    if stat_val is not None:
                        desc = f"0x{stat_val & 0xFF:02X}  {stat_val:4d}  {stat_key}"
                    else:
                        desc = f"--    {'--':>4}  {stat_key}"
                else:  # None = unknown, no stat
                    desc = f"--    {'--':>4}  unknown"
            else:
                desc = ""

            stat_lbl = QLabel(desc)
            stat_lbl.setFont(QFont("Courier New", 9))
            self._header_grid.addWidget(stat_lbl, row, 1)

            # Column 3: OK/NOK status
            status_lbl = QLabel()
            status_font = QFont()
            status_font.setBold(True)
            status_lbl.setFont(status_font)

            if idx in mapping:
                stat_key, _ = mapping[idx]
                if stat_key is not None:
                    stat_val = stats.get(stat_key)
                    expected = stat_val & 0xFF if stat_val is not None else None
                else:
                    expected = None

                if expected is not None and h_byte is not None:
                    if h_byte == expected:
                        status_lbl.setText("OK")
                        status_lbl.setStyleSheet("color: green;")
                    else:
                        status_lbl.setText("NOK")
                        status_lbl.setStyleSheet("color: red;")
                else:
                    status_lbl.setText("--")
            else:
                # Unmapped bytes: expected to be 0
                if h_byte is not None:
                    if h_byte == 0:
                        status_lbl.setText("OK")
                        status_lbl.setStyleSheet("color: green;")
                    else:
                        status_lbl.setText("NOK")
                        status_lbl.setStyleSheet("color: red;")
                else:
                    status_lbl.setText("--")

            self._header_grid.addWidget(status_lbl, row, 2)
            row += 1

        self._header_grid.addItem(
            QSpacerItem(0, 0, QSizePolicy.Minimum, QSizePolicy.Expanding), row, 0, 1, 3)

    def _populate_pattern_grid(self):
        """Fill the Pattern tab with all pattern statistics."""
        while self._pattern_grid.count():
            item = self._pattern_grid.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        s = self.slot.get_pattern_stats()
        mono = QFont("Courier New", 9)
        bold_font = QFont()
        bold_font.setBold(True)

        rows = [
            ("n",                   s["n"],             None),
            ("x_min",               s["x_min"],         None),
            ("x_max",               s["x_max"],         None),
            ("y_min",               s["y_min"],         None),
            ("y_max",               s["y_max"],         None),
            ("y_min_to_bound",      s["y_min_to_bound"], None),
            ("span_x",              s["span_x"],        None),
            ("span_y",              s["span_y"],        None),
            ("dx_max",              s["dx_max"],        None),
            ("dx_min",              s["dx_min"],        None),
            ("dx_min_abs",          s["dx_min_abs"],    None),
            ("dx_abs_max",          s["dx_abs_max"],    None),
            ("dy_max",              s["dy_max"],        None),
            ("dy_min",              s["dy_min"],        None),
            ("dy_min_abs",          s["dy_min_abs"],    None),
            ("dy_abs_max",          s["dy_abs_max"],    None),
            ("is_reversed",         s["is_reversed"],   None),
            ("dx_0n",               s["dx_0n"],         None),
            ("dx_0n_abs",           s["dx_0n_abs"],     None),
            ("d0x_max",             s["d0x_max"],       None),
            ("d0x_min",             s["d0x_min"],       None),
            ("d0x_min_abs",         s["d0x_min_abs"],   None),
            ("d0y_max",             s["d0y_max"],       None),
            ("d0y_min",             s["d0y_min"],       None),
            ("d0y_min_abs",         s["d0y_min_abs"],   None),
            ("p0_x",                s["p0_x"],          None),
            ("p0_y",                s["p0_y"],          None),
            ("p1_x",                s["p1_x"],          None),
            ("p1_y",                s["p1_y"],          None),
            ("p1_dx",               s["p1_dx"],         None),
            ("p1_dy",               s["p1_dy"],         None),
            ("p1_dx_abs",           s["p1_dx_abs"],     None),
            ("p1_dy_abs",           s["p1_dy_abs"],     None),
            ("pn_x",                s["pn_x"],          None),
            ("pn_y",                s["pn_y"],          None),
            ("pn_dx",               s["pn_dx"],         None),
            ("pn_dy",               s["pn_dy"],         None),
            ("pn_dx_abs",           s["pn_dx_abs"],     None),
            ("pn_dy_abs",           s["pn_dy_abs"],     None),
            ("dnx_max",             s["dnx_max"],       None),
            ("dnx_min",             s["dnx_min"],       None),
            ("dnx_min_abs",         s["dnx_min_abs"],   None),
            ("dny_max",             s["dny_max"],       None),
            ("dny_min",             s["dny_min"],       None),
            ("dny_min_abs",         s["dny_min_abs"],   None),
            ("checksum",            s["checksum"],      None),
        ]

        # Header row
        for col, text in enumerate(("Statistic", "Hex", "Dec")):
            hdr = QLabel(text)
            hdr.setFont(bold_font)
            self._pattern_grid.addWidget(hdr, 0, col)

        for row_idx, (label, val, _) in enumerate(rows, start=1):
            lbl = QLabel(label)
            lbl.setFont(mono)
            self._pattern_grid.addWidget(lbl, row_idx, 0)

            if isinstance(val, bool) or val is None:
                hex_text = "--"
            else:
                hex_text = f"0x{val & 0xFF:02X}"
            hex_lbl = QLabel(hex_text)
            hex_lbl.setFont(mono)
            self._pattern_grid.addWidget(hex_lbl, row_idx, 1)

            dec_text = str(val) if val is not None else "--"
            dec_lbl = QLabel(dec_text)
            dec_lbl.setFont(mono)
            self._pattern_grid.addWidget(dec_lbl, row_idx, 2)

        self._pattern_grid.addItem(
            QSpacerItem(0, 0, QSizePolicy.Minimum, QSizePolicy.Expanding), len(rows) + 1, 0, 1, 3)


    def _populate_points_grid(self):
        """Fill the Pattern tab with point list: index, (x,y) Dec, (x,y) Hex."""
        # Populate the points table
        data = list(self.slot.data)
        xs = data[0::2]
        ys = data[1::2]
        n = min(len(xs), len(ys))

        mono = QFont("Courier New", 9)

        if n == 0:
            self._points_table.setRowCount(1)
            it = QTableWidgetItem("--")
            it.setFont(mono)
            self._points_table.setItem(0, 0, it)
            self._points_table.setItem(0, 1, QTableWidgetItem(""))
            self._points_table.setItem(0, 2, QTableWidgetItem(""))
            self._points_table.setItem(0, 3, QTableWidgetItem(""))
            self._points_table.setItem(0, 4, QTableWidgetItem(""))
            return

        self._points_table.setRowCount(n)
        for i in range(n):
            idx_it = QTableWidgetItem(str(i))
            idx_it.setFont(mono)

            dec_text = f"({xs[i]}, {ys[i]})"
            dec_it = QTableWidgetItem(dec_text)
            dec_it.setFont(mono)

            hex_text = f"(0x{xs[i] & 0xFF:02X}, 0x{ys[i] & 0xFF:02X})"
            hex_it = QTableWidgetItem(hex_text)
            hex_it.setFont(mono)

            if i == 0:
                diff_dec_text = "--"
                diff_hex_text = "--"
            else:
                dx = xs[i] - xs[i-1]
                dy = ys[i] - ys[i-1]
                diff_dec_text = f"({dx}, {dy})"
                diff_hex_text = f"(0x{(dx & 0xFF):02X}, 0x{(dy & 0xFF):02X})"

            diff_dec_it = QTableWidgetItem(diff_dec_text)
            diff_dec_it.setFont(mono)
            diff_hex_it = QTableWidgetItem(diff_hex_text)
            diff_hex_it.setFont(mono)

            self._points_table.setItem(i, 0, idx_it)
            self._points_table.setItem(i, 1, hex_it)
            self._points_table.setItem(i, 2, dec_it)
            self._points_table.setItem(i, 3, diff_hex_it)
            self._points_table.setItem(i, 4, diff_dec_it)

        # Ensure table expands to fill available space and size columns
        self._points_table.setSizePolicy(self._points_table.sizePolicy().horizontalPolicy(), QSizePolicy.Expanding)
        try:
            self._points_table.resizeColumnsToContents()
        except Exception:
            pass

    def _grid_to_text(self, grid: QGridLayout) -> str:
        """Extract all text from a QGridLayout of QLabels as tab-separated rows."""
        lines = []
        for row in range(grid.rowCount()):
            parts = []
            for col in range(grid.columnCount()):
                item = grid.itemAtPosition(row, col)
                if item and item.widget() and isinstance(item.widget(), QLabel):
                    parts.append(item.widget().text())
                else:
                    parts.append("")
            if any(p for p in parts):
                lines.append("\t".join(parts))
        return "\n".join(lines)

    def _on_header_grid_context_menu(self, pos):
        menu = QMenu(self)
        copy_action = menu.addAction("Copy")
        action = menu.exec_(self._header_grid_widget.mapToGlobal(pos))
        if action == copy_action:
            QApplication.clipboard().setText(self._grid_to_text(self._header_grid))

    def _on_stats_grid_context_menu(self, pos):
        menu = QMenu(self)
        copy_action = menu.addAction("Copy")
        action = menu.exec_(self._pattern_grid_widget.mapToGlobal(pos))
        if action == copy_action:
            QApplication.clipboard().setText(self._grid_to_text(self._pattern_grid))

    def _on_points_table_context_menu(self, pos):
        menu = QMenu(self)
        copy_action = menu.addAction("Copy")
        action = menu.exec_(self._points_table.viewport().mapToGlobal(pos))
        if action == copy_action:
            col_count = self._points_table.columnCount()
            row_count = self._points_table.rowCount()
            lines = []
            headers = []
            for c in range(col_count):
                h = self._points_table.horizontalHeaderItem(c)
                headers.append(h.text() if h else "")
            lines.append("\t".join(headers))
            for row in range(row_count):
                parts = []
                for col in range(col_count):
                    item = self._points_table.item(row, col)
                    parts.append(item.text() if item else "")
                lines.append("\t".join(parts))
            QApplication.clipboard().setText("\n".join(lines))

    def _on_point_row_clicked(self, idx: int):
        """Handle user clicking a point row: highlight preview and row."""
        # legacy handler for ClickableLabel rows (kept for compatibility)
        try:
            self._preview.selected_point = idx
            self._preview.update()
        except Exception:
            pass

    def _on_point_table_selection_changed(self):
        """Handle selection change in the points `QTableWidget`."""
        sels = self._points_table.selectionModel().selectedRows()
        if not sels:
            self._preview.selected_point = None
            self._preview.update()
            return
        row = sels[0].row()
        # ensure within range
        try:
            self._preview.selected_point = int(self._points_table.item(row, 0).text())
        except Exception:
            self._preview.selected_point = row
        self._preview.update()

    # ------------------------------------------------------------------
    # Data helpers (load/refresh)
    # ------------------------------------------------------------------

    def _load_slot(self):
        """Populate all fields from the current slot data."""
        self._slot_label.setText(f"Slot:  P {self.slot.slot_id}")
        self._type_label.setText(f"    Type:  {self.slot.slot_type}")
        self._bytes_label.setText(f"    Bytes:  {self.slot.get_size_bytes()}")
        stitches = len(self.slot.data) // 2 if self.slot.slot_type != "Empty" else 0
        self._stitches_label.setText(f"    Stitches:  {stitches}")
        self._preview.slot_data = list(self.slot.data)
        self._preview.slot_type = self.slot.slot_type
        self._preview.update()
        self._header_edit.setPlainText(self.slot.header_raw)
        self._header_edit_2.setPlainText(self.slot.header_raw)
        self._pattern_edit.setPlainText(self.slot.pattern_raw)
        self._populate_header_grid()
        self._populate_pattern_grid()
        self._populate_points_grid()
        self._clear_btn.setEnabled(self.slot.slot_type != "Empty")
        self._update_nav_buttons()

    def refresh(self):
        """Re-read from the slot and update all displayed fields."""
        self._load_slot()

    def _clear_slot(self):
        """Clear the slot, refresh display, and notify the main window."""
        self.slot.slot_type = "Empty"
        self.slot.data = []
        self.slot.header_raw = ""
        self.slot.pattern_raw = ""
        self.refresh()
        if self._on_clear_callback:
            self._on_clear_callback()
