"""
Card Slot detail window - shows full information about a single memory card slot.

This is adapted from slot_detail_window.py for card (memory card) slot data.
"""

from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel,
                             QPushButton, QTextEdit, QGroupBox, QTabWidget, QWidget,
                             QScrollArea, QGridLayout, QSizePolicy, QSpacerItem,
                             QTableWidget, QTableWidgetItem, QMenu, QApplication,
                             QCheckBox)
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtWidgets import QShortcut
from PyQt5.QtGui import QKeySequence
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


class CardSlotDetailWindow(QDialog):
    """Non-modal window showing detailed information about a memory card slot."""

    def __init__(self, slots: list, slot_id: int, on_clear=None, on_navigate=None,
                 machine_model: str = None, parent=None):
        super().__init__(parent)
        self._slots = slots
        self.slot = slots[slot_id]
        self._on_clear_callback = on_clear
        self._on_navigate = on_navigate
        self._machine_model = machine_model or ""
        self.setWindowTitle(f"Card Slot C {slot_id} - Details")
        self.setWindowFlags(Qt.Window)
        self.setAttribute(Qt.WA_DeleteOnClose)
        self.setMinimumWidth(720)
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

        QShortcut(QKeySequence("Right"), self).activated.connect(lambda: self._navigate(+1))
        QShortcut(QKeySequence("Left"), self).activated.connect(lambda: self._navigate(-1))

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

        # Pattern preview (reuse PatternPreview for now)
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
        self._show_canvas_cb = QCheckBox("Show canvas")
        self._show_canvas_cb.stateChanged.connect(self._on_show_canvas_changed)
        split_row.addWidget(self._show_canvas_cb)
        self._hide_points_cb = QCheckBox("Hide points")
        self._hide_points_cb.stateChanged.connect(self._on_hide_points_changed)
        split_row.addWidget(self._hide_points_cb)
        split_row.addStretch()
        layout.addLayout(split_row)

        # Tabbed widget
        tabs = QTabWidget()

        # Tab 1: Raw data (adds preview_image field)
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

        # New field for memory card slots: preview_image (raw representation)
        preview_image_label = QLabel("Preview image (raw)")
        preview_image_label.setFont(header_font)
        raw_layout.addWidget(preview_image_label)

        self._preview_image_edit = QTextEdit()
        self._preview_image_edit.setReadOnly(True)
        self._preview_image_edit.setFont(QFont("Courier New", 9))
        self._preview_image_edit.setFixedHeight(80)
        raw_layout.addWidget(self._preview_image_edit)

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
        self._points_table.setColumnCount(7)
        self._points_table.setHorizontalHeaderLabels(["#", "Dec (x, y)", "Dec diff(n, n-1)", "Dec (x, y, t)", "Dec (x, y, tacc)", "Hex (x,y)", "Hex diff(n, n-1)"])
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
        # find current index in the slots list
        try:
            current_idx = next(i for i, s in enumerate(self._slots) if s.slot_id == self.slot.slot_id)
        except StopIteration:
            return
        new_idx = current_idx + delta
        if not (0 <= new_idx < len(self._slots)):
            return
        if self._on_navigate and not self._on_navigate(current_idx, new_idx):
            return
        self.slot = self._slots[new_idx]
        self.setWindowTitle(f"Card Slot C {self.slot.slot_id} - Details")
        self._load_slot()

    def _update_nav_buttons(self):
        try:
            idx = next(i for i, s in enumerate(self._slots) if s.slot_id == self.slot.slot_id)
        except StopIteration:
            idx = 0
        self._prev_btn.setEnabled(idx > 0)
        self._next_btn.setEnabled(idx < len(self._slots) - 1)

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
        """Placeholder: populate header grid. Keep behavior similar to P-Memory window for now."""
        try:
            return self._populate_header_grid_75xx()
        except Exception:
            # If slot doesn't provide expected stats, just clear the grid
            while self._header_grid.count():
                item = self._header_grid.takeAt(0)
                if item.widget():
                    item.widget().deleteLater()

    # Reuse the existing implementations from slot_detail_window for now
    def _populate_header_grid_75xx(self):
        # Copying behavior from slot_detail_window but safe-guarded by try/except
        while self._header_grid.count():
            item = self._header_grid.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        header_bytes = self._get_header_bytes()
        try:
            stats = self.slot.get_pattern_stats()
        except Exception:
            stats = {}

        if getattr(self.slot, 'pattern_type', None) in ("9mm", "MAXI"):
            mapping = {
                2: ("fix_0x10", "Fixed byte?"),               # DONE
                3: ("fix_0x02", "Fixed byte?"),               # DONE
                4: ("bank", "Bank number?"),                  # DONE
                5: ("slot_no", "Slot number"),                # DONE
                6: ("type", "Pattern type"),                  # DONE
                9: ("d0x_min_abs", "abs(min(dxs)-xs[0]))"),   # DONE
                11: ("pn_x", "xs[end]"),                      # DONE
                13: ("span_x", "max(xs) - min(xs)"),          # DONE
                15: ("y_min_to_bound", "0x36 - min(ys)"),     # DONE
                17: ("span_y", "max(ys) - min(ys)"),          # DONE
                19: (None, "Unknown"),                        # DONE (scaling)
                22: ("dx_abs_max", "max(abs(dxs))"),          # DONE
                24: ("size_preview", "size(preview_image)"),  # DONE
                26: (None, "Unknown"),                        # DONE (uknown)
                27: ("size_pattern", "size(pattern_raw)"),    # DONE
                29: ("size_name", "size(filename)"),          # DONE
            }
        else:
            mapping = {
                2: ("fix_0x10", "Fixed byte?"),               # DONE
                3: ("fix_0x02", "Fixed byte?"),               # DONE
                4: ("bank", "Bank number?"),                  # DONE
                5: ("slot_no", "Slot number"),                # DONE
                6: ("type", "Pattern type"),                  # DONE
                24: ("size_preview", "size(preview_image)"),  # DONE
                26: (None, "Unknown"),                        # DONE (uknown)
                27: ("size_pattern", "size(pattern_raw)"),    # DONE
                29: ("size_name", "size(filename)"),          # DONE
            }

        two_byte_pairs = {7: 8, 9: 10, 11: 12, 13: 14, 15: 16, 17: 18, 20: 21, 22: 23, 24: 25, 27:28}
        skip_indices = set(two_byte_pairs.values())

        mono = QFont("Courier New", 9)
        bold_font = QFont()
        bold_font.setBold(True)

        # Header row — col 0..6
        for col, text in enumerate(("Byte", "Hex", "Dec", "Stat hex", "Stat dec", "Stat name", "OK/NOK")):
            hdr = QLabel(text)
            hdr.setFont(bold_font)
            self._header_grid.addWidget(hdr, 0, col)

        def _add_row(grid_row, byte_label, h_val, combined, stat_key, stat_val_raw, is_two_byte):
            idx_lbl = QLabel(byte_label)
            idx_lbl.setFont(mono)
            self._header_grid.addWidget(idx_lbl, grid_row, 0)

            if combined is not None:
                hex_str = f"0x{combined & 0xFFFF:04X}" if is_two_byte else f"0x{combined:02X}"
            else:
                hex_str = "--"
            hex_lbl = QLabel(hex_str)
            hex_lbl.setFont(mono)
            self._header_grid.addWidget(hex_lbl, grid_row, 1)

            dec_str = str(combined) if combined is not None else "--"
            dec_lbl = QLabel(dec_str)
            dec_lbl.setFont(mono)
            self._header_grid.addWidget(dec_lbl, grid_row, 2)

            if stat_key is not None and stat_val_raw is not None:
                if is_two_byte:
                    stat_hex_str = f"0x{stat_val_raw & 0xFFFF:04X}"
                else:
                    stat_hex_str = f"0x{stat_val_raw & 0xFF:02X}"
                stat_dec_str = str(stat_val_raw)
            else:
                stat_hex_str = "--"
                stat_dec_str = "--"

            stat_hex_lbl = QLabel(stat_hex_str)
            stat_hex_lbl.setFont(mono)
            self._header_grid.addWidget(stat_hex_lbl, grid_row, 3)

            stat_dec_lbl = QLabel(stat_dec_str)
            stat_dec_lbl.setFont(mono)
            self._header_grid.addWidget(stat_dec_lbl, grid_row, 4)

            if h_val is not None:
                name_str = stat_key if stat_key is not None else ("unknown" if h_val in mapping else "")
            else:
                name_str = stat_key or ""
            name_lbl = QLabel(name_str or "")
            name_lbl.setFont(mono)
            self._header_grid.addWidget(name_lbl, grid_row, 5)

            status_lbl = QLabel()
            status_lbl.setFont(bold_font)
            if stat_key is not None and stat_val_raw is not None and combined is not None:
                expected = stat_val_raw if is_two_byte else (stat_val_raw & 0xFF)
                if combined == expected:
                    status_lbl.setText("OK")
                    status_lbl.setStyleSheet("color: green;")
                else:
                    status_lbl.setText("NOK")
                    status_lbl.setStyleSheet("color: red;")
            elif stat_key is None and combined is not None:
                if combined == 0:
                    status_lbl.setText("OK")
                    status_lbl.setStyleSheet("color: green;")
                else:
                    status_lbl.setText("NOK")
                    status_lbl.setStyleSheet("color: red;")
            else:
                status_lbl.setText("--")
            self._header_grid.addWidget(status_lbl, grid_row, 6)

        grid_row = 1
        for idx in range(max(16, len(header_bytes))):
            if idx in skip_indices:
                continue

            if idx in two_byte_pairs:
                idx2 = two_byte_pairs[idx]
                h_hi = header_bytes[idx]  if idx  < len(header_bytes) else None
                h_lo = header_bytes[idx2] if idx2 < len(header_bytes) else None
                combined = ((h_hi << 8) | h_lo) if (h_hi is not None and h_lo is not None) else None
                if combined is not None:
                    combined = combined - 0x10000 if combined >= 0x8000 else combined
                byte_label = f"H[{idx}-{idx2}]"

                stat_key, stat_label = mapping[idx] if idx in mapping else (None, "")
                stat_val_raw = stats.get(stat_key) if stat_key else None
                _add_row(grid_row, byte_label, idx, combined, stat_key, stat_val_raw, is_two_byte=True)
                if idx in mapping:
                    w = self._header_grid.itemAtPosition(grid_row, 0)
                    if w and w.widget():
                        w.widget().setToolTip(stat_label)
            else:
                h_byte = header_bytes[idx] if idx < len(header_bytes) else None
                byte_label = f"H[{idx}]"

                stat_key, stat_label = mapping[idx] if idx in mapping else (None, "")
                stat_val_raw = stats.get(stat_key) if stat_key else None
                _add_row(grid_row, byte_label, idx, h_byte, stat_key, stat_val_raw, is_two_byte=False)
                if idx in mapping:
                    w = self._header_grid.itemAtPosition(grid_row, 0)
                    if w and w.widget():
                        w.widget().setToolTip(stat_label)

            grid_row += 1

        self._header_grid.addItem(
            QSpacerItem(0, 0, QSizePolicy.Minimum, QSizePolicy.Expanding), grid_row, 0, 1, 7)

    def _populate_pattern_grid(self):
        """Fill the Pattern tab with all pattern statistics."""
        while self._pattern_grid.count():
            item = self._pattern_grid.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        try:
            s = self.slot.get_pattern_stats()
        except Exception:
            s = {}
        mono = QFont("Courier New", 9)
        bold_font = QFont()
        bold_font.setBold(True)

        rows = [
            ("n",                   s.get("n"),             None),
            ("x_min",               s.get("x_min"),         None),
            ("x_max",               s.get("x_max"),         None),
            ("y_min",               s.get("y_min"),         None),
            ("y_max",               s.get("y_max"),         None),
            ("y_min_norm",          s.get("y_min_norm"),    None),
            ("y_max_norm",          s.get("y_max_norm"),    None),
            ("y_max_norm_div_2",    s.get("y_max_norm_div_2"),   None),
            ("y_min_to_bound",      s.get("y_min_to_bound"), None),
            ("span_x",              s.get("span_x"),        None),
            ("span_y",              s.get("span_y"),        None),
            ("dx_max",              s.get("dx_max"),        None),
            ("dx_min",              s.get("dx_min"),        None),
            ("dx_min_abs",          s.get("dx_min_abs"),    None),
            ("dx_abs_max",          s.get("dx_abs_max"),    None),
            ("dy_max",              s.get("dy_max"),        None),
            ("dy_min",              s.get("dy_min"),        None),
            ("dy_min_abs",          s.get("dy_min_abs"),    None),
            ("dy_abs_max",          s.get("dy_abs_max"),    None),
            ("is_reversed",         s.get("is_reversed"),   None),
            ("dx_0n",               s.get("dx_0n"),         None),
            ("dx_0n_abs",           s.get("dx_0n_abs"),     None),
            ("dy_0n",               s.get("dy_0n"),         None),
            ("dy_0n_abs",           s.get("dy_0n_abs"),     None),
            ("d0x_max",             s.get("d0x_max"),       None),
            ("d0x_min",             s.get("d0x_min"),       None),
            ("d0x_min_abs",         s.get("d0x_min_abs"),   None),
            ("d0y_max",             s.get("d0y_max"),       None),
            ("d0y_min",             s.get("d0y_min"),       None),
            ("d0y_min_abs",         s.get("d0y_min_abs"),   None),
            ("p0_x",                s.get("p0_x"),          None),
            ("p0_y",                s.get("p0_y"),          None),
            ("p1_x",                s.get("p1_x"),          None),
            ("p1_y",                s.get("p1_y"),          None),
            ("p1_dx",               s.get("p1_dx"),         None),
            ("p1_dy",               s.get("p1_dy"),         None),
            ("p1_dx_abs",           s.get("p1_dx_abs"),     None),
            ("p1_dy_abs",           s.get("p1_dy_abs"),     None),
            ("pn_x",                s.get("pn_x"),          None),
            ("pn_y",                s.get("pn_y"),          None),
            ("pn_dx",               s.get("pn_dx"),         None),
            ("pn_dy",               s.get("pn_dy"),         None),
            ("pn_dx_abs",           s.get("pn_dx_abs"),     None),
            ("pn_dy_abs",           s.get("pn_dy_abs"),     None),
            ("dnx_max",             s.get("dnx_max"),       None),
            ("dnx_min",             s.get("dnx_min"),       None),
            ("dnx_min_abs",         s.get("dnx_min_abs"),   None),
            ("dny_max",             s.get("dny_max"),       None),
            ("dny_min",             s.get("dny_min"),       None),
            ("dny_min_abs",         s.get("dny_min_abs"),   None),
            ("checksum",            s.get("checksum"),      None),
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
        data = list(getattr(self.slot, 'pattern_xy', []))
        xs = data[0::2]
        ys = data[1::2]
        n = min(len(xs), len(ys))

        xyt = list(getattr(self.slot, 'pattern_xyt', []))
        xytacc = list(getattr(self.slot, 'pattern_xytacc', []))

        mono = QFont("Courier New", 9)

        if n == 0:
            self._points_table.setRowCount(1)
            it = QTableWidgetItem("--")
            it.setFont(mono)
            self._points_table.setItem(0, 0, it)
            for col in range(1, 7):
                self._points_table.setItem(0, col, QTableWidgetItem(""))
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

            base = i * 3
            if base + 2 < len(xyt):
                xyt_text = f"({xyt[base]}, {xyt[base+1]}, {xyt[base+2]})"
            else:
                xyt_text = "--"
            xyt_it = QTableWidgetItem(xyt_text)
            xyt_it.setFont(mono)

            base = i * 3
            if base + 2 < len(xytacc):
                xytacc_text = f"({xytacc[base]}, {xytacc[base+1]}, {xytacc[base+2]})"
            else:
                xytacc_text = "--"
            xytacc_it = QTableWidgetItem(xytacc_text)
            xytacc_it.setFont(mono)

            self._points_table.setItem(i, 0, idx_it)
            self._points_table.setItem(i, 1, dec_it)
            self._points_table.setItem(i, 2, diff_dec_it)
            self._points_table.setItem(i, 3, xyt_it)
            self._points_table.setItem(i, 4, xytacc_it)
            self._points_table.setItem(i, 5, hex_it)
            self._points_table.setItem(i, 6, diff_hex_it)

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
        try:
            self._preview.selected_point = idx
            self._preview.update()
        except Exception:
            pass

    def _on_point_table_selection_changed(self):
        sels = self._points_table.selectionModel().selectedRows()
        if not sels:
            self._preview.selected_point = None
            self._preview.update()
            return
        row = sels[0].row()
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
        self._slot_label.setText(f"Slot:  C {self.slot.slot_id}")
        self._type_label.setText(f"    Type:  {getattr(self.slot, 'pattern_type', '')}")
        self._bytes_label.setText(f"    Bytes:  {getattr(self.slot, 'get_size_bytes', lambda: 'N/A')()}" )
        self._stitches_label.setText(f"    Stitches:  {getattr(self.slot, 'get_size_stitches', lambda: 'N/A')()}" )
        self._preview.pattern_xy = list(getattr(self.slot, 'pattern_xy', []))
        self._preview.pattern_type = getattr(self.slot, 'pattern_type', "")
        self._preview.update()
        self._refresh_raw_display()
        self._populate_header_grid()
        self._populate_pattern_grid()
        self._populate_points_grid()
        self._clear_btn.setEnabled(getattr(self.slot, 'pattern_type', "Empty") != "Empty")
        self._update_nav_buttons()

    def _format_header_raw(self, raw: str) -> str:
        if not self._logical_split_cb.isChecked() or not raw:
            return raw
        return ' '.join(raw[i:i+2] for i in range(0, len(raw), 2))

    def _format_pattern_raw(self, raw: str, pattern_type: str) -> str:
        if not self._logical_split_cb.isChecked() or not raw:
            return raw
        lines = []
        i = 0
        if pattern_type == "9mm":
            group = 5
            while i < len(raw):
                a = raw[i:i+3]
                b = raw[i+3:i+5]
                line = a + (' ' + b if b else '')
                lines.append(line)
                i += group
        else:
            group = 7
            while i < len(raw):
                a = raw[i:i+3]
                b = raw[i+3:i+5]
                c = raw[i+5:i+7]
                line = a + (' ' + b if b else '') + (' ' + c if c else '')
                lines.append(line)
                i += group
        return '\n'.join(lines)

    def _on_show_canvas_changed(self):
        self._preview.show_canvas = self._show_canvas_cb.isChecked()
        self._preview.update()

    def _on_hide_points_changed(self):
        self._preview.show_points = not self._hide_points_cb.isChecked()
        self._preview.update()

    def _refresh_raw_display(self):
        formatted_header = self._format_header_raw(getattr(self.slot, 'header_raw', ''))
        self._header_edit.setPlainText(formatted_header)
        self._header_edit_2.setPlainText(formatted_header)
        self._pattern_edit.setPlainText(
            self._format_pattern_raw(getattr(self.slot, 'pattern_raw', ''), getattr(self.slot, 'pattern_type', '')))
        # preview_raw is the raw preview field for card slots
        self._preview_image_edit.setPlainText(getattr(self.slot, 'preview_raw', ''))

    def refresh(self):
        """Re-read from the slot and update all displayed fields."""
        self._load_slot()

    def _clear_slot(self):
        """Clear the slot, refresh display, and notify the main window."""
        try:
            self.slot.clear()
        except Exception:
            # If slot doesn't implement clear(), try to reset common fields
            for attr in ('header_raw', 'pattern_raw', 'pattern_xy', 'pattern_xyt', 'pattern_xytacc', 'preview_raw'):
                if hasattr(self.slot, attr):
                    try:
                        setattr(self.slot, attr, '' if isinstance(getattr(self.slot, attr), str) else [])
                    except Exception:
                        pass
        self.refresh()
        if self._on_clear_callback:
            self._on_clear_callback()
