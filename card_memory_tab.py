"""
Card Memory tab widget
"""

from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
                             QLabel, QFrame, QScrollArea, QTabWidget,
                             QCheckBox, QSpinBox, QPushButton, QFileDialog,
                             QInputDialog, QMessageBox)
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QFont, QColor, QImage, QPixmap, QPainter, QTransform

from machine_state import MachineState, CardMemorySpace, CardMemorySlot
from pmemory_tab import PatternPreview


class CardPreviewWidget(QWidget):
    """Renders the 1bpp preview bitmap stored in CardMemorySlot.preview_raw."""

    _TARGET_HEIGHT = 48  # display height in screen pixels

    def __init__(self, preview_raw: str, pattern_type: str, is_embroidery: bool = False):
        super().__init__()
        self._pixmap = self._build_pixmap(preview_raw, pattern_type, is_embroidery)
        if self._pixmap and not self._pixmap.isNull():
            orig_h = self._pixmap.height()
            scale = max(1, self._TARGET_HEIGHT // orig_h) if orig_h > 0 else 1
            self._display_pixmap = self._pixmap.scaled(
                self._pixmap.width() * scale,
                orig_h * scale,
                Qt.KeepAspectRatio,
                Qt.FastTransformation,
            )
        else:
            self._display_pixmap = None
        if self._display_pixmap:
            self.setFixedSize(self._display_pixmap.width(), self._display_pixmap.height())
        else:
            self.setFixedSize(50, self._TARGET_HEIGHT)
        self.setStyleSheet("border: 1px solid #ccc; background-color: white;")

    def _build_pixmap(self, preview_raw: str, pattern_type: str, is_embroidery: bool):
        if not preview_raw:
            return None
        try:
            data = bytes.fromhex(preview_raw)
        except ValueError:
            return None

        col_height = 24 if pattern_type == "9mm" else 48
        bytes_per_col = col_height // 8  # 3 for 9mm, 6 for MAXI/Embroidery

        if len(data) < bytes_per_col:
            return None

        num_cols = len(data) // bytes_per_col
        img = QImage(num_cols, col_height, QImage.Format_RGB32)
        img.fill(QColor(255, 255, 255).rgb())
        black = QColor(0, 0, 0).rgb()

        for col in range(num_cols):
            for byte_idx in range(bytes_per_col):
                byte_val = data[col * bytes_per_col + byte_idx]
                # byte_idx=0 is the bottom-most 8-pixel group; last byte_idx is the top group
                y_base = col_height - 8 - byte_idx * 8
                for bit in range(8):
                    # MSB (bit 7) is the topmost pixel of this 8-pixel segment
                    if (byte_val >> (7 - bit)) & 1:
                        img.setPixel(col, y_base + bit, black)

        if is_embroidery:
            img = img.transformed(QTransform().rotate(180))

        return QPixmap.fromImage(img)

    def paintEvent(self, event):
        if not self._display_pixmap:
            return
        painter = QPainter(self)
        painter.drawPixmap(0, 0, self._display_pixmap)
        painter.end()


class CardSlotWidget(QFrame):
    """Widget representing a single card memory slot"""

    def __init__(self, slot: CardMemorySlot, on_click=None, is_embroidery: bool = False, display_index: int = 0):
        super().__init__()
        self.slot = slot
        self._on_click = on_click
        self._is_embroidery = is_embroidery
        self._display_index = display_index
        self.setup_ui()
        self.setStyleSheet("border: 1px solid #ddd; padding: 5px;")
        if on_click:
            self.setCursor(Qt.PointingHandCursor)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton and self._on_click:
            self._on_click(self.slot)
        super().mousePressEvent(event)

    def setup_ui(self):
        main_layout = QHBoxLayout()
        main_layout.setContentsMargins(5, 5, 5, 5)
        main_layout.setSpacing(5)

        # Slot number on the left (dynamic: recalculated when slots are added/removed)
        slot_label = QLabel(f"{self._display_index}")
        slot_label_font = QFont()
        slot_label_font.setBold(True)
        slot_label_font.setPointSize(10)
        slot_label.setFont(slot_label_font)
        slot_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(slot_label)

        right_layout = QVBoxLayout()
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(3)

        size_info = f"{self.slot.get_size_bytes()} B"
        if self.slot.get_size_stitches() > 0:
            size_info += f"  {self.slot.get_size_stitches()} St."
        info_label = QLabel(f"{self.slot.pattern_type}   {size_info}")
        info_font = QFont()
        info_font.setPointSize(8)
        info_label.setFont(info_font)
        right_layout.addWidget(info_label)

        if self.slot.preview_raw:
            preview = CardPreviewWidget(
                self.slot.preview_raw, self.slot.pattern_type, self._is_embroidery
            )
        else:
            preview = PatternPreview(self.slot.pattern_xy, self.slot.pattern_type)
        right_layout.addWidget(preview)

        main_layout.addLayout(right_layout)
        self.setLayout(main_layout)


class CardSpaceTab(QWidget):
    """Tab widget showing all occupied slots in one card memory space"""

    slot_clicked = pyqtSignal(object)

    def __init__(self, space: CardMemorySpace):
        super().__init__()
        self.space = space
        self._slot_widgets = []
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout()
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        self._scroll_widget = QWidget()
        self._grid_layout = QGridLayout(self._scroll_widget)
        self._grid_layout.setSpacing(5)
        scroll.setWidget(self._scroll_widget)
        layout.addWidget(scroll)
        self.setLayout(layout)
        self._populate()

    def _populate(self):
        self._slot_widgets.clear()
        while self._grid_layout.count():
            item = self._grid_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        slots = self.space.sorted_slots()

        if not slots:
            empty_label = QLabel("No patterns stored")
            empty_label.setAlignment(Qt.AlignCenter)
            font = QFont()
            font.setItalic(True)
            empty_label.setFont(font)
            self._grid_layout.addWidget(empty_label, 0, 0)
            self._grid_layout.setRowStretch(1, 1)
            return

        columns = 8
        for i, slot in enumerate(slots):
            widget = CardSlotWidget(
                slot,
                on_click=self.slot_clicked.emit,
                is_embroidery=self.space.space_name == "Embroidery",
                display_index=i,
            )
            self._slot_widgets.append(widget)
            self._grid_layout.addWidget(widget, i // columns, i % columns,
                                        Qt.AlignLeft | Qt.AlignTop)

        self._grid_layout.setColumnStretch(columns, 1)
        self._grid_layout.setRowStretch(len(slots) // columns + 1, 1)

    def update_space(self, space: CardMemorySpace):
        self.space = space
        self._populate()


class CardMemoryTab(QWidget):
    """Card Memory tab split into 9mm, MAXI and Embroidery spaces"""

    # Signals emitted to main window
    card_inserted = pyqtSignal(str)     # path to card file
    card_ejected = pyqtSignal()
    card_created = pyqtSignal(str)      # path to new card file
    card_modified = pyqtSignal()        # emitted when card data changed (for auto-save)
    auto_save_changed = pyqtSignal(bool)  # emitted when auto-save checkbox toggles
    card_state_changed = pyqtSignal(bool)  # emitted on any card insert/eject/load; True if card is inserted

    def __init__(self, machine_state: MachineState):
        super().__init__()
        self.machine_state = machine_state
        self._auto_save = False
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout()

        # ---- Top controls ----
        controls_layout = QHBoxLayout()
        controls_layout.setContentsMargins(0, 0, 0, 0)
        controls_layout.setSpacing(8)

        # Insert Card button
        self._insert_card_btn = QPushButton("Insert Card")
        self._insert_card_btn.setToolTip("Choose a memory card JSON file to insert")
        self._insert_card_btn.clicked.connect(self._on_insert_card)
        controls_layout.addWidget(self._insert_card_btn)

        # Eject Card button
        self._eject_card_btn = QPushButton("Eject Card")
        self._eject_card_btn.setToolTip("Remove the current memory card")
        self._eject_card_btn.setEnabled(False)
        self._eject_card_btn.clicked.connect(self._on_eject_card)
        controls_layout.addWidget(self._eject_card_btn)

        # Create New Card button
        self._create_card_btn = QPushButton("Create New Card")
        self._create_card_btn.setToolTip("Create a new empty memory card file")
        self._create_card_btn.clicked.connect(self._on_create_card)
        controls_layout.addWidget(self._create_card_btn)

        # Separator
        controls_layout.addSpacing(12)

        # Card number label
        self._card_number_label = QLabel("Card: None")
        font = QFont()
        font.setBold(True)
        self._card_number_label.setFont(font)
        controls_layout.addWidget(self._card_number_label)

        # Spacer to push remaining controls to the right
        controls_layout.addStretch(1)

        # Save Card state button
        self._save_card_btn = QPushButton("Save Card state")
        self._save_card_btn.setToolTip("Save current card state to file")
        self._save_card_btn.setEnabled(False)
        self._save_card_btn.clicked.connect(self._on_save_card)
        controls_layout.addWidget(self._save_card_btn)

        # Save automatically checkbox
        self._auto_save_checkbox = QCheckBox("Save automatically")
        self._auto_save_checkbox.setToolTip("Automatically save card file after each pattern store/delete operation")
        self._auto_save_checkbox.toggled.connect(self._on_auto_save_toggled)
        controls_layout.addWidget(self._auto_save_checkbox)

        layout.addLayout(controls_layout)

        # ---- Card tabs ----
        self._tab_widget = QTabWidget()

        self._tab_9mm = CardSpaceTab(self.machine_state.card_9mm)
        self._tab_maxi = CardSpaceTab(self.machine_state.card_maxi)
        self._tab_embroidery = CardSpaceTab(self.machine_state.card_embroidery)

        self._tab_widget.addTab(self._tab_9mm, "9mm")
        self._tab_widget.addTab(self._tab_maxi, "MAXI")
        self._tab_widget.addTab(self._tab_embroidery, "Embroidery")

        layout.addWidget(self._tab_widget)
        self.setLayout(layout)

        # Initial UI state
        self._update_card_ui()

    def update_ui(self, machine_state: MachineState):
        """Update UI with new machine state"""
        self.machine_state = machine_state
        self._tab_9mm.update_space(machine_state.card_9mm)
        self._tab_maxi.update_space(machine_state.card_maxi)
        self._tab_embroidery.update_space(machine_state.card_embroidery)
        self._update_card_ui()

    def _update_card_ui(self):
        """Update button states and card number label based on current machine state."""
        ms = self.machine_state
        if ms is None:
            return

        card_inserted = ms.card_inserted
        self._eject_card_btn.setEnabled(card_inserted)
        self._save_card_btn.setEnabled(card_inserted)

        if card_inserted and ms.card_file_path:
            self._card_number_label.setText(f"Card: #{ms.card_number}")
        else:
            self._card_number_label.setText("Card: None")

        # Notify main window so menu actions stay in sync
        self.card_state_changed.emit(card_inserted)

    # ------------------------------------------------------------------
    # Button handlers
    # ------------------------------------------------------------------

    def _maybe_save_card(self) -> bool:
        """If the current card has unsaved changes, ask the user what to do.
        
        Returns:
            True  — proceed (saved or discarded)
            False — cancelled (user wants to stay)
        """
        ms = self.machine_state
        if not ms or not ms.card_inserted or not ms.card_modified:
            return True

        reply = QMessageBox.question(
            self,
            "Unsaved Card Changes",
            f"Card #{ms.card_number} has unsaved changes.\nDo you want to save them?",
            QMessageBox.Save | QMessageBox.Discard | QMessageBox.Cancel,
            QMessageBox.Save,
        )
        if reply == QMessageBox.Save:
            try:
                ms.save_card_file()
                return True
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to save card:\n{str(e)}")
                return False
        elif reply == QMessageBox.Discard:
            return True
        else:  # Cancel
            return False

    def _on_insert_card(self):
        """Show file dialog to choose a memory card JSON file."""
        # If a card is already inserted, check for unsaved changes first
        if self.machine_state.card_inserted:
            if not self._maybe_save_card():
                return

        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Insert Memory Card",
            "",
            "Memory Card JSON (*.json);;All files (*.*)"
        )
        if not file_path:
            return

        try:
            self.machine_state.load_card_file(file_path)
            self.update_ui(self.machine_state)
            self.card_inserted.emit(file_path)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load card file:\n{str(e)}")

    def _on_eject_card(self):
        """Eject the current card."""
        if not self._maybe_save_card():
            return
        self.machine_state.eject_card()
        self.update_ui(self.machine_state)
        self.card_ejected.emit()

    def _on_create_card(self):
        """Create a new empty memory card file."""
        # If a card is already inserted, check for unsaved changes first
        if self.machine_state.card_inserted:
            if not self._maybe_save_card():
                return

        # Ask for card number
        card_num, ok = QInputDialog.getInt(
            self,
            "Create New Card",
            "Card number:",
            value=1, min=0, max=255
        )
        if not ok:
            return

        # Ask for save location
        import os, sys
        from pathlib import Path
        base_dir = Path(sys.executable).parent if getattr(sys, 'frozen', False) else Path(__file__).parent
        default_dir = str(base_dir / "memory_cards")
        os.makedirs(default_dir, exist_ok=True)

        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Create New Memory Card",
            str(Path(default_dir) / f"card_{card_num}.json"),
            "Memory Card JSON (*.json);;All files (*.*)"
        )
        if not file_path:
            return

        try:
            # Create empty card data
            card_data = {
                "card_number": card_num,
                "patterns": {
                    "9mm": [],
                    "MAXI": [],
                    "Embroidery": [],
                },
            }
            import json
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(card_data, f, indent=2)

            # Load the new card
            self.machine_state.load_card_file(file_path)
            self.update_ui(self.machine_state)
            self.card_created.emit(file_path)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to create card file:\n{str(e)}")

    def _on_save_card(self):
        """Save current card state to its file."""
        if not self.machine_state.card_file_path:
            return
        try:
            self.machine_state.save_card_file()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save card file:\n{str(e)}")

    def _on_auto_save_toggled(self, checked: bool):
        """Toggle auto-save on card modifications."""
        self._auto_save = bool(checked)
        self.auto_save_changed.emit(bool(checked))

    @property
    def auto_save_enabled(self) -> bool:
        return self._auto_save

    # ------------------------------------------------------------------
    # Public API for menu actions
    # ------------------------------------------------------------------

    def insert_card(self):
        """Public wrapper for Insert Card action (usable from menu)."""
        self._on_insert_card()

    def eject_card(self):
        """Public wrapper for Eject Card action (usable from menu)."""
        self._on_eject_card()

    def save_card(self):
        """Public wrapper for Save Card action (usable from menu)."""
        self._on_save_card()

    def set_auto_save(self, enabled: bool):
        """Set auto-save state and update checkbox."""
        self._auto_save_checkbox.setChecked(bool(enabled))
