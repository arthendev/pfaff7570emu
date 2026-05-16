"""
Card Memory tab widget
"""

from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
                             QLabel, QFrame, QScrollArea, QTabWidget)
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

    def __init__(self, slot: CardMemorySlot, on_click=None, is_embroidery: bool = False):
        super().__init__()
        self.slot = slot
        self._on_click = on_click
        self._is_embroidery = is_embroidery
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

        # Slot number label
        slot_label = QLabel(str(self.slot.slot_id))
        slot_label_font = QFont()
        slot_label_font.setBold(True)
        slot_label_font.setPointSize(10)
        slot_label.setFont(slot_label_font)
        slot_label.setAlignment(Qt.AlignCenter)
        slot_label.setFixedWidth(28)
        main_layout.addWidget(slot_label)

        right_layout = QVBoxLayout()
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(3)

        size_info = f"{self.slot.get_size_bytes()} bytes"
        if self.slot.get_size_stitches() > 0:
            size_info += f"  {self.slot.get_size_stitches()} stitches"
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

    def __init__(self, machine_state: MachineState):
        super().__init__()
        self.machine_state = machine_state
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout()

        self._tab_widget = QTabWidget()

        self._tab_9mm = CardSpaceTab(self.machine_state.card_9mm)
        self._tab_maxi = CardSpaceTab(self.machine_state.card_maxi)
        self._tab_embroidery = CardSpaceTab(self.machine_state.card_embroidery)

        self._tab_widget.addTab(self._tab_9mm, "9mm")
        self._tab_widget.addTab(self._tab_maxi, "MAXI")
        self._tab_widget.addTab(self._tab_embroidery, "Embroidery")

        layout.addWidget(self._tab_widget)
        self.setLayout(layout)

    def update_ui(self, machine_state: MachineState):
        """Update UI with new machine state"""
        self.machine_state = machine_state
        self._tab_9mm.update_space(machine_state.card_9mm)
        self._tab_maxi.update_space(machine_state.card_maxi)
        self._tab_embroidery.update_space(machine_state.card_embroidery)
