"""
P-Memory tab widget
"""

from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, 
                             QLabel, QFrame, QScrollArea)
from PyQt5.QtCore import Qt, QSize, pyqtSignal
from PyQt5.QtGui import QFont, QColor, QPalette, QPixmap, QPainter, QBrush, QPen
from machine_state import MachineState


class PatternPreview(QFrame):
    """Widget to display pattern preview as a horizontal rectangle"""
    
    def __init__(self, slot_data, slot_type="", show_points=False):
        super().__init__()
        self.slot_data = slot_data
        self.slot_type = slot_type
        self.show_points = show_points
        self.selected_point = None
        self.setFixedHeight(45)
        self.setStyleSheet("border: 1px solid #ccc; background-color: white;")
        self.setMinimumWidth(100)
    
    def paintEvent(self, event):
        """Paint the pattern preview"""
        super().paintEvent(event)
        
        if not self.slot_data or len(self.slot_data) == 0:
            return
        
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        
        rect = self.rect()
        width = rect.width()
        height = rect.height()
        padding = 4

        if self.slot_type in ("9mm", "MAXI") and len(self.slot_data) >= 2:
            # Interpret flat list as interleaved x, y coordinates
            xs = self.slot_data[0::2]
            ys = self.slot_data[1::2]
            n = min(len(xs), len(ys))
            if n >= 2:
                x_min, x_max = min(xs), max(xs)
                y_min, y_max = min(ys), max(ys)
                x_range = x_max - x_min or 1
                y_range = y_max - y_min or 1

                draw_w = width - 2 * padding
                draw_h = height - 2 * padding
                scale = min(draw_w / x_range, draw_h / y_range)
                x_offset = padding + (draw_w - x_range * scale) / 2
                y_offset = padding + (draw_h - y_range * scale) / 2

                def to_screen(xi, yi):
                    sx = x_offset + (xi - x_min) * scale
                    sy = height - y_offset - (yi - y_min) * scale
                    return int(sx), int(sy)

                painter.setPen(QPen(QColor(0, 0, 180), 1))
                for i in range(n - 1):
                    x1, y1 = to_screen(xs[i], ys[i])
                    x2, y2 = to_screen(xs[i + 1], ys[i + 1])
                    painter.drawLine(x1, y1, x2, y2)

                # Draw intermediate points as small black dots
                if self.show_points:
                    painter.setPen(Qt.NoPen)
                    painter.setBrush(QBrush(QColor(0, 0, 0)))
                    for i in range(1, n - 1):
                        cx, cy = to_screen(xs[i], ys[i])
                        painter.drawEllipse(cx - 2, cy - 2, 4, 4)

                # Highlight selected point: small filled black rectangle (drawn before colored dots)
                if self.show_points and self.selected_point is not None:
                    sel = self.selected_point
                    if 0 <= sel < n:
                        sx, sy = to_screen(xs[sel], ys[sel])
                        painter.setPen(Qt.NoPen)
                        painter.setBrush(QBrush(QColor(0, 0, 0)))
                        size = 6
                        painter.drawRect(sx - size // 2, sy - size // 2, size, size)

                # First point — green (drawn on top)
                if self.show_points:
                    painter.setPen(Qt.NoPen)
                    painter.setBrush(QBrush(QColor(0, 180, 0)))
                    cx, cy = to_screen(xs[0], ys[0])
                    painter.drawEllipse(cx - 3, cy - 3, 6, 6)

                # Last point — red (drawn on top)
                if self.show_points:
                    painter.setBrush(QBrush(QColor(200, 0, 0)))
                    cx, cy = to_screen(xs[n - 1], ys[n - 1])
                    painter.drawEllipse(cx - 3, cy - 3, 6, 6)
        else:
            # Fallback: simple byte intensity bars
            num_bytes = len(self.slot_data)
            byte_width = max(1, width // max(num_bytes, 1))
            for i, byte_val in enumerate(self.slot_data[:width // byte_width]):
                x = i * byte_width
                gray_val = byte_val % 256
                color = QColor(gray_val, gray_val, gray_val)
                painter.fillRect(x, 0, byte_width, height, color)
        
        painter.end()


class SlotWidget(QFrame):
    """Widget representing a single memory slot"""
    
    def __init__(self, slot, on_click=None):
        super().__init__()
        self.slot = slot
        self._on_click = on_click
        self.setup_ui()
        self.setStyleSheet("border: 1px solid #ddd; padding: 5px;")
        if on_click:
            self.setCursor(Qt.PointingHandCursor)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton and self._on_click:
            self._on_click(self.slot)
        super().mousePressEvent(event)
    
    def setup_ui(self):
        """Setup slot display UI"""
        main_layout = QHBoxLayout()
        main_layout.setContentsMargins(5, 5, 5, 5)
        main_layout.setSpacing(5)
        
        # Slot number on the left
        slot_label = QLabel(f"P {self.slot.slot_id}")
        slot_label_font = QFont()
        slot_label_font.setBold(True)
        slot_label_font.setPointSize(10)
        slot_label.setFont(slot_label_font)
        slot_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(slot_label)
        
        # Info and preview on the right (vertical layout)
        right_layout = QVBoxLayout()
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(3)
        
        # Slot type and size
        info_label = QLabel(f"{self.slot.slot_type}   {self.slot.get_size_bytes()} bytes")
        info_font = QFont()
        info_font.setPointSize(8)
        info_label.setFont(info_font)
        right_layout.addWidget(info_label)
        
        # Pattern preview
        preview = PatternPreview(self.slot.data, self.slot.slot_type)
        right_layout.addWidget(preview)
        
        main_layout.addLayout(right_layout)
        
        self.setLayout(main_layout)


class PMemoryTab(QWidget):
    """P-Memory tab showing all 30 memory slots"""

    slot_clicked = pyqtSignal(object)
    
    def __init__(self, machine_state: MachineState):
        super().__init__()
        self.machine_state = machine_state
        self.slot_widgets = []
        self.setup_ui()
    
    def setup_ui(self):
        """Setup P-Memory tab UI"""
        layout = QVBoxLayout()
        
        # Create scroll area for all slots
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        self.scroll_widget = QWidget()
        self.grid_layout = QGridLayout(self.scroll_widget)
        self.grid_layout.setSpacing(5)
        
        scroll.setWidget(self.scroll_widget)
        layout.addWidget(scroll)
        self.setLayout(layout)
        
        self._populate_slots()
    
    def _populate_slots(self):
        """Clear and repopulate the grid with current machine state slots"""
        self.slot_widgets.clear()
        while self.grid_layout.count():
            item = self.grid_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        
        columns = 8
        for i in range(30):
            slot = self.machine_state.get_p_memory_slot(i)
            slot_widget = SlotWidget(slot, on_click=self.slot_clicked.emit)
            self.slot_widgets.append(slot_widget)
            row = i // columns
            col = i % columns
            self.grid_layout.addWidget(slot_widget, row, col)
        
        self.grid_layout.setRowStretch(30 // columns + 1, 1)
    
    def update_ui(self, machine_state: MachineState):
        """Update UI with new machine state"""
        self.machine_state = machine_state
        self._populate_slots()
