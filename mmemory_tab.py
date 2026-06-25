"""
M-Memory tab widget
"""

from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
                             QLabel, QFrame, QScrollArea)
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QFont

from machine_state import MachineState, MMemorySlot
from pmemory_tab import PatternPreview


class MMemorySlotWidget(QFrame):
    """Widget representing a single M-Memory slot"""

    def __init__(self, slot: MMemorySlot, on_click=None):
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
        slot_label = QLabel(f"M {self.slot.slot_id}")
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

        # Size info (number of stitch patterns in sequence)
        info_label = QLabel(f"Patterns: {self.slot.get_size_patterns()}")
        info_font = QFont()
        info_font.setPointSize(8)
        info_label.setFont(info_font)
        right_layout.addWidget(info_label)

        # Pattern preview (first pattern in sequence for visual reference)
        preview = PatternPreview(self.slot.pattern_xy, "")
        right_layout.addWidget(preview)

        main_layout.addLayout(right_layout)

        self.setLayout(main_layout)


class MMemoryTab(QWidget):
    """M-Memory tab showing all 32 memory slots (M0..M31)"""

    slot_clicked = pyqtSignal(object)

    def __init__(self, machine_state: MachineState):
        super().__init__()
        self.machine_state = machine_state
        self.slot_widgets = []
        self.setup_ui()

    def setup_ui(self):
        """Setup M-Memory tab UI"""
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

        # Ensure 32 slots exist
        if not self.machine_state.m_memory_slots:
            self.machine_state.init_m_memory_slots()

        columns = 8
        num_slots = len(self.machine_state.m_memory_slots)
        for i in range(num_slots):
            slot = self.machine_state.get_m_memory_slot(i)
            slot_widget = MMemorySlotWidget(slot, on_click=self.slot_clicked.emit)
            self.slot_widgets.append(slot_widget)
            row = i // columns
            col = i % columns
            self.grid_layout.addWidget(slot_widget, row, col)

        self.grid_layout.setRowStretch(num_slots // columns + 1, 1)

    def update_ui(self, machine_state: MachineState):
        """Update UI with new machine state"""
        self.machine_state = machine_state
        self._populate_slots()
