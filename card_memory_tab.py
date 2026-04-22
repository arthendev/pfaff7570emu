"""
Card Memory tab widget
"""

from PyQt5.QtWidgets import QWidget, QVBoxLayout, QLabel
from PyQt5.QtGui import QFont
from machine_state import MachineState


class CardMemoryTab(QWidget):
    """Card Memory tab (placeholder for now)"""
    
    def __init__(self, machine_state: MachineState):
        super().__init__()
        self.machine_state = machine_state
        self.setup_ui()
    
    def setup_ui(self):
        """Setup Card Memory tab UI"""
        layout = QVBoxLayout()
        
        label = QLabel("Card Memory")
        font = QFont()
        font.setBold(True)
        font.setPointSize(14)
        label.setFont(font)
        
        layout.addWidget(label)
        layout.addStretch()
        
        self.setLayout(layout)
    
    def update_ui(self, machine_state: MachineState):
        """Update UI with new machine state"""
        self.machine_state = machine_state
