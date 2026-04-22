"""
Machine state model and data persistence
"""

import json
from pathlib import Path
from typing import List, Dict, Any
from dataclasses import dataclass, asdict, field


@dataclass
class MemorySlot:
    """Represents a single memory slot"""
    slot_id: int
    slot_type: str = "Empty"  # Empty, 9mm, MAXI
    data: List[int] = field(default_factory=list)
    
    def get_size_bytes(self) -> int:
        """Get size of data in bytes"""
        return len(self.data)
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary for JSON serialization"""
        return {
            "slot_id": self.slot_id,
            "slot_type": self.slot_type,
            "data": self.data
        }
    
    @staticmethod
    def from_dict(data: Dict[str, Any]) -> 'MemorySlot':
        """Create MemorySlot from dictionary"""
        return MemorySlot(
            slot_id=data.get("slot_id", 0),
            slot_type=data.get("slot_type", "Empty"),
            data=data.get("data", [])
        )


class MachineState:
    """Manages the state of the sewing machine"""
    
    def __init__(self):
        self.p_memory_total_size = 40710  # total bytes available in P-Memory (sum of all slots)
        self.p_memory_slots: List[MemorySlot] = []
        self.m_memory: List[int] = []
        self.card_memory: List[int] = []
        
        # Initialize P-Memory with 30 empty slots
        for i in range(30):
            self.p_memory_slots.append(MemorySlot(slot_id=i, slot_type="Empty"))
    
    
    def get_p_memory_slot(self, slot_id: int) -> MemorySlot:
        """Get P-Memory slot by ID"""
        if 0 <= slot_id < len(self.p_memory_slots):
            return self.p_memory_slots[slot_id]
        raise IndexError(f"Invalid slot ID: {slot_id}")
    
    def set_p_memory_slot(self, slot: MemorySlot):
        """Set P-Memory slot"""
        if 0 <= slot.slot_id < len(self.p_memory_slots):
            self.p_memory_slots[slot.slot_id] = slot
        else:
            raise IndexError(f"Invalid slot ID: {slot.slot_id}")
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert machine state to dictionary for JSON serialization"""
        return {
            "p_memory_slots": [slot.to_dict() for slot in self.p_memory_slots],
            "m_memory": self.m_memory,
            "card_memory": self.card_memory
        }
    
    def from_dict(self, data: Dict[str, Any]):
        """Load machine state from dictionary"""
        if "p_memory_slots" in data:
            self.p_memory_slots = [
                MemorySlot.from_dict(slot) for slot in data["p_memory_slots"]
            ]
        
        if "m_memory" in data:
            self.m_memory = data["m_memory"]
        
        if "card_memory" in data:
            self.card_memory = data["card_memory"]
    
    def save_to_file(self, file_path: str):
        """Save machine state to JSON file"""
        file_path = Path(file_path)
        file_path.parent.mkdir(parents=True, exist_ok=True)
        
        with open(file_path, 'w') as f:
            json.dump(self.to_dict(), f, indent=2)
    
    def load_from_file(self, file_path: str):
        """Load machine state from JSON file"""
        file_path = Path(file_path)
        
        if not file_path.exists():
            raise FileNotFoundError(f"File not found: {file_path}")
        
        with open(file_path, 'r') as f:
            data = json.load(f)
        
        self.from_dict(data)
