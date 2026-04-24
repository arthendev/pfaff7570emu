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
    header_raw: str = ""  # raw header as ASCII string (hex-encoded pairs from machine)
    pattern_raw: str = ""  # pattern as ASCII string (3-digit x + 2-digit y per point)
    
    def get_size_bytes(self) -> int:
        """Get size of data in bytes"""
        return len(self.pattern_raw)

    def get_pattern_stats(self) -> dict:
        """Compute statistics from the pattern data (x,y interleaved)."""
        stats = {
            "n": 0,
            "x_min": None, "x_max": None,
            "y_min": None, "y_max": None,
            "y_min_to_bound": None,  # 0x36 - y_min
            "span_x": None, "span_y": None,
            "dx_max": None, "dx_min": None,
            "dx_min_abs": None, "dx_abs_max": None,
            "dy_max": None, "dy_min": None,
            "dy_min_abs": None, "dy_abs_max": None,
            "is_reversed": False,
            "dx_0n": None, "dx_0n_abs": None,
            "d0x_max": None, "d0x_min": None, "d0x_min_abs": None,
            "d0y_max": None, "d0y_min": None, "d0y_min_abs": None,
            "p0_x": None, "p0_y": None,
            "p1_x": None, "p1_y": None,
            "p1_dx": None, "p1_dy": None,
            "p1_dx_abs": None, "p1_dy_abs": None,
            "pn_x": None, "pn_y": None,
            "pn_dx": None, "pn_dy": None,
            "pn_dx_abs": None, "pn_dy_abs": None,
            "dnx_max": None, "dnx_min": None, "dnx_min_abs": None,
            "dny_max": None, "dny_min": None, "dny_min_abs": None,
            "checksum": None,
        }
        if self.slot_type == "Empty" or len(self.data) < 2:
            return stats
        xs = self.data[0::2]
        ys = self.data[1::2]
        xs_reversed = list(reversed(xs));
        ys_reversed = list(reversed(ys));
        if not xs or not ys:
            return stats
        stats["n"] = min(len(xs), len(ys))
        stats["x_min"] = min(xs)
        stats["x_max"] = max(xs)
        stats["y_min"] = min(ys)
        stats["y_max"] = max(ys)
        stats["y_min_to_bound"] = 0x36 - stats["y_min"] if stats["y_min"] is not None else None
        stats["span_x"] = stats["x_max"] - stats["x_min"]
        stats["span_y"] = stats["y_max"] - stats["y_min"]
        if len(xs) >= 2:
            dxs = [xs[i + 1] - xs[i] for i in range(len(xs) - 1)]
            dys = [ys[i + 1] - ys[i] for i in range(len(ys) - 1)]
            stats["dx_max"] = max(d for d in dxs)
            stats["dx_min"] = min(dxs)
            stats["dx_min_abs"] = abs(stats["dx_min"])
            stats["dx_abs_max"] = max(abs(d) for d in dxs)
            stats["dy_max"] = max(d for d in dys)
            stats["dy_min"] = min(dys)
            stats["dy_min_abs"] = abs(stats["dy_min"])
            stats["dy_abs_max"] = max(abs(d) for d in dys)
            stats["is_reversed"] = xs[-1] < xs[0]
            stats["dx_0n"] = xs[-1] - xs[0]
            stats["dx_0n_abs"] = abs(stats["dx_0n"])
            stats["d0x_max"] = stats["x_max"] - xs[0]
            stats["d0x_min"] = stats["x_min"] - xs[0]
            stats["d0x_min_abs"] = abs(stats["d0x_min"])
        stats["d0y_max"] = stats["y_max"] - ys[0]
        stats["d0y_min"] = stats["y_min"] - ys[0]
        stats["d0y_min_abs"] = abs(stats["d0y_min"])
        stats["p0_x"] = xs[0]
        stats["p0_y"] = ys[0]
        stats["p1_x"] = xs[1] if len(xs) > 1 else None
        stats["p1_y"] = ys[1] if len(ys) > 1 else None
        stats["p1_dx"] = xs[1] - xs[0] if len(xs) > 1 else None
        stats["p1_dy"] = ys[1] - ys[0] if len(ys) > 1 else None
        stats["p1_dx_abs"] = abs(stats["p1_dx"]) if stats["p1_dx"] is not None else None
        stats["p1_dy_abs"] = abs(stats["p1_dy"]) if stats["p1_dy"] is not None else None
        stats["pn_x"] = xs[-1]
        stats["pn_y"] = ys[-1]
        if len(xs) >= 2:
            dxs_reversed = [xs_reversed[i + 1] - xs_reversed[i] for i in range(len(xs_reversed) - 1)]
            dys_reversed = [ys_reversed[i + 1] - ys_reversed[i] for i in range(len(ys_reversed) - 1)]
            stats["pn_dx"] = xs[-2] - xs[-1]
            stats["pn_dy"] = ys[-2] - ys[-1]
            stats["pn_dx_abs"] = abs(stats["pn_dx"])
            stats["pn_dy_abs"] = abs(stats["pn_dy"])
            stats["dnx_max"] = stats["x_max"] - xs[-1]
            stats["dnx_min"] = stats["x_min"] - xs[-1]
            stats["dnx_min_abs"] = abs(stats["dnx_min"])
        stats["dny_max"] = stats["y_max"] - ys[-1]
        stats["dny_min"] = stats["y_min"] - ys[-1]
        stats["dny_min_abs"] = abs(stats["dny_min"])
        stats["checksum"] = sum(self.data) % 256
        return stats
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary for JSON serialization"""
        return {
            "slot_id": self.slot_id,
            "slot_type": self.slot_type,
            "data": self.data,
            "header_raw": self.header_raw,
            "pattern_raw": self.pattern_raw
        }
    
    @staticmethod
    def from_dict(data: Dict[str, Any]) -> 'MemorySlot':
        """Create MemorySlot from dictionary"""
        return MemorySlot(
            slot_id=data.get("slot_id", 0),
            slot_type=data.get("slot_type", "Empty"),
            data=data.get("data", []),
            header_raw=data.get("header_raw", ""),
            pattern_raw=data.get("pattern_raw", "")
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
