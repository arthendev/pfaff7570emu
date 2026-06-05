"""
Machine state model and data persistence
"""

import json
import logging
from pathlib import Path
from typing import List, Dict, Any, Optional
from dataclasses import dataclass, asdict, field

logger = logging.getLogger(__name__)

# ToDo: consider common get_pattern_stats() method that can be used by both MemorySlot and CardMemorySlot

@dataclass
class MemorySlot:
    """Represents a single memory slot"""
    slot_id: int
    pattern_type: str = "Empty"  # Empty, 9mm, MAXI
    header_raw: str = ""  # raw header as ASCII string (hex-encoded pairs from machine)
    pattern_raw: str = ""  # pattern as ASCII string (as received from machine, e.g. groups of 5 or 7 chars per stitch)
    pattern_bytes: List[int] = field(default_factory=list)
    pattern_xy: List[int] = field(default_factory=list)
    pattern_xyt: List[int] = field(default_factory=list)  # for MAXI: x,y,transport
    pattern_xytacc: List[int] = field(default_factory=list)  # for MAXI: x,y,transport_accumulated
    
    def clear(self):
        """Clear the slot data, resetting to Empty"""
        self.pattern_type = "Empty"
        self.header_raw = ""
        self.pattern_raw = ""
        self.pattern_bytes = []
        self.pattern_xy = []
        self.pattern_xyt = []
        self.pattern_xytacc = []
        
    def get_size_bytes(self) -> int:
        """Get size of data in bytes"""
        return len(self.pattern_bytes)
    
    def get_size_stitches(self) -> int:
        """Get number of stitches (points) in the pattern"""
        return len(self.pattern_xy) // 2  # x,y pairs

    def set_slot_data(self, pattern_type: str, header_raw: str, pattern_raw: str) -> None:
        """Set slot data and parse the pattern"""
        self.pattern_type = pattern_type
        self.header_raw = header_raw
        self.pattern_raw = pattern_raw
        self.parse_pattern_data()

        if self.pattern_type == "Empty" or len(self.pattern_xy) < 2:
            return
        
        xs = self.pattern_xy[0::2]
        ys = self.pattern_xy[1::2]
        pairs_str = " ".join(f"({xs[i]},{ys[i]})" for i in range(len(xs)))
        logger.debug(f"[{self.pattern_type} stitch: {pairs_str}]")
        stats = self.get_pattern_stats()
        logger.debug(
            f"Pattern stats: n={stats['n']}, n_bytes={len(self.pattern_bytes)}, "
            f"x_min={stats['x_min']} x_max={stats['x_max']}, span_x={stats['span_x']}, "
            f"y_min={stats['y_min']} y_max={stats['y_max']}, span_y={stats['span_y']}, "
            f"p0_x={stats['p0_x']}, p0_y={stats['p0_y']}, pn_x={stats['pn_x']}, pn_y={stats['pn_y']}"
        )

    def parse_pattern_data(self) -> None:
        """Parse self.pattern_raw into self.pattern_bytes and self.pattern_xy based on self.pattern_type.

        9mm:  groups of 5 chars — 3-digit x + 2-digit y
        MAXI: groups of 7 chars — 3-digit x + 2-digit y + sign char + 1-digit side transport;
              side transport is accumulated (maxi_transport); effective x = raw_x + maxi_transport
        """
        pattern_bytes = []
        pattern_xy = []
        pattern_xyt = []
        pattern_xytacc = []
        
        if self.pattern_type == "9mm":
            raw = self.pattern_raw
            for i in range(0, len(raw), 5):
                group = raw[i:i+5]
                if len(group) < 5:
                    break
                try:
                    x = int(group[0:3])
                    y = int(group[3:5])
                except ValueError:
                    break
                pattern_bytes.append(x)
                pattern_bytes.append(y)
                pattern_xy.append(x)
                pattern_xy.append(y)
                pattern_xyt.append(x)
                pattern_xyt.append(y)
                pattern_xyt.append(0) # no transport for 9mm
                pattern_xytacc.append(x)
                pattern_xytacc.append(y)
                pattern_xytacc.append(0) # no transport for 9mm
            self.pattern_bytes = pattern_bytes
            self.pattern_xy = pattern_xy
            self.pattern_xyt = pattern_xyt
            self.pattern_xytacc = pattern_xytacc

        elif self.pattern_type == "MAXI":
            side_transport_acc = 0
            raw = self.pattern_raw
            for i in range(0, len(raw), 7):
                group = raw[i:i+7]
                if len(group) < 7:
                    break
                try:
                    x = int(group[0:3])
                    y = int(group[3:5])
                    side_transport = int(group[5:7])
                    side_transport_acc += side_transport
                except (ValueError, IndexError):
                    break
                pattern_bytes.append(x)
                pattern_bytes.append(y)
                pattern_bytes.append(side_transport)
                pattern_xy.append(x)
                pattern_xy.append(y + side_transport_acc)
                pattern_xyt.append(x)
                pattern_xyt.append(y)
                pattern_xyt.append(side_transport)
                pattern_xytacc.append(x)
                pattern_xytacc.append(y)
                pattern_xytacc.append(side_transport_acc)
            self.pattern_bytes = pattern_bytes
            self.pattern_xy = pattern_xy
            self.pattern_xyt = pattern_xyt
            self.pattern_xytacc = pattern_xytacc

        self.pattern_bytes = pattern_bytes
        self.pattern_xy = pattern_xy
        self.pattern_xyt = pattern_xyt
        self.pattern_xytacc = pattern_xytacc

    def get_pattern_stats(self) -> dict:
        """Compute statistics from the pattern data (x,y interleaved)."""
        stats = {
            "n": 0,
            "x_min": None, "x_max": None,
            "y_min": None, "y_max": None,
            "y_min_abs": None,
            "y_min_neg": None,
            "y_min_norm": None, "y_max_norm": None,
            "y_max_norm_div_2": None,
            "y_min_to_bound": None,  # 0x36 - y_min
            "y_min_symmetry": None,  # this one is tricky; used for memory card upload
            "span_x": None, "span_y": None,
            "dx_max": None, "dx_min": None,
            "dx_min_abs": None, "dx_abs_max": None,
            "dy_max": None, "dy_min": None,
            "dy_min_abs": None, "dy_abs_max": None,
            "is_reversed": False,
            "dx_0n": None, "dx_0n_abs": None,
            "dy_0n": None, "dy_0n_abs": None,
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
        if self.pattern_type == "Empty" or len(self.pattern_xy) < 2:
            return stats
        xs = self.pattern_xy[0::2]
        ys = self.pattern_xy[1::2]
        xs_reversed = list(reversed(xs));
        ys_reversed = list(reversed(ys));
        if not xs or not ys:
            return stats
        stats["n"] = min(len(xs), len(ys))
        stats["x_min"] = min(xs)
        stats["x_max"] = max(xs)
        stats["y_min"] = min(ys)
        stats["y_min_abs"] = abs(stats["y_min"])
        stats["y_min_neg"] = -stats["y_min"]
        stats["y_max"] = max(ys)
        stats["y_min_norm"] = 0
        stats["y_max_norm"] = stats["y_max"] - stats["y_min"]
        stats["y_max_norm_div_2"] = stats["y_max_norm"] // 2
        stats["y_min_to_bound"] = 0x36 - stats["y_min"]

        # y_min_symmetry
        # based on y_min but with additional flag (MSB set) if the pattern is "top-heavy" (y_max farther from 27 than y_min)
        # if pattern is "symmetrical" (i.e. y_max and y_min are equally distant from 27, or y_max=53 and y_min=0), then y_min_symmetry is 0 (with no extra flag)
        if self.pattern_type == "9mm":
                
            y_min_symmetry = stats["y_min"]
            
            # Check if top-heavy
            if stats["y_max"] - 27 > 27 - stats["y_min"]:
                y_min_symmetry |= 0x80  # set MSB to indicate top-heavy
            
            # If pattern is symmetrical, set y_min_symmetry to 0 (no extra flag)
            if stats["y_max"] >= 27 and stats["y_min"] <= 27:
                dymax_27 = stats["y_max"] - 27
                dymin_27 = 27 - stats["y_min"]
                is_max_width = stats["y_max"] == 53 and stats["y_min"] == 0

                if dymax_27 == dymin_27 or is_max_width:
                    y_min_symmetry = 0
        else:
            y_min_symmetry = None

        stats["y_min_symmetry"] = y_min_symmetry

        stats["span_x"] = stats["x_max"] - stats["x_min"]
        stats["span_y"] = stats["y_max"] - stats["y_min"]
        if len(xs) > 1:
            dxs = [xs[i + 1] - xs[i] for i in range(len(xs) - 1)]
            dys = [ys[i + 1] - ys[i] for i in range(len(ys) - 1)]
        else:
            dxs = [0]
            dys = [0]
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
        stats["dy_0n"] = ys[-1] - ys[0]
        stats["dy_0n_abs"] = abs(stats["dy_0n"])
        stats["d0x_max"] = stats["x_max"] - xs[0]
        stats["d0x_min"] = stats["x_min"] - xs[0]
        stats["d0x_min_abs"] = abs(stats["d0x_min"])
        stats["d0y_max"] = stats["y_max"] - ys[0]
        stats["d0y_min"] = stats["y_min"] - ys[0]
        stats["d0y_min_abs"] = abs(stats["d0y_min"])
        stats["p0_x"] = xs[0]
        stats["p0_y"] = ys[0]
        stats["p1_x"] = xs[1] if len(xs) > 1 else 0
        stats["p1_y"] = ys[1] if len(ys) > 1 else 0
        stats["p1_dx"] = xs[1] - xs[0] if len(xs) > 1 else 0
        stats["p1_dy"] = ys[1] - ys[0] if len(ys) > 1 else 0
        stats["p1_dx_abs"] = abs(stats["p1_dx"])
        stats["p1_dy_abs"] = abs(stats["p1_dy"])
        stats["pn_x"] = xs[-1]
        stats["pn_y"] = ys[-1]
        if len(xs) > 1:
            dxs_reversed = [xs_reversed[i + 1] - xs_reversed[i] for i in range(len(xs_reversed) - 1)]
            dys_reversed = [ys_reversed[i + 1] - ys_reversed[i] for i in range(len(ys_reversed) - 1)]
        else:
            dxs_reversed = [0]
            dys_reversed = [0]
        stats["pn_dx"] = xs[-2] - xs[-1] if len(xs) > 1 else 0
        stats["pn_dy"] = ys[-2] - ys[-1] if len(ys) > 1 else 0
        stats["pn_dx_abs"] = abs(stats["pn_dx"])
        stats["pn_dy_abs"] = abs(stats["pn_dy"])
        stats["dnx_max"] = stats["x_max"] - xs[-1]
        stats["dnx_min"] = stats["x_min"] - xs[-1]
        stats["dnx_min_abs"] = abs(stats["dnx_min"])
        stats["dny_max"] = stats["y_max"] - ys[-1]
        stats["dny_min"] = stats["y_min"] - ys[-1]
        stats["dny_min_abs"] = abs(stats["dny_min"])
        stats["checksum"] = sum(self.pattern_xy) % 256
        return stats
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary for JSON serialization"""
        return {
            "slot_id": self.slot_id,
            "pattern_type": self.pattern_type,
            "header_raw": self.header_raw,
            "pattern_raw": self.pattern_raw,
        }
    
    @staticmethod
    def from_dict(data: Dict[str, Any]) -> 'MemorySlot':
        """Create MemorySlot from dictionary"""
        slot = MemorySlot(
            slot_id=data.get("slot_id", 0),
            pattern_type=data.get("pattern_type", "Empty"),
            header_raw=data.get("header_raw", ""),
            pattern_raw=data.get("pattern_raw", ""),
        )
        slot.parse_pattern_data()
        return slot


@dataclass
class CardMemorySlot:
    """Represents a single slot on a memory card"""
    slot_id: int
    pattern_type: str = ""  # "9mm", "MAXI", "Small hoop", "Large hoop"
    header_raw: str = ""
    preview_raw: str = ""
    pattern_raw: str = ""
    filename: str = ""
    pattern_bytes: List[int] = field(default_factory=list)
    pattern_xy: List[int] = field(default_factory=list)
    pattern_xyt: List[int] = field(default_factory=list)
    pattern_xytacc: List[int] = field(default_factory=list)

    def clear(self) -> None:
        """Remove this slot from its parent CardMemorySpace (if set) and clear data.

        This is the single place that implements clearing/removal of a card slot's
        pattern. It assumes the slot was previously added to a CardMemorySpace via
        `CardMemorySpace.set_slot()` which sets a `_parent` attribute on the slot.
        """
        self.pattern_type = "Empty"
        self.header_raw = ""
        self.preview_raw = ""
        self.pattern_raw = ""
        self.filename = ""
        self.pattern_bytes = []
        self.pattern_xy = []
        self.pattern_xyt = []
        self.pattern_xytacc = []
        parent = self._parent
        parent.slots.remove(self)

    def get_size_bytes(self) -> int:
        """Get size of data in bytes"""
        return len(self.pattern_raw) // 2

    def get_size_stitches(self) -> int:
        """Get number of stitches in the pattern"""
        return len(self.pattern_xy) // 2

    def set_slot_data(self, pattern_type: str, header_raw: str, preview_raw: str, pattern_raw: str) -> None:
        """Set slot data and parse the pattern"""
        self.pattern_type = pattern_type
        self.header_raw = header_raw
        self.preview_raw = preview_raw
        self.pattern_raw = pattern_raw
        self.parse_pattern_data()

    def parse_pattern_data(self) -> None:
        """Parse pattern_raw into pattern_bytes and pattern_xy.

        9mm:  groups of 2 bytes — dx (with 0x5B offset) and y absolute
        MAXI: groups of 3 bytes — dy_acc transport (with 0xC6 offset), dx (with 0x5B offset), and y absolute;
        Embroidery (Small hoop / Large hoop): format unknown, no parsing yet.
        """
        pattern_bytes: List[int] = []
        pattern_xy: List[int] = []
        pattern_xyt: List[int] = []
        pattern_xytacc: List[int] = []

        if self.pattern_type == "9mm":
            # Card 9mm encoding: raw is hex bytes. First/last byte may be 0x80/0x8A markers.
            raw = self.pattern_raw.strip()
            # convert to list of byte ints
            bytes_list: List[int] = []
            for i in range(0, len(raw), 2):
                chunk = raw[i:i+2]
                if len(chunk) < 2:
                    break
                try:
                    bytes_list.append(int(chunk, 16))
                except ValueError:
                    break

            if not bytes_list:
                # nothing to decode
                self.pattern_bytes = []
                self.pattern_xy = []
                self.pattern_xyt = []
                self.pattern_xytacc = []
            else:
                specials = {0x80, 0x8A}
                start = 1 if bytes_list[0] in specials else 0
                end = len(bytes_list) - 1 if bytes_list[-1] in specials else len(bytes_list)
                body = bytes_list[start:end]
                # store raw pattern bytes as the body
                pattern_bytes = list(body)

                # decode pairs: (dx_encoded, y_abs)
                xs: List[int] = []
                ys: List[int] = []
                prev_x = 0
                pair_count = len(body) // 2
                for i in range(pair_count):
                    dx_enc = body[i*2]
                    y_abs = body[i*2 + 1]
                    dx = dx_enc - 0x5B
                    x = prev_x - dx
                    xs.append(x)
                    ys.append(y_abs)
                    prev_x = x

                # shift if negative
                if xs:
                    min_x = min(xs)
                    if min_x < 0:
                        shift = -min_x
                        xs = [x + shift for x in xs]

                # interleave
                pattern_xy = []
                for x, y in zip(xs, ys):
                    pattern_xy.extend([int(x), int(y)])
                # pattern_xyt and pattern_xytacc keep y transport as 0 for 9mm
                for x, y in zip(xs, ys):
                    pattern_xyt.extend([int(x), int(y), 0])
                    pattern_xytacc.extend([int(x), int(y), 0])

        elif self.pattern_type == "MAXI":
            # Card-encoded MAXI hex triplets: pattern_raw is hex bytes.
            # Each triplet is (b0, b1, b2): transport diff, dx encoded, y absolute.
            raw = (self.pattern_raw or "").strip()
            bytes_list: List[int] = []
            for i in range(0, len(raw), 2):
                chunk = raw[i:i+2]
                if len(chunk) < 2:
                    break
                try:
                    bytes_list.append(int(chunk, 16))
                except ValueError:
                    break

            if bytes_list:
                specials = {0x80, 0x8A}
                start = 1 if bytes_list[0] in specials else 0
                end = len(bytes_list) - 1 if bytes_list[-1] in specials else len(bytes_list)
                body = bytes_list[start:end]
                pattern_bytes = list(body)

                triplets = len(body) // 3
                xs: List[int] = []
                ys: List[int] = []
                prev_x = 0
                side_transport_acc = 0
                for i in range(triplets):
                    b0 = body[i*3]
                    b1 = body[i*3 + 1]
                    b2 = body[i*3 + 2]

                    side_transport = b0 - 0xC6
                    side_transport_acc += side_transport
                    dx = b1 - 0x5B
                    x = prev_x - dx
                    y = b2
                    xs.append(x)
                    ys.append(y + side_transport_acc)
                    # store immediate transport and accumulated transport
                    pattern_xyt.extend([int(x), int(y), int(side_transport)])
                    pattern_xytacc.extend([int(x), int(y), int(side_transport_acc)])
                    prev_x = x

                # Shift x-coordinates if there are negative values, to ensure all x are non-negative.
                if xs:
                    min_x = min(xs)
                    if min_x < 0:
                        shift = -min_x
                        xs = [x + shift for x in xs]
                        # shift stored x values in xyt/xytacc
                        for idx in range(0, len(pattern_xyt), 3):
                            pattern_xyt[idx] = int(pattern_xyt[idx]) + shift
                        for idx in range(0, len(pattern_xytacc), 3):
                            pattern_xytacc[idx] = int(pattern_xytacc[idx]) + shift

                # interleave into pattern_xy
                for x, y in zip(xs, ys):
                    pattern_xy.extend([int(x), int(y)])

        self.pattern_bytes = pattern_bytes
        self.pattern_xy = pattern_xy
        self.pattern_xyt = pattern_xyt
        self.pattern_xytacc = pattern_xytacc

    def get_pattern_stats(self) -> dict:
        """Compute statistics from the pattern data (x,y interleaved).

        Copied from MemorySlot.get_pattern_stats to provide the same analysis for
        CardMemorySlot instances.
        """
        stats = {
            "n": 0,
            "x_min": None, "x_max": None,
            "y_min": None, "y_max": None,
            "y_min_abs": None,
            "y_min_neg": None,
            "y_min_norm": None, "y_max_norm": None,
            "y_max_norm_div_2": None,
            "y_min_to_bound": None,  # 0x36 - y_min
            "y_min_symmetry": None,  # this one is tricky; used for memory card upload
            "span_x": None, "span_y": None,
            "dx_max": None, "dx_min": None,
            "dx_min_abs": None, "dx_abs_max": None,
            "dy_max": None, "dy_min": None,
            "dy_min_abs": None, "dy_abs_max": None,
            "is_reversed": False,
            "dx_0n": None, "dx_0n_abs": None,
            "dy_0n": None, "dy_0n_abs": None,
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
        if self.pattern_type == "Empty" or len(self.pattern_xy) < 2:
            return stats
        xs = self.pattern_xy[0::2]
        ys = self.pattern_xy[1::2]
        xs_reversed = list(reversed(xs));
        ys_reversed = list(reversed(ys));
        if not xs or not ys:
            return stats
        stats["n"] = min(len(xs), len(ys))
        stats["x_min"] = min(xs)
        stats["x_max"] = max(xs)
        stats["y_min"] = min(ys)
        stats["y_min_abs"] = abs(stats["y_min"])
        stats["y_min_neg"] = -stats["y_min"]
        stats["y_max"] = max(ys)
        stats["y_min_norm"] = 0
        stats["y_max_norm"] = stats["y_max"] - stats["y_min"]
        stats["y_max_norm_div_2"] = stats["y_max_norm"] // 2
        stats["y_min_to_bound"] = 0x36 - stats["y_min"]

        # y_min_symmetry
        # based on y_min but with additional flag (MSB set) if the pattern is "top-heavy" (y_max farther from 27 than y_min)
        # if pattern is "symmetrical" (i.e. y_max and y_min are equally distant from 27, or y_max=53 and y_min=0), then y_min_symmetry is 0 (with no extra flag)
        if self.pattern_type == "9mm":
                
            y_min_symmetry = stats["y_min"]
            
            # Check if top-heavy
            if stats["y_max"] - 27 > 27 - stats["y_min"]:
                y_min_symmetry |= 0x80  # set MSB to indicate top-heavy
            
            # If pattern is symmetrical, set y_min_symmetry to 0 (no extra flag)
            if stats["y_max"] >= 27 and stats["y_min"] <= 27:
                dymax_27 = stats["y_max"] - 27
                dymin_27 = 27 - stats["y_min"]
                is_max_width = stats["y_max"] == 53 and stats["y_min"] == 0

                if dymax_27 == dymin_27 or is_max_width:
                    y_min_symmetry = 0
        else:
            y_min_symmetry = None

        stats["y_min_symmetry"] = y_min_symmetry

        stats["span_x"] = stats["x_max"] - stats["x_min"]
        stats["span_y"] = stats["y_max"] - stats["y_min"]
        if len(xs) > 1:
            dxs = [xs[i + 1] - xs[i] for i in range(len(xs) - 1)]
            dys = [ys[i + 1] - ys[i] for i in range(len(ys) - 1)]
        else:
            dxs = [0]
            dys = [0]
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
        stats["dy_0n"] = ys[-1] - ys[0]
        stats["dy_0n_abs"] = abs(stats["dy_0n"])
        stats["d0x_max"] = stats["x_max"] - xs[0]
        stats["d0x_min"] = stats["x_min"] - xs[0]
        stats["d0x_min_abs"] = abs(stats["d0x_min"])
        stats["d0y_max"] = stats["y_max"] - ys[0]
        stats["d0y_min"] = stats["y_min"] - ys[0]
        stats["d0y_min_abs"] = abs(stats["d0y_min"])
        stats["p0_x"] = xs[0]
        stats["p0_y"] = ys[0]
        stats["p1_x"] = xs[1] if len(xs) > 1 else 0
        stats["p1_y"] = ys[1] if len(ys) > 1 else 0
        stats["p1_dx"] = xs[1] - xs[0] if len(xs) > 1 else 0
        stats["p1_dy"] = ys[1] - ys[0] if len(ys) > 1 else 0
        stats["p1_dx_abs"] = abs(stats["p1_dx"])
        stats["p1_dy_abs"] = abs(stats["p1_dy"])
        stats["pn_x"] = xs[-1]
        stats["pn_y"] = ys[-1]
        if len(xs) > 1:
            dxs_reversed = [xs_reversed[i + 1] - xs_reversed[i] for i in range(len(xs_reversed) - 1)]
            dys_reversed = [ys_reversed[i + 1] - ys_reversed[i] for i in range(len(ys_reversed) - 1)]
        else:
            dxs_reversed = [0]
            dys_reversed = [0]
        stats["pn_dx"] = xs[-2] - xs[-1] if len(xs) > 1 else 0
        stats["pn_dy"] = ys[-2] - ys[-1] if len(ys) > 1 else 0
        stats["pn_dx_abs"] = abs(stats["pn_dx"])
        stats["pn_dy_abs"] = abs(stats["pn_dy"])
        stats["dnx_max"] = stats["x_max"] - xs[-1]
        stats["dnx_min"] = stats["x_min"] - xs[-1]
        stats["dnx_min_abs"] = abs(stats["dnx_min"])
        stats["dny_max"] = stats["y_max"] - ys[-1]
        stats["dny_min"] = stats["y_min"] - ys[-1]
        stats["dny_min_abs"] = abs(stats["dny_min"])
        stats["checksum"] = sum(self.pattern_xy) % 256
        return stats

    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary for JSON serialization.
        
        Note: slot_id is not stored — slot positions are dynamic and derived
        from list order at load time.
        """
        return {
            "pattern_type": self.pattern_type,
            "header_raw": self.header_raw,
            "preview_raw": self.preview_raw,
            "pattern_raw": self.pattern_raw,
            "filename": self.filename,
        }

    @staticmethod
    def from_dict(data: Dict[str, Any]) -> 'CardMemorySlot':
        """Create CardMemorySlot from dictionary.
        
        slot_id is ignored from persisted data — it is assigned dynamically
        based on list position by the caller.
        """
        slot = CardMemorySlot(
            slot_id=0,
            pattern_type=data.get("pattern_type", ""),
            header_raw=data.get("header_raw", ""),
            preview_raw=data.get("preview_raw", ""),
            pattern_raw=data.get("pattern_raw", ""),
            filename=data.get("filename", ""),
        )
        slot.parse_pattern_data()
        return slot


class CardMemorySpace:
    """One addressable space on a memory card, holding only occupied slots."""

    def __init__(self, space_name: str):
        self.space_name = space_name
        # store as a simple list: dynamic positions, no persistent slot numbers
        self.slots: List[CardMemorySlot] = []

    def get_slot(self, position: int) -> Optional[CardMemorySlot]:
        """Return the slot at the given position (0-based index), or None if out of range."""
        if position < 0 or position >= len(self.slots):
            return None
        return self.slots[position]

    def set_slot(self, slot: CardMemorySlot) -> None:
        """Append or replace a slot. We prefer to replace by matching slot_id
        if present (for compatibility), otherwise append to the dynamic list."""
        # assign parent reference and append to the dynamic list
        slot._parent = self
        self.slots.append(slot)

    def delete_slot(self, position: int) -> None:
        """Delete the slot at the given position (0-based index). If out of range, do nothing."""
        if position < 0 or position >= len(self.slots):
            return
        self.slots.pop(position)

    def clear(self) -> None:
        self.slots.clear()

    def sorted_slots(self) -> List[CardMemorySlot]:
        """Return all occupied slots in stored order (dynamic list)."""
        return list(self.slots)

    def to_dict(self) -> List[Dict[str, Any]]:
        return [slot.to_dict() for slot in self.slots]

    def from_dict(self, data: List[Dict[str, Any]]) -> None:
        # load as a simple ordered list to preserve stored ordering
        self.slots.clear()
        for item in data:
            slot = CardMemorySlot.from_dict(item)
            slot._parent = self
            self.slots.append(slot)


class MachineState:
    """Manages the state of the sewing machine"""

    # Model definitions: name -> (p_memory_total_size, num_slots)
    MODELS = {
        "PFAFF Creative 7570":    (40710, 30), # From real machine
        "PFAFF Creative 7550":    (40710, 30), # From real machine
        "PFAFF Creative 1475 CD": (5000, 16),  # Arbitrary pick, need real ones
    }

    def __init__(self, model_name: str = None):
        self.p_memory_total_size = None
        self.p_memory_slots: List[MemorySlot] = []
        self.m_memory: List[int] = []
        self.card_9mm = CardMemorySpace("9mm")
        self.card_maxi = CardMemorySpace("MAXI")
        self.card_embroidery = CardMemorySpace("Embroidery")
        # Card presence and file path (controlled by UI)
        self.card_file_path: Optional[str] = None
        self.card_number: int = 1
        self._card_modified: bool = False
        self.machine_model = None
        
        if model_name is not None:
            self.configure_model(model_name)

    def configure_model(self, model_name: str):
        """Reconfigure machine parameters for the given model name."""
        if model_name not in self.MODELS:
            raise ValueError(f"Unknown model: {model_name}")
        
        current_model = self.machine_model
        if (current_model == "PFAFF Creative 1475 CD" and model_name != "PFAFF Creative 1475 CD") \
            or (current_model != "PFAFF Creative 1475 CD" and model_name == "PFAFF Creative 1475 CD"):
            logger.warning(f"Switching from {current_model} to {model_name} - resetting all P-Memory slots to Empty to avoid stale data issues.")
            self.p_memory_slots: List[MemorySlot] = []
        
        total_size, num_slots = self.MODELS[model_name]
        self.machine_model = model_name
        self.p_memory_total_size = total_size
        current = len(self.p_memory_slots)
        if num_slots == 0:
            for i in range(1, num_slots):
                self.p_memory_slots.append(MemorySlot(slot_id=i, pattern_type="Empty"))
        if num_slots > current:
            for i in range(current, num_slots):
                self.p_memory_slots.append(MemorySlot(slot_id=i, pattern_type="Empty"))
        elif num_slots < current:
            self.p_memory_slots = self.p_memory_slots[:num_slots]
    
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
    
    def clear_card_memory(self) -> None:
        """Clear all card memory spaces."""
        self.card_9mm.clear()
        self.card_maxi.clear()
        self.card_embroidery.clear()

    @property
    def card_inserted(self) -> bool:
        """Card is inserted when card_file_path is set and the file exists."""
        if not self.card_file_path:
            return False
        resolved = self._resolve_card_path(self.card_file_path)
        return Path(resolved).exists()

    @property
    def card_modified(self) -> bool:
        """True if the card has unsaved changes."""
        return self._card_modified

    @property
    def supports_card(self) -> bool:
        """True if the current machine model supports Memory Cards (only 7570)."""
        return self.machine_model == "PFAFF Creative 7570"

    def mark_card_modified(self):
        """Mark the card state as modified (unsaved changes)."""
        self._card_modified = True

    def load_card_file(self, file_path: str) -> None:
        """Load card patterns from a memory card JSON file.
        
        The card file format:
        {
            "card_number": <int>,
            "patterns": {
                "9mm": [...],
                "MAXI": [...],
                "Embroidery": [...]
            }
        }
        """
        file_path = str(Path(file_path).resolve())
        with open(file_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        self.card_number = data.get("card_number", 1)
        self.card_file_path = self._make_card_path_relative(file_path)
        self.clear_card_memory()
        
        patterns = data.get("patterns", {})
        if isinstance(patterns, dict):
            if "9mm" in patterns:
                self.card_9mm.from_dict(patterns["9mm"])
            if "MAXI" in patterns:
                self.card_maxi.from_dict(patterns["MAXI"])
            if "Embroidery" in patterns:
                self.card_embroidery.from_dict(patterns["Embroidery"])
        
        self._card_modified = False
        logger.info(f"Loaded card #{self.card_number} from {file_path}")

    def save_card_file(self, file_path: str = None) -> bool:
        """Save card patterns to a memory card JSON file.
        
        If file_path is None, uses the current card_file_path.
        Returns True on success.
        """
        if file_path is None:
            if not self.card_file_path:
                logger.warning("No card file path set, cannot save")
                return False
            file_path = self._resolve_card_path(self.card_file_path)
        
        file_path = Path(file_path)
        file_path.parent.mkdir(parents=True, exist_ok=True)
        
        data = {
            "card_number": self.card_number,
            "patterns": {
                "9mm": self.card_9mm.to_dict(),
                "MAXI": self.card_maxi.to_dict(),
                "Embroidery": self.card_embroidery.to_dict(),
            },
        }
        
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2)
        
        self.card_file_path = self._make_card_path_relative(str(file_path.resolve()))
        self._card_modified = False
        logger.info(f"Saved card #{self.card_number} to {file_path}")
        return True

    def eject_card(self) -> None:
        """Eject the current card: clear card data and reset card_file_path."""
        self.clear_card_memory()
        self.card_file_path = None
        self.card_number = 1
        self._card_modified = False
        logger.info("Card ejected")

    def _resolve_card_path(self, path: str) -> str:
        """Resolve a card file path: relative paths are resolved relative to the app directory."""
        import sys
        base_dir = Path(sys.executable).parent if getattr(sys, 'frozen', False) else Path(__file__).parent
        p = Path(path)
        if p.is_absolute():
            return str(p)
        return str((base_dir / p).resolve())

    def _make_card_path_relative(self, absolute_path: str) -> str:
        """Convert an absolute card path to relative if it's under the app directory."""
        import sys
        base_dir = Path(sys.executable).parent if getattr(sys, 'frozen', False) else Path(__file__).parent
        try:
            return str(Path(absolute_path).resolve().relative_to(base_dir.resolve()))
        except ValueError:
            return absolute_path

    def to_dict(self) -> Dict[str, Any]:
        """Convert machine state to dictionary for JSON serialization.
        
        Card patterns are NOT stored here — they live in a separate memory card JSON file.
        Only the path to the card file is stored.  The card number lives in the card file.
        """
        return {
            "machine_model": self.machine_model,
            "p_memory_total_size": self.p_memory_total_size,
            "p_memory_slots": [slot.to_dict() for slot in self.p_memory_slots],
            "m_memory": self.m_memory,
            "card_file_path": self.card_file_path,
        }
    
    def from_dict(self, data: Dict[str, Any]):
        """Load machine state from dictionary"""
        # Restore model first so slot count is correct before loading slots
        if "machine_model" in data:
            saved_model = data["machine_model"]
            if saved_model in self.MODELS:
                self.configure_model(saved_model)
            else:
                logger.warning(f"Unknown machine_model in saved state: {saved_model!r} - ignoring")

        if "p_memory_total_size" in data:
            self.p_memory_total_size = data["p_memory_total_size"]

        # Reset all slots to Empty first so stale data from a previous state
        # never bleeds into the newly loaded state.
        for slot in self.p_memory_slots:
            slot.clear()

        if "p_memory_slots" in data:
            for slot_data in data["p_memory_slots"]:
                loaded = MemorySlot.from_dict(slot_data)
                idx = loaded.slot_id
                if 0 <= idx < len(self.p_memory_slots):
                    self.p_memory_slots[idx] = loaded
                else:
                    self.p_memory_slots.append(loaded)

        if "m_memory" in data:
            self.m_memory = data["m_memory"]

        # card_number is no longer stored in machine state — it lives in the card file.
        # Ignore any legacy "card_number" key.

        # Clear any currently loaded card data
        self.clear_card_memory()

        # Load card memory from separate card file if path is present
        card_path = data.get("card_file_path")
        if card_path:
            resolved = self._resolve_card_path(card_path)
            if Path(resolved).exists():
                self.card_file_path = card_path
                self.load_card_file(resolved)
            else:
                logger.warning(f"Card file not found: {resolved}")
                self.card_file_path = None
        else:
            self.card_file_path = None
            # Old-format card_memory key is ignored — card patterns now live in separate files.
    
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
