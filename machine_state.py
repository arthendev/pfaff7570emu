"""
Machine state model and data persistence
"""

import json
import logging
from pathlib import Path
from typing import List, Dict, Any, Optional
from dataclasses import dataclass, asdict, field

logger = logging.getLogger(__name__)


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
            "y_min_norm": None, "y_max_norm": None,
            "y_max_norm_div_2": None,
            "y_min_to_bound": None,  # 0x36 - y_min
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
        stats["y_max"] = max(ys)
        stats["y_min_norm"] = 0
        stats["y_max_norm"] = stats["y_max"] - stats["y_min"]
        stats["y_max_norm_div_2"] = stats["y_max_norm"] // 2
        stats["y_min_to_bound"] = 0x36 - stats["y_min"]
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

    def get_size_bytes(self) -> int:
        """Get size of data in bytes"""
        return len(self.pattern_bytes)

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

        9mm:  groups of 5 chars — 3-digit x + 2-digit y
        MAXI: groups of 7 chars — 3-digit x + 2-digit y + sign char + 1-digit side transport
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
                dy_acc = 0
                for i in range(triplets):
                    b0 = body[i*3]
                    b1 = body[i*3 + 1]
                    b2 = body[i*3 + 2]

                    transport = b0 - 0xC6
                    dy_acc += transport
                    dx = b1 - 0x5B
                    x = prev_x - dx
                    y = b2 + dy_acc
                    xs.append(x)
                    ys.append(y)
                    # store immediate transport and accumulated transport
                    pattern_xyt.extend([int(x), int(y), int(transport)])
                    pattern_xytacc.extend([int(x), int(y), int(dy_acc)])
                    prev_x = x

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
            "y_min_norm": None, "y_max_norm": None,
            "y_max_norm_div_2": None,
            "y_min_to_bound": None,  # 0x36 - y_min
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
        if self.pattern_type == "" or len(self.pattern_xy) < 2:
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
        stats["y_min_abs"] = abs(min(ys))
        stats["y_max"] = max(ys)
        stats["y_min_norm"] = 0
        stats["y_max_norm"] = stats["y_max"] - stats["y_min"]
        stats["y_max_norm_div_2"] = stats["y_max_norm"] // 2
        stats["y_min_to_bound"] = 0x36 - stats["y_min"]
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
        stats["dx_min_abs"] = abs(stats["dx_min"]) if stats["dx_min"] is not None else None
        stats["dx_abs_max"] = max(abs(d) for d in dxs)
        stats["dy_max"] = max(d for d in dys)
        stats["dy_min"] = min(dys)
        stats["dy_min_abs"] = abs(stats["dy_min"]) if stats["dy_min"] is not None else None
        stats["dy_abs_max"] = max(abs(d) for d in dys)
        stats["is_reversed"] = xs[-1] < xs[0]
        stats["dx_0n"] = xs[-1] - xs[0]
        stats["dx_0n_abs"] = abs(stats["dx_0n"]) if stats["dx_0n"] is not None else None
        stats["dy_0n"] = ys[-1] - ys[0]
        stats["dy_0n_abs"] = abs(stats["dy_0n"]) if stats["dy_0n"] is not None else None
        stats["d0x_max"] = stats["x_max"] - xs[0]
        stats["d0x_min"] = stats["x_min"] - xs[0]
        stats["d0x_min_abs"] = abs(stats["d0x_min"]) if stats["d0x_min"] is not None else None
        stats["d0y_max"] = stats["y_max"] - ys[0]
        stats["d0y_min"] = stats["y_min"] - ys[0]
        stats["d0y_min_abs"] = abs(stats["d0y_min"]) if stats["d0y_min"] is not None else None
        stats["p0_x"] = xs[0]
        stats["p0_y"] = ys[0]
        stats["p1_x"] = xs[1] if len(xs) > 1 else 0
        stats["p1_y"] = ys[1] if len(ys) > 1 else 0
        stats["p1_dx"] = xs[1] - xs[0] if len(xs) > 1 else 0
        stats["p1_dy"] = ys[1] - ys[0] if len(ys) > 1 else 0
        stats["p1_dx_abs"] = abs(stats["p1_dx"]) if stats["p1_dx"] is not None else None
        stats["p1_dy_abs"] = abs(stats["p1_dy"]) if stats["p1_dy"] is not None else None
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
        stats["pn_dx_abs"] = abs(stats["pn_dx"]) if stats["pn_dx"] is not None else None
        stats["pn_dy_abs"] = abs(stats["pn_dy"]) if stats["pn_dy"] is not None else None
        stats["dnx_max"] = stats["x_max"] - xs[-1]
        stats["dnx_min"] = stats["x_min"] - xs[-1]
        stats["dnx_min_abs"] = abs(stats["dnx_min"]) if stats["dnx_min"] is not None else None
        stats["dny_max"] = stats["y_max"] - ys[-1]
        stats["dny_min"] = stats["y_min"] - ys[-1]
        stats["dny_min_abs"] = abs(stats["dny_min"]) if stats["dny_min"] is not None else None
        stats["checksum"] = sum(self.pattern_xy) % 256 if self.pattern_xy else None
        return stats

    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary for JSON serialization"""
        return {
            "slot_id": self.slot_id,
            "pattern_type": self.pattern_type,
            "header_raw": self.header_raw,
            "preview_raw": self.preview_raw,
            "pattern_raw": self.pattern_raw,
            "filename": self.filename,
        }

    @staticmethod
    def from_dict(data: Dict[str, Any]) -> 'CardMemorySlot':
        """Create CardMemorySlot from dictionary"""
        slot = CardMemorySlot(
            slot_id=data.get("slot_id", 0),
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
        self.slots: Dict[int, CardMemorySlot] = {}

    def get_slot(self, slot_id: int) -> Optional[CardMemorySlot]:
        return self.slots.get(slot_id)

    def set_slot(self, slot: CardMemorySlot) -> None:
        self.slots[slot.slot_id] = slot

    def delete_slot(self, slot_id: int) -> None:
        self.slots.pop(slot_id, None)

    def clear(self) -> None:
        self.slots.clear()

    def sorted_slots(self) -> List[CardMemorySlot]:
        """Return all occupied slots sorted by slot_id."""
        return sorted(self.slots.values(), key=lambda s: s.slot_id)

    def to_dict(self) -> List[Dict[str, Any]]:
        return [slot.to_dict() for slot in self.sorted_slots()]

    def from_dict(self, data: List[Dict[str, Any]]) -> None:
        self.slots.clear()
        for item in data:
            slot = CardMemorySlot.from_dict(item)
            self.slots[slot.slot_id] = slot


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

    def to_dict(self) -> Dict[str, Any]:
        """Convert machine state to dictionary for JSON serialization"""
        return {
            "machine_model": self.machine_model,
            "p_memory_total_size": self.p_memory_total_size,
            "p_memory_slots": [slot.to_dict() for slot in self.p_memory_slots],
            "m_memory": self.m_memory,
            "card_memory": {
                "9mm": self.card_9mm.to_dict(),
                "MAXI": self.card_maxi.to_dict(),
                "Embroidery": self.card_embroidery.to_dict(),
            },
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

        # Load card memory — new format is a dict with space keys; old format was a plain
        # list (ignored for backwards compatibility).
        self.clear_card_memory()
        card_data = data.get("card_memory")
        if isinstance(card_data, dict):
            if "9mm" in card_data:
                self.card_9mm.from_dict(card_data["9mm"])
            if "MAXI" in card_data:
                self.card_maxi.from_dict(card_data["MAXI"])
            if "Embroidery" in card_data:
                self.card_embroidery.from_dict(card_data["Embroidery"])
    
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
