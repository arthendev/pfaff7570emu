#!/usr/bin/env python3
"""
Convert default_state.json data fields from decimal integers to hex strings.
Writes the result to default_state_hex.json.
"""

import json

INPUT_FILE = "default_state.json"
OUTPUT_FILE = "default_state_hex.json"

with open(INPUT_FILE, "r") as f:
    state = json.load(f)

for slot in state.get("p_memory_slots", []):
    if "data" in slot and isinstance(slot["data"], list):
        slot["data"] = [format(v, "02X") for v in slot["data"]]

with open(OUTPUT_FILE, "w") as f:
    json.dump(state, f, indent=2)

print(f"Written: {OUTPUT_FILE}")

# Quick verification
with open(OUTPUT_FILE, "r") as f:
    verify = json.load(f)
for slot in verify["p_memory_slots"]:
    if slot.get("data"):
        print(f"Slot {slot['slot_id']}: {len(slot['data'])} entries, sample: {slot['data'][:8]}")
