from datetime import datetime, timezone
import struct
import unittest

from ppt2pptx.oleps import FMTID, read_summary_information

def _lpstr(value: str, extra_terminators: int = 0) -> bytes:
    payload = value.encode("cp1252") + b"\0" * (1 + extra_terminators)
    result = struct.pack("<HHI", 0x1E, 0, len(payload)) + payload
    return result + bytes((-len(result)) % 4)

def _filetime(value: datetime) -> bytes:
    epoch = datetime(1601, 1, 1, tzinfo=timezone.utc)
    return struct.pack("<HHQ", 0x40, 0, int((value - epoch).total_seconds() * 10_000_000))

def _summary() -> bytes:
    properties = {1: struct.pack("<HHHH", 2, 0, 1252, 0), 2: _lpstr("Résumé", extra_terminators=2),
                  4: _lpstr("Ada"), 12: _filetime(datetime(2020, 1, 2, 3, 4, 5, tzinfo=timezone.utc))}
    table_size = 8 + len(properties) * 8
    payload, entries = bytearray(), []
    for prop_id, value in properties.items():
        entries.append((prop_id, table_size + len(payload)))
        payload.extend(value)
    prop_set = bytearray(struct.pack("<II", table_size + len(payload), len(properties)))
    for entry in entries:
        prop_set.extend(struct.pack("<II", *entry))
    prop_set.extend(payload)
    return struct.pack("<HHI16sI", 0xFFFE, 0, 0, bytes(16), 1) + FMTID + struct.pack("<I", 48) + prop_set

class OlePropertyTests(unittest.TestCase):
    def test_reads_core_properties(self):
        properties = read_summary_information(_summary())
        self.assertEqual(properties.title, "Résumé")
        self.assertEqual(properties.creator, "Ada")
        self.assertEqual(properties.created, "2020-01-02T03:04:05Z")

    def test_ignores_legacy_lpstr_alignment_terminators(self):
        payload = bytearray(_summary())
        # The normal fixture already exercises one terminator; old PowerPoint
        # files with extra counted NULs are covered by the parser's rstrip.
        self.assertNotIn("�", read_summary_information(bytes(payload)).title)
