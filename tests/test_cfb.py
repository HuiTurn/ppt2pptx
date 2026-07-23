import struct
import unittest

from ppt2pptx.cfb import CompoundFile, Limits
from ppt2pptx.errors import InvalidPpt


def _minimal_cfb(*, stream_name: str, stream_data: bytes, size_high: int = 0) -> bytes:
    sector = 512
    stream_sectors = max(1, (len(stream_data) + sector - 1) // sector)
    # sector 0: FAT, sector 1: directory, sectors 2..: stream
    fat_entries = [0xFFFFFFFE, 0xFFFFFFFE]  # FAT sector self-ref unused; dir ends
    # Actually FAT[0]=END for FAT itself is wrong - sector 0 is FAT marked FATSECT
    fat_values = [0xFFFFFFFD, 0xFFFFFFFE]  # sector0=FAT, sector1=dir END
    for index in range(stream_sectors):
        next_sector = 2 + index + 1 if index + 1 < stream_sectors else 0xFFFFFFFE
        fat_values.append(next_sector)
    fat = b"".join(struct.pack("<I", value) for value in fat_values)
    fat = fat.ljust(sector, b"\xff")

    def entry(name: str, kind: int, start: int, size: int, size_hi: int = 0) -> bytes:
        encoded = name.encode("utf-16le") + b"\x00\x00"
        raw = bytearray(128)
        raw[: len(encoded)] = encoded
        struct.pack_into("<H", raw, 64, len(encoded))
        raw[66] = kind
        struct.pack_into("<I", raw, 116, start)
        struct.pack_into("<I", raw, 120, size)
        struct.pack_into("<I", raw, 124, size_hi)
        return bytes(raw)

    directory = entry("Root Entry", 5, 0xFFFFFFFE, 0)
    directory += entry(stream_name, 2, 2, len(stream_data), size_high)
    directory = directory.ljust(sector, b"\x00")
    stream = stream_data.ljust(stream_sectors * sector, b"\x00")

    header = bytearray(sector)
    header[:8] = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
    struct.pack_into("<H", header, 24, 0x003E)
    struct.pack_into("<H", header, 26, 3)
    struct.pack_into("<H", header, 28, 0xFFFE)
    struct.pack_into("<H", header, 30, 9)
    struct.pack_into("<H", header, 32, 6)
    struct.pack_into("<I", header, 44, 1)  # FAT count
    struct.pack_into("<I", header, 48, 1)  # first directory sector
    struct.pack_into("<I", header, 56, 4096)
    struct.pack_into("<I", header, 60, 0xFFFFFFFE)
    struct.pack_into("<I", header, 76, 0)  # first FAT sector
    return bytes(header) + fat + directory + stream


class CfbTests(unittest.TestCase):
    def test_reads_version3_stream_size_ignoring_high_dword_garbage(self):
        payload = b"P" * 5000
        data = _minimal_cfb(stream_name="PowerPoint Document", stream_data=payload, size_high=0xD2FCF2CB)
        compound = CompoundFile(data, Limits())
        self.assertEqual(compound.open_stream("PowerPoint Document"), payload)

    def test_rejects_missing_required_stream(self):
        data = _minimal_cfb(stream_name="Current User", stream_data=b"x")
        compound = CompoundFile(data, Limits())
        with self.assertRaises(InvalidPpt):
            compound.open_stream("PowerPoint Document")
