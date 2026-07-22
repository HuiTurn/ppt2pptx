"""Small bounded reader for CFB/OLE compound files used by legacy PPT."""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import struct

from .errors import InvalidPpt

FREE, END, FAT, DIF = 0xFFFFFFFF, 0xFFFFFFFE, 0xFFFFFFFD, 0xFFFFFFFC

@dataclass(frozen=True, slots=True)
class Limits:
    max_input_bytes: int = 512 * 1024 * 1024
    max_stream_bytes: int = 256 * 1024 * 1024

class CompoundFile:
    def __init__(self, data: bytes, limits: Limits | None = None) -> None:
        self.limits = limits or Limits()
        if len(data) > self.limits.max_input_bytes or len(data) < 512:
            raise InvalidPpt("invalid or oversized compound file")
        if data[:8] != b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
            raise InvalidPpt("not a Compound File Binary container")
        self.data = data
        major = struct.unpack_from("<H", data, 26)[0]
        self.sector_size = 512 if major == 3 else 4096 if major == 4 else 0
        if not self.sector_size or len(data) < self.sector_size:
            raise InvalidPpt("unsupported compound file version")
        self.mini_sector_size = 64
        self.mini_cutoff = struct.unpack_from("<I", data, 56)[0]
        self.first_dir = struct.unpack_from("<I", data, 48)[0]
        self.first_mini_fat = struct.unpack_from("<I", data, 60)[0]
        self.num_mini_fat = struct.unpack_from("<I", data, 64)[0]
        self._sectors = (len(data) - self.sector_size) // self.sector_size
        self.fat = self._load_fat()
        self.entries = self._read_directory()
        self.by_name = {name.casefold(): entry for name, entry in self.entries.items()}
        self.mini_fat = self._read_chain_values(self.first_mini_fat, self.num_mini_fat)
        self._mini_stream: bytes | None = None

    @classmethod
    def from_path(cls, path: str | Path, limits: Limits | None = None) -> "CompoundFile":
        p = Path(path)
        actual = limits or Limits()
        if p.stat().st_size > actual.max_input_bytes:
            raise InvalidPpt("input exceeds configured size limit")
        return cls(p.read_bytes(), actual)

    def _sector(self, index: int) -> bytes:
        if index >= self._sectors or index in (FREE, END, FAT, DIF):
            raise InvalidPpt("compound file sector reference is invalid")
        start = (index + 1) * self.sector_size
        return self.data[start:start + self.sector_size]

    def _load_fat(self) -> tuple[int, ...]:
        count = struct.unpack_from("<I", self.data, 44)[0]
        ids = [x for x in struct.unpack_from("<109I", self.data, 76) if x != FREE]
        next_difat, difat_count = struct.unpack_from("<II", self.data, 68)
        for _ in range(difat_count):
            block = self._sector(next_difat)
            values = struct.unpack("<%dI" % (self.sector_size // 4), block)
            ids.extend(x for x in values[:-1] if x != FREE)
            next_difat = values[-1]
        if len(ids) < count:
            raise InvalidPpt("compound file FAT is truncated")
        values: list[int] = []
        for index in ids[:count]: values.extend(struct.unpack("<%dI" % (self.sector_size // 4), self._sector(index)))
        return tuple(values)

    def _chain(self, start: int, table: tuple[int, ...], limit: int | None = None) -> list[int]:
        result, seen, current = [], set(), start
        while current != END:
            if current in seen or current >= len(table) or current in (FREE, FAT, DIF):
                raise InvalidPpt("compound file sector chain is invalid")
            seen.add(current); result.append(current)
            if len(result) > (limit or len(table) + 1): raise InvalidPpt("compound file sector chain is unbounded")
            current = table[current]
        return result

    def _read_chain_values(self, start: int, declared: int) -> tuple[int, ...]:
        if not declared: return ()
        raw = b"".join(self._sector(i) for i in self._chain(start, self.fat, declared))
        return struct.unpack("<%dI" % (len(raw) // 4), raw)

    def _read_directory(self) -> dict[str, tuple[int, int, int]]:
        raw = b"".join(self._sector(i) for i in self._chain(self.first_dir, self.fat))
        entries: list[tuple[str, int, int, int]] = []
        for offset in range(0, len(raw) - 127, 128):
            name_len = struct.unpack_from("<H", raw, offset + 64)[0]
            kind = raw[offset + 66]
            if kind not in (2, 5) or name_len < 2 or name_len > 64: continue
            name = raw[offset:offset + name_len - 2].decode("utf-16le", "replace")
            start, size = struct.unpack_from("<IQ", raw, offset + 116)
            entries.append((name, kind, start, size))
        return {name: (kind, start, size) for name, kind, start, size in entries}

    def open_stream(self, name: str) -> bytes:
        entry = self.by_name.get(name.casefold())
        if entry is None: raise InvalidPpt(f"required stream is missing: {name}")
        kind, start, size = entry
        if kind != 2 or size > self.limits.max_stream_bytes: raise InvalidPpt(f"invalid stream: {name}")
        if not size: return b""
        if size < self.mini_cutoff:
            if self._mini_stream is None:
                root = next((x for x in self.entries.values() if x[0] == 5), None)
                if root is None: raise InvalidPpt("compound file root storage is missing")
                self._mini_stream = self._read_regular(root[1], root[2])
            chunks = []
            for idx in self._chain(start, self.mini_fat):
                pos = idx * self.mini_sector_size; chunks.append(self._mini_stream[pos:pos+self.mini_sector_size])
            return b"".join(chunks)[:size]
        return self._read_regular(start, size)

    def _read_regular(self, start: int, size: int) -> bytes:
        return b"".join(self._sector(i) for i in self._chain(start, self.fat))[:size]
