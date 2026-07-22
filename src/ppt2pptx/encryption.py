"""RC4 CryptoAPI decryption for password-protected PowerPoint streams."""
from __future__ import annotations

import hashlib
import hmac
import struct

from .errors import EncryptedPresentationError, InvalidPpt

RT_USER_EDIT_ATOM = 4085
RT_PERSIST_DIRECTORY_ATOM = 6002
RT_DOCUMENT_ENCRYPTION_ATOM = 12052


def _rc4(key: bytes, data: bytes) -> bytes:
    if not key:
        raise InvalidPpt("PowerPoint encryption key is empty")
    state = list(range(256))
    j = 0
    for index in range(256):
        j = (j + state[index] + key[index % len(key)]) & 0xFF
        state[index], state[j] = state[j], state[index]
    output = bytearray(len(data))
    i = j = 0
    for position, value in enumerate(data):
        i = (i + 1) & 0xFF
        j = (j + state[i]) & 0xFF
        state[i], state[j] = state[j], state[i]
        output[position] = value ^ state[(state[i] + state[j]) & 0xFF]
    return bytes(output)


def _cryptoapi_key(secret: bytes, block: int, key_bits: int) -> bytes:
    digest = hashlib.sha1(secret + struct.pack("<I", block)).digest()
    if key_bits == 40:
        return digest[:5] + bytes(11)
    return digest[: key_bits // 8]


def _record(data: bytes, offset: int, expected_type: int | None = None) -> tuple[int, bytes]:
    if offset < 0 or offset + 8 > len(data):
        raise InvalidPpt("PowerPoint record offset is outside the document stream")
    _, record_type, length = struct.unpack_from("<HHI", data, offset)
    end = offset + 8 + length
    if end > len(data):
        raise InvalidPpt("PowerPoint record extends beyond the document stream")
    if expected_type is not None and record_type != expected_type:
        raise InvalidPpt(f"expected PowerPoint record {expected_type} at offset {offset}, found {record_type}")
    return record_type, data[offset + 8:end]


def _persist_mappings(data: bytes, current_user: bytes) -> tuple[dict[int, int], int]:
    if len(current_user) < 20:
        raise InvalidPpt("encrypted Current User stream is truncated")
    user_edit_offset = struct.unpack_from("<I", current_user, 16)[0]
    mappings: dict[int, int] = {}
    encryption_persist_id: int | None = None
    seen: set[int] = set()
    while user_edit_offset:
        if user_edit_offset in seen or len(seen) >= 4096:
            raise InvalidPpt("PowerPoint user edit chain is cyclic or unbounded")
        seen.add(user_edit_offset)
        _, edit = _record(data, user_edit_offset, RT_USER_EDIT_ATOM)
        if len(edit) not in (28, 32):
            raise InvalidPpt("encrypted PowerPoint UserEditAtom has an invalid size")
        previous_edit, persist_offset = struct.unpack_from("<II", edit, 8)
        if encryption_persist_id is None and len(edit) == 32:
            candidate = struct.unpack_from("<I", edit, 28)[0]
            if candidate != 0xFFFFFFFF:
                encryption_persist_id = candidate
        _, persist = _record(data, persist_offset, RT_PERSIST_DIRECTORY_ATOM)
        cursor = 0
        while cursor < len(persist):
            if cursor + 4 > len(persist):
                raise InvalidPpt("PowerPoint persist directory is truncated")
            info = struct.unpack_from("<I", persist, cursor)[0]
            cursor += 4
            count, first = info >> 20, info & 0xFFFFF
            if not count or cursor + count * 4 > len(persist):
                raise InvalidPpt("PowerPoint persist directory entry is invalid")
            for index in range(count):
                offset = struct.unpack_from("<I", persist, cursor)[0]
                cursor += 4
                mappings.setdefault(first + index, offset)
        user_edit_offset = previous_edit
    if encryption_persist_id is None:
        raise InvalidPpt("encrypted PowerPoint file has no encryption session reference")
    return mappings, encryption_persist_id


def _encryption_parameters(payload: bytes, password: str) -> tuple[bytes, int]:
    if len(payload) < 72:
        raise InvalidPpt("PowerPoint encryption information is truncated")
    major, minor, outer_flags, header_size = struct.unpack_from("<HHII", payload)
    if (major, minor) not in ((2, 2), (3, 2), (4, 2)):
        raise EncryptedPresentationError(f"unsupported PowerPoint encryption version {major}.{minor}")
    header_end = 12 + header_size
    if header_size < 32 or header_end + 60 > len(payload):
        raise InvalidPpt("PowerPoint CryptoAPI encryption header is truncated")
    flags, size_extra, algorithm_id, hash_id, key_bits, provider_type, _, reserved2 = struct.unpack_from(
        "<8I", payload, 12
    )
    if flags != outer_flags or not flags & 0x04 or flags & 0x30:
        raise InvalidPpt("PowerPoint CryptoAPI flags are inconsistent")
    if size_extra != 0 or algorithm_id not in (0, 0x6801) or hash_id not in (0, 0x8004):
        raise EncryptedPresentationError("unsupported PowerPoint CryptoAPI cipher or hash algorithm")
    key_bits = key_bits or 40
    if key_bits < 40 or key_bits > 128 or key_bits % 8 or provider_type not in (0, 1) or reserved2:
        raise InvalidPpt("PowerPoint CryptoAPI key or provider fields are invalid")
    salt_size = struct.unpack_from("<I", payload, header_end)[0]
    if salt_size != 16:
        raise InvalidPpt("PowerPoint CryptoAPI salt must contain 16 bytes")
    salt = payload[header_end + 4:header_end + 20]
    encrypted_verifier = payload[header_end + 20:header_end + 36]
    hash_size = struct.unpack_from("<I", payload, header_end + 36)[0]
    encrypted_hash = payload[header_end + 40:header_end + 60]
    if hash_size != 20 or len(encrypted_hash) != 20:
        raise InvalidPpt("PowerPoint CryptoAPI verifier hash is invalid")
    secret = hashlib.sha1(salt + password[:255].encode("utf-16le")).digest()
    verifier_data = _rc4(_cryptoapi_key(secret, 0, key_bits), encrypted_verifier + encrypted_hash)
    if not hmac.compare_digest(hashlib.sha1(verifier_data[:16]).digest(), verifier_data[16:]):
        raise EncryptedPresentationError("incorrect password for encrypted PowerPoint presentation")
    return secret, key_bits


def decrypt_powerpoint_document(data: bytes, current_user: bytes, password: str | None) -> bytes:
    """Decrypt persisted records while leaving edit and directory records intact."""
    if password is None:
        raise EncryptedPresentationError("password-protected PowerPoint presentation; provide --password")
    mappings, encryption_persist_id = _persist_mappings(data, current_user)
    encryption_offset = mappings.get(encryption_persist_id)
    if encryption_offset is None:
        raise InvalidPpt("PowerPoint encryption record is missing from the persist directory")
    _, encryption_payload = _record(data, encryption_offset, RT_DOCUMENT_ENCRYPTION_ATOM)
    secret, key_bits = _encryption_parameters(encryption_payload, password)
    output = bytearray(data)
    occupied_until = 0
    for persist_id, offset in sorted(mappings.items(), key=lambda item: item[1]):
        if persist_id == encryption_persist_id:
            continue
        if offset < occupied_until:
            continue
        key = _cryptoapi_key(secret, persist_id, key_bits)
        header = _rc4(key, data[offset:offset + 8])
        if len(header) != 8:
            raise InvalidPpt("encrypted PowerPoint record header is truncated")
        length = struct.unpack_from("<I", header, 4)[0]
        end = offset + 8 + length
        if length > len(data) or end > len(data):
            raise InvalidPpt("decrypted PowerPoint record has an invalid size")
        output[offset:end] = _rc4(key, data[offset:end])
        occupied_until = end
    return bytes(output)
