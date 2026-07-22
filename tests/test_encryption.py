import hashlib
import struct
import unittest

from ppt2pptx.encryption import (
    _cryptoapi_key,
    _rc4,
    decrypt_powerpoint_document,
)
from ppt2pptx.errors import EncryptedPresentationError


def rec(kind, payload=b"", version=0, instance=0):
    return struct.pack("<HHI", (instance << 4) | version, kind, len(payload)) + payload


class EncryptionTests(unittest.TestCase):
    def _fixture(self, password="secret"):
        salt = bytes(range(16))
        verifier = bytes(range(16, 32))
        secret = hashlib.sha1(salt + password.encode("utf-16le")).digest()
        verifier_cipher = _rc4(
            _cryptoapi_key(secret, 0, 40),
            verifier + hashlib.sha1(verifier).digest(),
        )
        header = struct.pack("<HHII8I", 2, 2, 4, 32, 4, 0, 0x6801, 0x8004, 40, 1, 0, 0)
        encryption_payload = header + struct.pack("<I", 16) + salt + verifier_cipher[:16] + struct.pack("<I", 20) + verifier_cipher[16:]

        clear = rec(1000, rec(1001, bytes(40), version=1), version=0xF)
        encrypted = _rc4(_cryptoapi_key(secret, 1, 40), clear)
        encryption_offset = len(encrypted)
        encryption_record = rec(12052, encryption_payload, version=0xF)
        persist_offset = encryption_offset + len(encryption_record)
        persist_payload = struct.pack("<I2I", (2 << 20) | 1, 0, encryption_offset)
        persist_record = rec(6002, persist_payload)
        edit_offset = persist_offset + len(persist_record)
        edit_payload = struct.pack("<6I2HI", 0, 0, 0, persist_offset, 1, 2, 0, 0, 2)
        edit_record = rec(4085, edit_payload)
        current_user = bytearray(20)
        current_user[12:16] = b"\xdf\xc4\xd1\xf3"
        struct.pack_into("<I", current_user, 16, edit_offset)
        return encrypted + encryption_record + persist_record + edit_record, bytes(current_user), clear

    def test_decrypts_cryptoapi_persist_record(self):
        stream, current_user, clear = self._fixture()
        decrypted = decrypt_powerpoint_document(stream, current_user, "secret")
        self.assertEqual(decrypted[:len(clear)], clear)

    def test_rejects_incorrect_password(self):
        stream, current_user, _ = self._fixture()
        with self.assertRaisesRegex(EncryptedPresentationError, "incorrect password"):
            decrypt_powerpoint_document(stream, current_user, "wrong")
