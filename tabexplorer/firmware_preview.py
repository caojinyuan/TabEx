"""Bounded, read-only views of binary images and ASCII firmware records."""

import binascii
from dataclasses import dataclass
import os

from .i18n import tr


BINARY_EXTENSIONS = {'.bin', '.rom', '.eeprom', '.eep', '.img', '.raw', '.dump'}
SREC_EXTENSIONS = {'.s19', '.s28', '.s37', '.s1', '.s2', '.s3', '.srec', '.srecord', '.mot'}
IHEX_EXTENSIONS = {'.hex', '.ihex', '.ihx'}
HEX_PAGE_BYTES = 4096
RECORD_BYTE_LIMIT = 256 * 1024
RECORD_LINE_LIMIT = 1000


@dataclass(frozen=True)
class HexPage:
    start: int
    size: int
    data: bytes

    def text(self):
        width = max(8, len(f'{self.size - 1:X}'))
        lines = [f'{"OFFSET":<{width}}  00 01 02 03 04 05 06 07  08 09 0A 0B 0C 0D 0E 0F  ASCII']
        for i in range(0, len(self.data), 16):
            row = self.data[i:i + 16]
            digits = f'{row[:8].hex(" "):<23}  {row[8:].hex(" "):<23}'.upper()
            ascii_text = ''.join(chr(b) if 32 <= b < 127 else '.' for b in row)
            lines.append(f'{self.start + i:0{width}X}  {digits}  {ascii_text}')
        return '\n'.join(lines)


def read_hex_page(path, offset=0):
    with open(path, 'rb') as stream:
        size = os.fstat(stream.fileno()).st_size
        if not 0 <= offset < max(1, size):
            raise ValueError(tr("偏移超出文件范围。"))
        start = offset // HEX_PAGE_BYTES * HEX_PAGE_BYTES
        stream.seek(start)
        data = stream.read(HEX_PAGE_BYTES)
        if len(data) != min(HEX_PAGE_BYTES, size - start):
            raise OSError(tr("文件大小已变化，请重新打开预览。"))
    return HexPage(start, size, data)


_SREC_WIDTHS = {'S0': 2, 'S1': 2, 'S2': 3, 'S3': 4, 'S5': 2, 'S6': 3, 'S7': 4, 'S8': 3, 'S9': 2}


def _record(line, motorola, base):
    """Return (type, address/parameter, payload, data record?, next Intel address base)."""
    if motorola:
        kind = line[:2].decode('ascii').upper()
        width = _SREC_WIDTHS.get(kind)
        if not width:
            raise ValueError(tr("记录类型无效"))
        raw = binascii.unhexlify(line[2:])
        if len(raw) < width + 2 or raw[0] != len(raw) - 1:
            raise ValueError(tr("记录长度不符"))
        if (sum(raw) & 0xff) != 0xff:
            raise ValueError(tr("校验和不符"))
        address = int.from_bytes(raw[1:1 + width], 'big')
        payload = raw[1 + width:-1]
        if kind in ('S5', 'S6', 'S7', 'S8', 'S9') and payload:
            raise ValueError(tr("记录长度不符"))
        return kind, address, payload, kind in ('S1', 'S2', 'S3'), base

    if not line.startswith(b':'):
        raise ValueError(tr("记录类型无效"))
    raw = binascii.unhexlify(line[1:])
    if len(raw) < 5 or raw[0] != len(raw) - 5:
        raise ValueError(tr("记录长度不符"))
    if sum(raw) & 0xff:
        raise ValueError(tr("校验和不符"))
    address, kind, payload = int.from_bytes(raw[1:3], 'big'), raw[3], raw[4:-1]
    if kind == 0:
        address = base + address if base is not None else None
    else:
        lengths = {1: 0, 2: 2, 3: 4, 4: 2, 5: 4}
        if kind not in lengths or address != 0 or len(payload) != lengths[kind]:
            raise ValueError(tr("记录类型或参数无效"))
        operand = int.from_bytes(payload, 'big')
        if kind in (2, 4):
            base = address = operand << (4 if kind == 2 else 16)
        elif kind == 3:
            address = ((operand >> 16) << 4) + (operand & 0xffff)
        elif kind == 5:
            address = operand
    return f'{kind:02X}', address, payload, kind == 0, base


def read_records(path, motorola):
    with open(path, 'rb') as stream:
        sample = stream.read(RECORD_BYTE_LIMIT + 1)
    clipped = len(sample) > RECORD_BYTE_LIMIT
    sample = sample[:RECORD_BYTE_LIMIT].removeprefix(b'\xef\xbb\xbf')
    if clipped and not sample.endswith((b'\r', b'\n')):
        # A record split by the preview limit is not a corrupt record.
        sample = sample[:max(sample.rfind(b'\r'), sample.rfind(b'\n')) + 1]
    lines = sample.splitlines()
    clipped = clipped or len(lines) > RECORD_LINE_LIMIT
    rows, errors, data_bytes, base = [], 0, 0, 0
    for number, raw in enumerate(lines[:RECORD_LINE_LIMIT], 1):
        line = raw.strip()
        if not line:
            continue
        try:
            if len(line) > (514 if motorola else 521):
                raise ValueError(tr("记录长度不符"))
            kind, address, payload, is_data, base = _record(line, motorola, base)
        except ValueError as error:
            errors += 1
            if not motorola and line[7:9] in (b'02', b'04'):
                base = None  # Do not present subsequent offsets as known absolute addresses.
            reason = tr("十六进制字符无效") if isinstance(error, (binascii.Error, UnicodeError)) else str(error)
            snippet = line[:80].decode('ascii', errors='backslashreplace') + ('…' if len(line) > 80 else '')
            rows.append((str(number), '—', '—', '—', snippet, reason))
        else:
            if is_data:
                data_bytes += len(payload)
            rows.append((str(number), kind, f'0x{address:08X}' if address is not None else '?',
                         str(len(payload)), payload.hex(' ').upper(), 'OK'))
    summary = tr("预览范围：{} 条记录 · {} 个数据字节 · {} 条错误").format(len(rows), data_bytes, errors)
    if clipped:
        summary += ' · ' + tr("内容已截断。")
    return rows, summary
