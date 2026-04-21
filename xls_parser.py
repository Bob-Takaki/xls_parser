"""
xls_parser: A lightweight, dependency-free Excel (.xls) parser for Python.
Pandasなどの巨大な外部ライブラリに依存せず、標準ライブラリのみで軽量に動作します。
"""

import struct
import os
from typing import List, Any, Dict, Tuple, Optional

# --- 定数定義 (Constants) ---
OLE2_SIGNATURE = (0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1)

# BIFF (Binary Interchange File Format) Record IDs
RECORD_SST = 0x00FC        # Shared String Table (文字列辞書)
RECORD_LABELSST = 0x00FD   # Cell with String from SST (文字列セル)
RECORD_NUMBER = 0x0203     # Cell with Floating Point Number (数値セル)


def read_xls_as_array(file_path: str) -> List[List[Any]]:
    """
    指定された .xls ファイルを読み込み、データを2次元配列(リストのリスト)として返します。

    Args:
        file_path (str): 読み込む .xls ファイルのパス

    Returns:
        List[List[Any]]: 行と列で構成された2次元配列のデータ
                         (空のセルは空文字列 "" として埋められます)

    Raises:
        FileNotFoundError: ファイルが存在しない場合
        ValueError: 有効な .xls (OLE2) ファイルではない場合
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Error: '{file_path}' not found.")

    with open(file_path, 'rb') as f:
        # 1. OLE2ファイルシステムからWorkbookストリーム(バイナリ)を抽出
        workbook_data = _extract_workbook_stream(f)

    # 2. 抽出したバイナリからセルデータを解析して配列化
    return _parse_biff_data(workbook_data)


def _extract_workbook_stream(f) -> bytearray:
    """OLE2コンテナを解析し、Workbookデータストリームを抽出する内部関数"""
    header_data = f.read(512)
    magic = struct.unpack('<8B', header_data[:8])
    if magic != OLE2_SIGNATURE:
        raise ValueError("Invalid OLE2 signature. Not a valid .xls file.")

    sector_shift = struct.unpack('<H', header_data[30:32])[0]
    sector_size = 2 ** sector_shift
    dir_start_sector = struct.unpack('<I', header_data[48:52])[0]

    # ディレクトリの解析
    dir_offset = 512 + (dir_start_sector * sector_size)
    f.seek(dir_offset)

    wb_start_sector: Optional[int] = None
    wb_size: Optional[int] = None

    for _ in range(100):
        entry_data = f.read(128)
        if len(entry_data) < 128:
            break
        name_len = struct.unpack('<H', entry_data[64:66])[0]
        if name_len == 0:
            continue

        raw_name = entry_data[:name_len].decode('utf-16le')
        name = raw_name.rstrip('\x00')

        if name in ("Workbook", "Book"):
            wb_start_sector = struct.unpack('<I', entry_data[116:120])[0]
            wb_size = struct.unpack('<I', entry_data[120:124])[0]
            break

    if wb_start_sector is None or wb_size is None:
        raise ValueError("Workbook stream not found in OLE2 container.")

    # FATテーブルの構築
    difat_entries = struct.unpack('<109I', header_data[76:512])
    fat_table = []
    for sector_num in difat_entries:
        if sector_num == 0xFFFFFFFF:
            break
        f.seek(512 + sector_num * sector_size)
        fat_sector_data = f.read(sector_size)
        entries_per_sector = sector_size // 4
        fat_table.extend(struct.unpack(f'<{entries_per_sector}I', fat_sector_data))

    # FATチェーンを辿ってデータを結合
    workbook_chain = []
    current_sector = wb_start_sector
    while current_sector != 0xFFFFFFFE and current_sector != 0xFFFFFFFF:
        workbook_chain.append(current_sector)
        current_sector = fat_table[current_sector]

    workbook_data = bytearray()
    for sec in workbook_chain:
        f.seek(512 + sec * sector_size)
        workbook_data.extend(f.read(sector_size))

    return workbook_data[:wb_size]


def _parse_biff_data(workbook_data: bytearray) -> List[List[Any]]:
    """BIFFレコードを解析し、2次元配列を構築する内部関数"""
    offset = 0
    shared_strings: List[str] = []
    cells_dict: Dict[Tuple[int, int], Any] = {}
    max_row, max_col = 0, 0

    while offset < len(workbook_data):
        if offset + 4 > len(workbook_data):
            break

        record_id, record_size = struct.unpack('<HH', workbook_data[offset:offset+4])
        record_data = workbook_data[offset+4 : offset+4+record_size]

        # 1. SST (文字列辞書) の構築
        if record_id == RECORD_SST:
            unique_strings = struct.unpack('<I', record_data[4:8])[0]
            str_offset = 8
            for _ in range(unique_strings):
                if str_offset >= len(record_data): break
                char_count = struct.unpack('<H', record_data[str_offset:str_offset+2])[0]
                flags = record_data[str_offset+2]
                str_offset += 3

                has_rich = flags & 0x08
                has_ext = flags & 0x04
                run_count, ext_size = 0, 0

                if has_rich:
                    run_count = struct.unpack('<H', record_data[str_offset:str_offset+2])[0]
                    str_offset += 2
                if has_ext:
                    ext_size = struct.unpack('<I', record_data[str_offset:str_offset+4])[0]
                    str_offset += 4

                is_16bit = flags & 0x01
                if is_16bit:
                    byte_len = char_count * 2
                    text = record_data[str_offset:str_offset+byte_len].decode('utf-16le', errors='replace')
                else:
                    byte_len = char_count
                    text = record_data[str_offset:str_offset+byte_len].decode('latin-1', errors='replace')

                str_offset += byte_len
                if has_rich: str_offset += 4 * run_count
                if has_ext: str_offset += ext_size
                shared_strings.append(text)

        # 2. 文字列セル
        elif record_id == RECORD_LABELSST:
            row, col = struct.unpack('<HH', record_data[0:4])
            sst_index = struct.unpack('<I', record_data[6:10])[0]
            if sst_index < len(shared_strings):
                cells_dict[(row, col)] = shared_strings[sst_index]
                max_row, max_col = max(max_row, row), max(max_col, col)

        # 3. 数値セル
        elif record_id == RECORD_NUMBER:
            row, col = struct.unpack('<HH', record_data[0:4])
            value = struct.unpack('<d', record_data[6:14])[0]
            if value.is_integer():
                value = int(value)
            cells_dict[(row, col)] = value
            max_row, max_col = max(max_row, row), max(max_col, col)

        offset += 4 + record_size

    if not cells_dict:
        return []

    # 2次元配列化
    sheet_array = [["" for _ in range(max_col + 1)] for _ in range(max_row + 1)]
    for (r, c), value in cells_dict.items():
        sheet_array[r][c] = value

    return sheet_array
