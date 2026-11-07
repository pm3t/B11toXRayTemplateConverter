#!/usr/bin/env python3
"""
Breakdown Report Cargo Generator
Konversi file manifes BC 1.1 (format .xlsx maupun .xls) menjadi template Breakdown Report (.xlsx).
"""

import os
import sys
import argparse
from datetime import datetime, date
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Coba import xlrd untuk mendukung format file Excel lama (.xls)
try:
    import xlrd
    HAS_XLRD = True
except ImportError:
    HAS_XLRD = False


# Daftar header kolom untuk file output
OUTPUT_HEADERS = [
    'No',
    'MAWB',
    'HAWB',
    'Act Pcs',
    'of',
    'Tot Pcs',
    'Act Weight',
    'of',
    'Tot Weight',
    'Nature of Goods',
    'Actual Times',
    'Remarks',
    'Status',
    'BC11',
    'TanggalBC11',
    'FlightNumber',
    'FlightDate',
    'Origin',
    'Destination',
    'Shipper ',
    'AlamatShipper ',
    'Consignee',
    'AlamatConsignee'
]

# Lebar kolom default jika membuat dari nol tanpa template
COLUMN_WIDTHS = {
    'A': 5.0,
    'B': 14.0,
    'C': 14.0,
    'D': 8.0,
    'E': 5.0,
    'F': 8.0,
    'G': 10.0,
    'H': 5.0,
    'I': 12.0,
    'J': 20.0,
    'K': 12.0,
    'L': 10.0,
    'M': 8.0,
    'N': 10.0,
    'O': 12.0,
    'P': 12.0,
    'Q': 12.0,
    'R': 10.0,
    'S': 10.0,
    'T': 35.0,
    'U': 50.0,
    'V': 35.0,
    'W': 55.0,
}


def parse_date(date_val):
    """Parse berbagai format tanggal ke objek datetime."""
    if not date_val:
        return None
    if isinstance(date_val, (datetime, date)):
        return datetime(date_val.year, date_val.month, date_val.day)
    
    s = str(date_val).strip()
    date_formats = (
        '%d-%m-%Y', '%d/%m/%Y', '%Y-%m-%d', '%Y/%m/%d',
        '%d-%m-%y', '%d/%m/%y', '%m/%d/%Y', '%m-%d-%Y'
    )
    for fmt in date_formats:
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            pass
    return s


def get_nature_of_goods(id_detil, barang_list, max_length=50):
    """
    Mengambil Nature of Goods / Uraian Barang berdasarkan ID DETIL.
    Memfilter entri 'CONSOL' / HS CODE '0000' jika ada uraian barang spesifik lainnya.
    Dibatasi maksimal max_length karakter (default: 50 karakter).
    """
    if id_detil is None:
        return None

    # Normalisasi id_detil ke string
    id_detil_str = str(id_detil).strip()
    # Jika float seperti '1.0', ubah ke '1'
    if id_detil_str.endswith('.0'):
        id_detil_str = id_detil_str[:-2]

    matched = []
    for b in barang_list:
        b_id = str(b.get('ID DETIL', '')).strip()
        if b_id.endswith('.0'):
            b_id = b_id[:-2]
        if b_id == id_detil_str:
            matched.append(b)

    if not matched:
        return None
    
    # Filter item yang bukan 'CONSOL' dan bukan kode HS '0000' / '0'
    non_consol = [
        b for b in matched 
        if str(b.get('URAIAN BARANG', '')).strip().upper() != 'CONSOL'
        and str(b.get('HS CODE', '')).strip() not in ('0000', '0')
    ]
    
    if non_consol:
        descriptions = [str(b.get('URAIAN BARANG', '')).strip() for b in non_consol if b.get('URAIAN BARANG')]
        res = '; '.join(descriptions)
    else:
        # Jika hanya ada CONSOL atau deskripsi umum
        res = str(matched[0].get('URAIAN BARANG', '') or '').strip()

    if res and max_length is not None and len(res) > max_length:
        res = res[:max_length]

    return res


def _read_input_xlsx(input_path):
    """Membaca file format modern .xlsx menggunakan openpyxl."""
    wb = openpyxl.load_workbook(input_path, data_only=True)
    
    # 1. Sheet Header
    header_data = {}
    if 'Header' in wb.sheetnames:
        ws_hdr = wb['Header']
        keys = [ws_hdr.cell(1, c).value for c in range(1, ws_hdr.max_column + 1)]
        values = [ws_hdr.cell(2, c).value for c in range(1, ws_hdr.max_column + 1)]
        header_data = dict(zip(keys, values))
    
    # 2. Sheet Detil / Detail
    detil_rows = []
    detil_sheet_name = next((s for s in wb.sheetnames if s.lower() in ('detil', 'detail')), None)
    if detil_sheet_name:
        ws_det = wb[detil_sheet_name]
        keys = [ws_det.cell(1, c).value for c in range(1, ws_det.max_column + 1)]
        for r in range(2, ws_det.max_row + 1):
            row_vals = [ws_det.cell(r, c).value for c in range(1, ws_det.max_column + 1)]
            if any(v is not None and str(v).strip() != '' for v in row_vals):
                detil_rows.append(dict(zip(keys, row_vals)))

    # 3. Sheet Barang
    barang_rows = []
    if 'Barang' in wb.sheetnames:
        ws_brg = wb['Barang']
        keys = [ws_brg.cell(1, c).value for c in range(1, ws_brg.max_column + 1)]
        for r in range(2, ws_brg.max_row + 1):
            row_vals = [ws_brg.cell(r, c).value for c in range(1, ws_brg.max_column + 1)]
            if any(v is not None and str(v).strip() != '' for v in row_vals):
                barang_rows.append(dict(zip(keys, row_vals)))

    wb.close()
    return header_data, detil_rows, barang_rows


def _read_input_xls(input_path):
    """Membaca file format lama .xls menggunakan xlrd."""
    if not HAS_XLRD:
        raise ImportError(
            "Modul 'xlrd' belum terpasang untuk membaca file format .xls lama. "
            "Silakan pasang dengan menjalankan: pip install xlrd"
        )
    
    wb = xlrd.open_workbook(input_path)

    def get_row(sheet, r):
        row_vals = []
        for c in range(sheet.ncols):
            val = sheet.cell_value(r, c)
            ctype = sheet.cell_type(r, c)
            if ctype == xlrd.XL_CELL_DATE:
                try:
                    dt = xlrd.xldate_as_tuple(val, wb.datemode)
                    val = datetime(*dt)
                except Exception:
                    pass
            elif ctype == xlrd.XL_CELL_NUMBER:
                if isinstance(val, float) and val.is_integer():
                    val = int(val)
            row_vals.append(val)
        return row_vals

    # 1. Sheet Header
    header_data = {}
    if 'Header' in wb.sheet_names():
        ws_hdr = wb.sheet_by_name('Header')
        if ws_hdr.nrows >= 1:
            keys = [str(ws_hdr.cell_value(0, c)).strip() for c in range(ws_hdr.ncols)]
            vals = get_row(ws_hdr, 1) if ws_hdr.nrows > 1 else []
            header_data = dict(zip(keys, vals))

    # 2. Sheet Detil / Detail
    detil_rows = []
    detil_sheet_name = next((s for s in wb.sheet_names() if s.lower() in ('detil', 'detail')), None)
    if detil_sheet_name:
        ws_det = wb.sheet_by_name(detil_sheet_name)
        if ws_det.nrows >= 1:
            keys = [str(ws_det.cell_value(0, c)).strip() for c in range(ws_det.ncols)]
            for r in range(1, ws_det.nrows):
                vals = get_row(ws_det, r)
                if any(v not in (None, '') for v in vals):
                    detil_rows.append(dict(zip(keys, vals)))

    # 3. Sheet Barang
    barang_rows = []
    if 'Barang' in wb.sheet_names():
        ws_brg = wb.sheet_by_name('Barang')
        if ws_brg.nrows >= 1:
            keys = [str(ws_brg.cell_value(0, c)).strip() for c in range(ws_brg.ncols)]
            for r in range(1, ws_brg.nrows):
                vals = get_row(ws_brg, r)
                if any(v not in (None, '') for v in vals):
                    barang_rows.append(dict(zip(keys, vals)))

    return header_data, detil_rows, barang_rows


def read_input_excel(input_path):
    """
    Membaca file input excel (otomatis mendeteksi .xlsx maupun .xls).
    """
    if not os.path.exists(input_path):
        raise FileNotFoundError(f"File input '{input_path}' tidak ditemukan.")

    ext = os.path.splitext(input_path)[1].lower()
    if ext == '.xls':
        return _read_input_xls(input_path)
    elif ext == '.xlsx':
        try:
            return _read_input_xlsx(input_path)
        except Exception:
            # Fallback jika file berekstensi .xlsx ternyata berformat .xls di dalamnya
            if HAS_XLRD:
                try:
                    return _read_input_xls(input_path)
                except Exception:
                    pass
            raise
    else:
        # Ekstensi lain, coba .xlsx terlebih dahulu lalu fallback ke .xls
        try:
            return _read_input_xlsx(input_path)
        except Exception:
            if HAS_XLRD:
                return _read_input_xls(input_path)
            raise


def create_or_load_template(template_path='template.xlsx'):
    """Memuat template jika ada, atau membuat workbook baru dengan styling yang sesuai."""
    if template_path and os.path.exists(template_path):
        wb = openpyxl.load_workbook(template_path)
        ws = wb.active
        # Pastikan baris data lama (mulai baris 3) bersih
        if ws.max_row > 2:
            ws.delete_rows(3, ws.max_row - 2)
        return wb, ws
    
    # Buat dari nol jika template.xlsx tidak ada
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Sheet1'
    
    # Baris 1: Judul
    ws.merge_cells('A1:V1')
    cell_title = ws['A1']
    cell_title.value = 'Breakdown Report Cargo International'
    cell_title.font = Font(name='Aptos Narrow', size=12, bold=True)
    cell_title.alignment = Alignment(horizontal='center', vertical='center')
    
    # Baris 2: Header kolom
    header_font = Font(name='Tahoma', size=8, bold=False)
    header_fill = PatternFill(start_color='708090', end_color='708090', fill_type='solid')
    header_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    
    for col_idx, header_text in enumerate(OUTPUT_HEADERS, start=1):
        cell = ws.cell(row=2, column=col_idx, value=header_text)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        
        col_letter = get_column_letter(col_idx)
        if col_letter in COLUMN_WIDTHS:
            ws.column_dimensions[col_letter].width = COLUMN_WIDTHS[col_letter]
            
    return wb, ws


def is_transshipment(det_row, header_data=None):
    """
    Mendeteksi apakah baris detil merupakan kargo transshipment (hanya transit / lanjut ke luar negeri).
    Kriteria transshipment:
    1. KD KELOMPOK POS bernilai '02' (kargo transshipment / pos lanjutan)
    2. ATAU PELABUHAN BONGKAR / PELABUHAN AKHIR bukan pelabuhan kedatangan / tujuan impor di Header
    """
    # 1. Cek kode kelompok pos (Pos 01 = Impor lokal, Pos 02 = Transshipment / Transit)
    kd_pos = str(det_row.get('KD KELOMPOK POS') or '').strip()
    if kd_pos.endswith('.0'):
        kd_pos = kd_pos[:-2]
    # Format dua digit: '1' -> '01', '2' -> '02'
    if len(kd_pos) == 1:
        kd_pos = '0' + kd_pos
        
    if kd_pos == '02':
        return True

    # 2. Cek pelabuhan bongkar / pelabuhan akhir jika pelabuhan bongkar di header diketahui
    header_bongkar = ''
    if header_data:
        header_bongkar = str(header_data.get('PELABUHAN BONGKAR') or '').strip().upper()

    det_bongkar = str(det_row.get('PELABUHAN BONGKAR') or '').strip().upper()
    det_akhir = str(det_row.get('PELABUHAN AKHIR') or '').strip().upper()

    if header_bongkar and det_bongkar and det_bongkar != header_bongkar:
        return True

    return False


def convert_manifest_to_breakdown(input_path, output_path, template_path='template.xlsx'):
    """
    Fungsi utama untuk membaca input manifes dan menulis output template breakdown report.
    """
    # Otomatis koreksi typo ekstensi output seperti .xlxs -> .xlsx
    if output_path.lower().endswith('.xlxs'):
        output_path = output_path[:-5] + '.xlsx'
        print(f"[i] Nama file output disesuaikan menjadi: {output_path}")

    print(f"[*] Membaca file input: {input_path}")
    header_data, detil_rows, barang_rows = read_input_excel(input_path)
    print(f"    - Ditemukan {len(detil_rows)} baris data pada Detil.")
    print(f"    - Ditemukan {len(barang_rows)} entri barang pada Barang.")

    # Filter kargo transshipment (hanya ambil kargo impor lokal / screening)
    filtered_detil_rows = []
    skipped_count = 0
    for det in detil_rows:
        if is_transshipment(det, header_data):
            skipped_count += 1
            mawb = det.get('NO MASTER BLAWB')
            bongkar = det.get('PELABUHAN BONGKAR') or det.get('PELABUHAN AKHIR')
            print(f"    [i] Melewati kargo transshipment: MAWB {mawb} (Tujuan: {bongkar})")
        else:
            filtered_detil_rows.append(det)

    if skipped_count > 0:
        print(f"    - {skipped_count} kargo transshipment dilewati. {len(filtered_detil_rows)} kargo akan digenerate.")
    
    detil_rows = filtered_detil_rows
    
    # Header umum
    bc11_no = header_data.get('NO BC 11')
    if bc11_no is not None:
        bc11_no = str(bc11_no).strip()
        if bc11_no.endswith('.0'):
            bc11_no = bc11_no[:-2]

    bc11_date = parse_date(header_data.get('TGL BC 11'))
    flight_number = header_data.get('NO VOYAGE / ARRIVAL') or header_data.get('NO IMO') or header_data.get('CALL SIGN')
    flight_date = parse_date(header_data.get('TGL TIBA') or header_data.get('TGL BC 11'))
    default_origin = header_data.get('PELABUHAN ASAL')
    default_dest = header_data.get('PELABUHAN BONGKAR')

    # Siapkan workbook output
    wb, ws = create_or_load_template(template_path)
    
    # Styling baris data
    data_font = Font(name='Aptos Narrow', size=11, bold=False)
    date_font = Font(name='Arial', size=10, bold=False)
    data_align = Alignment(vertical='top', wrap_text=True)
    
    # Tulis data
    for idx, det in enumerate(detil_rows, start=1):
        row_num = idx + 2
        
        # Act Pcs (diambil dari JUMLAH KEMASAN)
        act_pcs = det.get('JUMLAH KEMASAN')
        if act_pcs is not None and str(act_pcs).strip() != '':
            try:
                act_pcs = int(float(act_pcs))
            except (ValueError, TypeError):
                pass
        else:
            act_pcs = None

        # Tot Pcs
        tot_pcs = det.get('TOTAL KEMASAN')
        if tot_pcs is None or str(tot_pcs).strip() == '':
            tot_pcs = det.get('JUMLAH KEMASAN')
        if tot_pcs is not None:
            try:
                tot_pcs = int(float(tot_pcs))
            except (ValueError, TypeError):
                pass
        
        # Tot Weight
        tot_weight = det.get('BRUTO')
        if tot_weight is not None:
            tot_weight = str(tot_weight).strip()
            
        nature_of_goods = get_nature_of_goods(det.get('ID DETIL'), barang_rows)
        origin = det.get('PELABUHAN ASAL') or default_origin
        destination = det.get('PELABUHAN BONGKAR') or default_dest
        
        row_values = [
            idx,                                        # 1: No
            det.get('NO MASTER BLAWB'),                 # 2: MAWB
            det.get('NO HOST BLAWB'),                   # 3: HAWB
            act_pcs,                                    # 4: Act Pcs
            None,                                       # 5: of
            tot_pcs,                                    # 6: Tot Pcs
            None,                                       # 7: Act Weight
            None,                                       # 8: of
            tot_weight,                                 # 9: Tot Weight
            nature_of_goods,                            # 10: Nature of Goods
            None,                                       # 11: Actual Times
            None,                                       # 12: Remarks
            None,                                       # 13: Status
            bc11_no,                                    # 14: BC11
            bc11_date,                                  # 15: TanggalBC11
            flight_number,                              # 16: FlightNumber
            flight_date,                                # 17: FlightDate
            origin,                                     # 18: Origin
            destination,                                # 19: Destination
            det.get('NAMA SHIPPER'),                    # 20: Shipper
            det.get('ALMT SHIPPER'),                    # 21: AlamatShipper
            det.get('NAMA CONSIGNEE'),                  # 22: Consignee
            det.get('ALMT CONSIGNEE')                   # 23: AlamatConsignee
        ]
        
        for col_idx, val in enumerate(row_values, start=1):
            cell = ws.cell(row=row_num, column=col_idx, value=val)
            cell.alignment = data_align
            
            # Khusus kolom tanggal (TanggalBC11 di col 15, FlightDate di col 17)
            if col_idx in (15, 17):
                cell.font = date_font
                if isinstance(val, (datetime, date)):
                    cell.number_format = 'mm-dd-yy'
            elif col_idx == 14: # BC11
                cell.font = date_font
            else:
                cell.font = data_font

    wb.save(output_path)
    wb.close()
    print(f"[✓] Berhasil membuat file output: {output_path} ({len(detil_rows)} baris data)")


def main():
    parser = argparse.ArgumentParser(description="Program Generator Template XRay Cargo dari Manifes (.xlsx / .xls) ke OUT.xlsx")
    parser.add_argument("input", nargs="?", default="IN.xlsx", help="Path file input excel (default: IN.xlsx)")
    parser.add_argument("output", nargs="?", default="OUT.xlsx", help="Path file output excel (default: OUT.xlsx)")
    parser.add_argument("--template", default="template.xlsx", help="Path file template excel (default: template.xlsx)")
    
    args = parser.parse_args()
    
    # Jalankan konversi
    convert_manifest_to_breakdown(args.input, args.output, args.template)


if __name__ == "__main__":
    main()
