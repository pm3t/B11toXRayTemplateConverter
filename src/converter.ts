import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';

export interface HeaderData {
  noBc11?: string;
  tglBc11?: Date | string;
  flightNumber?: string;
  flightDate?: Date | string;
  origin?: string;
  destination?: string;
  [key: string]: any;
}

export interface DetilRow {
  no: number;
  mawb: string;
  hawb: string;
  actPcs: number | null;
  totPcs: number | null;
  actWeight?: string | null;
  totWeight: string;
  natureOfGoods: string;
  bc11: string;
  tanggalBc11: Date | null;
  flightNumber: string;
  flightDate: Date | null;
  origin: string;
  destination: string;
  shipper: string;
  alamatShipper: string;
  consignee: string;
  alamatConsignee: string;
  idDetil?: string;
}

export interface ConversionResult {
  header: HeaderData;
  rows: DetilRow[];
  totalRecords: number;
  transshipmentCount: number;
  formatType?: 'BC11' | 'NVOCC';
  fileName: string;
}

export const OUTPUT_HEADERS = [
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
];

export const COLUMN_WIDTHS: Record<string, number> = {
  A: 5.0,
  B: 14.0,
  C: 14.0,
  D: 8.0,
  E: 5.0,
  F: 8.0,
  G: 10.0,
  H: 5.0,
  I: 12.0,
  J: 20.0,
  K: 12.0,
  L: 10.0,
  M: 8.0,
  N: 10.0,
  O: 12.0,
  P: 12.0,
  Q: 12.0,
  R: 10.0,
  S: 10.0,
  T: 35.0,
  U: 50.0,
  V: 35.0,
  W: 55.0,
};

export function parseDate(val: any): Date | null {
  if (!val) return null;
  if (val instanceof Date) return val;
  const s = String(val).trim();
  
  // Format DD-MM-YYYY or DD/MM/YYYY
  const m1 = s.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/);
  if (m1) {
    return new Date(parseInt(m1[3]), parseInt(m1[2]) - 1, parseInt(m1[1]));
  }
  
  // Format YYYY-MM-DD
  const m2 = s.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/);
  if (m2) {
    return new Date(parseInt(m2[1]), parseInt(m2[2]) - 1, parseInt(m2[3]));
  }

  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

export function getNatureOfGoods(idDetil: any, barangRows: any[], maxLength = 50): string {
  if (idDetil === null || idDetil === undefined) return '';
  let idStr = String(idDetil).trim();
  if (idStr.endsWith('.0')) idStr = idStr.slice(0, -2);

  const matched = barangRows.filter((b) => {
    let bId = String(b['ID DETIL'] ?? '').trim();
    if (bId.endsWith('.0')) bId = bId.slice(0, -2);
    return bId === idStr;
  });

  if (matched.length === 0) return '';

  const nonConsol = matched.filter((b) => {
    const desc = String(b['URAIAN BARANG'] ?? '').trim().toUpperCase();
    const hs = String(b['HS CODE'] ?? '').trim();
    return desc !== 'CONSOL' && hs !== '0000' && hs !== '0';
  });

  let res = '';
  if (nonConsol.length > 0) {
    res = nonConsol
      .map((b) => String(b['URAIAN BARANG'] ?? '').trim())
      .filter(Boolean)
      .join('; ');
  } else {
    res = String(matched[0]['URAIAN BARANG'] ?? '').trim();
  }

  if (res && maxLength && res.length > maxLength) {
    res = res.slice(0, maxLength);
  }
  return res;
}

export async function parseManifestFile(file: File): Promise<ConversionResult> {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: 'array', cellDates: true });

  const blSheetName = workbook.SheetNames.find((s) => ['bl', 'b/l'].includes(s.toLowerCase().trim()));
  const isNvocc = Boolean(blSheetName);

  if (isNvocc) {
    // ==========================================
    // TIPE FORMAT 2: NVOCC (BL, BL HS, nvocc)
    // ==========================================
    let headerData: HeaderData = {
      noBc11: '',
      tglBc11: null,
      flightNumber: '',
      flightDate: null,
      origin: '',
      destination: '',
    };

    const nvoccSheetName = workbook.SheetNames.find((s) =>
      ['nvocc', 'nvocc header', 'header nvocc', 'header'].includes(s.toLowerCase().trim())
    );

    if (nvoccSheetName) {
      // Baca baris secara array of arrays untuk akses indeks kolom yang pasti
      const rawAoa = XLSX.utils.sheet_to_json<any[]>(workbook.Sheets[nvoccSheetName], { header: 1 });
      if (rawAoa.length > 1) {
        const row1 = rawAoa[1];
        // FlightNo: sheet nvocc kolom I (index 8)
        const flightNo = row1[8] !== undefined ? String(row1[8]).trim() : '';
        // FlightDate: sheet nvocc kolom tanggal tiba (K) (index 10)
        const flightDate = parseDate(row1[10]);

        headerData.flightNumber = flightNo;
        headerData.flightDate = flightDate;
      }
    }

    // Baca Sheet BL HS
    const hsSheetName = workbook.SheetNames.find((s) =>
      ['bl hs', 'bl_hs', 'hs', 'blhs'].includes(s.toLowerCase().trim())
    );
    const barangRaw: any[] = [];
    if (hsSheetName) {
      const hsAoa = XLSX.utils.sheet_to_json<any[]>(workbook.Sheets[hsSheetName], { header: 1 });
      for (let r = 1; r < hsAoa.length; r++) {
        const row = hsAoa[r];
        if (!row || row.length === 0) continue;
        const idBl = row[1] !== undefined ? String(row[1]).trim() : ''; // Kolom B: Id Bl
        const uraian = row[3] !== undefined ? String(row[3]).trim() : ''; // Kolom D: Uraian Barang
        const kodeHs = row[2] !== undefined ? String(row[2]).trim() : ''; // Kolom C: Kode Hs
        const seri = row[4] !== undefined ? String(row[4]).trim() : ''; // Kolom E: Seri Hs
        if (idBl || uraian) {
          barangRaw.push({
            'ID DETIL': idBl,
            'URAIAN BARANG': uraian,
            'HS CODE': kodeHs,
            'SERI BARANG': seri,
          });
        }
      }
    }

    // Baca Sheet BL
    const blAoa = blSheetName
      ? XLSX.utils.sheet_to_json<any[]>(workbook.Sheets[blSheetName], { header: 1 })
      : [];

    const validDetil: any[] = [];
    let transshipmentCount = 0;

    for (let r = 1; r < blAoa.length; r++) {
      const row = blAoa[r];
      if (!row || row.length === 0 || !row.some((c) => c !== undefined && c !== null && String(c).trim() !== '')) {
        continue;
      }

      let kdPos = row[3] !== undefined ? String(row[3]).trim() : ''; // Kolom D: Kode Kelompok Pos
      if (kdPos.endsWith('.0')) kdPos = kdPos.slice(0, -2);
      if (kdPos.length === 1) kdPos = '0' + kdPos;

      const dest = row[25] !== undefined && String(row[25]).trim() !== ''
        ? String(row[25]).trim()
        : (row[23] !== undefined ? String(row[23]).trim() : ''); // Kolom Z (Akhir), fallback X (Bongkar)

      const isTransshipment = kdPos === '02';
      if (isTransshipment) {
        transshipmentCount++;
      } else {
        validDetil.push(row);
      }
    }

    const rows: DetilRow[] = validDetil.map((row, index) => {
      const idx = index + 1;
      const idBl = row[0] !== undefined ? String(row[0]).trim() : ''; // Kolom A: Id Bl
      const mawb = row[5] !== undefined ? String(row[5]).trim() : ''; // Kolom F: Nomor Bl
      const hawb = row[7] !== undefined ? String(row[7]).trim() : ''; // Kolom H: Nomor Host Bl

      // Act PCs & Total PCs: sheet BL kolom U (index 20)
      let pcs: number | null = null;
      const rawPcs = row[20];
      if (rawPcs !== undefined && rawPcs !== null && String(rawPcs).trim() !== '') {
        const n = Number(rawPcs);
        if (!isNaN(n)) pcs = Math.floor(n);
      }

      // Act weight: sheet BL kolom Q (index 16)
      let actWeight: string | null = null;
      if (row[16] !== undefined && row[16] !== null && String(row[16]).trim() !== '') {
        actWeight = String(row[16]).trim();
      }

      const natureOfGoods = getNatureOfGoods(idBl, barangRaw, 50);

      // Origin: sheet BL kolom Y (index 24), fallback W (index 22)
      const origin = row[24] !== undefined && String(row[24]).trim() !== ''
        ? String(row[24]).trim()
        : (row[22] !== undefined ? String(row[22]).trim() : '');

      // Destination: sheet BL kolom Z (index 25), fallback X (index 23)
      const destination = row[25] !== undefined && String(row[25]).trim() !== ''
        ? String(row[25]).trim()
        : (row[23] !== undefined ? String(row[23]).trim() : '');

      // Shipper: sheet BL kolom K (index 10)
      const shipper = row[10] !== undefined ? String(row[10]).trim() : '';
      // Alamat Shipper: sheet BL kolom L (index 11)
      const alamatShipper = row[11] !== undefined ? String(row[11]).trim() : '';
      // Consignee: sheet BL kolom M (index 12)
      const consignee = row[12] !== undefined ? String(row[12]).trim() : '';
      // Alamat Consignee: sheet BL kolom N (index 13)
      const alamatConsignee = row[13] !== undefined ? String(row[13]).trim() : '';

      return {
        no: idx,
        mawb,
        hawb,
        actPcs: pcs,
        totPcs: pcs,
        actWeight,
        totWeight: actWeight || '',
        natureOfGoods,
        bc11: '',
        tanggalBc11: null,
        flightNumber: headerData.flightNumber || '',
        flightDate: headerData.flightDate instanceof Date ? headerData.flightDate : null,
        origin,
        destination,
        shipper,
        alamatShipper,
        consignee,
        alamatConsignee,
        idDetil: idBl,
      };
    });

    return {
      header: headerData,
      rows,
      totalRecords: rows.length,
      transshipmentCount,
      formatType: 'NVOCC',
      fileName: file.name,
    };
  }

  // ==========================================
  // TIPE FORMAT 1: BC 1.1 (Header, Detil, Barang)
  // ==========================================
  // 1. Header Sheet
  let headerData: HeaderData = {};
  const headerSheetName = workbook.SheetNames.find((s) => s.toLowerCase() === 'header');
  if (headerSheetName) {
    const rawHeader = XLSX.utils.sheet_to_json<any>(workbook.Sheets[headerSheetName]);
    if (rawHeader.length > 0) {
      const h = rawHeader[0];
      let bc11 = h['NO BC 11'] !== undefined ? String(h['NO BC 11']).trim() : '';
      if (bc11.endsWith('.0')) bc11 = bc11.slice(0, -2);

      headerData = {
        noBc11: bc11,
        tglBc11: parseDate(h['TGL BC 11']),
        flightNumber: String(h['NO VOYAGE / ARRIVAL'] || h['NO IMO'] || h['CALL SIGN'] || '').trim(),
        flightDate: parseDate(h['TGL TIBA'] || h['TGL BC 11']),
        origin: String(h['PELABUHAN ASAL'] || '').trim(),
        destination: String(h['PELABUHAN BONGKAR'] || '').trim(),
        ...h,
      };
    }
  }

  // 2. Detil Sheet
  const detilSheetName = workbook.SheetNames.find((s) => ['detil', 'detail'].includes(s.toLowerCase()));
  const detilRaw = detilSheetName
    ? XLSX.utils.sheet_to_json<any>(workbook.Sheets[detilSheetName])
    : [];

  // 3. Barang Sheet
  const barangSheetName = workbook.SheetNames.find((s) => s.toLowerCase() === 'barang');
  const barangRaw = barangSheetName
    ? XLSX.utils.sheet_to_json<any>(workbook.Sheets[barangSheetName])
    : [];

  // Filter kargo transshipment (KD KELOMPOK POS == '02' atau PELABUHAN BONGKAR bukan tujuan di Header)
  const headerBongkar = String(headerData.destination || headerData['PELABUHAN BONGKAR'] || '').trim().toUpperCase();

  const validDetil: any[] = [];
  let transshipmentCount = 0;

  detilRaw.forEach((det) => {
    let kdPos = String(det['KD KELOMPOK POS'] || '').trim();
    if (kdPos.endsWith('.0')) kdPos = kdPos.slice(0, -2);
    if (kdPos.length === 1) kdPos = '0' + kdPos;

    const detBongkar = String(det['PELABUHAN BONGKAR'] || '').trim().toUpperCase();

    const isTransshipment =
      kdPos === '02' ||
      (headerBongkar && detBongkar && detBongkar !== headerBongkar);

    if (isTransshipment) {
      transshipmentCount++;
    } else {
      validDetil.push(det);
    }
  });

  const rows: DetilRow[] = validDetil.map((det, index) => {
    const idx = index + 1;

    // Act Pcs dari JUMLAH KEMASAN
    let actPcs: number | null = null;
    const rawAct = det['JUMLAH KEMASAN'];
    if (rawAct !== undefined && rawAct !== null && String(rawAct).trim() !== '') {
      const n = Number(rawAct);
      if (!isNaN(n)) actPcs = Math.floor(n);
    }

    // Tot Pcs dari TOTAL KEMASAN (fallback ke JUMLAH KEMASAN)
    let totPcs: number | null = null;
    let rawTot = det['TOTAL KEMASAN'];
    if (rawTot === undefined || rawTot === null || String(rawTot).trim() === '') {
      rawTot = det['JUMLAH KEMASAN'];
    }
    if (rawTot !== undefined && rawTot !== null && String(rawTot).trim() !== '') {
      const n = Number(rawTot);
      if (!isNaN(n)) totPcs = Math.floor(n);
    }

    // Tot Weight dari BRUTO
    let totWeight = '';
    if (det['BRUTO'] !== undefined && det['BRUTO'] !== null) {
      totWeight = String(det['BRUTO']).trim();
    }

    const natureOfGoods = getNatureOfGoods(det['ID DETIL'], barangRaw, 50);

    const origin = String(det['PELABUHAN ASAL'] || headerData.origin || '').trim();
    const destination = String(det['PELABUHAN BONGKAR'] || headerData.destination || '').trim();

    return {
      no: idx,
      mawb: String(det['NO MASTER BLAWB'] || '').trim(),
      hawb: String(det['NO HOST BLAWB'] || '').trim(),
      actPcs,
      totPcs,
      totWeight,
      natureOfGoods,
      bc11: headerData.noBc11 || '',
      tanggalBc11: headerData.tglBc11 instanceof Date ? headerData.tglBc11 : null,
      flightNumber: headerData.flightNumber || '',
      flightDate: headerData.flightDate instanceof Date ? headerData.flightDate : null,
      origin,
      destination,
      shipper: String(det['NAMA SHIPPER'] || '').trim(),
      alamatShipper: String(det['ALMT SHIPPER'] || '').trim(),
      consignee: String(det['NAMA CONSIGNEE'] || '').trim(),
      alamatConsignee: String(det['ALMT CONSIGNEE'] || '').trim(),
      idDetil: String(det['ID DETIL'] || ''),
    };
  });

  return {
    header: headerData,
    rows,
    totalRecords: rows.length,
    transshipmentCount,
    formatType: 'BC11',
    fileName: file.name,
  };
}

export async function generateExcelBlob(conversion: ConversionResult): Promise<Blob> {
  const workbook = new ExcelJS.Workbook();
  const ws = workbook.addWorksheet('Sheet1', {
    views: [{ showGridLines: true }],
  });

  // Baris 1: Judul
  ws.mergeCells('A1:V1');
  const titleCell = ws.getCell('A1');
  titleCell.value = 'Breakdown Report Cargo International';
  titleCell.font = { name: 'Aptos Narrow', size: 12, bold: true };
  titleCell.alignment = { horizontal: 'center', vertical: 'middle' };

  // Baris 2: Header Kolom
  const headerRow = ws.getRow(2);
  OUTPUT_HEADERS.forEach((h, i) => {
    const colIdx = i + 1;
    const cell = headerRow.getCell(colIdx);
    cell.value = h;
    cell.font = { name: 'Tahoma', size: 8, bold: false };
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF708090' }, // Slate gray
    };
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  });

  // Lebar kolom
  Object.entries(COLUMN_WIDTHS).forEach(([letter, width]) => {
    ws.getColumn(letter).width = width;
  });

  // Tulis baris data
  conversion.rows.forEach((row, i) => {
    const rowNum = i + 3;
    const r = ws.getRow(rowNum);

    const values = [
      row.no, // 1: No
      row.mawb, // 2: MAWB
      row.hawb, // 3: HAWB
      row.actPcs, // 4: Act Pcs
      null, // 5: of
      row.totPcs, // 6: Tot Pcs
      row.actWeight || null, // 7: Act Weight
      null, // 8: of
      row.totWeight, // 9: Tot Weight
      row.natureOfGoods, // 10: Nature of Goods
      null, // 11: Actual Times
      null, // 12: Remarks
      null, // 13: Status
      row.bc11 || null, // 14: BC11
      row.tanggalBc11, // 15: TanggalBC11
      row.flightNumber, // 16: FlightNumber
      row.flightDate, // 17: FlightDate
      row.origin, // 18: Origin
      row.destination, // 19: Destination
      row.shipper, // 20: Shipper
      row.alamatShipper, // 21: AlamatShipper
      row.consignee, // 22: Consignee
      row.alamatConsignee, // 23: AlamatConsignee
    ];

    values.forEach((v, colIdx0) => {
      const colIdx = colIdx0 + 1;
      const cell = r.getCell(colIdx);
      cell.value = v;
      cell.alignment = { vertical: 'top', wrapText: true };

      if (colIdx === 15 || colIdx === 17) {
        cell.font = { name: 'Arial', size: 10 };
        if (v instanceof Date) {
          cell.numFmt = 'mm-dd-yy';
        }
      } else if (colIdx === 14) {
        cell.font = { name: 'Arial', size: 10 };
      } else {
        cell.font = { name: 'Aptos Narrow', size: 11 };
      }
    });
  });

  const buffer = await workbook.xlsx.writeBuffer();
  return new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
}

