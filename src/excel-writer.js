const XLSX = require('xlsx');
const path = require('path');
const { KISO_COLUMNS, MOKUTEKICHI_COLUMNS } = require('./fields');

const KISO_SHEET_NAME = '【基礎】案件管理表 ';
const MOKUTEKICHI_SHEET_NAME = '【目的地】案件管理表';
const HEADER_ROW_COUNT = 2; // ヘッダーは2行

function getNextRowNumber(ws) {
  const ref = ws['!ref'];
  if (!ref) return HEADER_ROW_COUNT + 1;
  const range = XLSX.utils.decode_range(ref);
  // データ行を探索して最後の行を見つける
  let lastDataRow = HEADER_ROW_COUNT;
  for (let r = HEADER_ROW_COUNT; r <= range.e.r; r++) {
    const cellA = ws[XLSX.utils.encode_cell({ r, c: 0 })];
    if (cellA && cellA.v !== undefined && cellA.v !== '') {
      lastDataRow = r;
    }
  }
  return lastDataRow + 1;
}

function getNextNo(ws) {
  const ref = ws['!ref'];
  if (!ref) return 1;
  const range = XLSX.utils.decode_range(ref);
  let maxNo = 0;
  for (let r = HEADER_ROW_COUNT; r <= range.e.r; r++) {
    const cell = ws[XLSX.utils.encode_cell({ r, c: 0 })];
    if (cell && typeof cell.v === 'number' && cell.v > maxNo) {
      maxNo = cell.v;
    }
  }
  return maxNo + 1;
}

function setCellValue(ws, row, col, value) {
  if (value === undefined || value === null || value === '') return;
  const cellRef = XLSX.utils.encode_cell({ r: row, c: col });
  ws[cellRef] = { t: typeof value === 'number' ? 'n' : 's', v: value };
}

function updateRange(ws, row, col) {
  if (!ws['!ref']) {
    ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: row, c: col } });
    return;
  }
  const range = XLSX.utils.decode_range(ws['!ref']);
  if (row > range.e.r) range.e.r = row;
  if (col > range.e.c) range.e.c = col;
  ws['!ref'] = XLSX.utils.encode_range(range);
}

function writeKisoData(excelPath, data) {
  const wb = XLSX.readFile(excelPath);
  const ws = wb.Sheets[KISO_SHEET_NAME];
  if (!ws) throw new Error(`シート "${KISO_SHEET_NAME}" が見つかりません`);

  const row = getNextRowNumber(ws);
  const no = getNextNo(ws);
  const cols = KISO_COLUMNS;

  // マンション種別マッピング
  const mansionTypeMap = {
    '(賃貸)マンション・アパート': '賃貸',
    '(分譲)マンション': '分譲',
    '月極駐車場': '月極駐車場',
  };

  // kWタイプマッピング
  const kwMap = {
    '6kW普通充電器': '6kW',
    '3kWコンセント': '3kW',
  };

  // プランマッピング
  const planMap = {
    '6kW普通充電器': 'マンションゼロプラン',
    '3kWコンセント': 'マンションゼロプラン',
  };

  setCellValue(ws, row, cols.no, no);
  setCellValue(ws, row, cols.propertyName, data.facilityName);
  setCellValue(ws, row, cols.address, data.facilityAddress);
  setCellValue(ws, row, cols.mansionType, mansionTypeMap[data.industry] || data.industry);
  setCellValue(ws, row, cols.landOwner, data.landOwner || '');
  setCellValue(ws, row, cols.applicationCount, parseInt(data.applicationCount) || data.applicationCount);
  setCellValue(ws, row, cols.parkingCount, parseInt(data.parkingCount) || data.parkingCount);
  setCellValue(ws, row, cols.applicationPeriod, data.applicationPeriod);
  setCellValue(ws, row, cols.units, parseInt(data.mansionUnits) || data.mansionUnits);
  setCellValue(ws, row, cols.kwNumber, kwMap[data.kwType] || data.kwType);
  setCellValue(ws, row, cols.plan, planMap[data.kwType] || 'マンションゼロプラン');
  setCellValue(ws, row, cols.newOrExisting, '既築');
  setCellValue(ws, row, cols.companyName, data.companyName);
  setCellValue(ws, row, cols.salesPerson, data.salesPerson);

  // 取引先住所
  if (data.companyAddress) {
    setCellValue(ws, row, cols.companyAddress, data.companyAddress);
  }

  // クラウドサイン送付先
  if (data.contactCompany) {
    setCellValue(ws, row, cols.csCompany, data.contactCompany);
  }
  if (data.contactName) {
    setCellValue(ws, row, cols.csPersonName, data.contactName);
  }
  if (data.contactEmail) {
    setCellValue(ws, row, cols.csEmail, data.contactEmail);
  }

  // リード獲得パートナー
  if (data.accompanyType === 'ストア単独') {
    setCellValue(ws, row, cols.leadPartner, 'ストア');
  } else if (data.referralSource) {
    setCellValue(ws, row, cols.leadPartner, data.referralSource);
  }

  // 範囲を更新
  const maxCol = Math.max(...Object.values(cols));
  updateRange(ws, row, maxCol);

  XLSX.writeFile(wb, excelPath);
  return { sheetName: KISO_SHEET_NAME, row: row + 1, no };
}

function writeMokutekichiData(excelPath, data) {
  const wb = XLSX.readFile(excelPath);
  const ws = wb.Sheets[MOKUTEKICHI_SHEET_NAME];
  if (!ws) throw new Error(`シート "${MOKUTEKICHI_SHEET_NAME}" が見つかりません`);

  const row = getNextRowNumber(ws);
  const no = getNextNo(ws);
  const cols = MOKUTEKICHI_COLUMNS;

  // kWタイプマッピング
  const kwMap = {
    '6kW普通充電器': '6kW',
    '3kWコンセント': '3kW',
  };

  setCellValue(ws, row, cols.no, no);
  setCellValue(ws, row, cols.facilityName, data.facilityName);
  setCellValue(ws, row, cols.address, data.facilityAddress);
  setCellValue(ws, row, cols.industry, data.industryOther || data.industry);
  setCellValue(ws, row, cols.landOwner, data.landOwner || '');
  setCellValue(ws, row, cols.applicationCount, parseInt(data.applicationCount) || data.applicationCount);
  setCellValue(ws, row, cols.parkingCount, parseInt(data.parkingCount) || data.parkingCount);
  setCellValue(ws, row, cols.applicationPeriod, data.applicationPeriod);
  setCellValue(ws, row, cols.kwNumber, kwMap[data.kwType] || data.kwType);
  setCellValue(ws, row, cols.plan, data.contractPlan || 'プライムゼロプラン');
  setCellValue(ws, row, cols.companyName, data.companyName);
  setCellValue(ws, row, cols.companyAddress, data.companyAddress || '');
  setCellValue(ws, row, cols.salesPerson, data.salesPerson);

  // クラウドサイン送付先
  if (data.contactCompany) {
    setCellValue(ws, row, cols.csCompany, data.contactCompany);
  }
  if (data.contactName) {
    setCellValue(ws, row, cols.csPersonName, data.contactName);
  }
  if (data.contactEmail) {
    setCellValue(ws, row, cols.csEmail, data.contactEmail);
  }

  // リード獲得パートナー
  if (data.accompanyType === 'ストア単独') {
    setCellValue(ws, row, cols.leadPartner, 'ストア');
  } else if (data.referralSource) {
    setCellValue(ws, row, cols.leadPartner, data.referralSource);
  }

  // 範囲を更新
  const maxCol = Math.max(...Object.values(cols));
  updateRange(ws, row, maxCol);

  XLSX.writeFile(wb, excelPath);
  return { sheetName: MOKUTEKICHI_SHEET_NAME, row: row + 1, no };
}

module.exports = { writeKisoData, writeMokutekichiData };
