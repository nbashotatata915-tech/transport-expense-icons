const ICON_URL = 'https://raw.githubusercontent.com/nbashotatata915-tech/transport-expense-icons2/main/7E090507-5563-4FC4-BAF7-468B70D55772.png';

const SHEET_NAMES = {
  responses: 'フォーム回答',
  drivers: 'ドライバーマスタ',
  places: '場所マスタ',
  gas: 'ガソリン単価マスタ',
  results: '精算結果',
  summary: '月別集計',
  errors: 'エラーログ',
  statusHistory: 'ステータス変更履歴',
};

const APPLICATION_TYPES = {
  pickup: '新入生迎え',
  luggage: 'リーグ戦・荷物車出し',
};

const DISTANCE_TYPES = {
  roundTrip: '往復扱い',
  oneWay: '片道扱い',
};

const INITIAL_APPLICATION_STATUS = '審査中';
const STATUS_APPROVED = '承認済';
const STATUS_REJECTED = '差し戻し';

const INITIAL_PAYMENT_STATUS = '未払い';
const STATUS_PAID = '支払済';

const HISTORY_KIND_STATUS = 'ステータス';
const HISTORY_KIND_PAYMENT = '支払ステータス';

function getSpreadsheet_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function doGet(e) {
  const page = e && e.parameter && e.parameter.page;

  if (page === 'admin') {
    return HtmlService
      .createHtmlOutputFromFile('Admin')
      .setTitle('交通費精算 管理画面')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
      .setFaviconUrl(ICON_URL)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  return HtmlService
    .createHtmlOutputFromFile('Index')
    .setTitle('交通費申請フォーム')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setFaviconUrl(ICON_URL)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('交通費精算')
    .addItem('月別集計を更新', 'updateMonthlySummary')
    .addItem('精算結果を全件再作成', 'rebuildAllResults')
    .addToUi();
}

/**
 * Googleフォーム送信時用
 */
function onFormSubmit(e) {
  const ss = getSpreadsheet_();
  const values = e.values;

  const row = {
    timestamp: values[0],
    date: values[1],
    name: normalizeName_(values[2]),
    applicationType: values[3],
    pickupPlace: values[4],
    distanceType: values[5],
    gym: values[6],
  };

  const meta = createInitialApplicationMeta_();
  const formAppendedRow = e && e.range && typeof e.range.getRow === 'function'
    ? e.range.getRow()
    : null;
  fillFormResponseMeta_(ss, meta, formAppendedRow);

  const result = buildSettlementResult_(ss, row);
  appendSettlementResult_(ss, result, meta.applicationId);
}

/**
 * Webアプリ画面表示時に、プルダウン用データを取得
 */
function getWebAppInitialData() {
  const ss = getSpreadsheet_();

  return {
    drivers: getDriverNames_(ss),
    pickupPlaces: getPlaceNamesByType_(ss, '新入生迎え場所'),
    gyms: getPlaceNamesByType_(ss, '体育館'),
  };
}

/**
 * Webアプリから申請されたときの処理
 */
function submitWebApplication(data) {
  const ss = getSpreadsheet_();

  const row = {
    timestamp: new Date(),
    date: data.date,
    name: normalizeName_(data.name),
    applicationType: data.applicationType,
    pickupPlace: data.pickupPlace,
    distanceType: data.distanceType,
    gym: data.gym,
  };

  const meta = createInitialApplicationMeta_();

  const result = buildSettlementResult_(ss, row);

  appendWebApplicationToResponseSheet_(ss, row, meta);
  appendSettlementResult_(ss, result, meta.applicationId);

  return {
    applicationId: meta.applicationId,
    status: result.status,
    payment: result.payment,
    targetKm: result.targetKm,
    errorMessage: result.errorMessage,
  };
}

/**
 * 精算結果を作る共通処理
 */
function buildSettlementResult_(ss, row) {
  const errors = [];

  const month = formatMonth_(row.date);
  if (!month) errors.push('日付が正しくありません');

  const driver = findDriver_(ss, row.name);
  if (!driver) errors.push(`ドライバーマスタに氏名「${row.name}」がありません`);

  let placeName = '';
  let distanceTypeLabel = '';
  let multiplier = 0;

  if (row.applicationType === APPLICATION_TYPES.pickup) {
    placeName = row.pickupPlace;
    distanceTypeLabel = row.distanceType;

    if (!placeName) errors.push('新入生迎えですが、迎え場所が未入力です');
    if (!distanceTypeLabel) errors.push('新入生迎えですが、距離区分が未入力です');

    if (distanceTypeLabel === DISTANCE_TYPES.roundTrip) {
      multiplier = 2;
    } else if (distanceTypeLabel === DISTANCE_TYPES.oneWay) {
      multiplier = 1;
    } else {
      errors.push(`距離区分「${distanceTypeLabel}」が不正です`);
    }
  } else if (row.applicationType === APPLICATION_TYPES.luggage) {
    placeName = row.gym;
    distanceTypeLabel = '往復固定';
    multiplier = 2;

    if (!placeName) errors.push('リーグ戦・荷物車出しですが、実施体育館が未入力です');
  } else {
    errors.push(`申請種別「${row.applicationType}」が不正です`);
  }

  const place = placeName ? findPlace_(ss, placeName) : null;
  if (placeName && !place) errors.push(`場所マスタに「${placeName}」がありません`);

  const gasPrice = driver && month ? findGasPrice_(ss, month, driver.fuelType) : null;
  if (driver && month && gasPrice === null) {
    errors.push(`ガソリン単価マスタに「${month}」の「${driver.fuelType}」価格がありません`);
  }

  if (driver && (!driver.fuelEfficiency || driver.fuelEfficiency <= 0)) {
    errors.push(`氏名「${row.name}」の想定燃費が未入力または不正です`);
  }

  const oneWayKm = place ? place.oneWayKm : 0;
  const targetKm = oneWayKm * multiplier;

  let payment = '';
  if (errors.length === 0) {
    payment = roundToNearestYen_(targetKm * gasPrice / driver.fuelEfficiency);
  }

  return {
    processedAt: new Date(),
    applicationDate: row.date,
    month,
    name: row.name,
    studentId: driver ? driver.studentId : '',
    applicationType: row.applicationType,
    place: placeName,
    distanceType: distanceTypeLabel,
    targetKm,
    fuelType: driver ? driver.fuelType : '',
    gasPrice: gasPrice !== null ? gasPrice : '',
    fuelEfficiency: driver ? driver.fuelEfficiency : '',
    payment,
    status: errors.length === 0 ? 'OK' : '要確認',
    errorMessage: errors.join(' / '),
  };
}

/**
 * 精算結果シートに追加
 */
function appendSettlementResult_(ss, result, applicationId) {
  const sheet = ss.getSheetByName(SHEET_NAMES.results);
  if (!sheet) throw new Error(`シート「${SHEET_NAMES.results}」がありません`);

  sheet.appendRow([
    result.processedAt,
    result.applicationDate,
    result.month,
    result.name,
    result.studentId,
    result.applicationType,
    result.place,
    result.distanceType,
    result.targetKm,
    result.fuelType,
    result.gasPrice,
    result.fuelEfficiency,
    result.payment,
    result.status,
    result.errorMessage,
    applicationId || '',
  ]);

  if (result.status !== 'OK') {
    appendErrorLog_(ss, result);
  }
}

/**
 * エラーログに追加
 */
function appendErrorLog_(ss, result) {
  const sheet = ss.getSheetByName(SHEET_NAMES.errors);
  if (!sheet) return;

  sheet.appendRow([
    new Date(),
    result.applicationDate,
    result.name,
    result.studentId,
    result.applicationType,
    result.errorMessage,
  ]);
}

/**
 * Webアプリからの申請をフォーム回答シートにも記録
 */
function appendWebApplicationToResponseSheet_(ss, row, meta) {
  const sheet = ss.getSheetByName(SHEET_NAMES.responses);
  if (!sheet) throw new Error(`シート「${SHEET_NAMES.responses}」がありません`);

  sheet.appendRow([
    row.timestamp,
    row.date,
    row.name,
    row.applicationType,
    row.pickupPlace,
    row.distanceType,
    row.gym,
    meta.applicationId,
    meta.status,
    meta.rejectReason,
    meta.approvedAt,
    meta.paymentStatus,
  ]);
}

/**
 * 新規申請の初期メタ情報（申請ID・ステータス等）を作成
 */
function createInitialApplicationMeta_() {
  return {
    applicationId: generateApplicationId_(),
    status: INITIAL_APPLICATION_STATUS,
    rejectReason: '',
    approvedAt: '',
    paymentStatus: INITIAL_PAYMENT_STATUS,
  };
}

/**
 * 申請IDを発行：APP-YYYYMMDD-HHmmss-XXXX
 */
function generateApplicationId_() {
  const tz = Session.getScriptTimeZone();
  const datePart = Utilities.formatDate(new Date(), tz, 'yyyyMMdd-HHmmss');
  const randomPart = Math.random().toString(36).slice(2, 6).toUpperCase().padEnd(4, '0');
  return `APP-${datePart}-${randomPart}`;
}

/**
 * フォーム回答シートのうち、Googleフォームが直前に追記した行へ
 * 申請ID・ステータス等のメタ情報を書き込む
 */
function fillFormResponseMeta_(ss, meta, rowIndex) {
  const sheet = ss.getSheetByName(SHEET_NAMES.responses);
  if (!sheet) throw new Error(`シート「${SHEET_NAMES.responses}」がありません`);

  const targetRow = rowIndex || sheet.getLastRow();
  if (targetRow < 2) return;

  sheet.getRange(targetRow, 8, 1, 5).setValues([[
    meta.applicationId,
    meta.status,
    meta.rejectReason,
    meta.approvedAt,
    meta.paymentStatus,
  ]]);
}

/**
 * 月別集計を更新
 */
function updateMonthlySummary() {
  const ss = getSpreadsheet_();
  const resultSheet = ss.getSheetByName(SHEET_NAMES.results);
  const summarySheet = ss.getSheetByName(SHEET_NAMES.summary);

  if (!resultSheet) throw new Error(`シート「${SHEET_NAMES.results}」がありません`);
  if (!summarySheet) throw new Error(`シート「${SHEET_NAMES.summary}」がありません`);

  const values = resultSheet.getDataRange().getValues();
  const rows = values.slice(1);

  const summary = new Map();

  rows.forEach(row => {
    const month = normalizeMonth_(row[2]);
    const name = normalizeName_(row[3]);
    const studentId = row[4];
    const targetKm = Number(row[8]) || 0;
    const payment = Number(row[12]) || 0;
    const status = row[13];

    if (status !== 'OK') return;
    if (!month || !name) return;

    const key = `${month}__${name}`;

    if (!summary.has(key)) {
      summary.set(key, {
        month,
        name,
        studentId,
        totalKm: 0,
        totalPayment: 0,
      });
    }

    const item = summary.get(key);
    item.totalKm += targetKm;
    item.totalPayment += payment;
  });

  summarySheet.clearContents();
  summarySheet.appendRow(['月', '氏名', '学籍番号', '合計対象距離km', '支給額合計']);

  const output = Array.from(summary.values())
    .sort((a, b) => {
      if (a.month !== b.month) return String(a.month).localeCompare(String(b.month));
      return String(a.name).localeCompare(String(b.name));
    })
    .map(item => [
      item.month,
      item.name,
      item.studentId,
      roundToOneDecimal_(item.totalKm),
      item.totalPayment,
    ]);

  if (output.length > 0) {
    summarySheet.getRange(2, 1, output.length, output[0].length).setValues(output);
  }
}

/**
 * 全フォーム回答から精算結果を再作成
 */
function rebuildAllResults() {
  const ss = getSpreadsheet_();
  const responseSheet = ss.getSheetByName(SHEET_NAMES.responses);
  const resultSheet = ss.getSheetByName(SHEET_NAMES.results);
  const errorSheet = ss.getSheetByName(SHEET_NAMES.errors);

  if (!responseSheet) throw new Error(`シート「${SHEET_NAMES.responses}」がありません`);
  if (!resultSheet) throw new Error(`シート「${SHEET_NAMES.results}」がありません`);

  resultSheet.clearContents();
  resultSheet.appendRow([
    '処理日時',
    '申請日',
    '月',
    '氏名',
    '学籍番号',
    '申請種別',
    '場所',
    '距離区分',
    '対象距離km',
    '燃料種別',
    'ガソリン単価',
    '想定燃費',
    '支給額',
    'ステータス',
    'エラー内容',
    '申請ID',
  ]);

  if (errorSheet) {
    errorSheet.clearContents();
    errorSheet.appendRow([
      '記録日時',
      '申請日',
      '氏名',
      '学籍番号',
      '申請種別',
      'エラー内容',
    ]);
  }

  const values = responseSheet.getDataRange().getValues();
  const rows = values.slice(1);

  rows.forEach((values, rowIndex) => {
    const row = {
      timestamp: values[0],
      date: values[1],
      name: normalizeName_(values[2]),
      applicationType: values[3],
      pickupPlace: values[4],
      distanceType: values[5],
      gym: values[6],
    };

    let applicationId = String(values[7] || '').trim();
    if (!applicationId) {
      applicationId = generateApplicationId_();
      responseSheet.getRange(rowIndex + 2, 8).setValue(applicationId);
    }

    const result = buildSettlementResult_(ss, row);
    appendSettlementResult_(ss, result, applicationId);
  });

  updateMonthlySummary();
}

/**
 * ドライバーマスタから氏名一覧を取得
 */
function getDriverNames_(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.drivers);
  if (!sheet) throw new Error(`シート「${SHEET_NAMES.drivers}」がありません`);

  const values = sheet.getDataRange().getValues();

  return values
    .slice(1)
    .map(row => normalizeName_(row[0]))
    .filter(name => name);
}

/**
 * 場所マスタから種別ごとの場所一覧を取得
 */
function getPlaceNamesByType_(ss, type) {
  const sheet = ss.getSheetByName(SHEET_NAMES.places);
  if (!sheet) throw new Error(`シート「${SHEET_NAMES.places}」がありません`);

  const values = sheet.getDataRange().getValues();

  return values
    .slice(1)
    .filter(row => String(row[1]).trim() === type)
    .map(row => String(row[0]).trim())
    .filter(name => name);
}

/**
 * 氏名からドライバー情報を取得
 */
function findDriver_(ss, name) {
  const sheet = ss.getSheetByName(SHEET_NAMES.drivers);
  if (!sheet) throw new Error(`シート「${SHEET_NAMES.drivers}」がありません`);

  const values = sheet.getDataRange().getValues();
  const rows = values.slice(1);
  const targetName = normalizeName_(name);

  for (const row of rows) {
    if (normalizeName_(row[0]) === targetName) {
      return {
        name: normalizeName_(row[0]),
        studentId: row[1],
        carType: row[2],
        fuelType: String(row[3] || '').trim(),
        fuelEfficiency: Number(row[4]),
      };
    }
  }

  return null;
}

/**
 * 場所名から距離情報を取得
 */
function findPlace_(ss, placeName) {
  const sheet = ss.getSheetByName(SHEET_NAMES.places);
  if (!sheet) throw new Error(`シート「${SHEET_NAMES.places}」がありません`);

  const values = sheet.getDataRange().getValues();
  const rows = values.slice(1);

  for (const row of rows) {
    if (String(row[0]).trim() === String(placeName).trim()) {
      return {
        name: row[0],
        type: row[1],
        oneWayKm: Number(row[2]),
      };
    }
  }

  return null;
}

/**
 * 月と燃料種別からガソリン単価を取得
 *
 * ガソリン単価マスタの想定列：
 * A列：月
 * B列：レギュラー価格
 * C列：ハイオク価格
 * D列：軽油価格
 *
 * ドライバーマスタの燃料種別：
 * レギュラー / ハイオク / 軽油
 */
function findGasPrice_(ss, month, fuelType) {
  const sheet = ss.getSheetByName(SHEET_NAMES.gas);
  if (!sheet) throw new Error(`シート「${SHEET_NAMES.gas}」がありません`);

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return null;

  const header = values[0].map(value => String(value || '').trim());
  const rows = values.slice(1);

  const fuelTypeText = String(fuelType || '').trim();

  const monthCol = findHeaderIndex_(header, ['月', '対象月']);
  const regularCol = findHeaderIndex_(header, ['レギュラー価格', 'レギュラー']);
  const premiumCol = findHeaderIndex_(header, ['ハイオク価格', 'ハイオク']);
  const dieselCol = findHeaderIndex_(header, ['軽油価格', '軽油']);

  const actualMonthCol = monthCol >= 0 ? monthCol : 0;
  const actualRegularCol = regularCol >= 0 ? regularCol : 1;
  const actualPremiumCol = premiumCol >= 0 ? premiumCol : 2;
  const actualDieselCol = dieselCol >= 0 ? dieselCol : 3;

  for (const row of rows) {
    const rowMonth = normalizeMonth_(row[actualMonthCol]);

    if (rowMonth === month) {
      if (fuelTypeText === 'レギュラー') return parsePriceOrNull_(row[actualRegularCol]);
      if (fuelTypeText === 'ハイオク') return parsePriceOrNull_(row[actualPremiumCol]);
      if (fuelTypeText === '軽油') return parsePriceOrNull_(row[actualDieselCol]);
      return null;
    }
  }

  return null;
}

/**
 * ヘッダー名から列番号を探す
 */
function findHeaderIndex_(header, candidates) {
  for (const candidate of candidates) {
    const index = header.findIndex(name => String(name || '').trim() === candidate);
    if (index >= 0) return index;
  }

  return -1;
}

/**
 * 金額・単価を数値化する
 */
function parsePriceOrNull_(value) {
  if (value === null || value === undefined || value === '') return null;

  const number = Number(String(value).replace(/,/g, '').trim());

  if (!number || number <= 0) return null;

  return number;
}

/**
 * 日付から YYYY-MM を作成
 */
function formatMonth_(dateValue) {
  if (!dateValue) return '';

  let date;

  if (Object.prototype.toString.call(dateValue) === '[object Date]') {
    date = dateValue;
  } else {
    date = new Date(dateValue);
  }

  if (isNaN(date.getTime())) return '';

  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');

  return `${year}-${month}`;
}

/**
 * ガソリン単価マスタや集計シートの月表記を YYYY-MM に正規化
 */
function normalizeMonth_(value) {
  if (!value) return '';

  if (Object.prototype.toString.call(value) === '[object Date]') {
    return formatMonth_(value);
  }

  const text = String(value).trim();

  if (/^\d{4}-\d{2}$/.test(text)) {
    return text;
  }

  const date = new Date(text);
  if (!isNaN(date.getTime())) {
    return formatMonth_(date);
  }

  return text;
}

/**
 * 氏名比較用：空白や全角スペースを整える
 */
function normalizeName_(value) {
  return String(value || '')
    .replace(/\u3000/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * 支給額の端数処理：1円単位四捨五入
 */
function roundToNearestYen_(value) {
  return Math.round(Number(value));
}

/**
 * 距離表示用：小数1桁
 */
function roundToOneDecimal_(value) {
  return Math.round(Number(value) * 10) / 10;
}

/**
 * 初期データ読み込み確認用
 */
function testGetInitialData() {
  const data = getWebAppInitialData();
  console.log(JSON.stringify(data));
}

/**
 * 初期データをエラーログに書き出す確認用
 */
function testGetInitialDataToCell() {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(SHEET_NAMES.errors) || ss.insertSheet(SHEET_NAMES.errors);

  try {
    sheet.getRange('H1').setValue('初期データテスト');

    const data = getWebAppInitialData();

    sheet.getRange('H2').setValue(JSON.stringify(data));
    sheet.getRange('H3').setValue('成功');
    sheet.getRange('H4').clearContent();
  } catch (error) {
    sheet.getRange('H1').setValue('初期データテスト');
    sheet.getRange('H2').setValue('エラー発生');
    sheet.getRange('H3').setValue(error.message);
    sheet.getRange('H4').setValue(error.stack);
  }
}

/**
 * 接続確認用
 */
function testBoundSpreadsheet() {
  const ss = getSpreadsheet_();

  if (!ss) {
    throw new Error('このApps Scriptはスプレッドシートに紐づいていません');
  }

  const sheet = ss.getSheetByName(SHEET_NAMES.errors) || ss.insertSheet(SHEET_NAMES.errors);
  sheet.getRange('H1').setValue('接続テスト');
  sheet.getRange('H2').setValue(ss.getName());
  sheet.getRange('H3').setValue(ss.getSheets().map(s => s.getName()).join(', '));
}

/**
 * 管理画面用：月別集計を取得
 */
function getMonthlySummaryForWeb(month) {
  updateMonthlySummary();

  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(SHEET_NAMES.summary);

  if (!sheet) {
    throw new Error(`シート「${SHEET_NAMES.summary}」がありません`);
  }

  const values = sheet.getDataRange().getValues();
  const rows = values.slice(1);

  const targetMonth = normalizeMonth_(month);

  return rows
    .filter(row => normalizeMonth_(row[0]) === targetMonth)
    .map(row => ({
      month: normalizeMonth_(row[0]),
      name: normalizeName_(row[1]),
      studentId: row[2],
      totalKm: row[3],
      totalPayment: row[4],
    }));
}

/**
 * 管理画面用：月の候補を取得
 */
function getAvailableMonthsForWeb() {
  const ss = getSpreadsheet_();
  const resultSheet = ss.getSheetByName(SHEET_NAMES.results);

  if (!resultSheet) {
    throw new Error(`シート「${SHEET_NAMES.results}」がありません`);
  }

  const values = resultSheet.getDataRange().getDisplayValues();
  const rows = values.slice(1);

  const months = rows
    .map(row => normalizeMonth_(row[2]))
    .filter(month => month);

  return [...new Set(months)].sort().reverse();
}

/**
 * 管理画面用：部員別申請履歴を取得
 * getDisplayValues() を使い、シート上の表示文字列で照合する
 */
function getApplicationHistoryForWeb(month, name) {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(SHEET_NAMES.results);

  if (!sheet) {
    throw new Error(`シート「${SHEET_NAMES.results}」がありません`);
  }

  const values = sheet.getDataRange().getDisplayValues();
  const rows = values.slice(1);

  const targetMonth = normalizeMonth_(month);
  const targetName = normalizeName_(name);

  return rows
    .filter(row => {
      const rowMonth = normalizeMonth_(row[2]);
      const rowName = normalizeName_(row[3]);

      return rowMonth === targetMonth && rowName === targetName;
    })
    .map(row => ({
      processedAt: row[0],
      applicationDate: row[1],
      month: normalizeMonth_(row[2]),
      name: normalizeName_(row[3]),
      studentId: row[4],
      applicationType: row[5],
      place: row[6],
      distanceType: row[7],
      targetKm: row[8],
      fuelType: row[9],
      gasPrice: row[10],
      fuelEfficiency: row[11],
      payment: row[12],
      status: row[13],
      errorMessage: row[14],
    }));
}

/**
 * 管理画面用：申請履歴の氏名候補を取得
 * getDisplayValues() を使い、表示上の月・氏名で取得する
 */
function getHistoryNamesForWeb(month) {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(SHEET_NAMES.results);

  if (!sheet) {
    throw new Error(`シート「${SHEET_NAMES.results}」がありません`);
  }

  const values = sheet.getDataRange().getDisplayValues();
  const rows = values.slice(1);

  const targetMonth = normalizeMonth_(month);

  const names = rows
    .filter(row => normalizeMonth_(row[2]) === targetMonth)
    .map(row => normalizeName_(row[3]))
    .filter(name => name);

  return [...new Set(names)].sort();
}

/**
 * 管理画面用：審査中の申請一覧を取得
 *
 * フォーム回答シートのステータスが「審査中」の行を抽出し、
 * 同じ申請IDを持つ精算結果から対象距離・支給額・計算ステータスを付与する
 */
function getPendingApplicationsForWeb() {
  const ss = getSpreadsheet_();
  const responseSheet = ss.getSheetByName(SHEET_NAMES.responses);
  const resultSheet = ss.getSheetByName(SHEET_NAMES.results);

  if (!responseSheet) throw new Error(`シート「${SHEET_NAMES.responses}」がありません`);
  if (!resultSheet) throw new Error(`シート「${SHEET_NAMES.results}」がありません`);

  const resultMap = new Map();
  const resultValues = resultSheet.getDataRange().getValues();
  resultValues.slice(1).forEach(row => {
    const id = String(row[15] || '').trim();
    if (id) {
      resultMap.set(id, {
        place: row[6],
        distanceType: row[7],
        targetKm: row[8],
        payment: row[12],
        calcStatus: row[13],
        errorMessage: row[14],
      });
    }
  });

  const responseValues = responseSheet.getDataRange().getValues();
  const pending = [];

  responseValues.slice(1).forEach(row => {
    const status = String(row[8] || '').trim();
    if (status !== INITIAL_APPLICATION_STATUS) return;

    const applicationId = String(row[7] || '').trim();
    if (!applicationId) return;

    const result = resultMap.get(applicationId);

    pending.push({
      applicationId,
      applicationDate: formatDateForDisplay_(row[1]),
      name: normalizeName_(row[2]),
      applicationType: row[3],
      place: result ? result.place : (row[3] === APPLICATION_TYPES.pickup ? row[4] : row[6]),
      distanceType: result ? result.distanceType : row[5],
      targetKm: result ? result.targetKm : '',
      payment: result ? result.payment : '',
      calcStatus: result ? result.calcStatus : '',
      calcError: result ? result.errorMessage : '',
    });
  });

  return pending;
}

/**
 * 管理画面用：複数の申請を一括で承認
 *
 * 戻り値：{ successCount, errors: [{ applicationId, message }] }
 */
function approveApplications(applicationIds) {
  if (!Array.isArray(applicationIds) || applicationIds.length === 0) {
    return { successCount: 0, errors: [] };
  }

  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(SHEET_NAMES.responses);
  if (!sheet) throw new Error(`シート「${SHEET_NAMES.responses}」がありません`);

  const targetIds = new Set(
    applicationIds.map(id => String(id || '').trim()).filter(id => id)
  );
  const values = sheet.getDataRange().getValues();
  const now = new Date();

  let successCount = 0;
  const errors = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const id = String(row[7] || '').trim();
    if (!targetIds.has(id)) continue;

    const currentStatus = String(row[8] || '').trim();
    if (currentStatus !== INITIAL_APPLICATION_STATUS) {
      errors.push({
        applicationId: id,
        message: `ステータスが「${currentStatus || '空'}」のため承認できません`,
      });
      continue;
    }

    const rowIndex = i + 1;
    sheet.getRange(rowIndex, 9).setValue(STATUS_APPROVED);
    sheet.getRange(rowIndex, 11).setValue(now);

    appendStatusHistory_(ss, id, HISTORY_KIND_STATUS, currentStatus, STATUS_APPROVED, '');
    successCount++;
  }

  return { successCount, errors };
}

/**
 * 管理画面用：申請を1件差し戻し（理由必須）
 */
function rejectApplication(applicationId, reason) {
  const id = String(applicationId || '').trim();
  const reasonText = String(reason || '').trim();

  if (!id) throw new Error('申請IDが必要です');
  if (!reasonText) throw new Error('差し戻し理由を入力してください');

  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(SHEET_NAMES.responses);
  if (!sheet) throw new Error(`シート「${SHEET_NAMES.responses}」がありません`);

  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const rowId = String(row[7] || '').trim();
    if (rowId !== id) continue;

    const currentStatus = String(row[8] || '').trim();
    if (currentStatus !== INITIAL_APPLICATION_STATUS) {
      throw new Error(`ステータスが「${currentStatus || '空'}」のため差し戻しできません`);
    }

    const rowIndex = i + 1;
    sheet.getRange(rowIndex, 9).setValue(STATUS_REJECTED);
    sheet.getRange(rowIndex, 10).setValue(reasonText);

    appendStatusHistory_(ss, id, HISTORY_KIND_STATUS, currentStatus, STATUS_REJECTED, reasonText);
    return { applicationId: id };
  }

  throw new Error(`申請ID「${id}」が見つかりません`);
}

/**
 * 管理画面用：承認済かつ未払いの申請一覧を取得
 *
 * フォーム回答シートで ステータス=承認済 かつ 支払ステータス=未払い の行を抽出し、
 * 申請IDで精算結果と join して返す。
 */
function getApprovedUnpaidApplicationsForWeb() {
  const ss = getSpreadsheet_();
  const responseSheet = ss.getSheetByName(SHEET_NAMES.responses);
  const resultSheet = ss.getSheetByName(SHEET_NAMES.results);

  if (!responseSheet) throw new Error(`シート「${SHEET_NAMES.responses}」がありません`);
  if (!resultSheet) throw new Error(`シート「${SHEET_NAMES.results}」がありません`);

  const resultMap = new Map();
  const resultValues = resultSheet.getDataRange().getValues();
  resultValues.slice(1).forEach(row => {
    const id = String(row[15] || '').trim();
    if (id) {
      resultMap.set(id, {
        place: row[6],
        distanceType: row[7],
        targetKm: row[8],
        payment: row[12],
        calcStatus: row[13],
        errorMessage: row[14],
      });
    }
  });

  const responseValues = responseSheet.getDataRange().getValues();
  const list = [];

  responseValues.slice(1).forEach(row => {
    const status = String(row[8] || '').trim();
    if (status !== STATUS_APPROVED) return;

    const paymentStatus = String(row[11] || '').trim();
    if (paymentStatus !== INITIAL_PAYMENT_STATUS) return;

    const applicationId = String(row[7] || '').trim();
    if (!applicationId) return;

    const result = resultMap.get(applicationId);

    list.push({
      applicationId,
      applicationDate: formatDateForDisplay_(row[1]),
      name: normalizeName_(row[2]),
      applicationType: row[3],
      place: result ? result.place : (row[3] === APPLICATION_TYPES.pickup ? row[4] : row[6]),
      targetKm: result ? result.targetKm : '',
      payment: result ? result.payment : '',
      calcStatus: result ? result.calcStatus : '',
      calcError: result ? result.errorMessage : '',
      approvedAt: formatDateForDisplay_(row[10]),
    });
  });

  return list;
}

/**
 * 管理画面用：複数の申請を一括で支払済みに設定
 *
 * ステータス=承認済 かつ 支払ステータス=未払い のみ遷移可能。
 * 戻り値：{ successCount, errors: [{ applicationId, message }] }
 */
function markAsPaid(applicationIds) {
  if (!Array.isArray(applicationIds) || applicationIds.length === 0) {
    return { successCount: 0, errors: [] };
  }

  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(SHEET_NAMES.responses);
  if (!sheet) throw new Error(`シート「${SHEET_NAMES.responses}」がありません`);

  const targetIds = new Set(
    applicationIds.map(id => String(id || '').trim()).filter(id => id)
  );
  const values = sheet.getDataRange().getValues();

  let successCount = 0;
  const errors = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const id = String(row[7] || '').trim();
    if (!targetIds.has(id)) continue;

    const currentAppStatus = String(row[8] || '').trim();
    if (currentAppStatus !== STATUS_APPROVED) {
      errors.push({
        applicationId: id,
        message: `ステータスが「${currentAppStatus || '空'}」のため支払処理できません（承認済が必要）`,
      });
      continue;
    }

    const currentPayStatus = String(row[11] || '').trim();
    if (currentPayStatus !== INITIAL_PAYMENT_STATUS) {
      errors.push({
        applicationId: id,
        message: `支払ステータスが「${currentPayStatus || '空'}」のため変更できません`,
      });
      continue;
    }

    const rowIndex = i + 1;
    sheet.getRange(rowIndex, 12).setValue(STATUS_PAID);

    appendStatusHistory_(ss, id, HISTORY_KIND_PAYMENT, currentPayStatus, STATUS_PAID, '');
    successCount++;
  }

  return { successCount, errors };
}

/**
 * ステータス変更履歴シートを取得（無ければ作成 / 旧スキーマなら自動マイグレーション）
 *
 * 列構成：
 * A 変更日時 | B 申請ID | C 変更種別 | D 変更前ステータス | E 変更後ステータス | F 操作者 | G 備考
 *
 * Step 2 で作られた 6 列スキーマ（変更種別なし）を検出した場合は、C列を自動挿入する。
 */
function getOrCreateStatusHistorySheet_(ss) {
  let sheet = ss.getSheetByName(SHEET_NAMES.statusHistory);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.statusHistory);
    sheet.appendRow([
      '変更日時',
      '申請ID',
      '変更種別',
      '変更前ステータス',
      '変更後ステータス',
      '操作者',
      '備考',
    ]);
    return sheet;
  }

  const lastCol = sheet.getLastColumn();
  if (lastCol === 6) {
    const headers = sheet.getRange(1, 1, 1, 6).getValues()[0];
    if (headers[0] === '変更日時' && headers[1] === '申請ID') {
      sheet.insertColumnBefore(3);
      sheet.getRange(1, 3).setValue('変更種別');
    }
  }

  return sheet;
}

/**
 * ステータス変更履歴に1件追記
 */
function appendStatusHistory_(ss, applicationId, kind, fromStatus, toStatus, note) {
  const sheet = getOrCreateStatusHistorySheet_(ss);
  sheet.appendRow([
    new Date(),
    applicationId,
    kind,
    fromStatus,
    toStatus,
    getCurrentUserEmail_(),
    note || '',
  ]);
}

/**
 * 操作中のユーザーメールを取得（取得不可なら空）
 */
function getCurrentUserEmail_() {
  try {
    return Session.getActiveUser().getEmail() || '';
  } catch (e) {
    return '';
  }
}

/**
 * 日付表示用：YYYY-MM-DD に整形
 */
function formatDateForDisplay_(value) {
  if (!value) return '';

  let date;

  if (Object.prototype.toString.call(value) === '[object Date]') {
    date = value;
  } else {
    date = new Date(value);
  }

  if (isNaN(date.getTime())) {
    return String(value);
  }

  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');

  return `${year}-${month}-${day}`;
}
