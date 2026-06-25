function doGet(e) {
  const data = readStudioData();
  const json = JSON.stringify(data);
  const callback = e && e.parameter && e.parameter.callback;

  if (callback) {
    return ContentService
      .createTextOutput(callback + "(" + json + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  writeStudioData(data);

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, savedAt: new Date().toISOString() }))
    .setMimeType(ContentService.MimeType.JSON);
}

function writeStudioData(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  writeSheet(ss, "Customers", [
    ["고객번호", "고객명", "전화번호", "아이이름", "아이정보", "메모", "등록일"]
  ], (data.customers || []).map(c => [
    c.id || "",
    c.name || "",
    c.phone || "",
    c.childName || "",
    c.childInfo || "",
    c.memo || "",
    c.createdAt || ""
  ]));

  writeSheet(ss, "Visits", [
    ["방문ID", "고객번호", "방문회차", "촬영일", "촬영종류", "예약금", "잔금", "결제수단", "결제직원", "메모", "사진수", "등록일"]
  ], (data.visits || []).map(v => [
    v.id || "",
    v.customerId || "",
    v.visitNo || "",
    v.date || "",
    v.shootType || "",
    v.deposit || 0,
    v.balance || 0,
    v.paymentMethod || "",
    v.paymentStaff || "",
    v.memo || "",
    (v.photos || []).length,
    v.createdAt || ""
  ]));

  writeSheet(ss, "Reservations", [
    ["예약ID", "고객번호", "고객명", "예약일", "시간", "촬영종류", "담당직원", "상태", "메모", "등록일"]
  ], (data.reservations || []).map(r => [
    r.id || "",
    r.customerId || "",
    r.customerName || "",
    r.date || "",
    r.time || "",
    r.shootType || "",
    r.staff || "",
    r.status || "",
    r.memo || "",
    r.createdAt || ""
  ]));
}

function readStudioData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  return {
    customers: readRows(ss, "Customers").map(row => ({
      id: row["고객번호"] || "",
      name: row["고객명"] || "",
      phone: row["전화번호"] || "",
      childName: row["아이이름"] || "",
      childInfo: row["아이정보"] || "",
      memo: row["메모"] || "",
      createdAt: row["등록일"] || new Date().toISOString()
    })).filter(c => c.id || c.name || c.phone),

    visits: readRows(ss, "Visits").map(row => ({
      id: row["방문ID"] || "",
      customerId: row["고객번호"] || "",
      visitNo: Number(row["방문회차"] || 0),
      date: normalizeDate(row["촬영일"]),
      shootType: row["촬영종류"] || "",
      deposit: Number(row["예약금"] || 0),
      balance: Number(row["잔금"] || 0),
      paymentMethod: row["결제수단"] || "",
      paymentStaff: row["결제직원"] || "",
      memo: row["메모"] || "",
      photos: [],
      createdAt: row["등록일"] || new Date().toISOString()
    })).filter(v => v.customerId || v.date || v.shootType),

    reservations: readRows(ss, "Reservations").map(row => ({
      id: row["예약ID"] || "",
      customerId: row["고객번호"] || "",
      date: normalizeDate(row["예약일"]),
      time: normalizeTime(row["시간"]),
      shootType: row["촬영종류"] || "",
      staff: row["담당직원"] || "",
      status: row["상태"] || "예약",
      memo: row["메모"] || "",
      createdAt: row["등록일"] || new Date().toISOString()
    })).filter(r => r.customerId || r.date || r.shootType),

    exportedAt: new Date().toISOString()
  };
}

function writeSheet(ss, name, header, rows) {
  const sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  sheet.clearContents();
  const values = header.concat(rows);
  sheet.getRange(1, 1, values.length, values[0].length).setValues(values);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, values[0].length);
}

function readRows(ss, name) {
  const sheet = ss.getSheetByName(name);
  if (!sheet) return [];

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];

  const headers = values[0].map(String);
  return values.slice(1).map(row => {
    const item = {};
    headers.forEach((header, index) => {
      item[header] = row[index];
    });
    return item;
  });
}

function normalizeDate(value) {
  if (!value) return "";
  if (Object.prototype.toString.call(value) === "[object Date]") {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return String(value).slice(0, 10);
}

function normalizeTime(value) {
  if (!value) return "";
  if (Object.prototype.toString.call(value) === "[object Date]") {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "HH:mm");
  }
  return String(value).slice(0, 5);
}
