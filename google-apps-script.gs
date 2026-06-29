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

  if (data.action === "syncCalendar") {
    const result = syncReservationsToCalendar(data);
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, result: result }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  writeStudioData(data);

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, savedAt: new Date().toISOString() }))
    .setMimeType(ContentService.MimeType.JSON);
}

function syncReservationsToCalendar(data) {
  const calendar = getTargetCalendar(data.calendarId);
  const reservations = data.reservations || [];
  const durationMinutes = Number(data.eventDurationMinutes || 60);
  const eventMap = getExistingReservationEvents(calendar);
  const result = { created: 0, updated: 0, deleted: 0, skipped: 0 };

  reservations.forEach(reservation => {
    if (!reservation.id || !reservation.date || !reservation.time) {
      result.skipped += 1;
      return;
    }

    const existing = eventMap[reservation.id];

    if (reservation.status === "취소") {
      if (existing) {
        existing.deleteEvent();
        result.deleted += 1;
      } else {
        result.skipped += 1;
      }
      return;
    }

    const start = parseDateTime(reservation.date, reservation.time);
    const end = new Date(start.getTime() + durationMinutes * 60 * 1000);
    const title = buildCalendarTitle(reservation);
    const description = buildCalendarDescription(reservation);

    if (existing) {
      existing.setTitle(title);
      existing.setTime(start, end);
      existing.setDescription(description);
      existing.setTag("bosoReservationId", reservation.id);
      result.updated += 1;
    } else {
      const event = calendar.createEvent(title, start, end, { description: description });
      event.setTag("bosoReservationId", reservation.id);
      result.created += 1;
    }
  });

  return result;
}

function getTargetCalendar(calendarId) {
  if (!calendarId || calendarId === "primary") {
    return CalendarApp.getDefaultCalendar();
  }

  const calendar = CalendarApp.getCalendarById(calendarId);
  if (!calendar) {
    throw new Error("Calendar not found: " + calendarId);
  }
  return calendar;
}

function getExistingReservationEvents(calendar) {
  const now = new Date();
  const rangeStart = new Date(now.getFullYear() - 1, 0, 1);
  const rangeEnd = new Date(now.getFullYear() + 2, 11, 31);
  const events = calendar.getEvents(rangeStart, rangeEnd);
  const map = {};

  events.forEach(event => {
    const reservationId = event.getTag("bosoReservationId");
    if (reservationId) {
      map[reservationId] = event;
    }
  });

  return map;
}

function buildCalendarTitle(reservation) {
  const name = reservation.customerName || "고객";
  const type = [reservation.shootType, reservation.productName].filter(Boolean).join(" · ") || "촬영";
  return "[보소사진관] " + name + " - " + type;
}

function buildCalendarDescription(reservation) {
  return [
    "고객명: " + (reservation.customerName || ""),
    "전화번호: " + (reservation.customerPhone || ""),
    "아이 이름: " + (reservation.childName || ""),
    "촬영종류: " + (reservation.shootType || ""),
    "촬영상품: " + (reservation.productName || ""),
    "담당 직원: " + (reservation.staff || ""),
    "상태: " + (reservation.status || ""),
    "예약ID: " + (reservation.id || ""),
    "",
    "메모:",
    reservation.memo || ""
  ].join("\n");
}

function parseDateTime(dateValue, timeValue) {
  const date = normalizeDate(dateValue);
  const time = normalizeTime(timeValue) || "00:00";
  const parts = date.split("-").map(Number);
  const timeParts = time.split(":").map(Number);
  return new Date(parts[0], parts[1] - 1, parts[2], timeParts[0] || 0, timeParts[1] || 0);
}

function writeStudioData(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  writeSheet(ss, "Customers", [
    ["고객번호", "고객명", "전화번호", "아이이름", "아이정보", "주소", "메모", "등록일"]
  ], (data.customers || []).map(c => [
    c.id || "",
    c.name || "",
    c.phone || "",
    c.childName || "",
    c.childInfo || "",
    c.address || "",
    c.memo || "",
    c.createdAt || ""
  ]));

  writeSheet(ss, "Visits", [
    ["방문ID", "고객번호", "방문회차", "촬영일", "촬영종류", "촬영상품", "총금액", "계약금받은금액", "계약금결제방법", "잔금받은금액", "잔금결제방법", "잔금직원", "총받은금액", "남은금액", "정산상태", "택배여부", "메모", "사진수", "등록일"]
  ], (data.visits || []).map(v => [
    v.id || "",
    v.customerId || "",
    v.visitNo || "",
    v.date || "",
    v.shootType || "",
    v.productName || "",
    v.totalAmount || 0,
    v.deposit || 0,
    Number(v.deposit || 0) > 0 ? "계좌" : "미결제",
    v.balance || 0,
    v.balancePaymentMethod || "",
    v.balancePaymentStaff || "",
    getPaidAmount(v),
    getRemainingAmount(v),
    getSettlementStatus(v),
    v.deliveryStatus || "없음",
    v.memo || "",
    (v.photos || []).length,
    v.createdAt || ""
  ]));

  writeSheet(ss, "Reservations", [
    ["예약ID", "고객번호", "고객명", "예약일", "시간", "촬영종류", "촬영상품", "담당직원", "상태", "메모", "등록일"]
  ], (data.reservations || []).map(r => [
    r.id || "",
    r.customerId || "",
    r.customerName || "",
    r.date || "",
    r.time || "",
    r.shootType || "",
    r.productName || "",
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
      address: row["주소"] || "",
      memo: row["메모"] || "",
      createdAt: row["등록일"] || new Date().toISOString()
    })).filter(c => c.id || c.name || c.phone),

    visits: readRows(ss, "Visits").map(row => ({
      id: row["방문ID"] || "",
      customerId: row["고객번호"] || "",
      visitNo: Number(row["방문회차"] || 0),
      date: normalizeDate(row["촬영일"]),
      shootType: row["촬영종류"] || "",
      productName: row["촬영상품"] || "",
      totalAmount: Number(row["총금액"] || 0),
      deposit: Number(row["계약금"] || row["예약금"] || 0),
      depositPaymentMethod: Number(row["계약금"] || row["예약금"] || 0) > 0 ? "계좌" : "미결제",
      balance: Number(row["잔금"] || 0),
      balancePaymentMethod: row["잔금결제방법"] || "",
      balancePaymentStaff: row["잔금직원"] || "",
      deliveryStatus: row["택배여부"] || "없음",
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
      productName: row["촬영상품"] || "",
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

function getPaidAmount(visit) {
  return Number(visit.deposit || 0) + Number(visit.balance || 0);
}

function getRemainingAmount(visit) {
  return Math.max(Number(visit.totalAmount || 0) - getPaidAmount(visit), 0);
}

function getSettlementStatus(visit) {
  return getRemainingAmount(visit) <= 0 ? "정산완료" : "잔금있음";
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
