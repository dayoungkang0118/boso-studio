const STORAGE_KEY = "boso-studio-manager-v1";
const DB_NAME = "boso-studio-manager-db";
const DB_STORE = "app-state";
const DB_VERSION = 1;

const state = {
  customers: [],
  visits: [],
  reservations: [],
  settings: {
    sheetWebhookUrl: "",
    calendarId: "",
    calendarDuration: 60,
  },
  selectedCustomerId: null,
};

const titles = {
  dashboard: ["대시보드", "오늘 예약과 최근 방문 고객을 확인하세요."],
  customers: ["고객관리", "이름, 전화번호, 아이 이름으로 빠르게 찾아보세요."],
  reservations: ["예약관리", "예약 일정과 촬영 상태를 관리하세요."],
  settings: ["연동/백업", "Google Sheets 동기화와 백업을 관리하세요."],
};

const appsScriptSample = `function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  writeSheet(ss, "Customers", [
    ["고객번호", "고객명", "전화번호", "아이이름", "아이정보", "메모", "등록일"]
  ], data.customers.map(c => [
    c.id, c.name, c.phone, c.childName, c.childInfo, c.memo, c.createdAt
  ]));

  writeSheet(ss, "Visits", [
    ["고객번호", "방문회차", "촬영일", "촬영종류", "예약금", "잔금", "결제수단", "결제직원", "메모", "사진수"]
  ], data.visits.map(v => [
    v.customerId, v.visitNo, v.date, v.shootType, v.deposit, v.balance,
    v.paymentMethod, v.paymentStaff, v.memo, (v.photos || []).length
  ]));

  writeSheet(ss, "Reservations", [
    ["고객번호", "고객명", "예약일", "시간", "촬영종류", "담당직원", "상태", "메모"]
  ], data.reservations.map(r => [
    r.customerId, r.customerName, r.date, r.time, r.shootType, r.staff, r.status, r.memo
  ]));

  return ContentService.createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function writeSheet(ss, name, header, rows) {
  const sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  sheet.clearContents();
  const values = header.concat(rows);
  sheet.getRange(1, 1, values.length, values[0].length).setValues(values);
}`;

const $ = (selector) => document.querySelector(selector);
const $$ = (selector) => Array.from(document.querySelectorAll(selector));

async function loadState() {
  const saved = await readPersistedState();
  if (saved) {
    Object.assign(state, saved);
    return;
  }

  seedSampleData();
  saveState();
}

function saveState() {
  persistState(state).catch(() => {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  });
}

function seedSampleData() {
  const today = toDateInput(new Date());
  state.customers = [
    {
      id: "BOSO-2026-0001",
      name: "김하늘",
      phone: "010-1234-5678",
      childName: "서아",
      childInfo: "2025.02.10",
      memo: "돌사진 문의 많음",
      createdAt: new Date().toISOString(),
    },
  ];
  state.visits = [
    {
      id: newId(),
      customerId: "BOSO-2026-0001",
      visitNo: 1,
      date: today,
      shootType: "아기사진",
      productName: "돌상 패키지",
      deposit: 50000,
      balance: 150000,
      paymentMethod: "계좌",
      paymentStaff: "대표",
      memo: "첫 방문. 밝은 배경 선호.",
      photos: [],
      createdAt: new Date().toISOString(),
    },
  ];
  state.reservations = [
    {
      id: newId(),
      customerId: "BOSO-2026-0001",
      date: today,
      time: "14:00",
      shootType: "가족사진",
      productName: "가족사진 20R",
      staff: "대표",
      status: "예약",
      memo: "부모님 동반",
      createdAt: new Date().toISOString(),
    },
  ];
}

async function init() {
  await loadState();
  loadAppsScriptSample();
  $("#sheetWebhookUrl").value = state.settings.sheetWebhookUrl || "";
  $("#calendarId").value = state.settings.calendarId || "";
  $("#calendarDuration").value = state.settings.calendarDuration || 60;
  $("#reservationDateFilter").value = "";

  bindEvents();
  renderAll();
}

function bindEvents() {
  $$(".nav-button").forEach((button) => {
    button.addEventListener("click", () => switchView(button.dataset.view));
  });

  $$("[data-view-target]").forEach((button) => {
    button.addEventListener("click", () => switchView(button.dataset.viewTarget));
  });

  $("#openCustomerModal").addEventListener("click", () => {
    $("#customerForm").reset();
    $("#customerModal").showModal();
  });

  $("#openReservationModal").addEventListener("click", () => {
    $("#reservationForm").reset();
    fillCustomerSelect();
    $("#reservationForm").date.value = toDateInput(new Date());
    $("#reservationModal").showModal();
  });

  $$("[data-close-modal]").forEach((button) => {
    button.addEventListener("click", () => $(`#${button.dataset.closeModal}`).close());
  });

  $("#customerSearch").addEventListener("input", renderCustomers);
  $("#shootTypeFilter").addEventListener("change", renderCustomers);
  $("#reservationSearch").addEventListener("input", renderReservations);
  $("#reservationDateFilter").addEventListener("change", renderReservations);

  $("#customerForm").addEventListener("submit", handleCustomerSubmit);
  $("#visitForm").addEventListener("submit", handleVisitSubmit);
  $("#reservationForm").addEventListener("submit", handleReservationSubmit);

  $("#saveWebhook").addEventListener("click", saveWebhook);
  $("#saveCalendarSettings").addEventListener("click", saveCalendarSettings);
  $("#pushSheets").addEventListener("click", pushSheets);
  $("#pullSheets").addEventListener("click", pullSheets);
  $("#pushCalendar").addEventListener("click", pushCalendar);
  $("#exportJson").addEventListener("click", exportJson);
  $("#exportCsv").addEventListener("click", exportCsv);
  $("#importJson").addEventListener("change", importJson);
}

function switchView(view) {
  $$(".nav-button").forEach((button) => button.classList.toggle("active", button.dataset.view === view));
  $$(".view").forEach((section) => section.classList.remove("active"));
  $(`#${view}View`).classList.add("active");
  $("#pageTitle").textContent = titles[view][0];
  $("#pageSubTitle").textContent = titles[view][1];
}

function renderAll() {
  renderDashboard();
  renderCustomers();
  renderReservations();
  fillCustomerSelect();
}

function renderDashboard() {
  const revenue = getRevenueSummary();
  $("#totalCustomers").textContent = state.customers.length;
  $("#totalVisits").textContent = state.visits.length;
  $("#upcomingReservations").textContent = state.reservations.filter((r) => r.status === "예약" && r.date >= toDateInput(new Date())).length;
  $("#unpaidBalance").textContent = formatWon(state.visits.reduce((sum, visit) => sum + Number(visit.balance || 0), 0));
  $("#currentMonthRevenue").textContent = `이번 달 ${formatWon(revenue.currentMonthTotal)}`;
  $("#currentMonthVisitCount").textContent = `${revenue.currentMonthVisitCount}건`;
  $("#monthlyRevenueChart").innerHTML = renderMonthlyRevenueChart(revenue.monthly);
  $("#shootTypeRevenueList").innerHTML = renderShootTypeRevenue(revenue.byShootType);

  const today = toDateInput(new Date());
  const todays = state.reservations
    .filter((r) => r.date === today)
    .sort((a, b) => a.time.localeCompare(b.time));

  $("#todayReservationList").innerHTML = todays.length
    ? todays.map(renderReservationItem).join("")
    : `<div class="empty-state">오늘 예약이 없습니다.</div>`;

  const recent = [...state.visits].sort((a, b) => b.date.localeCompare(a.date)).slice(0, 6);
  $("#recentVisitList").innerHTML = recent.length
    ? recent.map(renderVisitSummary).join("")
    : `<div class="empty-state">방문 기록이 없습니다.</div>`;
}

function getRevenueSummary() {
  const nowMonth = toMonthKey(new Date());
  const monthlyMap = new Map();
  const shootTypeMap = new Map();
  let currentMonthVisitCount = 0;

  state.visits.forEach((visit) => {
    const month = (visit.date || "").slice(0, 7);
    if (!month) return;

    const revenue = getVisitRevenue(visit);
    monthlyMap.set(month, (monthlyMap.get(month) || 0) + revenue);

    if (month === nowMonth) {
      currentMonthVisitCount += 1;
      shootTypeMap.set(visit.shootType, (shootTypeMap.get(visit.shootType) || 0) + revenue);
    }
  });

  const months = lastMonths(12);
  const monthly = months.map((month) => ({
    month,
    total: monthlyMap.get(month) || 0,
  }));

  return {
    monthly,
    currentMonthTotal: monthlyMap.get(nowMonth) || 0,
    currentMonthVisitCount,
    byShootType: Array.from(shootTypeMap, ([type, total]) => ({ type, total }))
      .sort((a, b) => b.total - a.total),
  };
}

function renderMonthlyRevenueChart(monthly) {
  const max = Math.max(...monthly.map((item) => item.total), 1);
  return monthly.map((item) => {
    const height = Math.max(4, Math.round((item.total / max) * 128));
    return `
      <div class="revenue-bar-item">
        <div class="revenue-value">${item.total ? formatCompactWon(item.total) : "0"}</div>
        <div class="revenue-bar-track">
          <div class="revenue-bar" style="height:${height}px"></div>
        </div>
        <div class="revenue-month">${item.month.slice(5)}월</div>
      </div>`;
  }).join("");
}

function renderShootTypeRevenue(items) {
  if (!items.length) return `<div class="empty-state compact">이번 달 방문 매출이 없습니다.</div>`;

  return items.map((item) => `
    <article class="list-item">
      <div class="item-top">
        <div>
          <div class="item-title">${escapeHtml(item.type)}</div>
          <div class="item-meta">이번 달 촬영 매출</div>
        </div>
        <strong>${formatWon(item.total)}</strong>
      </div>
    </article>`).join("");
}

function renderCustomers() {
  const query = normalize($("#customerSearch").value);
  const shootType = $("#shootTypeFilter").value;

  const customers = state.customers.filter((customer) => {
    const visitText = getVisits(customer.id).map((visit) => [visit.date, visit.shootType, visit.productName, visit.memo].join(" ")).join(" ");
    const haystack = normalize([customer.id, customer.name, customer.phone, customer.childName, customer.childInfo, customer.memo, visitText].join(" "));
    const matchesQuery = !query || haystack.includes(query);
    const matchesShoot = !shootType || state.visits.some((visit) => visit.customerId === customer.id && visit.shootType === shootType);
    return matchesQuery && matchesShoot;
  });

  $("#customerList").innerHTML = customers.length
    ? customers.map(renderCustomerListItem).join("")
    : `<div class="empty-state">검색 결과가 없습니다.</div>`;

  $$(".customer-card").forEach((card) => {
    card.addEventListener("click", () => {
      state.selectedCustomerId = card.dataset.customerId;
      renderCustomerDetail();
      renderCustomers();
    });
  });

  renderCustomerDetail();
}

function renderCustomerListItem(customer) {
  const visits = getVisits(customer.id).sort((a, b) => b.date.localeCompare(a.date));
  const latestVisit = visits[0];
  const selected = state.selectedCustomerId === customer.id ? " selected" : "";
  return `
    <article class="list-item clickable customer-card${selected}" data-customer-id="${customer.id}">
      <div class="item-top">
        <div>
          <div class="item-title">${escapeHtml(customer.name)} <span class="badge">${customer.id}</span></div>
          <div class="item-meta">${escapeHtml(customer.phone)} · 아이: ${escapeHtml(customer.childName || "-")}</div>
          <div class="item-meta">최근 촬영: ${latestVisit ? `${formatDate(latestVisit.date)} · ${escapeHtml(latestVisit.shootType)}${latestVisit.productName ? ` · ${escapeHtml(latestVisit.productName)}` : ""}` : "기록 없음"}</div>
        </div>
        <span class="badge ${visits.length > 1 ? "done" : ""}">${visits.length}회</span>
      </div>
    </article>`;
}

function renderCustomerDetail() {
  const customer = state.customers.find((item) => item.id === state.selectedCustomerId);
  if (!customer) {
    $("#customerDetail").innerHTML = `<div class="empty-state">고객을 선택하면 방문 기록과 결제 내역이 보입니다.</div>`;
    return;
  }

  const visits = getVisits(customer.id).sort((a, b) => b.date.localeCompare(a.date));
  const reservations = state.reservations
    .filter((item) => item.customerId === customer.id)
    .sort((a, b) => `${b.date} ${b.time}`.localeCompare(`${a.date} ${a.time}`));

  $("#customerDetail").innerHTML = `
    <div class="detail-header">
      <div class="detail-title">
        <h2>${escapeHtml(customer.name)} <span class="badge">${customer.id}</span></h2>
        <p class="muted">${escapeHtml(customer.phone)} · 아이: ${escapeHtml(customer.childName || "-")}</p>
      </div>
      <button class="primary-button" id="addVisit">+ 방문 기록</button>
    </div>
    <div class="detail-grid">
      <div class="info-box"><span>총 방문</span><strong>${visits.length}회</strong></div>
      <div class="info-box"><span>최근 촬영</span><strong>${visits[0] ? `${formatDate(visits[0].date)} · ${escapeHtml(visits[0].shootType)}` : "-"}</strong></div>
      <div class="info-box"><span>아이 정보</span><strong>${escapeHtml(customer.childInfo || "-")}</strong></div>
    </div>
    ${customer.memo ? `<div class="list-item"><strong>고객 메모</strong><div class="item-meta">${escapeHtml(customer.memo)}</div></div>` : ""}
    <div class="panel-head"><h2>방문/촬영 기록</h2></div>
    <div class="list">${visits.length ? visits.map(renderVisitDetail).join("") : `<div class="empty-state">방문 기록이 없습니다.</div>`}</div>
    <div class="panel-head" style="margin-top:18px"><h2>예약 이력</h2></div>
    <div class="list">${reservations.length ? reservations.map(renderReservationItem).join("") : `<div class="empty-state">예약 이력이 없습니다.</div>`}</div>
  `;

  $("#addVisit").addEventListener("click", () => {
    $("#visitForm").reset();
    $("#visitForm").customerId.value = customer.id;
    $("#visitForm").date.value = toDateInput(new Date());
    $("#visitModal").showModal();
  });
}

function renderVisitDetail(visit) {
  const photoMarkup = visit.photos?.length
    ? `<div class="photo-grid">${visit.photos.map((photo) => `<img src="${photo.dataUrl}" alt="${escapeHtml(photo.name)}" />`).join("")}</div>`
    : "";
  return `
    <article class="list-item visit-card">
      <div class="item-top">
        <div>
          <div class="item-title">${visit.visitNo}번째 방문 · ${escapeHtml(visit.shootType)}${visit.productName ? ` · ${escapeHtml(visit.productName)}` : ""}</div>
          <div class="item-meta">${formatDate(visit.date)} · 예약금 ${formatWon(visit.deposit)} · 잔금 ${formatWon(visit.balance)} · ${escapeHtml(visit.paymentMethod)} · 담당 ${escapeHtml(visit.paymentStaff || "-")}</div>
        </div>
        <span class="badge ${Number(visit.balance) > 0 ? "warning" : "done"}">${Number(visit.balance) > 0 ? "잔금있음" : "정산완료"}</span>
      </div>
      ${visit.memo ? `<div class="item-meta">${escapeHtml(visit.memo)}</div>` : ""}
      ${photoMarkup}
    </article>`;
}

function renderVisitSummary(visit) {
  const customer = getCustomer(visit.customerId);
  return `
    <article class="list-item">
      <div class="item-top">
        <div>
          <div class="item-title">${escapeHtml(customer?.name || "삭제된 고객")} · ${visit.visitNo}번째 방문</div>
          <div class="item-meta">${formatDate(visit.date)} · ${escapeHtml(visit.shootType)}${visit.productName ? ` · ${escapeHtml(visit.productName)}` : ""} · 사진 ${visit.photos?.length || 0}장</div>
        </div>
        <span class="badge">${escapeHtml(visit.customerId)}</span>
      </div>
    </article>`;
}

function renderReservations() {
  const query = normalize($("#reservationSearch").value);
  const date = $("#reservationDateFilter").value;
  const reservations = state.reservations
    .filter((reservation) => {
      const customer = getCustomer(reservation.customerId);
      const haystack = normalize([customer?.name, customer?.phone, reservation.shootType, reservation.productName, reservation.staff, reservation.memo].join(" "));
      return (!query || haystack.includes(query)) && (!date || reservation.date === date);
    })
    .sort((a, b) => `${a.date} ${a.time}`.localeCompare(`${b.date} ${b.time}`));

  $("#reservationList").innerHTML = reservations.length
    ? reservations.map(renderReservationItem).join("")
    : `<div class="empty-state">예약이 없습니다.</div>`;
}

function renderReservationItem(reservation) {
  const customer = getCustomer(reservation.customerId);
  const statusClass = reservation.status === "예약" ? "" : reservation.status === "촬영완료" ? "done" : "warning";
  return `
    <article class="list-item">
      <div class="item-top">
        <div>
          <div class="item-title">${formatDate(reservation.date)} ${reservation.time} · ${escapeHtml(customer?.name || "삭제된 고객")}</div>
          <div class="item-meta">${escapeHtml(customer?.phone || "-")} · ${escapeHtml(reservation.shootType)}${reservation.productName ? ` · ${escapeHtml(reservation.productName)}` : ""} · 담당 ${escapeHtml(reservation.staff || "-")}</div>
          ${reservation.memo ? `<div class="item-meta">${escapeHtml(reservation.memo)}</div>` : ""}
        </div>
        <span class="badge ${statusClass}">${escapeHtml(reservation.status)}</span>
      </div>
    </article>`;
}

function handleCustomerSubmit(event) {
  event.preventDefault();
  const form = new FormData(event.currentTarget);
  const customer = {
    id: nextCustomerId(),
    name: form.get("name").trim(),
    phone: form.get("phone").trim(),
    childName: form.get("childName").trim(),
    childInfo: form.get("childInfo").trim(),
    memo: form.get("memo").trim(),
    createdAt: new Date().toISOString(),
  };

  state.customers.unshift(customer);
  if (form.get("firstVisitDate") && form.get("firstShootType")) {
    state.visits.unshift({
      id: newId(),
      customerId: customer.id,
      visitNo: 1,
      date: form.get("firstVisitDate"),
      shootType: form.get("firstShootType"),
      productName: form.get("firstProductName").trim(),
      deposit: 0,
      balance: 0,
      paymentMethod: "미결제",
      paymentStaff: "",
      memo: "고객 등록 시 입력한 첫 촬영 기록",
      photos: [],
      createdAt: new Date().toISOString(),
    });
  }
  state.selectedCustomerId = customer.id;
  saveState();
  $("#customerModal").close();
  renderAll();
  switchView("customers");
  showToast("고객이 등록되었습니다.");
}

async function handleVisitSubmit(event) {
  event.preventDefault();
  const form = new FormData(event.currentTarget);
  const customerId = form.get("customerId");
  const photos = await filesToDataUrls(event.currentTarget.photos.files);
  const visit = {
    id: newId(),
    customerId,
    visitNo: getVisits(customerId).length + 1,
    date: form.get("date"),
    shootType: form.get("shootType"),
    productName: form.get("productName").trim(),
    deposit: Number(form.get("deposit") || 0),
    balance: Number(form.get("balance") || 0),
    paymentMethod: form.get("paymentMethod"),
    paymentStaff: form.get("paymentStaff").trim(),
    memo: form.get("memo").trim(),
    photos,
    createdAt: new Date().toISOString(),
  };

  state.visits.unshift(visit);
  saveState();
  $("#visitModal").close();
  renderAll();
  showToast("방문 기록이 저장되었습니다.");
}

function handleReservationSubmit(event) {
  event.preventDefault();
  const form = new FormData(event.currentTarget);
  const reservation = {
    id: newId(),
    customerId: form.get("customerId"),
    date: form.get("date"),
    time: form.get("time"),
    shootType: form.get("shootType"),
    productName: form.get("productName").trim(),
    staff: form.get("staff").trim(),
    status: form.get("status"),
    memo: form.get("memo").trim(),
    createdAt: new Date().toISOString(),
  };

  state.reservations.push(reservation);
  saveState();
  $("#reservationModal").close();
  renderAll();
  switchView("reservations");
  showToast("예약이 등록되었습니다.");
}

function fillCustomerSelect() {
  const select = $("#reservationCustomerSelect");
  select.innerHTML = state.customers
    .map((customer) => `<option value="${customer.id}">${escapeHtml(customer.name)} · ${escapeHtml(customer.phone)} · ${customer.id}</option>`)
    .join("");
}

function saveWebhook() {
  state.settings.sheetWebhookUrl = $("#sheetWebhookUrl").value.trim();
  saveCalendarSettings(false);
  saveState();
  showToast("연동 URL이 저장되었습니다.");
}

function saveCalendarSettings(showMessage = true) {
  state.settings.calendarId = $("#calendarId").value.trim();
  state.settings.calendarDuration = Number($("#calendarDuration").value || 60);
  saveState();
  if (showMessage) showToast("캘린더 설정이 저장되었습니다.");
}

async function pushSheets() {
  const url = state.settings.sheetWebhookUrl || $("#sheetWebhookUrl").value.trim();
  if (!url) {
    showToast("Apps Script 웹앱 URL을 먼저 입력하세요.");
    return;
  }

  const payload = {
    customers: state.customers,
    visits: state.visits,
    reservations: state.reservations.map((reservation) => ({
      ...reservation,
      customerName: getCustomer(reservation.customerId)?.name || "",
    })),
    syncedAt: new Date().toISOString(),
  };

  try {
    await fetch(url, {
      method: "POST",
      mode: "no-cors",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify(payload),
    });
    showToast("Google Sheets로 전송했습니다.");
  } catch {
    showToast("전송에 실패했습니다. URL과 배포 권한을 확인하세요.");
  }
}

async function pullSheets() {
  const url = state.settings.sheetWebhookUrl || $("#sheetWebhookUrl").value.trim();
  if (!url) {
    showToast("Apps Script 웹앱 URL을 먼저 입력하세요.");
    return;
  }

  try {
    const data = await fetchSheetJsonp(url);
    state.customers = data.customers || [];
    state.visits = (data.visits || []).map((visit) => ({
      photos: [],
      createdAt: new Date().toISOString(),
      ...visit,
      id: visit.id || newId(),
      deposit: Number(visit.deposit || 0),
      balance: Number(visit.balance || 0),
    }));
    state.reservations = (data.reservations || []).map((reservation) => ({
      createdAt: new Date().toISOString(),
      ...reservation,
      id: reservation.id || newId(),
    }));
    state.settings.sheetWebhookUrl = url;
    state.selectedCustomerId = null;
    saveState();
    renderAll();
    showToast("Google Sheets 데이터를 가져왔습니다.");
  } catch {
    showToast("가져오기에 실패했습니다. Apps Script 배포 상태를 확인하세요.");
  }
}

async function pushCalendar() {
  const url = state.settings.sheetWebhookUrl || $("#sheetWebhookUrl").value.trim();
  if (!url) {
    showToast("Apps Script 웹앱 URL을 먼저 입력하세요.");
    return;
  }

  saveCalendarSettings(false);

  const payload = {
    action: "syncCalendar",
    calendarId: state.settings.calendarId,
    eventDurationMinutes: state.settings.calendarDuration || 60,
    reservations: state.reservations.map((reservation) => {
      const customer = getCustomer(reservation.customerId);
      return {
        ...reservation,
        customerName: customer?.name || "",
        customerPhone: customer?.phone || "",
        childName: customer?.childName || "",
      };
    }),
    syncedAt: new Date().toISOString(),
  };

  try {
    await fetch(url, {
      method: "POST",
      mode: "no-cors",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify(payload),
    });
    showToast("Google Calendar로 예약을 전송했습니다.");
  } catch {
    showToast("캘린더 전송에 실패했습니다. Apps Script 권한을 확인하세요.");
  }
}

function fetchSheetJsonp(url) {
  return new Promise((resolve, reject) => {
    const callbackName = `bosoSheetCallback${Date.now()}`;
    const script = document.createElement("script");
    const separator = url.includes("?") ? "&" : "?";
    const timer = window.setTimeout(() => {
      cleanup();
      reject(new Error("Google Sheets request timed out"));
    }, 15000);

    window[callbackName] = (data) => {
      cleanup();
      resolve(data);
    };

    function cleanup() {
      window.clearTimeout(timer);
      delete window[callbackName];
      script.remove();
    }

    script.onerror = () => {
      cleanup();
      reject(new Error("Google Sheets request failed"));
    };
    script.src = `${url}${separator}callback=${callbackName}&ts=${Date.now()}`;
    document.body.appendChild(script);
  });
}

function loadAppsScriptSample() {
  fetch("google-apps-script.gs")
    .then((response) => response.text())
    .then((text) => {
      $("#appsScriptCode").textContent = text;
    })
    .catch(() => {
      $("#appsScriptCode").textContent = appsScriptSample;
    });
}

function exportJson() {
  downloadFile(`boso-backup-${toDateInput(new Date())}.json`, JSON.stringify(state, null, 2), "application/json");
}

function exportCsv() {
  const rows = [["고객번호", "고객명", "전화번호", "아이이름", "방문회차", "촬영일", "촬영종류", "촬영상품", "예약금", "잔금", "결제수단", "결제직원"]];
  state.visits.forEach((visit) => {
    const customer = getCustomer(visit.customerId) || {};
    rows.push([
      visit.customerId,
      customer.name || "",
      customer.phone || "",
      customer.childName || "",
      visit.visitNo,
      visit.date,
      visit.shootType,
      visit.productName || "",
      visit.deposit,
      visit.balance,
      visit.paymentMethod,
      visit.paymentStaff,
    ]);
  });
  const csv = rows.map((row) => row.map(csvCell).join(",")).join("\n");
  downloadFile(`boso-visits-${toDateInput(new Date())}.csv`, "\ufeff" + csv, "text/csv;charset=utf-8");
}

function importJson(event) {
  const file = event.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = () => {
    try {
      const imported = JSON.parse(reader.result);
      state.customers = imported.customers || [];
      state.visits = imported.visits || [];
      state.reservations = imported.reservations || [];
      state.settings = {
        sheetWebhookUrl: "",
        calendarId: "",
        calendarDuration: 60,
        ...(imported.settings || {}),
      };
      state.selectedCustomerId = null;
      saveState();
      $("#sheetWebhookUrl").value = state.settings.sheetWebhookUrl || "";
      $("#calendarId").value = state.settings.calendarId || "";
      $("#calendarDuration").value = state.settings.calendarDuration || 60;
      renderAll();
      showToast("백업을 복원했습니다.");
    } catch {
      showToast("JSON 파일을 읽지 못했습니다.");
    }
  };
  reader.readAsText(file);
  event.target.value = "";
}

function nextCustomerId() {
  const year = new Date().getFullYear();
  const prefix = `BOSO-${year}-`;
  const last = state.customers
    .map((customer) => customer.id)
    .filter((id) => id.startsWith(prefix))
    .map((id) => Number(id.split("-").pop()))
    .filter(Number.isFinite)
    .sort((a, b) => b - a)[0] || 0;
  return `${prefix}${String(last + 1).padStart(4, "0")}`;
}

function newId() {
  if (globalThis.crypto?.randomUUID) return globalThis.crypto.randomUUID();
  return `${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function getCustomer(id) {
  return state.customers.find((customer) => customer.id === id);
}

function getVisits(customerId) {
  return state.visits.filter((visit) => visit.customerId === customerId);
}

function normalize(value) {
  return String(value || "").replace(/\s/g, "").toLowerCase();
}

function formatWon(value) {
  return `${Number(value || 0).toLocaleString("ko-KR")}원`;
}

function formatCompactWon(value) {
  const amount = Number(value || 0);
  if (amount >= 10000) return `${Math.round(amount / 10000).toLocaleString("ko-KR")}만`;
  return amount.toLocaleString("ko-KR");
}

function getVisitRevenue(visit) {
  return Number(visit.deposit || 0) + Number(visit.balance || 0);
}

function formatDate(value) {
  if (!value) return "-";
  const [year, month, day] = value.split("-");
  return `${year}.${month}.${day}`;
}

function toMonthKey(date) {
  return toDateInput(date).slice(0, 7);
}

function lastMonths(count) {
  const result = [];
  const cursor = new Date();
  cursor.setDate(1);

  for (let index = count - 1; index >= 0; index -= 1) {
    const date = new Date(cursor);
    date.setMonth(cursor.getMonth() - index);
    result.push(toMonthKey(date));
  }

  return result;
}

function toDateInput(date) {
  const local = new Date(date.getTime() - date.getTimezoneOffset() * 60000);
  return local.toISOString().slice(0, 10);
}

function escapeHtml(value) {
  return String(value || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function csvCell(value) {
  return `"${String(value ?? "").replaceAll('"', '""')}"`;
}

function downloadFile(filename, content, type) {
  const blob = new Blob([content], { type });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  link.click();
  URL.revokeObjectURL(url);
}

function filesToDataUrls(files) {
  return Promise.all(
    Array.from(files).map((file) => new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => {
        const image = new Image();
        image.onload = () => {
          const maxSide = 1600;
          const scale = Math.min(1, maxSide / Math.max(image.width, image.height));
          const canvas = document.createElement("canvas");
          canvas.width = Math.round(image.width * scale);
          canvas.height = Math.round(image.height * scale);
          const context = canvas.getContext("2d");
          context.drawImage(image, 0, 0, canvas.width, canvas.height);
          resolve({
            name: file.name,
            type: "image/jpeg",
            dataUrl: canvas.toDataURL("image/jpeg", 0.82),
          });
        };
        image.onerror = reject;
        image.src = reader.result;
      };
      reader.onerror = reject;
      reader.readAsDataURL(file);
    })),
  );
}

function openDatabase() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);
    request.onupgradeneeded = () => {
      request.result.createObjectStore(DB_STORE);
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

async function readPersistedState() {
  try {
    const db = await openDatabase();
    const value = await idbRequest(db.transaction(DB_STORE, "readonly").objectStore(DB_STORE).get(STORAGE_KEY));
    db.close();
    if (value) return value;
  } catch {
    const saved = localStorage.getItem(STORAGE_KEY);
    return saved ? JSON.parse(saved) : null;
  }

  const saved = localStorage.getItem(STORAGE_KEY);
  return saved ? JSON.parse(saved) : null;
}

async function persistState(value) {
  const db = await openDatabase();
  await idbRequest(db.transaction(DB_STORE, "readwrite").objectStore(DB_STORE).put(JSON.parse(JSON.stringify(value)), STORAGE_KEY));
  db.close();
}

function idbRequest(request) {
  return new Promise((resolve, reject) => {
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

function showToast(message) {
  const toast = $("#toast");
  toast.textContent = message;
  toast.classList.add("show");
  window.clearTimeout(showToast.timer);
  showToast.timer = window.setTimeout(() => toast.classList.remove("show"), 2400);
}

init();
