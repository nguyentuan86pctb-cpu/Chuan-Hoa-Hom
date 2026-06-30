const DB_NAME = "chuan-hoa-cong-to-db";
const DB_STORE = "sessions";
const INCOMING_DB_NAME = "chuan-hoa-incoming-files";
const INCOMING_DB_STORE = "files";
const INCOMING_FILE_KEY = "latest-excel";
const SESSION_KEY = "current";
const STANDARD_SHEET_NAME = "Sheet1";
const REQUIRED_COLUMNS = ["MA_KHANG", "TEN_KHANG", "SO_TBI", "SO_COT", "Tram", "CHUAN_HOA"];
const EXPORT_COLUMNS = [...REQUIRED_COLUMNS, "COT_MOI", "GHI_CHU"];
const BOX_TYPES = [
  { label: "H1", value: "1", capacity: 1, cols: 1 },
  { label: "H2", value: "2", capacity: 2, cols: 2 },
  { label: "H4", value: "4", capacity: 4, cols: 2 },
  { label: "H6", value: "6", capacity: 6, cols: 3 },
  { label: "Hòm 3 pha", value: "3", capacity: 1, cols: 1 },
];
const CABINET_TYPES = ["06", "08", "10", "12", "16", "20"];
const PHASES = ["A", "B", "C", "D"];

let state = {
  session: createEmptySession(),
  selectedColumn: "",
  view: "import",
  search: "",
  message: "",
  saveState: "Chưa có dữ liệu",
  showAddCustomerForm: false,
};
let saveTimer = null;

function openDb() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, 1);
    request.onupgradeneeded = () => request.result.createObjectStore(DB_STORE);
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

async function saveSession(session) {
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(DB_STORE, "readwrite");
    tx.objectStore(DB_STORE).put(session, SESSION_KEY);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}

async function loadSession() {
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(DB_STORE, "readonly");
    const request = tx.objectStore(DB_STORE).get(SESSION_KEY);
    request.onsuccess = () => resolve(request.result || null);
    request.onerror = () => reject(request.error);
  });
}

async function clearSession() {
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(DB_STORE, "readwrite");
    tx.objectStore(DB_STORE).delete(SESSION_KEY);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}

function openIncomingDb() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(INCOMING_DB_NAME, 1);
    request.onupgradeneeded = () => request.result.createObjectStore(INCOMING_DB_STORE);
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

async function readIncomingFile() {
  const db = await openIncomingDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(INCOMING_DB_STORE, "readonly");
    const request = tx.objectStore(INCOMING_DB_STORE).get(INCOMING_FILE_KEY);
    request.onsuccess = () => resolve(request.result || null);
    request.onerror = () => reject(request.error);
  });
}

async function clearIncomingFile() {
  const db = await openIncomingDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(INCOMING_DB_STORE, "readwrite");
    tx.objectStore(INCOMING_DB_STORE).delete(INCOMING_FILE_KEY);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}

function normalizeText(value) {
  return String(value ?? "").trim();
}

function sanitizeFilePart(value) {
  return normalizeText(value)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/đ/g, "d")
    .replace(/Đ/g, "D")
    .replace(/[^a-zA-Z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "")
    .slice(0, 50);
}

function pad2(value) {
  return String(value || 0).padStart(2, "0");
}

function deriveLoHaTheFromSoCot(soCot) {
  const value = normalizeText(soCot);
  if (!value || value.startsWith("(")) return "";
  const match = value.match(/^(\d+)/);
  return match ? pad2(match[1]) : "";
}

function uid() {
  return crypto.randomUUID ? crypto.randomUUID() : `${Date.now()}-${Math.random()}`;
}

function nowText() {
  return new Date().toLocaleString("vi-VN");
}

function todayKey() {
  const now = new Date();
  return `${now.getFullYear()}${pad2(now.getMonth() + 1)}${pad2(now.getDate())}_${pad2(now.getHours())}${pad2(now.getMinutes())}`;
}

function createEmptySession() {
  return {
    officerName: "",
    unitName: "",
    stationName: "",
    sourceFileName: "",
    customers: [],
    otherCustomers: [],
    deletedCustomers: [],
    columns: {},
    logs: [],
    importedAt: "",
    updatedAt: "",
  };
}

function normalizeSessionShape(session) {
  return {
    ...createEmptySession(),
    ...(session || {}),
    customers: session?.customers || [],
    otherCustomers: session?.otherCustomers || [],
    deletedCustomers: session?.deletedCustomers || [],
    columns: session?.columns || {},
    logs: session?.logs || [],
  };
}

function getStationName(session = state.session) {
  return normalizeText(session.stationName) || normalizeText(session.customers?.[0]?.tram) || "Chưa rõ trạm";
}

function deriveStationName(customers) {
  const counts = new Map();
  (customers || []).forEach((customer) => {
    const tram = normalizeText(customer.tram);
    if (tram) counts.set(tram, (counts.get(tram) || 0) + 1);
  });
  return [...counts.entries()].sort((a, b) => b[1] - a[1])[0]?.[0] || "";
}

function addLog(session, action, soCot = "") {
  return {
    ...session,
    logs: [
      { time: nowText(), officerName: session.officerName || "Chưa nhập", soCot, action },
      ...(session.logs || []),
    ].slice(0, 300),
  };
}

function parseStandardWorkbook(workbook) {
  const sheet = workbook.Sheets[STANDARD_SHEET_NAME] || workbook.Sheets[workbook.SheetNames?.[0]];
  if (!sheet) throw new Error("File Excel không có sheet dữ liệu. App cần file theo mẫu File mau chuan.xlsx.");
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
  if (!rows.length) throw new Error("Sheet dữ liệu đang trống, không có dữ liệu khách hàng.");

  const headers = Object.keys(rows[0]).map(normalizeText);
  const missing = REQUIRED_COLUMNS.filter((col) => !headers.includes(col));
  if (missing.length) throw new Error(`File chưa đúng mẫu chuẩn, thiếu cột: ${missing.join(", ")}.`);

  return rows
    .map((row, index) => ({
      id: `${normalizeText(row.MA_KHANG)}-${index + 1}`,
      order: index + 1,
      maKhang: normalizeText(row.MA_KHANG),
      tenKhang: normalizeText(row.TEN_KHANG),
      soTbi: normalizeText(row.SO_TBI),
      soCot: normalizeText(row.SO_COT),
      tram: normalizeText(row.Tram),
      chuanHoaCu: normalizeText(row.CHUAN_HOA),
      note: normalizeText(row.GHI_CHU),
      raw: row,
    }))
    .filter((row) => row.maKhang && row.soCot);
}

function createColumnState(customers) {
  const grouped = {};
  customers.forEach((customer) => {
    grouped[customer.soCot] ||= [];
    grouped[customer.soCot].push(customer.id);
  });
  return Object.fromEntries(
    Object.entries(grouped).map(([soCot, ids]) => [
      soCot,
      { soCot, loHaThe: deriveLoHaTheFromSoCot(soCot), defaultPhase: "A", locations: suggestLocations(ids.length) },
    ]),
  );
}

function suggestLocations(count) {
  const locations = [];
  let remaining = count;
  let stt = 1;
  while (remaining > 0) {
    const type = remaining >= 6 ? BOX_TYPES[3] : remaining >= 4 ? BOX_TYPES[2] : remaining >= 2 ? BOX_TYPES[1] : BOX_TYPES[0];
    locations.push(createBoxLocation(stt, type.value));
    remaining -= type.capacity;
    stt += 1;
  }
  return locations;
}

function createBoxLocation(stt, boxType = "4") {
  const type = BOX_TYPES.find((item) => item.value === boxType) || BOX_TYPES[2];
  return {
    id: uid(),
    kind: "box",
    stt,
    boxType: type.value,
    capacity: type.capacity,
    cols: type.cols,
    phase: type.value === "3" ? "D" : "A",
    loHaThe: "",
    assignments: Array(type.capacity).fill(null),
  };
}

function createCabinetLocation(stt, cabinetType = "12") {
  const capacity = Number(cabinetType);
  return {
    id: uid(),
    kind: "cabinet",
    stt,
    cabinetFaces: "1",
    cabinetType,
    capacity,
    cols: capacity >= 12 ? 4 : 3,
    phase: "A",
    loHaThe: "",
    assignments: Array(capacity).fill(null),
  };
}

function locationCapacity(location) {
  return location.kind === "box" ? Number(location.capacity) : Number(location.cabinetType);
}

function makeCode(column, location, positionIndex) {
  const phase = location.phase || column.defaultPhase || "A";
  const lo = pad2(location.loHaThe || column.loHaThe);
  const pos = positionIndex + 1;
  if (location.kind === "box") return `${phase}${lo}${pad2(location.stt)}${location.boxType}${pos}`;
  return `${phase}${lo}${location.cabinetFaces}${location.cabinetType}${pad2(pos)}`;
}

function getAssignedCodes(session) {
  const codes = [];
  Object.values(session.columns || {}).forEach((column) => {
    (column.locations || []).forEach((location) => {
      (location.assignments || []).forEach((customerId, index) => {
        if (customerId) {
          codes.push({
            customerId,
            soCot: column.soCot,
            location,
            position: index + 1,
            code: makeCode(column, location, index),
          });
        }
      });
    });
  });
  return codes;
}

function makeError(soCot, customer, type, desc, suggest) {
  return { soCot, maKhang: customer?.maKhang || "", type, desc, suggest };
}

function validateSession(session) {
  const errors = [];
  const assigned = getAssignedCodes(session);
  const assignedCustomers = new Map();

  assigned.forEach((item) => {
    assignedCustomers.set(item.customerId, (assignedCustomers.get(item.customerId) || 0) + 1);
  });

  (session.customers || []).forEach((customer) => {
    const count = assignedCustomers.get(customer.id) || 0;
    if (count === 0) errors.push(makeError(customer.soCot, customer, "Thiếu mã", "Khách hàng chưa được gán mã chuẩn hóa.", "Mở cột và gán khách vào hòm/tủ."));
    if (count > 1) errors.push(makeError(customer.soCot, customer, "Gán nhiều vị trí", "Một khách hàng đang được gán nhiều hơn một vị trí.", "Chỉ giữ lại một vị trí cho khách này."));
  });

  return errors;
}

function columnStatus(session, soCot) {
  const customers = session.customers.filter((item) => item.soCot === soCot);
  const ids = new Set(customers.map((item) => item.id));
  const assigned = getAssignedCodes(session).filter((item) => ids.has(item.customerId));
  if (!assigned.length) return "Chưa làm";
  if (assigned.length < customers.length) return "Đang làm";
  return "Đã xong";
}

function getCustomersByColumn(session) {
  const grouped = {};
  (session.customers || []).forEach((customer) => {
    grouped[customer.soCot] ||= [];
    grouped[customer.soCot].push(customer);
  });
  return grouped;
}

function findOtherCustomerByDevice(soTbi) {
  const device = normalizeText(soTbi);
  if (!device) return null;
  return (state.session.otherCustomers || []).find((customer) => normalizeText(customer.soTbi) === device) || null;
}

function insertCustomerNearColumn(customers, customer, targetSoCot) {
  const list = [...(customers || [])];
  const lastIndex = list.reduce((foundIndex, item, index) => (item.soCot === targetSoCot ? index : foundIndex), -1);
  if (lastIndex >= 0) {
    list.splice(lastIndex + 1, 0, customer);
    return list.map((item, index) => ({ ...item, order: index + 1 }));
  }
  return [...list, customer].map((item, index) => ({ ...item, order: index + 1 }));
}

function setState(patch) {
  if (patch.session) patch = { ...patch, session: normalizeSessionShape(patch.session) };
  state = { ...state, ...patch };
  scheduleSave();
  render();
}

function updateSession(updater) {
  const next = updater(state.session);
  setState({ session: next });
}

function scheduleSave() {
  if (!state.session.customers.length) return;
  clearTimeout(saveTimer);
  state.saveState = "Đang lưu...";
  saveTimer = setTimeout(() => {
    saveSession({ ...state.session, updatedAt: new Date().toISOString() })
      .then(() => {
        state.saveState = `Đã tự lưu lúc ${nowText()}`;
        render();
      })
      .catch(() => {
        state.saveState = "Lưu lỗi";
        render();
      });
  }, 350);
}

function updateColumn(soCot, patcher) {
  updateSession((prev) => {
    const nextColumn = patcher(prev.columns[soCot]);
    return addLog({ ...prev, columns: { ...prev.columns, [soCot]: nextColumn } }, "Cập nhật cấu trúc hòm/tủ", soCot);
  });
}

function commitImportField(fieldName, value) {
  const nextValue = value ?? "";
  if (state.session[fieldName] === nextValue) return;
  state.session = normalizeSessionShape({ ...state.session, [fieldName]: nextValue });
  scheduleSave();
}

function bindCommitOnDone(input, fieldName) {
  if (!input) return;
  input.addEventListener("change", (event) => commitImportField(fieldName, event.target.value));
  input.addEventListener("blur", (event) => commitImportField(fieldName, event.target.value));
  input.addEventListener("keydown", (event) => {
    if (event.key === "Enter") {
      event.preventDefault();
      input.blur();
    }
  });
}

function escapeHtml(value) {
  return String(value ?? "").replace(/[&<>"']/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#039;" })[char]);
}

function render() {
  const app = document.getElementById("app");
  const errors = validateSession(state.session);
  app.innerHTML = `
    <div class="app-shell">
      <header class="topbar">
        <div>
          <p class="eyebrow">PWA hiện trường</p>
          <h1>Chuẩn hóa mã hòm công tơ</h1>
          <p class="subtitle">Nguồn dữ liệu chuẩn: file theo mẫu File mau chuan.xlsx, sheet Sheet1.</p>
        </div>
        <div class="top-actions">
          <div class="station-pill">Trạm: <strong>${escapeHtml(getStationName())}</strong></div>
          <div class="save-pill">💾 ${escapeHtml(state.saveState)}</div>
        </div>
      </header>
      ${state.message ? `<div class="notice">${escapeHtml(state.message)}</div>` : ""}
      <nav class="tabs">
        ${tabButton("import", "Nhập file", true)}
        ${tabButton("columns", "Danh sách cột", state.session.customers.length)}
        ${tabButton("work", "Chuẩn hóa", state.selectedColumn)}
        ${tabButton("check", "Kiểm tra lỗi", state.session.customers.length)}
      </nav>
      ${renderView(errors)}
    </div>`;
  bindEvents();
}

function tabButton(view, label, enabled) {
  return `<button data-view="${view}" class="${state.view === view ? "active" : ""}" ${enabled ? "" : "disabled"}>${label}</button>`;
}

function renderView(errors) {
  if (state.view === "columns") return renderColumns(errors);
  if (state.view === "work") return renderWork();
  if (state.view === "check") return renderCheck(errors);
  return renderImport();
}

function renderImport() {
  return `
    <section class="panel import-panel">
      <div class="field-grid">
        <label>Tên cán bộ
          <input id="officerName" value="${escapeHtml(state.session.officerName)}" placeholder="Ví dụ: Nguyễn Văn A">
        </label>
        <label>Đơn vị/đội
          <input id="unitName" value="${escapeHtml(state.session.unitName)}" placeholder="Ví dụ: Đội quản lý Tây Giang">
        </label>
        <label>Tên trạm đang làm
          <input id="stationName" value="${escapeHtml(getStationName() === "Chưa rõ trạm" ? "" : getStationName())}" placeholder="Ví dụ: Tây Giang 12">
        </label>
      </div>
      <label class="upload-box">
        <strong>📄 Chọn file Excel</strong>
        <span>App đọc file theo mẫu File mau chuan.xlsx</span>
        <input id="excelInput" type="file" accept=".xlsx,.xls">
      </label>
      <div class="mobile-file-help">
        <strong>Nếu file Excel được gửi qua Zalo/điện thoại:</strong>
        <ol>
          <li>Mở file trong Zalo rồi bấm <b>Lưu về máy</b> hoặc <b>Chia sẻ</b>.</li>
          <li>Khi bấm <b>Chọn file Excel</b>, vào <b>Gần đây</b>, <b>Tải xuống/Download</b> hoặc thư mục <b>Zalo</b>.</li>
          <li>Nếu đã cài app lên màn hình chính Android, có thể mở file trong Zalo → <b>Chia sẻ</b> → chọn <b>Chuẩn hóa CT</b>.</li>
        </ol>
      </div>
      <div class="actions">
        <button class="secondary" id="exportBackup" ${state.session.customers.length ? "" : "disabled"}>Tải backup JSON</button>
        <label class="button secondary">Khôi phục backup<input id="backupInput" type="file" accept=".json"></label>
        <button class="danger" id="resetAll">Xóa phiên</button>
      </div>
    </section>`;
}

function renderColumns(errors) {
  const grouped = getCustomersByColumn(state.session);
  const columns = Object.keys(state.session.columns)
    .filter((soCot) => soCot.toLowerCase().includes(state.search.toLowerCase()))
    .sort((a, b) => a.localeCompare(b, "vi", { numeric: true }));
  return `
    <section class="panel">
      <div class="summary-row">
        ${metric("Khách hàng", state.session.customers.length)}
        ${metric("Số cột", Object.keys(state.session.columns).length)}
        ${metric("Lỗi hiện tại", errors.length)}
      </div>
      <label class="search-box">Tìm số cột
        <input id="searchInput" value="${escapeHtml(state.search)}" placeholder="Ví dụ: 1.10">
      </label>
      <div class="column-list">
        ${columns
          .map((soCot) => {
            const status = columnStatus(state.session, soCot);
            const statusClass = status === "Đã xong" ? "done-card" : status === "Đang làm" ? "doing-card" : "";
            return `<button class="column-card ${statusClass} ${state.selectedColumn === soCot ? "selected" : ""}" data-column="${escapeHtml(soCot)}">
              <span><strong>Cột ${escapeHtml(soCot)}</strong><small>${state.session.columns[soCot]?.newSoCot ? `Cột mới: ${escapeHtml(state.session.columns[soCot].newSoCot)} - ` : ""}${grouped[soCot]?.length || 0} khách hàng</small></span>
              <em class="${status === "Đã xong" ? "done" : status === "Đang làm" ? "doing" : ""}">${status}</em>
            </button>`;
          })
          .join("")}
        <button class="column-card add-column-card" id="addColumn" type="button">
          <span><strong>+ Thêm cột</strong><small>Cột phát sinh ngoài hiện trường</small></span>
          <em>Thêm mới</em>
        </button>
      </div>
    </section>`;
}

function metric(label, value) {
  return `<div class="metric"><strong>${escapeHtml(value)}</strong><span>${escapeHtml(label)}</span></div>`;
}

function renderWork() {
  const column = state.session.columns[state.selectedColumn];
  if (!column) return `<section class="panel">Chưa chọn cột.</section>`;
  const customers = getCustomersByColumn(state.session)[state.selectedColumn] || [];
  const assignedIds = new Set(getAssignedCodes(state.session).filter((item) => item.soCot === state.selectedColumn).map((item) => item.customerId));
  const unassignedCustomers = customers.filter((customer) => !assignedIds.has(customer.id));
  const otherCustomers = state.session.otherCustomers || [];
  return `
    <section class="work-layout">
      <div class="panel control-panel">
        <div class="section-title">
          <div><p class="eyebrow">Đang chuẩn hóa</p><h2>Cột ${escapeHtml(state.selectedColumn)}</h2></div>
          <span>${customers.length} khách</span>
        </div>
        <div class="field-grid compact">
          <label>Số cột mới
            <input id="soCotEdit" value="${escapeHtml(column.newSoCot || "")}" placeholder="Nếu cột thực tế khác số cột gốc">
          </label>
          <label>Lộ mặc định
            <input id="loHaThe" maxlength="2" value="${escapeHtml(column.loHaThe)}">
          </label>
          <label>Pha mặc định
            <select id="defaultPhase">${PHASES.map((p) => `<option ${column.defaultPhase === p ? "selected" : ""}>${p}</option>`).join("")}</select>
          </label>
        </div>
        <div class="actions vertical">
          <button class="secondary" id="clearAssign">Xóa gán cột này</button>
        </div>
        <div class="customer-manager">
          ${
            state.showAddCustomerForm
              ? `<div class="add-customer-form">
                  <label>Số thiết bị *
                    <input id="newCustomerDevice" placeholder="Nhập số thiết bị thực tế">
                  </label>
                  <label>Tên khách hàng
                    <input id="newCustomerName" placeholder="Nếu có">
                  </label>
                  <div class="customer-form-actions">
                    <button id="addCustomer">Thêm khách</button>
                    <button class="secondary" id="cancelAddCustomer">Đóng</button>
                  </div>
                </div>`
              : `<button class="secondary" id="showAddCustomer">Thêm khách thực tế</button>`
          }
        </div>
        <div class="customer-pool">
          <h3>Khách chưa gán của cột</h3>
          ${
            unassignedCustomers.length
              ? unassignedCustomers
                  .map((c) => `<div class="customer-row">
                    <span>${escapeHtml(c.tenKhang || c.maKhang || "Khách chưa có tên")}</span>
                    <strong>${escapeHtml(c.soTbi)}</strong>
                    <small>Chưa gán</small>
                    <button class="danger remove-customer" data-customer="${escapeHtml(c.id)}">Xóa khách</button>
                  </div>`)
                  .join("")
              : `<div class="empty-mini">Đã gán hết khách trong cột.</div>`
          }
        </div>
        <div class="customer-pool other-customer-pool">
          <h3>Khách hàng khác</h3>
          ${
            otherCustomers.length
              ? otherCustomers
                  .map((c) => `<div class="customer-row other-customer-row">
                    <span>${escapeHtml(c.tenKhang || c.maKhang || "Khách chưa có tên")}</span>
                    <strong>${escapeHtml(c.soTbi)}</strong>
                    <small>Từ cột ${escapeHtml(c.originalSoCot || c.soCot || "")}</small>
                    <button class="restore-other-customer" data-customer="${escapeHtml(c.id)}">Thêm vào cột này</button>
                  </div>`)
                  .join("")
              : `<div class="empty-mini">Chưa có khách nào chuyển sang mục khác.</div>`
          }
        </div>
      </div>
      <div class="panel grid-panel">
        ${column.locations.map((location) => renderLocation(column, location, customers)).join("")}
        <div class="location-actions">
          <button class="secondary" id="addBox">Thêm hòm nếu thiếu</button>
          <button class="secondary" id="addCabinet">Thêm tủ</button>
          <button id="confirmColumn">Xác nhận cột này</button>
        </div>
      </div>
    </section>`;
}

function renderLocation(column, location, customers) {
  const capacity = locationCapacity(location);
  const gridCols = Math.min(location.cols || 2, 4);
  const assignedIds = new Set((column.locations || []).flatMap((item) => item.assignments || []).filter(Boolean));
  const optionFor = (currentId) =>
    customers
      .filter((c) => !assignedIds.has(c.id) || c.id === currentId)
      .map((c) => `<option value="${escapeHtml(c.id)}" ${c.id === currentId ? "selected" : ""}>${escapeHtml(c.tenKhang || c.maKhang || "Khách thêm tay")} - ${escapeHtml(c.soTbi)}</option>`)
      .join("");
  return `
    <article class="location-card" data-location="${escapeHtml(location.id)}">
      <div class="location-head">
        <div><strong>${location.kind === "box" ? `Hòm ${pad2(location.stt)}` : `Tủ ${pad2(location.stt)}`}</strong><span>${capacity} vị trí</span></div>
        <button class="icon-danger remove-location">Xóa</button>
      </div>
      <div class="location-options">
        <label>Pha
          <select class="phase-select">${PHASES.map((p) => `<option ${location.phase === p ? "selected" : ""}>${p}</option>`).join("")}</select>
        </label>
        ${
          location.kind === "box"
            ? `<label>Loại hòm
                <select class="box-type">${BOX_TYPES.map((t) => `<option value="${t.value}" ${location.boxType === t.value ? "selected" : ""}>${t.label}</option>`).join("")}</select>
              </label>`
            : `<label>Mặt tủ
                <select class="cabinet-faces"><option value="1" ${location.cabinetFaces === "1" ? "selected" : ""}>1</option><option value="2" ${location.cabinetFaces === "2" ? "selected" : ""}>2</option></select>
              </label>
              <label>Loại tủ
                <select class="cabinet-type">${CABINET_TYPES.map((t) => `<option value="${t}" ${location.cabinetType === t ? "selected" : ""}>${t}</option>`).join("")}</select>
              </label>`
        }
      </div>
      <div class="z-grid" style="grid-template-columns: repeat(${gridCols}, minmax(0, 1fr))">
        ${Array.from({ length: capacity })
          .map((_, index) => `
            <div class="slot">
              <b>${location.kind === "box" ? index + 1 : pad2(index + 1)}</b>
              <code>${escapeHtml(makeCode(column, location, index))}</code>
              <select class="assign-select" data-slot="${index}">
                <option value="">Chưa gán</option>
                ${optionFor(location.assignments[index] || "")}
              </select>
            </div>`)
          .join("")}
      </div>
    </article>`;
}

function renderCheck(errors) {
  return `
    <section class="panel">
      <div class="station-banner">Đang xuất file cho trạm: <strong>${escapeHtml(getStationName())}</strong></div>
      <div class="summary-row">
        ${metric("Tổng khách", state.session.customers.length)}
        ${metric("Đã gán", getAssignedCodes(state.session).length)}
        ${metric("Lỗi", errors.length)}
      </div>
      <div class="actions">
        <button id="shareDraft">Xuất nháp & chia sẻ Zalo</button>
        <button id="shareDone" ${errors.length ? "disabled" : ""}>Xuất & chia sẻ Zalo</button>
        <button class="secondary" id="exportBackup2">Backup JSON</button>
      </div>
      <div class="error-list">
        ${
          errors.length === 0
            ? `<div class="empty-state">✅ Không còn lỗi. Có thể xuất file hoàn thành.</div>`
            : errors
                .slice(0, 150)
                .map((e) => `<div class="error-row"><strong>${escapeHtml(e.type)}</strong><span>Cột ${escapeHtml(e.soCot)} - ${escapeHtml(e.maKhang)}</span><p>${escapeHtml(e.desc)}</p><small>${escapeHtml(e.suggest)}</small></div>`)
                .join("")
        }
      </div>
    </section>`;
}

function bindEvents() {
  document.querySelectorAll("[data-view]").forEach((btn) => {
    btn.addEventListener("click", () => setState({ view: btn.dataset.view }));
  });

  const officerName = document.getElementById("officerName");
  const unitName = document.getElementById("unitName");
  const stationName = document.getElementById("stationName");
  bindCommitOnDone(officerName, "officerName");
  bindCommitOnDone(unitName, "unitName");
  bindCommitOnDone(stationName, "stationName");
  const excelInput = document.getElementById("excelInput");
  if (excelInput) excelInput.addEventListener("change", handleImport);
  const backupInput = document.getElementById("backupInput");
  if (backupInput) backupInput.addEventListener("change", importBackup);
  const resetAll = document.getElementById("resetAll");
  if (resetAll) resetAll.addEventListener("click", resetSession);
  const exportBackup2 = document.getElementById("exportBackup2");
  if (exportBackup2) exportBackup2.textContent = "Backup Zalo";
  const exportBackup = document.getElementById("exportBackup");
  if (exportBackup) exportBackup.addEventListener("click", () => downloadBackup(false));
  if (exportBackup2) exportBackup2.addEventListener("click", () => downloadBackup(true));

  const searchInput = document.getElementById("searchInput");
  if (searchInput) searchInput.addEventListener("input", (e) => setState({ search: e.target.value }));
  document.querySelectorAll("[data-column]").forEach((btn) => {
    btn.addEventListener("click", () => setState({ selectedColumn: btn.dataset.column, view: "work" }));
  });
  const addColumn = document.getElementById("addColumn");
  if (addColumn) addColumn.addEventListener("click", addManualColumn);

  const loHaThe = document.getElementById("loHaThe");
  if (loHaThe) loHaThe.addEventListener("input", (e) => updateColumn(state.selectedColumn, (col) => ({ ...col, loHaThe: e.target.value.replace(/\D/g, "").slice(0, 2) })));
  const soCotEdit = document.getElementById("soCotEdit");
  if (soCotEdit) {
    soCotEdit.addEventListener("change", (e) => renameSelectedColumn(e.target.value));
    soCotEdit.addEventListener("keydown", (e) => {
      if (e.key === "Enter") {
        e.preventDefault();
        renameSelectedColumn(e.target.value);
      }
    });
  }
  const defaultPhase = document.getElementById("defaultPhase");
  if (defaultPhase) defaultPhase.addEventListener("change", (e) => updateColumn(state.selectedColumn, (col) => ({ ...col, defaultPhase: e.target.value })));
  const autoAssign = document.getElementById("autoAssign");
  if (autoAssign) autoAssign.addEventListener("click", autoAssignColumn);
  const addBox = document.getElementById("addBox");
  if (addBox) addBox.addEventListener("click", () => addLocation("box"));
  const addCabinet = document.getElementById("addCabinet");
  if (addCabinet) addCabinet.addEventListener("click", () => addLocation("cabinet"));
  const clearAssign = document.getElementById("clearAssign");
  if (clearAssign) clearAssign.addEventListener("click", clearAssignments);
  const confirmColumn = document.getElementById("confirmColumn");
  if (confirmColumn) confirmColumn.addEventListener("click", confirmSelectedColumn);
  const showAddCustomer = document.getElementById("showAddCustomer");
  if (showAddCustomer) showAddCustomer.addEventListener("click", () => setState({ showAddCustomerForm: true }));
  const cancelAddCustomer = document.getElementById("cancelAddCustomer");
  if (cancelAddCustomer) cancelAddCustomer.addEventListener("click", () => setState({ showAddCustomerForm: false }));
  const addCustomer = document.getElementById("addCustomer");
  if (addCustomer) addCustomer.addEventListener("click", addManualCustomer);
  document.querySelectorAll(".remove-customer").forEach((btn) => {
    btn.addEventListener("click", () => removeCustomer(btn.dataset.customer));
  });
  document.querySelectorAll(".restore-other-customer").forEach((btn) => {
    btn.addEventListener("click", () => restoreOtherCustomer(btn.dataset.customer));
  });

  document.querySelectorAll(".location-card").forEach((card) => bindLocationEvents(card));

  const shareDraft = document.getElementById("shareDraft");
  if (shareDraft) shareDraft.addEventListener("click", () => exportWorkbook(true, true));
  const shareDone = document.getElementById("shareDone");
  if (shareDone) shareDone.addEventListener("click", () => exportWorkbook(false, true));
}

function bindLocationEvents(card) {
  const id = card.dataset.location;
  card.querySelector(".remove-location")?.addEventListener("click", () => removeLocation(id));
  card.querySelector(".location-lo")?.addEventListener("input", (e) => updateLocation(id, { loHaThe: e.target.value.replace(/\D/g, "").slice(0, 2) }));
  card.querySelector(".phase-select")?.addEventListener("change", (e) => updateLocation(id, { phase: e.target.value }));
  card.querySelector(".box-type")?.addEventListener("change", (e) => {
    const type = BOX_TYPES.find((item) => item.value === e.target.value);
    updateLocation(id, { boxType: type.value, capacity: type.capacity, cols: type.cols, ...(type.value === "3" ? { phase: "D" } : {}) });
  });
  card.querySelector(".cabinet-faces")?.addEventListener("change", (e) => updateLocation(id, { cabinetFaces: e.target.value }));
  card.querySelector(".cabinet-type")?.addEventListener("change", (e) => updateLocation(id, { cabinetType: e.target.value, capacity: Number(e.target.value), cols: Number(e.target.value) >= 12 ? 4 : 3 }));
  card.querySelectorAll(".assign-select").forEach((select) => {
    select.addEventListener("change", (e) => assignSlot(id, Number(select.dataset.slot), e.target.value));
  });
}

async function handleImport(event) {
  const file = event.target.files?.[0];
  if (!file) return;
  await importWorkbookFile(file);
  event.target.value = "";
}

async function importWorkbookFile(file) {
  setState({ message: "Đang đọc file theo mẫu chuẩn..." });
  try {
    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data, { type: "array" });
    const customers = parseStandardWorkbook(workbook);
    const stationName = normalizeText(state.session.stationName) || deriveStationName(customers);
    const nextSession = addLog(
      {
        ...createEmptySession(),
        officerName: state.session.officerName,
        unitName: state.session.unitName,
        stationName,
        sourceFileName: file.name,
        customers,
        columns: createColumnState(customers),
        importedAt: new Date().toISOString(),
      },
      `Import ${customers.length} khách hàng từ file mẫu chuẩn${stationName ? ` - trạm ${stationName}` : ""}`,
    );
    const firstColumn = Object.keys(nextSession.columns)[0] || "";
    setState({ session: nextSession, selectedColumn: firstColumn, view: "columns", message: `Đã nhập ${customers.length} khách hàng từ file mẫu chuẩn${stationName ? ` cho trạm ${stationName}` : ""}.` });
  } catch (error) {
    setState({ message: error.message || "Không đọc được file Excel." });
  }
}

async function importSharedFileIfAny() {
  if (!new URLSearchParams(window.location.search).has("shared")) return;
  try {
    const incoming = await readIncomingFile();
    if (!incoming?.blob) {
      setState({ message: "Chưa nhận được file Excel chia sẻ. Sếp mở lại file trong Zalo và chọn Chia sẻ vào app." });
      return;
    }
    const file = new File([incoming.blob], incoming.name || "file_chia_se.xlsx", {
      type: incoming.type || "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    await clearIncomingFile();
    await importWorkbookFile(file);
    window.history.replaceState({}, "", "/");
  } catch (error) {
    setState({ message: "Không nhận được file chia sẻ. Sếp thử bấm Chọn file Excel và vào mục Gần đây/Tải xuống." });
  }
}

function autoAssignColumn() {
  const customers = getCustomersByColumn(state.session)[state.selectedColumn] || [];
  updateColumn(state.selectedColumn, (column) => {
    const ids = customers.map((item) => item.id);
    let cursor = 0;
    return {
      ...column,
      locations: column.locations.map((location) => {
        const assignments = Array(locationCapacity(location)).fill(null).map(() => ids[cursor++] || null);
        return { ...location, assignments };
      }),
    };
  });
  state.message = `Đã tự gán khách hàng theo thứ tự Excel cho cột ${state.selectedColumn}.`;
}

function renameSelectedColumn(newSoCotRaw) {
  const oldSoCot = state.selectedColumn;
  const newSoCot = normalizeText(newSoCotRaw);
  if (!oldSoCot) {
    render();
    return;
  }

  const currentColumn = state.session.columns[oldSoCot];
  if (!newSoCot || newSoCot === oldSoCot) {
    const nextSession = addLog(
      {
        ...state.session,
        columns: {
          ...state.session.columns,
          [oldSoCot]: { ...currentColumn, newSoCot: "", loHaThe: deriveLoHaTheFromSoCot(oldSoCot) },
        },
      },
      `Giữ nguyên số cột ${oldSoCot}`,
      oldSoCot,
    );
    setState({ session: nextSession, selectedColumn: oldSoCot, message: `Đã giữ nguyên số cột ${oldSoCot}.` });
    return;
  }

  const nextSession = addLog(
    {
      ...state.session,
      columns: {
        ...state.session.columns,
        [oldSoCot]: {
          ...currentColumn,
          newSoCot,
          loHaThe: deriveLoHaTheFromSoCot(newSoCot),
          columnChangeNote: `Thay đổi số cột từ ${oldSoCot} sang ${newSoCot}`,
        },
      },
    },
    `Ghi số cột mới ${newSoCot} cho cột gốc ${oldSoCot}`,
    oldSoCot,
  );
  setState({ session: nextSession, selectedColumn: oldSoCot, message: `Đã ghi cột mới ${newSoCot}. Cột gốc ${oldSoCot} vẫn được giữ, khi xuất Excel sẽ có cột COT_MOI và ghi chú thay đổi số cột.` });
}

function addManualColumn() {
  const soCot = normalizeText(window.prompt("Nhập số cột mới cần thêm:", ""));
  if (!soCot) {
    setState({ message: "Chưa thêm cột vì chưa nhập số cột." });
    return;
  }
  if (state.session.columns[soCot]) {
    setState({ selectedColumn: soCot, view: "work", message: `Cột ${soCot} đã có trong danh sách, em mở cột này để sếp làm tiếp.` });
    return;
  }
  const column = {
    soCot,
    loHaThe: deriveLoHaTheFromSoCot(soCot),
    defaultPhase: "A",
    locations: [createBoxLocation(1, "1")],
    manual: true,
  };
  const nextSession = addLog(
    { ...state.session, columns: { ...state.session.columns, [soCot]: column } },
    `Them cot moi tai hien truong: ${soCot}`,
    soCot,
  );
  setState({ session: nextSession, selectedColumn: soCot, view: "work", showAddCustomerForm: true, message: `Đã thêm cột ${soCot}. Sếp nhập Số thiết bị ở phần Thêm khách thực tế.` });
}

function clearAssignments() {
  updateColumn(state.selectedColumn, (column) => ({
    ...column,
    locations: column.locations.map((location) => ({ ...location, assignments: Array(locationCapacity(location)).fill(null) })),
  }));
}

function addLocation(kind) {
  updateColumn(state.selectedColumn, (column) => {
    const stt = column.locations.length + 1;
    return { ...column, locations: [...column.locations, kind === "box" ? createBoxLocation(stt, "4") : createCabinetLocation(stt, "12")] };
  });
}

function confirmSelectedColumn() {
  setState({ view: "columns", message: `Đã xác nhận cột ${state.selectedColumn}. Sếp chọn cột tiếp theo để làm tiếp.` });
}

function moveOtherCustomerToColumn(session, customerId, targetSoCot, patch = {}) {
  const otherCustomer = (session.otherCustomers || []).find((item) => item.id === customerId);
  if (!otherCustomer) return session;
  const sourceSoCot = otherCustomer.originalSoCot || otherCustomer.soCot || "";
  const customer = {
    ...otherCustomer,
    ...patch,
    soCot: targetSoCot,
    originalSoCot: sourceSoCot,
    restoredAt: nowText(),
    note: normalizeText(patch.note) || `Chuyen tu cot ${sourceSoCot || "khac"} vao cot ${targetSoCot}`,
  };
  delete customer.movedToOtherAt;
  delete customer.movedReason;
  return {
    ...session,
    customers: insertCustomerNearColumn(session.customers || [], customer, targetSoCot),
    otherCustomers: (session.otherCustomers || []).filter((item) => item.id !== customerId),
  };
}

function restoreOtherCustomer(customerId) {
  const customer = (state.session.otherCustomers || []).find((item) => item.id === customerId);
  if (!customer || !state.selectedColumn) return;
  const nextSession = addLog(
    moveOtherCustomerToColumn(state.session, customerId, state.selectedColumn),
    `Dua khach STB ${customer.soTbi} tu muc khach hang khac vao cot ${state.selectedColumn}`,
    state.selectedColumn,
  );
  setState({ session: nextSession, message: `Da dua so thiet bi ${customer.soTbi} vao cot ${state.selectedColumn}.` });
}

function addManualCustomer() {
  const tenKhang = normalizeText(document.getElementById("newCustomerName")?.value);
  const soTbi = normalizeText(document.getElementById("newCustomerDevice")?.value);
  if (!soTbi) {
    setState({ message: "Vui lòng nhập Số thiết bị trước khi thêm khách." });
    return;
  }

  const matchedOtherCustomer = findOtherCustomerByDevice(soTbi);
  if (matchedOtherCustomer) {
    const nextSession = addLog(
      moveOtherCustomerToColumn(state.session, matchedOtherCustomer.id, state.selectedColumn, tenKhang ? { tenKhang } : {}),
      `Tu dong dua khach STB ${soTbi} tu muc khach hang khac vao cot ${state.selectedColumn}`,
      state.selectedColumn,
    );
    setState({ session: nextSession, showAddCustomerForm: false, message: `Da tim thay so thiet bi ${soTbi} trong muc khach hang khac va dua vao cot hien tai.` });
    return;
  }

  const customer = {
    id: `them-tay-${uid()}`,
    order: state.session.customers.length + 1,
    maKhang: `THEM_TAY_${todayKey()}`,
    tenKhang: tenKhang || "Khách thêm tay",
    soTbi,
    soCot: state.selectedColumn,
    tram: getStationName() === "Chưa rõ trạm" ? "" : getStationName(),
    manual: true,
    note: "Thêm mới",
    raw: {},
  };

  const nextSession = addLog(
    { ...state.session, customers: insertCustomerNearColumn(state.session.customers, customer, state.selectedColumn) },
    `Thêm khách thực tế: ${customer.maKhang} - STB ${customer.soTbi}`,
    state.selectedColumn,
  );
  setState({ session: nextSession, showAddCustomerForm: false, message: `Đã thêm khách thực tế với số thiết bị ${soTbi}.` });
}

function removeCustomer(customerId) {
  const customer = state.session.customers.find((item) => item.id === customerId);
  if (!customer) return;
  const moveReason = normalizeText(window.prompt(`Nhap ly do chuyen khach ${customer.tenKhang || customer.maKhang} sang muc khach hang khac:`, "Khong dung cot nay"));
  if (!moveReason) {
    setState({ message: "Chua chuyen khach vi chua nhap ly do." });
    return;
  }
  const columnsAfterMove = Object.fromEntries(
    Object.entries(state.session.columns).map(([soCot, column]) => [
      soCot,
      {
        ...column,
        locations: (column.locations || []).map((location) => ({
          ...location,
          assignments: (location.assignments || []).map((assignedId) => (assignedId === customerId ? null : assignedId)),
        })),
      },
    ]),
  );
  const sessionAfterMove = addLog(
    {
      ...state.session,
      customers: state.session.customers.filter((item) => item.id !== customerId),
      otherCustomers: [
        {
          ...customer,
          originalSoCot: customer.originalSoCot || customer.soCot,
          movedToOtherAt: nowText(),
          movedReason: moveReason,
          note: `Khach hang khac - ly do: ${moveReason}`,
        },
        ...(state.session.otherCustomers || []).filter((item) => item.id !== customerId),
      ],
      columns: columnsAfterMove,
    },
    `Chuyen khach ${customer.maKhang} - STB ${customer.soTbi} sang muc khach hang khac. Ly do: ${moveReason}`,
    customer.soCot,
  );
  setState({ session: sessionAfterMove, message: `Da chuyen khach ${customer.maKhang} sang muc khach hang khac. Sang cot dung co the bam Them vao cot nay.` });
}

function removeLocation(locationId) {
  updateColumn(state.selectedColumn, (column) => ({
    ...column,
    locations: column.locations.filter((item) => item.id !== locationId).map((item, index) => ({ ...item, stt: index + 1 })),
  }));
}

function updateLocation(locationId, patch) {
  updateColumn(state.selectedColumn, (column) => ({
    ...column,
    locations: column.locations.map((location) => {
      if (location.id !== locationId) return location;
      const next = { ...location, ...patch };
      const capacity = locationCapacity(next);
      const assignments = [...(location.assignments || [])].slice(0, capacity);
      while (assignments.length < capacity) assignments.push(null);
      return { ...next, capacity, assignments };
    }),
  }));
}

function assignSlot(locationId, slotIndex, customerId) {
  updateColumn(state.selectedColumn, (column) => ({
    ...column,
    locations: column.locations.map((location) => {
      if (location.id !== locationId) return location;
      const assignments = [...location.assignments];
      assignments[slotIndex] = customerId || null;
      return { ...location, assignments };
    }),
  }));
}

async function resetSession() {
  await clearSession();
  setState({ session: createEmptySession(), selectedColumn: "", view: "import", message: "Đã xóa phiên làm việc trên máy.", saveState: "Chưa có dữ liệu" });
}

function downloadBlob(blob, fileName) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = fileName;
  a.click();
  URL.revokeObjectURL(url);
}

async function shareOrDownloadBlob(blob, fileName, title, text) {
  const file = new File([blob], fileName, { type: blob.type || "application/octet-stream" });
  if (navigator.canShare?.({ files: [file] })) {
    try {
      await navigator.share({ title, text, files: [file] });
      setState({ message: "Đã mở bảng chia sẻ file. Sếp chọn Zalo để gửi file." });
      return true;
    } catch (error) {
      if (error?.name === "AbortError") {
        setState({ message: "Sếp đã đóng bảng chia sẻ, file chưa được gửi." });
        return false;
      }
    }
  }
  downloadBlob(blob, fileName);
  setState({ message: "Máy/trình duyệt này chưa hỗ trợ mở bảng chia sẻ file. Em đã tải file về máy, sếp mở Zalo rồi chọn file vừa tải để gửi." });
  return false;
}

async function downloadBackup(shareAfter = false) {
  const blob = new Blob([JSON.stringify(state.session, null, 2)], { type: "application/json" });
  const fileName = `backup_chuan_hoa_${Date.now()}.json`;
  if (shareAfter) {
    await shareOrDownloadBlob(blob, fileName, "Backup Zalo", "File backup chuan hoa cong to.");
    return;
  }
  downloadBlob(blob, fileName);
}

async function importBackup(event) {
  const file = event.target.files?.[0];
  if (!file) return;
  try {
    const data = JSON.parse(await file.text());
    if (!data.customers || !data.columns) throw new Error("File backup không đúng định dạng.");
    setState({ session: addLog(data, "Khôi phục từ file backup JSON"), selectedColumn: Object.keys(data.columns)[0] || "", view: "columns", message: "Đã khôi phục backup." });
  } catch (error) {
    setState({ message: error.message || "Không import được backup." });
  }
}

function markDeletedRowsRed(sheet, activeCount, deletedCount) {
  if (!deletedCount) return;
  const redFill = { patternType: "solid", fgColor: { rgb: "FFFFCDD2" } };
  const redFont = { color: { rgb: "FF991B1B" }, bold: true };
  for (let rowIndex = activeCount + 2; rowIndex <= activeCount + deletedCount + 1; rowIndex += 1) {
    for (let colIndex = 0; colIndex < EXPORT_COLUMNS.length; colIndex += 1) {
      const cellRef = XLSX.utils.encode_cell({ r: rowIndex - 1, c: colIndex });
      sheet[cellRef] ||= { t: "s", v: "" };
      sheet[cellRef].s = { fill: redFill, font: redFont };
    }
  }
}

function getColumnChange(customer) {
  const column = state.session.columns?.[customer.soCot];
  const newSoCot = normalizeText(column?.newSoCot);
  if (!newSoCot || newSoCot === customer.soCot) return { newSoCot: "", note: "" };
  return { newSoCot, note: column?.columnChangeNote || `Thay đổi số cột từ ${customer.soCot} sang ${newSoCot}` };
}

function buildCustomerNote(customer) {
  const columnChange = getColumnChange(customer);
  if (columnChange.note) return columnChange.note;
  if (customer.manual) return "Thêm mới";
  return normalizeText(customer.note);
}

function buildWorkbook(onlyDraft) {
  const errors = validateSession(state.session);
  const assigned = getAssignedCodes(state.session);
  const assignedByCustomer = new Map(assigned.map((item) => [item.customerId, item]));

  const activeStandardRows = state.session.customers.map((customer) => {
    const item = assignedByCustomer.get(customer.id);
    const columnChange = getColumnChange(customer);
    return {
      MA_KHANG: customer.maKhang,
      TEN_KHANG: customer.tenKhang,
      SO_TBI: customer.soTbi,
      SO_COT: customer.soCot,
      Tram: customer.tram,
      CHUAN_HOA: item?.code || customer.chuanHoaCu || "",
      COT_MOI: columnChange.newSoCot,
      GHI_CHU: buildCustomerNote(customer),
    };
  });
  const deletedStandardRows = (state.session.deletedCustomers || []).map((customer) => ({
    MA_KHANG: customer.maKhang,
    TEN_KHANG: customer.tenKhang,
    SO_TBI: customer.soTbi,
    SO_COT: customer.soCot,
    Tram: customer.tram,
    CHUAN_HOA: customer.chuanHoaCu || "",
    COT_MOI: "",
    GHI_CHU: customer.note || `Đã xóa - Lý do: ${customer.deleteReason || ""}`,
  }));
  const otherStandardRows = (state.session.otherCustomers || []).map((customer) => ({
    MA_KHANG: customer.maKhang,
    TEN_KHANG: customer.tenKhang,
    SO_TBI: customer.soTbi,
    SO_COT: customer.originalSoCot || customer.soCot,
    Tram: customer.tram,
    CHUAN_HOA: customer.chuanHoaCu || "",
    COT_MOI: "",
    GHI_CHU: customer.note || `Khach hang khac - ly do: ${customer.movedReason || ""}`,
  }));
  const standardRows = [...activeStandardRows, ...otherStandardRows, ...deletedStandardRows];

  const dataRows = state.session.customers.map((customer) => {
    const item = assignedByCustomer.get(customer.id);
    const location = item?.location;
    return {
      "Số cột": customer.soCot,
      "Số cột mới": getColumnChange(customer).newSoCot,
      "Mã khách hàng": customer.maKhang,
      "Tên khách hàng": customer.tenKhang,
      "Số thiết bị": customer.soTbi,
      "Trạm": customer.tram,
      "Pha": location?.phase || "",
      "Lộ hạ thế": item ? pad2(location?.loHaThe || state.session.columns[customer.soCot]?.loHaThe) : "",
      "Loại vị trí": location ? (location.kind === "box" ? "Hòm cột" : "Tủ thanh cái") : "",
      "STT hòm": location?.kind === "box" ? pad2(location.stt) : "",
      "Loại hòm": location?.kind === "box" ? location.boxType : "",
      "Số mặt tủ": location?.kind === "cabinet" ? location.cabinetFaces : "",
      "Loại tủ": location?.kind === "cabinet" ? location.cabinetType : "",
      "Vị trí khách hàng": item?.position || "",
      "Mã chuẩn hóa": item?.code || "",
      "Trạng thái": item ? "Đã gán" : "Chưa gán",
    };
  });

  const summaryRows = Object.keys(state.session.columns).map((soCot) => {
    const total = state.session.customers.filter((item) => item.soCot === soCot).length;
    const ids = new Set(state.session.customers.filter((item) => item.soCot === soCot).map((item) => item.id));
    const assignedCount = assigned.filter((item) => ids.has(item.customerId)).length;
    const columnErrors = errors.filter((item) => item.soCot === soCot).length;
    const locations = state.session.columns[soCot].locations;
    return {
      "Tên trạm": getStationName(),
      "Số cột": soCot,
      "Số cột mới": state.session.columns[soCot].newSoCot || "",
      "Tổng khách hàng": total,
      "Đã gán": assignedCount,
      "Chưa gán": total - assignedCount,
      "Số hòm": locations.filter((item) => item.kind === "box").length,
      "Số tủ": locations.filter((item) => item.kind === "cabinet").length,
      "Trạng thái": columnStatus(state.session, soCot),
      "Có lỗi hay không": columnErrors ? "Có" : "Không",
    };
  });

  const errorRows = errors.map((error) => ({
    "Tên trạm": getStationName(),
    "Số cột": error.soCot,
    "Mã khách hàng": error.maKhang,
    "Loại lỗi": error.type,
    "Mô tả lỗi": error.desc,
    "Gợi ý xử lý": error.suggest,
  }));

  const logRows = state.session.logs.map((log) => ({
    "Thời gian": log.time,
    "Tên trạm": getStationName(),
    "Tên cán bộ": log.officerName,
    "Số cột": log.soCot,
    "Hành động": log.action,
  }));

  const otherRows = (state.session.otherCustomers || []).map((customer) => ({
    "So cot goc": customer.originalSoCot || customer.soCot,
    "Ma khach hang": customer.maKhang,
    "Ten khach hang": customer.tenKhang,
    "So thiet bi": customer.soTbi,
    "Ly do": customer.movedReason || "",
    "Thoi gian chuyen": customer.movedToOtherAt || "",
  }));

  const wb = XLSX.utils.book_new();
  const standardSheet = XLSX.utils.json_to_sheet(standardRows, { header: EXPORT_COLUMNS });
  markDeletedRowsRed(standardSheet, activeStandardRows.length + otherStandardRows.length, deletedStandardRows.length);
  XLSX.utils.book_append_sheet(wb, standardSheet, STANDARD_SHEET_NAME);
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(dataRows), "DU_LIEU_CHUAN_HOA");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(summaryRows), "TONG_HOP_THEO_COT");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(errorRows), "DANH_SACH_LOI");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(otherRows), "KHACH_HANG_KHAC");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(logRows), "NHAT_KY_THAO_TAC");
  const suffix = onlyDraft ? "BAN_NHAP" : "HOAN_THANH";
  const stationPart = sanitizeFilePart(getStationName()) || "chua_ro_tram";
  return { wb, fileName: `chuan_hoa_${stationPart}_${suffix}_${todayKey()}.xlsx` };
}

async function exportWorkbook(onlyDraft, shareAfter = false) {
  const { wb, fileName } = buildWorkbook(onlyDraft);
  const arrayBufferForShare = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  const blobForShare = new Blob([arrayBufferForShare], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  if (!shareAfter) {
    downloadBlob(blobForShare, fileName);
    return;
  }
  await shareOrDownloadBlob(
    blobForShare,
    fileName,
    onlyDraft ? `Ban nhap tram ${getStationName()}` : `Hoan thanh tram ${getStationName()}`,
    `${onlyDraft ? "Ban nhap" : "File hoan thanh"} chuan hoa cong to - tram ${getStationName()}.`,
  );
}

window.addEventListener("beforeunload", (event) => {
  if (state.session.customers.length) {
    event.preventDefault();
    event.returnValue = "";
  }
});

if ("serviceWorker" in navigator) navigator.serviceWorker.register("/sw.js").catch(() => {});

loadSession()
  .then((saved) => {
    if (saved?.customers?.length) {
      state.session = saved;
      state.selectedColumn = Object.keys(saved.columns)[0] || "";
      state.view = "columns";
      state.message = "Đã khôi phục phiên làm việc trước.";
    }
    render();
    importSharedFileIfAny();
  })
  .catch(() => {
    state.message = "Không đọc được bản lưu cũ.";
    render();
  });
