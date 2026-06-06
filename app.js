const STATUS_COLUMN_INDEX = 11; // Excel L sütunu: A=0, L=11
const STORAGE_KEY = "assessor_list_imported_rows_v1";
const STORAGE_HEADERS_KEY = "assessor_list_headers_v1";

const state = {
  headers: [],
  rows: [],
  filter: "ALL",
  search: ""
};

const el = {
  fileInput: document.getElementById("fileInput"),
  dropZone: document.getElementById("dropZone"),
  fileName: document.getElementById("fileName"),
  totalCount: document.getElementById("totalCount"),
  activeCount: document.getElementById("activeCount"),
  cancelCount: document.getElementById("cancelCount"),
  suspendCount: document.getElementById("suspendCount"),
  searchInput: document.getElementById("searchInput"),
  filterButtons: document.querySelectorAll(".filter-btn"),
  clearBtn: document.getElementById("clearBtn"),
  emptyState: document.getElementById("emptyState"),
  tableWrap: document.getElementById("tableWrap"),
  tableHead: document.getElementById("tableHead"),
  tableBody: document.getElementById("tableBody")
};

boot();

function boot() {
  attachEvents();
  loadStoredData();
  render();
}

function attachEvents() {
  el.fileInput.addEventListener("change", event => {
    const file = event.target.files?.[0];
    if (file) importFile(file);
  });

  ["dragenter", "dragover"].forEach(eventName => {
    el.dropZone.addEventListener(eventName, event => {
      event.preventDefault();
      el.dropZone.classList.add("drag-over");
    });
  });

  ["dragleave", "drop"].forEach(eventName => {
    el.dropZone.addEventListener(eventName, event => {
      event.preventDefault();
      el.dropZone.classList.remove("drag-over");
    });
  });

  el.dropZone.addEventListener("drop", event => {
    const file = event.dataTransfer.files?.[0];
    if (file) importFile(file);
  });

  el.searchInput.addEventListener("input", event => {
    state.search = normalizeText(event.target.value);
    renderTable();
  });

  el.filterButtons.forEach(button => {
    button.addEventListener("click", () => {
      state.filter = button.dataset.filter;
      el.filterButtons.forEach(btn => btn.classList.remove("active"));
      button.classList.add("active");
      renderTable();
    });
  });

  el.clearBtn.addEventListener("click", () => {
    state.headers = [];
    state.rows = [];
    state.filter = "ALL";
    state.search = "";
    el.searchInput.value = "";
    el.fileName.textContent = "Liste temizlendi. Yeni Excel dosyası import edebilirsiniz.";
    localStorage.removeItem(STORAGE_KEY);
    localStorage.removeItem(STORAGE_HEADERS_KEY);
    resetFilterButtons();
    render();
  });
}

async function importFile(file) {
  if (!window.XLSX) {
    alert("XLSX kütüphanesi yüklenemedi. İnternet bağlantısını kontrol edin veya SheetJS dosyasını yerel olarak ekleyin.");
    return;
  }

  try {
    const arrayBuffer = await file.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: "array", cellDates: false });
    const firstSheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[firstSheetName];
    const matrix = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      defval: "",
      blankrows: false,
      raw: false
    });

    if (!matrix.length) {
      throw new Error("Excel sayfası boş görünüyor.");
    }

    const headers = matrix[0].map((value, index) => cleanHeader(value, index));
    const rows = matrix
      .slice(1)
      .filter(row => row.some(cell => String(cell).trim() !== ""))
      .map(row => normalizeRowLength(row, headers.length));

    if (!headers[STATUS_COLUMN_INDEX]) {
      throw new Error("L sütunu bulunamadı. Dosya formatını kontrol edin.");
    }

    state.headers = headers;
    state.rows = rows;
    state.filter = "ALL";
    state.search = "";
    el.searchInput.value = "";
    resetFilterButtons();

    persistData();
    el.fileName.textContent = `${file.name} import edildi. Kayıt sayısı: ${rows.length}`;
    render();
  } catch (error) {
    console.error(error);
    alert(`Dosya okunamadı: ${error.message}`);
  } finally {
    el.fileInput.value = "";
  }
}

function cleanHeader(value, index) {
  const text = String(value ?? "").trim();
  return text || `Sütun ${index + 1}`;
}

function normalizeRowLength(row, length) {
  const normalized = Array.from({ length }, (_, index) => row[index] ?? "");
  return normalized.map(value => String(value ?? "").trim());
}

function normalizeStatus(value) {
  const normalized = String(value ?? "").trim().toUpperCase();

  if (normalized === "ACTIVE") return "ACTIVE";
  if (normalized === "CANCEL" || normalized === "CANCELLED") return "CANCEL";
  if (normalized === "SUSPEND" || normalized === "SUSPENDED") return "SUSPEND";

  return "UNKNOWN";
}

function getStatusClass(status) {
  return `status-${status.toLowerCase()}`;
}

function normalizeText(value) {
  return String(value ?? "")
    .toLocaleLowerCase("tr-TR")
    .replace(/ı/g, "i")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim();
}

function persistData() {
  try {
    localStorage.setItem(STORAGE_HEADERS_KEY, JSON.stringify(state.headers));
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state.rows));
  } catch (error) {
    console.warn("Liste localStorage içine kaydedilemedi.", error);
  }
}

function loadStoredData() {
  try {
    const headers = JSON.parse(localStorage.getItem(STORAGE_HEADERS_KEY) || "[]");
    const rows = JSON.parse(localStorage.getItem(STORAGE_KEY) || "[]");

    if (Array.isArray(headers) && Array.isArray(rows) && headers.length && rows.length) {
      state.headers = headers;
      state.rows = rows;
      el.fileName.textContent = `Tarayıcıda kayıtlı son liste yüklendi. Kayıt sayısı: ${rows.length}`;
    }
  } catch (error) {
    console.warn("Kayıtlı liste okunamadı.", error);
  }
}

function resetFilterButtons() {
  el.filterButtons.forEach(button => {
    button.classList.toggle("active", button.dataset.filter === "ALL");
  });
}

function render() {
  renderStats();
  renderHeaders();
  renderTable();
}

function renderStats() {
  const counts = countStatuses(state.rows);

  el.totalCount.textContent = state.rows.length;
  el.activeCount.textContent = counts.ACTIVE;
  el.cancelCount.textContent = counts.CANCEL;
  el.suspendCount.textContent = counts.SUSPEND;
}

function countStatuses(rows) {
  return rows.reduce((acc, row) => {
    const status = normalizeStatus(row[STATUS_COLUMN_INDEX]);
    if (acc[status] !== undefined) acc[status] += 1;
    return acc;
  }, { ACTIVE: 0, CANCEL: 0, SUSPEND: 0, UNKNOWN: 0 });
}

function renderHeaders() {
  el.tableHead.innerHTML = "";

  if (!state.headers.length) return;

  const tr = document.createElement("tr");
  state.headers.forEach((header, index) => {
    const th = document.createElement("th");
    th.textContent = index === STATUS_COLUMN_INDEX ? `${header} / Durum` : header;
    tr.appendChild(th);
  });

  el.tableHead.appendChild(tr);
}

function renderTable() {
  const hasData = state.headers.length && state.rows.length;
  el.emptyState.classList.toggle("hidden", hasData);
  el.tableWrap.classList.toggle("hidden", !hasData);

  if (!hasData) {
    el.tableBody.innerHTML = "";
    return;
  }

  const filteredRows = getFilteredRows();
  const fragment = document.createDocumentFragment();

  filteredRows.forEach(row => {
    const status = normalizeStatus(row[STATUS_COLUMN_INDEX]);
    const tr = document.createElement("tr");
    tr.className = getStatusClass(status);

    state.headers.forEach((_, colIndex) => {
      const td = document.createElement("td");
      const value = row[colIndex] ?? "";

      if (colIndex === STATUS_COLUMN_INDEX) {
        const badge = document.createElement("span");
        badge.className = `status-badge ${getStatusClass(status)}`;
        badge.textContent = status;
        td.appendChild(badge);
      } else {
        td.textContent = value;
      }

      tr.appendChild(td);
    });

    fragment.appendChild(tr);
  });

  el.tableBody.innerHTML = "";
  el.tableBody.appendChild(fragment);
}

function getFilteredRows() {
  return state.rows.filter(row => {
    const status = normalizeStatus(row[STATUS_COLUMN_INDEX]);
    const statusMatch = state.filter === "ALL" || status === state.filter;
    const searchMatch = !state.search || normalizeText(row.join(" ")).includes(state.search);
    return statusMatch && searchMatch;
  });
}
