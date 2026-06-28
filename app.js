const STORAGE_KEY = "extintores-importados-records-v1";
const DB_NAME = "extintores-importados-db";
const DB_VERSION = 1;
const STORE_NAME = "state";

const defectOptions = [
  "Extintor caducado.",
  "Hay un obstáculo.",
  "Extintor descargado.",
  "Extintor sin presión.",
  "Extintor en el suelo.",
  "Cristal del extintor ausente o roto.",
  "Sin señal.",
  "Señal caducada.",
];

const fields = [
  "edificio",
  "cantidad",
  "ubicacion",
  "modelo",
  "numeroSerie",
  "fechaFabricacion",
  "fechaProximoRetimbrado",
  "observaciones",
  "senal",
];

let records = [];

const $ = (id) => document.getElementById(id);

function safeText(value) {
  return value === undefined || value === null ? "" : String(value);
}

function createId() {
  return `manual-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function normalizeKeyPart(value) {
  return safeText(value)
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

function recordKey(record) {
  return [
    normalizeKeyPart(record.cantidad),
    normalizeKeyPart(record.numeroSerie),
    normalizeKeyPart(record.ubicacion),
  ].join("|");
}

function cleanRecord(record) {
  return {
    id: record.id || createId(),
    edificio: safeText(record.edificio),
    cantidad: safeText(record.cantidad),
    ubicacion: safeText(record.ubicacion),
    modelo: safeText(record.modelo),
    numeroSerie: safeText(record.numeroSerie),
    fechaFabricacion: safeText(record.fechaFabricacion),
    fechaProximoRetimbrado: safeText(record.fechaProximoRetimbrado),
    observaciones: safeText(record.observaciones),
    senal: safeText(record.senal),
    defectos: Array.isArray(record.defectos) ? record.defectos : [],
    visto: Boolean(record.visto),
    origen: record.origen || "excel",
  };
}

function excelCellToText(value) {
  if (value === undefined || value === null) return "";
  if (value instanceof Date) return value.toLocaleDateString("es-ES");
  if (typeof value === "object") {
    if (value.text) return String(value.text);
    if (value.result !== undefined) return excelCellToText(value.result);
    if (Array.isArray(value.richText)) return value.richText.map((part) => part.text || "").join("");
  }
  return String(value);
}

function rowToImportedRecord(rowValues, index) {
  const values = [];
  for (let col = 1; col <= 10; col += 1) {
    values[col] = excelCellToText(rowValues[col]);
  }

  return cleanRecord({
    id: `import-${Date.now()}-${index}-${Math.random().toString(16).slice(2)}`,
    cantidad: values[2],
    edificio: values[3],
    ubicacion: values[4],
    modelo: values[6],
    numeroSerie: values[7],
    fechaFabricacion: values[9],
    fechaProximoRetimbrado: values[10],
    observaciones: "",
    senal: "",
    defectos: [],
    visto: false,
    origen: "importado",
  });
}

function openDatabase() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);
    request.onupgradeneeded = () => {
      request.result.createObjectStore(STORE_NAME);
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

async function readState() {
  const db = await openDatabase();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, "readonly");
    const request = tx.objectStore(STORE_NAME).get("records");
    request.onsuccess = () => resolve(request.result || null);
    request.onerror = () => reject(request.error);
  });
}

async function writeState(value) {
  const db = await openDatabase();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, "readwrite");
    tx.objectStore(STORE_NAME).put(value, "records");
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}

async function loadRecords() {
  try {
    const saved = await readState();
    if (Array.isArray(saved)) {
      records = saved.map(cleanRecord);
      return;
    }
  } catch {}

  const stored = localStorage.getItem(STORAGE_KEY);
  if (stored) {
    try {
      records = JSON.parse(stored).map(cleanRecord);
      await saveRecords();
      return;
    } catch {
      records = [];
    }
  }

  records = (window.INITIAL_EXTINTORES_LISTADOS || []).map(cleanRecord);
  await saveRecords();
}

async function saveRecords() {
  records = records.map(cleanRecord);
  updateStats();
  try {
    await writeState(records);
  } catch {
    alert("No he podido guardar los datos. Puede faltar espacio en el navegador.");
  }
}

async function importExcelFile(file) {
  if (!window.ExcelJS) {
    alert("No se ha cargado el lector de Excel. Cierra y vuelve a abrir la app.");
    return;
  }

  const buffer = await file.arrayBuffer();
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  const sheet = workbook.worksheets[0];
  if (!sheet) {
    alert("No encuentro ninguna hoja en ese Excel.");
    return;
  }

  const existingKeys = new Set(records.map(recordKey).filter((key) => key !== "||"));
  const imported = [];
  let repeated = 0;

  sheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    const record = rowToImportedRecord(row.values, rowNumber);
    const hasData = [
      record.edificio,
      record.cantidad,
      record.ubicacion,
      record.modelo,
      record.numeroSerie,
    ].some((value) => safeText(value).trim());
    if (!hasData) return;

    const key = recordKey(record);
    if (key !== "||" && existingKeys.has(key)) {
      repeated += 1;
      return;
    }
    existingKeys.add(key);
    imported.push(record);
  });

  if (!imported.length) {
    $("importStatus").textContent = `No se añadieron registros nuevos. Repetidos detectados: ${repeated}.`;
    alert(`No se añadieron registros nuevos.\nRepetidos detectados: ${repeated}.`);
    return;
  }

  records = [...imported, ...records];
  await saveRecords();
  $("importStatus").textContent = `Importación correcta: ${imported.length} nuevos. Repetidos ignorados: ${repeated}.`;
  alert(`Importación correcta.\nNuevos: ${imported.length}\nRepetidos ignorados: ${repeated}`);
}

function updateStats() {
  const total = records.length;
  const seen = records.filter((record) => record.visto).length;
  $("totalCount").textContent = total;
  $("seenCount").textContent = seen;
  $("pendingCount").textContent = total - seen;
}

function showView(name) {
  $("homeView").classList.toggle("hidden", name !== "home");
  $("listView").classList.toggle("hidden", name !== "list");
  $("formView").classList.toggle("hidden", name !== "form");
  if (name === "list") renderTable();
  window.scrollTo({ top: 0, behavior: "smooth" });
}

function compareText(a, b) {
  return safeText(a).localeCompare(safeText(b), "es", { numeric: true, sensitivity: "base" });
}

function filteredRecords() {
  const filterEdificio = $("filterEdificio").value.trim().toLowerCase();
  const filterNumero = $("filterNumero").value.trim().toLowerCase();
  const filterSerie = $("filterSerie").value.trim().toLowerCase();
  const seenFilter = $("seenFilter").value;
  const sortOrder = $("sortOrder").value;

  const rows = records.filter((record) => {
    if (seenFilter === "seen" && !record.visto) return false;
    if (seenFilter === "pending" && record.visto) return false;
    const edificioText = [record.edificio, record.ubicacion].join(" ").toLowerCase();
    const numeroText = safeText(record.cantidad).toLowerCase();
    const serieText = safeText(record.numeroSerie).toLowerCase();
    if (filterEdificio && !edificioText.includes(filterEdificio)) return false;
    if (filterNumero && !numeroText.includes(filterNumero)) return false;
    if (filterSerie && !serieText.includes(filterSerie)) return false;
    return true;
  });

  if (sortOrder === "edificio") {
    rows.sort((a, b) => compareText(a.edificio, b.edificio) || compareText(a.cantidad, b.cantidad));
  }
  if (sortOrder === "numero") {
    rows.sort((a, b) => compareText(a.cantidad, b.cantidad) || compareText(a.edificio, b.edificio));
  }

  return rows;
}

function renderTable() {
  const body = $("recordsBody");
  const rows = filteredRecords();
  body.innerHTML = "";

  if (!rows.length) {
    body.innerHTML = `<tr><td colspan="12">No hay registros con ese filtro.</td></tr>`;
    return;
  }

  for (const record of rows) {
    const tr = document.createElement("tr");
    const defects = (record.defectos || []).length ? record.defectos.join(" / ") : "-";
    tr.innerHTML = `
      <td>${safeText(record.edificio) || "-"}</td>
      <td><strong>${safeText(record.cantidad) || "-"}</strong></td>
      <td>${safeText(record.ubicacion) || "-"}</td>
      <td>${safeText(record.modelo) || "-"}</td>
      <td>${safeText(record.numeroSerie) || "-"}</td>
      <td>${safeText(record.fechaFabricacion) || "-"}</td>
      <td>${safeText(record.fechaProximoRetimbrado) || "-"}</td>
      <td>${safeText(record.observaciones) || "-"}</td>
      <td>${safeText(record.senal) || "-"}</td>
      <td>${defects}</td>
      <td><span class="${record.visto ? "ok" : "pending"}">${record.visto ? "Sí" : "No"}</span></td>
      <td><button class="editBtn" data-edit="${record.id}">Ver / corregir</button></td>
    `;
    body.appendChild(tr);
  }

  body.querySelectorAll("[data-edit]").forEach((button) => {
    button.addEventListener("click", () => openForm(button.dataset.edit));
  });
}

function renderDefects(selected = []) {
  const box = $("defectsList");
  box.innerHTML = "";
  for (const option of defectOptions) {
    const label = document.createElement("label");
    label.className = "checkItem";
    label.innerHTML = `
      <input type="checkbox" value="${option}">
      <span>${option}</span>
    `;
    const input = label.querySelector("input");
    input.checked = selected.includes(option);
    box.appendChild(label);
  }
}

function openForm(id = null) {
  const record = id ? records.find((item) => item.id === id) : null;
  $("recordId").value = record?.id || "";
  $("formTitle").textContent = record ? "Ver y corregir extintor" : "Meter dato nuevo";
  $("formKicker").textContent = record ? "REGISTRO EXISTENTE" : "NUEVO REGISTRO";
  $("deleteBtn").classList.toggle("hidden", !record);

  for (const key of fields) {
    $(key).value = safeText(record?.[key]);
  }
  $("visto").checked = Boolean(record?.visto);
  renderDefects(record?.defectos || []);
  showView("form");
}

function collectForm() {
  const id = $("recordId").value || createId();
  const record = { id, origen: $("recordId").value ? "editado" : "manual" };

  for (const key of fields) {
    record[key] = $(key).value.trim();
  }

  record.defectos = Array.from($("defectsList").querySelectorAll("input:checked")).map((input) => input.value);
  record.visto = $("visto").checked;
  return cleanRecord(record);
}

async function saveForm(event) {
  event.preventDefault();
  const record = collectForm();
  const index = records.findIndex((item) => item.id === record.id);
  if (index >= 0) {
    records[index] = { ...records[index], ...record };
  } else {
    records.unshift(record);
  }
  await saveRecords();
  showView("list");
}

async function deleteCurrent() {
  const id = $("recordId").value;
  if (!id) return;
  if (!confirm("¿Seguro que quieres eliminar este registro?")) return;
  records = records.filter((record) => record.id !== id);
  await saveRecords();
  showView("list");
}

async function downloadExcel() {
  if (!window.ExcelJS) {
    alert("No se ha cargado el generador de Excel. Cierra y vuelve a abrir la app.");
    return;
  }

  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Extintores Importados";
  workbook.created = new Date();
  const sheet = workbook.addWorksheet("Extintores");

  const columns = [
    ["edificio", "Edificio", 14],
    ["cantidad", "Número SYCo", 18],
    ["ubicacion", "Ubicación", 42],
    ["modelo", "Modelo", 28],
    ["numeroSerie", "Nº serie", 18],
    ["fechaFabricacion", "Fecha / año fabricación", 22],
    ["fechaProximoRetimbrado", "Fecha retimbrado", 22],
    ["observaciones", "Observaciones", 34],
    ["senal", "Señal", 14],
    ["defectos", "Defectos encontrados", 42],
    ["defectoExtintorCaducado", "Extintor caducado", 20],
    ["defectoObstaculo", "Hay un obstáculo", 20],
    ["defectoDescargado", "Extintor descargado", 22],
    ["defectoSinPresion", "Extintor sin presión", 22],
    ["defectoEnSuelo", "Extintor en el suelo", 22],
    ["defectoCristal", "Cristal ausente o roto", 26],
    ["defectoSinSenal", "Sin señal", 16],
    ["defectoSenalCaducada", "Señal caducada", 20],
    ["visto", "Visto", 10],
  ];

  sheet.columns = columns.map(([key, header, width]) => ({ key, header, width }));
  sheet.getRow(1).font = { bold: true, color: { argb: "FFFFFFFF" } };
  sheet.getRow(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF3B0764" } };
  sheet.getRow(1).alignment = { vertical: "middle", horizontal: "center", wrapText: true };
  sheet.getRow(1).height = 30;

  for (const record of filteredRecords()) {
    const selectedDefects = record.defectos || [];
    sheet.addRow({
      ...record,
      defectos: selectedDefects.join(" / "),
      defectoExtintorCaducado: selectedDefects.includes("Extintor caducado.") ? "Sí" : "",
      defectoObstaculo: selectedDefects.includes("Hay un obstáculo.") ? "Sí" : "",
      defectoDescargado: selectedDefects.includes("Extintor descargado.") ? "Sí" : "",
      defectoSinPresion: selectedDefects.includes("Extintor sin presión.") ? "Sí" : "",
      defectoEnSuelo: selectedDefects.includes("Extintor en el suelo.") ? "Sí" : "",
      defectoCristal: selectedDefects.includes("Cristal del extintor ausente o roto.") ? "Sí" : "",
      defectoSinSenal: selectedDefects.includes("Sin señal.") ? "Sí" : "",
      defectoSenalCaducada: selectedDefects.includes("Señal caducada.") ? "Sí" : "",
      visto: record.visto ? "Sí" : "No",
    });
  }

  sheet.views = [{ state: "frozen", ySplit: 1 }];
  sheet.autoFilter = {
    from: { row: 1, column: 1 },
    to: { row: 1, column: columns.length },
  };

  sheet.eachRow((row, rowNumber) => {
    row.eachCell((cell) => {
      cell.border = {
        top: { style: "thin", color: { argb: "FFE6E0DA" } },
        left: { style: "thin", color: { argb: "FFE6E0DA" } },
        bottom: { style: "thin", color: { argb: "FFE6E0DA" } },
        right: { style: "thin", color: { argb: "FFE6E0DA" } },
      };
      cell.alignment = { vertical: "top", wrapText: true };
      if (rowNumber > 1 && rowNumber % 2 === 0) {
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFAF8F5" } };
      }
    });
  });

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const a = document.createElement("a");
  const date = new Date().toISOString().slice(0, 10);
  a.href = URL.createObjectURL(blob);
  a.download = `Extintores_Importados_${date}.xlsx`;
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 1000);
}

function bindEvents() {
  $("openListBtn").addEventListener("click", () => showView("list"));
  $("newRecordBtn").addEventListener("click", () => openForm());
  $("newRecordFromListBtn").addEventListener("click", () => openForm());
  $("downloadExcelBtn").addEventListener("click", downloadExcel);
  $("downloadExcelFromTableBtn").addEventListener("click", downloadExcel);
  $("viewTableFromFormBtn").addEventListener("click", () => showView("list"));
  $("importExcelBtn").addEventListener("click", () => $("importExcelInput").click());
  $("importExcelInput").addEventListener("change", async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;
    try {
      $("importStatus").textContent = "Importando Excel...";
      await importExcelFile(file);
      renderTable();
    } catch (error) {
      console.error(error);
      $("importStatus").textContent = "No se ha podido importar el Excel.";
      alert("No se ha podido importar el Excel. Revisa que tenga el mismo formato del listado.");
    } finally {
      event.target.value = "";
    }
  });

  ["filterEdificio", "filterNumero", "filterSerie", "sortOrder", "seenFilter"].forEach((id) => {
    $(id).addEventListener("input", renderTable);
    $(id).addEventListener("change", renderTable);
  });

  $("recordForm").addEventListener("submit", saveForm);
  $("deleteBtn").addEventListener("click", deleteCurrent);

  document.querySelectorAll("[data-back]").forEach((button) => {
    button.addEventListener("click", () => showView(button.dataset.back));
  });
}

if ("serviceWorker" in navigator) {
  window.addEventListener("load", () => {
    navigator.serviceWorker.register("./sw.js").catch(() => {});
  });
}

async function init() {
  await loadRecords();
  bindEvents();
  updateStats();
}

init();
