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
  ["edificioCodigo", "Edificio / código"],
  ["cantidad", "Cantidad / número"],
  ["edificio", "Edificio"],
  ["ubicacion", "Ubicación"],
  ["marca", "Marca"],
  ["modelo", "Modelo"],
  ["numeroSerie", "Nº serie"],
  ["caracteristicas", "Características"],
  ["fechaFabricacion", "Fecha / año fabricación"],
  ["fechaProximoRetimbrado", "Fecha próximo retimbrado"],
  ["observaciones", "Observaciones"],
];

let records = [];
let currentPhotos = ["", ""];

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

function excelCellToText(value) {
  if (value === undefined || value === null) return "";
  if (value instanceof Date) return value.toLocaleDateString("es-ES");
  if (typeof value === "object") {
    if (value.text) return String(value.text);
    if (value.result !== undefined) return excelCellToText(value.result);
    if (Array.isArray(value.richText)) return value.richText.map((part) => part.text || "").join("");
    if (value.hyperlink && value.text) return String(value.text);
  }
  return String(value);
}

function rowToImportedRecord(rowValues, index) {
  const values = [];
  for (let col = 1; col <= 10; col += 1) {
    values[col] = excelCellToText(rowValues[col]);
  }

  const record = {
    id: `import-${Date.now()}-${index}-${Math.random().toString(16).slice(2)}`,
    edificioCodigo: values[1],
    cantidad: values[2],
    edificio: values[3],
    ubicacion: values[4],
    marca: values[5],
    modelo: values[6],
    numeroSerie: values[7],
    caracteristicas: values[8],
    fechaFabricacion: values[9],
    fechaProximoRetimbrado: values[10],
    defectos: [],
    observaciones: "",
    visto: false,
    photos: ["", ""],
    origen: "importado",
  };

  return record;
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
      records = saved;
      return;
    }
  } catch {}

  const stored = localStorage.getItem(STORAGE_KEY);
  if (stored) {
    try {
      records = JSON.parse(stored);
      await saveRecords();
      return;
    } catch {
      records = [];
    }
  }

  records = (window.INITIAL_EXTINTORES_LISTADOS || []).map((record) => ({ ...record }));
  await saveRecords();
}

async function saveRecords() {
  updateStats();
  try {
    await writeState(records);
  } catch {
    alert("No he podido guardar los datos. Si tienes muchas fotos, puede faltar espacio en el navegador.");
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
  let empty = 0;

  sheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    const record = rowToImportedRecord(row.values, rowNumber);
    const hasData = [
      record.edificioCodigo,
      record.cantidad,
      record.edificio,
      record.ubicacion,
      record.marca,
      record.modelo,
      record.numeroSerie,
    ].some((value) => safeText(value).trim());
    if (!hasData) {
      empty += 1;
      return;
    }
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

function normalizePhoto(photo) {
  return typeof photo === "string" ? photo : "";
}

function setPhotoPreview(index, dataUrl) {
  const img = $(`photoPreview${index + 1}`);
  const box = $(`photoBox${index + 1}`);
  const text = box.querySelector("span");
  const cleanPhoto = normalizePhoto(dataUrl);
  currentPhotos[index] = cleanPhoto;
  img.src = cleanPhoto;
  img.classList.toggle("hidden", !cleanPhoto);
  text.classList.toggle("hidden", Boolean(cleanPhoto));
  $(`deletePhoto${index + 1}`).disabled = !cleanPhoto;
}

function resizePhoto(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const img = new Image();
      img.onload = () => {
        const maxSide = 1400;
        const scale = Math.min(1, maxSide / Math.max(img.width, img.height));
        const canvas = document.createElement("canvas");
        canvas.width = Math.round(img.width * scale);
        canvas.height = Math.round(img.height * scale);
        const ctx = canvas.getContext("2d");
        ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
        resolve(canvas.toDataURL("image/jpeg", 0.78));
      };
      img.onerror = reject;
      img.src = reader.result;
    };
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
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

function filteredRecords() {
  const filterEdificio = $("filterEdificio").value.trim().toLowerCase();
  const filterNumero = $("filterNumero").value.trim().toLowerCase();
  const filterSerie = $("filterSerie").value.trim().toLowerCase();
  const seenFilter = $("seenFilter").value;

  return records.filter((record) => {
    if (seenFilter === "seen" && !record.visto) return false;
    if (seenFilter === "pending" && record.visto) return false;
    const edificioText = [record.edificioCodigo, record.edificio, record.ubicacion].join(" ").toLowerCase();
    const numeroText = safeText(record.cantidad).toLowerCase();
    const serieText = safeText(record.numeroSerie).toLowerCase();
    if (filterEdificio && !edificioText.includes(filterEdificio)) return false;
    if (filterNumero && !numeroText.includes(filterNumero)) return false;
    if (filterSerie && !serieText.includes(filterSerie)) return false;
    return true;
  });
}

function renderTable() {
  const body = $("recordsBody");
  const rows = filteredRecords();
  body.innerHTML = "";

  if (!rows.length) {
    body.innerHTML = `<tr><td colspan="13">No hay registros con ese filtro.</td></tr>`;
    return;
  }

  for (const record of rows) {
    const tr = document.createElement("tr");
    const defects = (record.defectos || []).length ? record.defectos.join(" / ") : "-";
    const photo1 = record.photos?.[0]
      ? `<img class="tablePhoto" src="${record.photos[0]}" alt="Foto 1">`
      : `<span class="noPhoto">—</span>`;
    const photo2 = record.photos?.[1]
      ? `<img class="tablePhoto" src="${record.photos[1]}" alt="Foto 2">`
      : `<span class="noPhoto">—</span>`;
    tr.innerHTML = `
      <td><strong>${safeText(record.cantidad) || "-"}</strong></td>
      <td>${safeText(record.edificioCodigo) || safeText(record.edificio) || "-"}</td>
      <td>${safeText(record.ubicacion) || "-"}</td>
      <td>${safeText(record.marca) || "-"}</td>
      <td>${safeText(record.modelo) || "-"}</td>
      <td>${safeText(record.numeroSerie) || "-"}</td>
      <td>${safeText(record.fechaFabricacion) || "-"}</td>
      <td>${safeText(record.fechaProximoRetimbrado) || "-"}</td>
      <td>${defects}</td>
      <td>${photo1}</td>
      <td>${photo2}</td>
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

  for (const [key] of fields) {
    $(key).value = safeText(record?.[key]);
  }
  $("visto").checked = Boolean(record?.visto);
  renderDefects(record?.defectos || []);
  const photos = Array.isArray(record?.photos) ? record.photos : ["", ""];
  setPhotoPreview(0, photos[0]);
  setPhotoPreview(1, photos[1]);
  showView("form");
}

function collectForm() {
  const id = $("recordId").value || createId();
  const record = { id, origen: $("recordId").value ? "editado" : "manual" };

  for (const [key] of fields) {
    record[key] = $(key).value.trim();
  }

  record.defectos = Array.from($("defectsList").querySelectorAll("input:checked")).map((input) => input.value);
  record.visto = $("visto").checked;
  record.photos = [normalizePhoto(currentPhotos[0]), normalizePhoto(currentPhotos[1])];
  return record;
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
    ["edificioCodigo", "Edificio / código", 18],
    ["cantidad", "Cantidad / número", 18],
    ["edificio", "Edificio", 14],
    ["ubicacion", "Ubicación", 42],
    ["marca", "Marca", 18],
    ["modelo", "Modelo", 28],
    ["numeroSerie", "Nº serie", 18],
    ["caracteristicas", "Características", 24],
    ["fechaFabricacion", "Fecha / año fabricación", 22],
    ["fechaProximoRetimbrado", "Fecha próximo retimbrado", 24],
    ["defectos", "Defectos encontrados", 42],
    ["defectoExtintorCaducado", "Extintor caducado", 20],
    ["defectoObstaculo", "Hay un obstáculo", 20],
    ["defectoDescargado", "Extintor descargado", 22],
    ["defectoSinPresion", "Extintor sin presión", 22],
    ["defectoEnSuelo", "Extintor en el suelo", 22],
    ["defectoCristal", "Cristal ausente o roto", 26],
    ["defectoSinSenal", "Sin señal", 16],
    ["defectoSenalCaducada", "Señal caducada", 20],
    ["observaciones", "Observaciones", 34],
    ["foto1", "Foto 1", 22],
    ["foto2", "Foto 2", 22],
    ["visto", "Visto", 10],
  ];

  sheet.columns = columns.map(([key, header, width]) => ({ key, header, width }));
  sheet.getRow(1).font = { bold: true, color: { argb: "FFFFFFFF" } };
  sheet.getRow(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF6D28D9" } };
  sheet.getRow(1).alignment = { vertical: "middle", horizontal: "center", wrapText: true };
  sheet.getRow(1).height = 30;

  for (const record of records) {
    const selectedDefects = record.defectos || [];
    const row = sheet.addRow({
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
      foto1: record.photos?.[0] ? "Foto 1" : "",
      foto2: record.photos?.[1] ? "Foto 2" : "",
      visto: record.visto ? "Sí" : "No",
    });
    const rowNumber = row.number;
    if (record.photos?.[0] || record.photos?.[1]) {
      row.height = 92;
    }
    [0, 1].forEach((photoIndex) => {
      const photo = record.photos?.[photoIndex];
      if (!photo) return;
      const imageId = workbook.addImage({
        base64: photo,
        extension: "jpeg",
      });
      const col = photoIndex === 0 ? 19 : 20;
      sheet.addImage(imageId, {
        tl: { col, row: rowNumber - 1 },
        ext: { width: 120, height: 85 },
        editAs: "oneCell",
      });
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
  $("downloadExcelBtn").addEventListener("click", downloadExcel);
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
  $("downloadExcelFromTableBtn").addEventListener("click", downloadExcel);
  $("viewTableFromFormBtn").addEventListener("click", () => showView("list"));
  $("filterEdificio").addEventListener("input", renderTable);
  $("filterNumero").addEventListener("input", renderTable);
  $("filterSerie").addEventListener("input", renderTable);
  $("seenFilter").addEventListener("change", renderTable);
  $("recordForm").addEventListener("submit", saveForm);
  $("deleteBtn").addEventListener("click", deleteCurrent);
  [0, 1].forEach((index) => {
    $(`photoInput${index + 1}`).addEventListener("change", async (event) => {
      const file = event.target.files?.[0];
      if (!file) return;
      try {
        const dataUrl = await resizePhoto(file);
        setPhotoPreview(index, dataUrl);
      } catch {
        alert("No he podido cargar esa foto. Prueba con otra.");
      } finally {
        event.target.value = "";
      }
    });
    $(`deletePhoto${index + 1}`).addEventListener("click", () => setPhotoPreview(index, ""));
  });

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
