/* =========================================================================
   CONFIGURACIÓN
   ========================================================================= */

// Pegá aquí la URL de tu Apps Script (la que termina en /exec)
const APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbysz3NVB0xTq2P6NAqIiBTKTgsI2GujNPQOYeSG7dex2YLiW4S8gHRF_hrLxBHOYmtGXA/exec";

// Cada cuánto se vuelve a sincronizar automáticamente (ms). 60000 = 1 minuto.
const AUTO_REFRESH_MS = 60000;

// Nombres de columnas que usamos para inferir el estado en cada hoja.
// Si en algún momento cambiás el nombre de una columna en el Excel,
// solo hay que actualizarlo aquí.
const COL = {
  hecha: "¿Está hecha?",
  ingresada: "¿La ingresamos?",
  estadoParlamentario: "¿Estado parlamentario?",
  aprobado: "¿Se aprobó en el recinto?",
  eje: "¿Con qué eje de trabajo se vincula?",
  diploma: "¿Pedimos el diploma?",
  fecha: "¿En qué fecha?",
};

// Valor especial para agrupar los proyectos sin año detectable.
const SIN_FECHA = "Sin fecha";

// Saca el año de un proyecto con esta prioridad:
// 1) la columna de fecha real, si tiene un valor parseable
// 2) el año dentro del número de expediente en el título (ej. "2431-D-2024")
// 3) si no hay ninguno de los dos, "Sin fecha"
function getAnio(row, titulo) {
  const fechaRaw = row[COL.fecha];
  if (fechaRaw) {
    const fecha = new Date(fechaRaw);
    if (!isNaN(fecha.getTime())) {
      const anio = fecha.getFullYear();
      if (anio > 1990 && anio < 2100) return String(anio);
    }
  }
  // Buscamos un patrón de expediente tipo NNNN-D-AAAA o NNNN-P-AAAA en el título
  const match = String(titulo || "").match(/\d+-[A-Za-z]-(\d{4})/);
  if (match) return match[1];
  return SIN_FECHA;
}

// Las 3 categorías simples para agrupar el campo "¿Pedimos el diploma?",
// que en la planilla real tiene muchos valores parecidos pero no idénticos.
const DIPLOMA_CATEGORIAS = [
  { id: "entregado", nombre: "Entregado", color: "#6B7A4F" },
  { id: "pendiente", nombre: "Pendiente de entrega", color: "#C08A2E" },
  { id: "no_corresponde", nombre: "No corresponde", color: "#9C8B5E" },
];

// Mapeo de cada valor real (normalizado, sin tildes/mayúsculas) a su categoría.
const DIPLOMA_MAPEO = {
  "entregado": "entregado",
  "no se entrega": "no_corresponde",
  "no pedir": "no_corresponde",
  "en la oficina": "pendiente",
  "en protocolo": "pendiente",
  "oficina": "pendiente",
  "": "pendiente",
};

function getDiplomaCategoria(row) {
  const v = normaliza(row[COL.diploma]);
  const id = DIPLOMA_MAPEO.hasOwnProperty(v) ? DIPLOMA_MAPEO[v] : "pendiente";
  return DIPLOMA_CATEGORIAS.find((c) => c.id === id);
}

// Las 4 etapas del Kanban, en orden, con su color y la lógica de detección.
// "match" decide a qué etapa pertenece una fila. Se evalúa de abajo hacia
// arriba (chequea primero si está aprobado, luego en comisión, etc.)
const ETAPAS = [
  {
    id: "redactado",
    nombre: "Redactado",
    color: "#9C8B5E",
    match: (row) => normaliza(row[COL.hecha]) === "hecho" && !esIngresada(row),
  },
  {
    id: "ingresado",
    nombre: "Ingresado",
    color: "#1E3A5F",
    match: (row) => esIngresada(row) && normaliza(row[COL.estadoParlamentario]) !== "si",
  },
  {
    id: "comision",
    nombre: "En comisión",
    color: "#C08A2E",
    match: (row) => normaliza(row[COL.estadoParlamentario]) === "si" && normaliza(row[COL.aprobado]) !== "aprobado",
  },
  {
    id: "aprobado",
    nombre: "Aprobado",
    color: "#6B7A4F",
    match: (row) => normaliza(row[COL.aprobado]) === "aprobado",
  },
];

// Qué columnas reales se escriben en el Sheet cuando el usuario mueve
// una tarjeta a cada etapa (movimiento manual desde la web).
const ESCRITURA_POR_ETAPA = {
  redactado: { [COL.hecha]: "Hecho", [COL.ingresada]: "Sin ingresar", [COL.estadoParlamentario]: "", [COL.aprobado]: "" },
  ingresado: { [COL.hecha]: "Hecho", [COL.ingresada]: "Ingresada", [COL.estadoParlamentario]: "No", [COL.aprobado]: "" },
  comision: { [COL.hecha]: "Hecho", [COL.ingresada]: "Ingresada", [COL.estadoParlamentario]: "Sí", [COL.aprobado]: "" },
  aprobado: { [COL.hecha]: "Hecho", [COL.ingresada]: "Ingresada", [COL.estadoParlamentario]: "Sí", [COL.aprobado]: "Aprobado" },
};

// Nombres "bonitos" para cada hoja, y a qué columna corresponde el
// nombre del proyecto/declaración, porque no se llaman igual en todas las hojas.
const HOJAS_CONFIG = {
  "PROYECTOS DE LEY": { etiqueta: "Proyectos de Ley", colTitulo: "PROYECTO" },
  "DECLARACIONES DE INTERÉS": { etiqueta: "Declaraciones de Interés", colTitulo: "DECLARACIÓN" },
  "RESOLUCIONES": { etiqueta: "Resoluciones", colTitulo: "PROYECTO" },
  "HECHOS PROPIOS": { etiqueta: "Hechos Propios", colTitulo: "PROYECTO" },
};

/* =========================================================================
   ESTADO GLOBAL
   ========================================================================= */

let RAW_DATA = null;       // lo que llega del Apps Script
let VISTA_ACTIVA = "seguimiento"; // "seguimiento" | "diplomas"
let FILTROS = {
  hoja: "TODAS",
  eje: "TODOS",
  etapas: [],         // array vacío = todas. Ej: ["redactado","aprobado"]
  anios: [],          // array vacío = todos. Ej: ["2024","2025"]
  busqueda: "",
};
let FILTROS_DIPLOMAS = {
  eje: "TODOS",
  diploma: [],        // array vacío = todas las categorías
  busqueda: "",
};
let autoRefreshTimer = null;
let EXPORTAR_SOLO_FILTRADO_SEG = true;
let EXPORTAR_SOLO_FILTRADO_DIP = true;

/* =========================================================================
   HELPERS
   ========================================================================= */

function normaliza(v) {
  if (v === null || v === undefined) return "";
  // sacamos tildes (NFD) y pasamos a minúsculas, para que "Sí" === "si"
  return String(v).trim().toLowerCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

function esIngresada(row) {
  const v = normaliza(row[COL.ingresada]);
  return v.startsWith("ingresada"); // cubre "Ingresada" y "Ingresada c/modif"
}

function getEtapa(row) {
  // de la última a la primera, así "aprobado" gana si se cumplen varias
  for (let i = ETAPAS.length - 1; i >= 0; i--) {
    if (ETAPAS[i].match(row)) return ETAPAS[i];
  }
  return ETAPAS[0]; // por defecto: redactado / sin info
}

function getTitulo(row, sheetName) {
  const cfg = HOJAS_CONFIG[sheetName];
  const raw = row[cfg.colTitulo] || "(sin título)";
  return String(raw);
}

function getEje(row) {
  return row[COL.eje] || row["Eje de trabajo"] || "Sin clasificar";
}

function showToast(msg, isError = false) {
  const t = document.getElementById("toast");
  t.textContent = msg;
  t.className = "toast show" + (isError ? " error" : "");
  setTimeout(() => { t.className = "toast"; }, 3200);
}

function setSyncStatus(state, text) {
  const dot = document.getElementById("sync-dot");
  const txt = document.getElementById("sync-text");
  dot.className = "sync-dot " + (state === "ok" ? "" : state);
  txt.textContent = text;
}

/* =========================================================================
   CARGA DE DATOS
   ========================================================================= */

async function cargarDatos(manual = false) {
  const btn = document.getElementById("refresh-btn");
  btn.classList.add("spinning");
  setSyncStatus("loading", manual ? "Actualizando…" : "Sincronizando…");

  try {
    const res = await fetch(APPS_SCRIPT_URL + "?action=getAll&_=" + Date.now());
    if (!res.ok) throw new Error("Error HTTP " + res.status);
    const json = await res.json();
    if (!json.success) throw new Error(json.error || "Error desconocido del script");

    RAW_DATA = json.data;
    document.getElementById("error-zone").innerHTML = "";
    document.getElementById("loading-screen").style.display = "none";
    document.getElementById("app-content").style.display = "block";

    setSyncStatus("ok", "Conectado");
    document.getElementById("footer-time").textContent = new Date().toLocaleString("es-AR");

    render();
  } catch (err) {
    console.error(err);
    setSyncStatus("error", "Sin conexión");
    document.getElementById("loading-screen").style.display = "none";
    document.getElementById("app-content").style.display = "block";
    document.getElementById("error-zone").innerHTML = `
      <div class="error-banner">
        <span>⚠️</span>
        <div>
          <b>No se pudo conectar con la planilla.</b><br>
          ${err.message}. Verificá que la URL del Apps Script sea correcta y que esté implementada como
          "Cualquier usuario" puede acceder, y probá actualizar de nuevo.
        </div>
      </div>`;
    if (RAW_DATA) {
      render(); // mantenemos lo último que funcionó
    }
  } finally {
    btn.classList.remove("spinning");
  }
}

function reiniciarAutoRefresh() {
  if (autoRefreshTimer) clearInterval(autoRefreshTimer);
  autoRefreshTimer = setInterval(() => cargarDatos(false), AUTO_REFRESH_MS);
}

/* =========================================================================
   TRANSFORMACIÓN: aplanar todas las hojas en una sola lista de "items"
   ========================================================================= */

function getAllItems() {
  if (!RAW_DATA) return [];
  const items = [];
  Object.keys(HOJAS_CONFIG).forEach((sheetName) => {
    const hoja = RAW_DATA[sheetName];
    if (!hoja) return;
    hoja.rows.forEach((row) => {
      const titulo = getTitulo(row, sheetName);
      items.push({
        sheet: sheetName,
        etapa: getEtapa(row),
        titulo: titulo,
        eje: getEje(row),
        anio: getAnio(row, titulo),
        row: row,
      });
    });
  });
  return items;
}

function getItemsFiltrados() {
  let items = getAllItems();
  if (FILTROS.hoja !== "TODAS") items = items.filter((i) => i.sheet === FILTROS.hoja);
  if (FILTROS.eje !== "TODOS") items = items.filter((i) => i.eje === FILTROS.eje);
  if (FILTROS.etapas.length > 0) items = items.filter((i) => FILTROS.etapas.includes(i.etapa.id));
  if (FILTROS.anios.length > 0) items = items.filter((i) => FILTROS.anios.includes(i.anio));
  if (FILTROS.busqueda.trim() !== "") {
    const q = normaliza(FILTROS.busqueda);
    items = items.filter((i) => normaliza(i.titulo).includes(q));
  }
  return items;
}

// Lista base para la pestaña de Diplomas: solo Declaraciones de Interés
// que ya están Aprobadas (porque el diploma solo aplica a esos casos).
function getAllItemsDiplomas() {
  if (!RAW_DATA) return [];
  const hoja = RAW_DATA["DECLARACIONES DE INTERÉS"];
  if (!hoja) return [];
  return hoja.rows
    .filter((row) => normaliza(row[COL.aprobado]) === "aprobado")
    .map((row) => ({
      sheet: "DECLARACIONES DE INTERÉS",
      titulo: getTitulo(row, "DECLARACIONES DE INTERÉS"),
      eje: getEje(row),
      diplomaCategoria: getDiplomaCategoria(row),
      row: row,
    }));
}

function getItemsDiplomasFiltrados() {
  let items = getAllItemsDiplomas();
  if (FILTROS_DIPLOMAS.eje !== "TODOS") items = items.filter((i) => i.eje === FILTROS_DIPLOMAS.eje);
  if (FILTROS_DIPLOMAS.diploma.length > 0) items = items.filter((i) => FILTROS_DIPLOMAS.diploma.includes(i.diplomaCategoria.id));
  if (FILTROS_DIPLOMAS.busqueda.trim() !== "") {
    const q = normaliza(FILTROS_DIPLOMAS.busqueda);
    items = items.filter((i) => normaliza(i.titulo).includes(q));
  }
  return items;
}

function getEjesUnicosDiplomas() {
  const set = new Set();
  getAllItemsDiplomas().forEach((i) => set.add(i.eje));
  return Array.from(set).filter(Boolean).sort();
}

function getEjesUnicos() {
  const set = new Set();
  getAllItems().forEach((i) => set.add(i.eje));
  return Array.from(set).filter(Boolean).sort();
}

function getAniosUnicos() {
  const set = new Set();
  getAllItems().forEach((i) => set.add(i.anio));
  const arr = Array.from(set).filter(Boolean);
  // Años más recientes primero, "Sin fecha" siempre al final
  arr.sort((a, b) => {
    if (a === SIN_FECHA) return 1;
    if (b === SIN_FECHA) return -1;
    return Number(b) - Number(a);
  });
  return arr;
}

/* =========================================================================
   RENDER: overview stats
   ========================================================================= */

function renderOverview() {
  const items = getItemsFiltrados();
  const total = items.length;
  const aprobados = items.filter((i) => i.etapa.id === "aprobado").length;
  const enComision = items.filter((i) => i.etapa.id === "comision").length;
  const ingresados = items.filter((i) => i.etapa.id === "ingresado").length;
  const pct = total ? Math.round((aprobados / total) * 100) : 0;

  const porEje = {};
  items.forEach((i) => { porEje[i.eje] = (porEje[i.eje] || 0) + 1; });
  const ejeTop = Object.entries(porEje).sort((a, b) => b[1] - a[1])[0];

  const html = `
    <div class="stat-card">
      <div class="label">Total de items</div>
      <div class="value">${total}</div>
      <div class="sub">en las 4 hojas combinadas</div>
    </div>
    <div class="stat-card">
      <div class="label">Aprobados</div>
      <div class="value" style="color:var(--olive)">${aprobados}</div>
      <div class="sub">${pct}% del total filtrado</div>
    </div>
    <div class="stat-card">
      <div class="label">En comisión</div>
      <div class="value" style="color:var(--ochre)">${enComision}</div>
      <div class="sub">con giro parlamentario activo</div>
    </div>
    <div class="stat-card">
      <div class="label">Ingresados</div>
      <div class="value" style="color:var(--navy)">${ingresados}</div>
      <div class="sub">esperando tratamiento</div>
    </div>
    <div class="stat-card chart-card">
      <div class="label">Eje con más actividad</div>
      <div class="value" style="font-size:19px; color:var(--ink);">${ejeTop ? ejeTop[0] : "—"}</div>
      <div class="sub">${ejeTop ? ejeTop[1] + " proyectos" : ""}</div>
      ${renderBarrasEjes(porEje, total)}
    </div>
  `;
  document.getElementById("overview-stats").innerHTML = html;
}

function renderBarrasEjes(porEje, total) {
  const entries = Object.entries(porEje).sort((a, b) => b[1] - a[1]).slice(0, 5);
  if (!entries.length) return "";
  const max = entries[0][1];
  return `<div style="margin-top:10px; display:flex; flex-direction:column; gap:5px;">
    ${entries.map(([eje, count]) => `
      <div style="display:flex; align-items:center; gap:7px;">
        <div style="flex:1; background:var(--paper-dim); border-radius:3px; height:7px; overflow:hidden;">
          <div style="width:${(count/max*100)}%; height:100%; background:var(--navy); border-radius:3px;"></div>
        </div>
        <div style="font-size:10.5px; color:var(--ink-soft); width:18px; text-align:right;">${count}</div>
      </div>
    `).join("")}
  </div>`;
}

/* =========================================================================
   RENDER: filtros
   ========================================================================= */

function renderFiltros() {
  const ejes = getEjesUnicos();
  const anios = getAniosUnicos();
  const html = `
    <div class="filter-group" style="flex:2; min-width:220px;">
      <label>Buscar proyecto</label>
      <input type="text" id="f-busqueda" placeholder="Escribí un nombre o expediente…" value="${FILTROS.busqueda.replace(/"/g,'&quot;')}">
    </div>
    <div class="filter-group">
      <label>Hoja</label>
      <select id="f-hoja">
        <option value="TODAS">Todas las hojas</option>
        ${Object.entries(HOJAS_CONFIG).map(([key, cfg]) =>
          `<option value="${key}" ${FILTROS.hoja===key?"selected":""}>${cfg.etiqueta}</option>`
        ).join("")}
      </select>
    </div>
    <div class="filter-group">
      <label>Eje de trabajo</label>
      <select id="f-eje">
        <option value="TODOS">Todos los ejes</option>
        ${ejes.map((e) => `<option value="${e.replace(/"/g,'&quot;')}" ${FILTROS.eje===e?"selected":""}>${e}</option>`).join("")}
      </select>
    </div>
    <div class="filter-group" style="min-width:320px;">
      <label>Etapa <span style="text-transform:none; font-weight:400; opacity:0.7;">(elegí una o varias)</span></label>
      <div class="filter-pills" id="f-etapas">
        ${ETAPAS.map((e) => {
          const activa = FILTROS.etapas.includes(e.id);
          return `<span class="pill ${activa?'active':''}" data-etapa="${e.id}" style="${activa?`background:${e.color};border-color:${e.color};`:''}">${e.nombre}</span>`;
        }).join("")}
      </div>
    </div>
    <div class="filter-group" style="min-width:320px;">
      <label>Año <span style="text-transform:none; font-weight:400; opacity:0.7;">(elegí uno o varios)</span></label>
      <div class="filter-pills" id="f-anios">
        ${anios.map((a) => {
          const activa = FILTROS.anios.includes(a);
          return `<span class="pill ${activa?'active':''}" data-anio="${a}">${a}</span>`;
        }).join("")}
      </div>
    </div>
    <button class="clear-filters" id="f-clear">Limpiar filtros</button>
  `;
  document.getElementById("filters-zone").innerHTML = html;

  document.getElementById("f-busqueda").addEventListener("input", (e) => {
    FILTROS.busqueda = e.target.value;
    renderOverview(); renderKanban(); renderTabla();
  });
  document.getElementById("f-hoja").addEventListener("change", (e) => {
    FILTROS.hoja = e.target.value;
    renderAll();
  });
  document.getElementById("f-eje").addEventListener("change", (e) => {
    FILTROS.eje = e.target.value;
    renderOverview(); renderKanban(); renderTabla();
  });
  document.querySelectorAll("#f-etapas .pill").forEach((pill) => {
    pill.addEventListener("click", () => {
      const id = pill.dataset.etapa;
      const idx = FILTROS.etapas.indexOf(id);
      if (idx === -1) {
        FILTROS.etapas.push(id); // la agrego a la selección
      } else {
        FILTROS.etapas.splice(idx, 1); // la saco de la selección
      }
      renderFiltros(); renderOverview(); renderKanban(); renderTabla();
    });
  });
  document.querySelectorAll("#f-anios .pill").forEach((pill) => {
    pill.addEventListener("click", () => {
      const anio = pill.dataset.anio;
      const idx = FILTROS.anios.indexOf(anio);
      if (idx === -1) {
        FILTROS.anios.push(anio);
      } else {
        FILTROS.anios.splice(idx, 1);
      }
      renderFiltros(); renderOverview(); renderKanban(); renderTabla();
    });
  });
  document.getElementById("f-clear").addEventListener("click", () => {
    FILTROS = { hoja: "TODAS", eje: "TODOS", etapas: [], anios: [], busqueda: "" };
    renderAll();
  });
}

/* =========================================================================
   RENDER: Kanban
   ========================================================================= */

function renderKanban() {
  const items = getItemsFiltrados();
  const board = document.getElementById("kanban-board");

  board.innerHTML = ETAPAS.map((etapa) => {
    const itemsEtapa = items.filter((i) => i.etapa.id === etapa.id);
    return `
      <div class="kcol state-${etapa.id}" style="--kcolor:${etapa.color}">
        <div class="kcol-head">
          <span class="title">${etapa.nombre}</span>
          <span class="count">${itemsEtapa.length}</span>
        </div>
        <div class="kcol-body" data-etapa="${etapa.id}">
          ${itemsEtapa.length ? itemsEtapa.map((item) => renderCard(item)).join("") : `<div class="empty-state">Sin items</div>`}
        </div>
      </div>
    `;
  }).join("");

  attachDragEvents();
}

function renderCard(item) {
  const cfg = HOJAS_CONFIG[item.sheet];
  const uid = item.row._rowNumber + "::" + item.sheet;
  return `
    <div class="kcard" draggable="true" data-uid="${uid}" data-sheet="${item.sheet}" data-row="${item.row._rowNumber}">
      <span class="title">${escapeHtml(item.titulo)}</span>
      <div class="meta">
        <span class="tag hoja-tag">${cfg.etiqueta}</span>
        <span class="tag">${escapeHtml(item.eje)}</span>
      </div>
    </div>
  `;
}

function escapeHtml(str) {
  if (str === null || str === undefined) return "";
  return String(str).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");
}

let draggedUid = null;

function attachDragEvents() {
  document.querySelectorAll(".kcard").forEach((card) => {
    card.addEventListener("dragstart", (e) => {
      draggedUid = card.dataset.uid;
      card.classList.add("dragging");
      e.dataTransfer.effectAllowed = "move";
    });
    card.addEventListener("dragend", () => {
      card.classList.remove("dragging");
    });
  });

  document.querySelectorAll(".kcol-body").forEach((col) => {
    col.addEventListener("dragover", (e) => {
      e.preventDefault();
      col.classList.add("drag-over");
    });
    col.addEventListener("dragleave", () => col.classList.remove("drag-over"));
    col.addEventListener("drop", (e) => {
      e.preventDefault();
      col.classList.remove("drag-over");
      const nuevaEtapa = col.dataset.etapa;
      if (draggedUid) moverItem(draggedUid, nuevaEtapa);
    });
  });
}

async function moverItem(uid, nuevaEtapaId) {
  const [rowNumberStr, sheetName] = uid.split("::");
  const rowNumber = parseInt(rowNumberStr, 10);
  const updates = ESCRITURA_POR_ETAPA[nuevaEtapaId];

  // Optimistic UI: actualizamos localmente antes de confirmar el guardado
  const hoja = RAW_DATA[sheetName];
  const rowObj = hoja.rows.find((r) => r._rowNumber === rowNumber);
  if (!rowObj) return;
  const backup = { ...rowObj };
  Object.assign(rowObj, updates);
  renderOverview(); renderKanban(); renderTabla();

  showToast("Guardando cambio…");
  try {
    const res = await fetch(APPS_SCRIPT_URL, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify({
        action: "updateMultiple",
        sheet: sheetName,
        rowNumber: rowNumber,
        updates: updates,
      }),
    });
    const json = await res.json();
    if (!json.success) throw new Error(json.error || "Error al guardar");
    showToast("Guardado en la planilla ✓");
  } catch (err) {
    console.error(err);
    // revertimos si falló
    Object.assign(rowObj, backup);
    renderOverview(); renderKanban(); renderTabla();
    showToast("No se pudo guardar: " + err.message, true);
  }
}

/* =========================================================================
   RENDER: tabla de detalle
   ========================================================================= */

function renderSheetTabs() {
  const counts = {};
  Object.keys(HOJAS_CONFIG).forEach((s) => {
    counts[s] = (RAW_DATA[s] ? RAW_DATA[s].rows.length : 0);
  });
  const html = `
    <div class="sheet-tab ${FILTROS.hoja==='TODAS'?'active':''}" data-sheet="TODAS">
      Todas <span class="tab-count">${Object.values(counts).reduce((a,b)=>a+b,0)}</span>
    </div>
    ${Object.entries(HOJAS_CONFIG).map(([key, cfg]) => `
      <div class="sheet-tab ${FILTROS.hoja===key?'active':''}" data-sheet="${key}">
        ${cfg.etiqueta} <span class="tab-count">${counts[key]}</span>
      </div>
    `).join("")}
  `;
  document.getElementById("sheet-tabs").innerHTML = html;
  document.querySelectorAll(".sheet-tab").forEach((tab) => {
    tab.addEventListener("click", () => {
      FILTROS.hoja = tab.dataset.sheet;
      renderAll();
    });
  });
}

function renderTabla() {
  const items = getItemsFiltrados();
  const head = document.getElementById("table-head");
  const body = document.getElementById("table-body");

  if (FILTROS.hoja !== "TODAS" && RAW_DATA[FILTROS.hoja]) {
    const headers = RAW_DATA[FILTROS.hoja].headers;
    head.innerHTML = "<tr>" + headers.map((h) => `<th>${escapeHtml(h)}</th>`).join("") + "<th>Etapa</th></tr>";
    body.innerHTML = items.map((item) => {
      return "<tr>" + headers.map((h) => {
        const cls = h === HOJAS_CONFIG[item.sheet].colTitulo ? ' class="proyecto-cell"' : "";
        return `<td${cls}>${escapeHtml(item.row[h] ?? "")}</td>`;
      }).join("") + `<td>${statusChip(item.etapa)}</td></tr>`;
    }).join("") || `<tr><td colspan="${headers.length+1}"><div class="empty-state">No hay resultados con estos filtros.</div></td></tr>`;
  } else {
    head.innerHTML = `<tr><th>Hoja</th><th>Proyecto / Declaración</th><th>Eje de trabajo</th><th>Etapa</th></tr>`;
    body.innerHTML = items.map((item) => `
      <tr>
        <td>${HOJAS_CONFIG[item.sheet].etiqueta}</td>
        <td class="proyecto-cell">${escapeHtml(item.titulo)}</td>
        <td>${escapeHtml(item.eje)}</td>
        <td>${statusChip(item.etapa)}</td>
      </tr>
    `).join("") || `<tr><td colspan="4"><div class="empty-state">No hay resultados con estos filtros.</div></td></tr>`;
  }
}

function statusChip(etapa) {
  return `<span class="status-chip state-${etapa.id}">${etapa.nombre}</span>`;
}

/* =========================================================================
   VISTA: selector (Seguimiento / Diplomas)
   ========================================================================= */

function renderSelectorVista() {
  const el = document.getElementById("vista-selector");
  if (!el) return;
  el.innerHTML = `
    <button class="vista-btn ${VISTA_ACTIVA==='seguimiento'?'active':''}" data-vista="seguimiento">Seguimiento</button>
    <button class="vista-btn ${VISTA_ACTIVA==='diplomas'?'active':''}" data-vista="diplomas">Diplomas</button>
  `;
  document.querySelectorAll(".vista-btn").forEach((btn) => {
    btn.addEventListener("click", () => {
      VISTA_ACTIVA = btn.dataset.vista;
      renderAll();
    });
  });

  document.getElementById("vista-seguimiento").style.display = VISTA_ACTIVA === "seguimiento" ? "block" : "none";
  document.getElementById("vista-diplomas").style.display = VISTA_ACTIVA === "diplomas" ? "block" : "none";
}

/* =========================================================================
   RENDER: pestaña de Diplomas
   ========================================================================= */

function renderFiltrosDiplomas() {
  const ejes = getEjesUnicosDiplomas();
  const html = `
    <div class="filter-group" style="flex:2; min-width:220px;">
      <label>Buscar declaración</label>
      <input type="text" id="fd-busqueda" placeholder="Escribí un nombre o expediente…" value="${FILTROS_DIPLOMAS.busqueda.replace(/"/g,'&quot;')}">
    </div>
    <div class="filter-group">
      <label>Eje de trabajo</label>
      <select id="fd-eje">
        <option value="TODOS">Todos los ejes</option>
        ${ejes.map((e) => `<option value="${e.replace(/"/g,'&quot;')}" ${FILTROS_DIPLOMAS.eje===e?"selected":""}>${e}</option>`).join("")}
      </select>
    </div>
    <div class="filter-group" style="min-width:320px;">
      <label>Estado del diploma <span style="text-transform:none; font-weight:400; opacity:0.7;">(elegí uno o varios)</span></label>
      <div class="filter-pills" id="fd-categorias">
        ${DIPLOMA_CATEGORIAS.map((c) => {
          const activa = FILTROS_DIPLOMAS.diploma.includes(c.id);
          return `<span class="pill ${activa?'active':''}" data-cat="${c.id}" style="${activa?`background:${c.color};border-color:${c.color};`:''}">${c.nombre}</span>`;
        }).join("")}
      </div>
    </div>
    <button class="clear-filters" id="fd-clear">Limpiar filtros</button>
  `;
  document.getElementById("filters-zone-diplomas").innerHTML = html;

  document.getElementById("fd-busqueda").addEventListener("input", (e) => {
    FILTROS_DIPLOMAS.busqueda = e.target.value;
    renderKanbanDiplomas();
  });
  document.getElementById("fd-eje").addEventListener("change", (e) => {
    FILTROS_DIPLOMAS.eje = e.target.value;
    renderKanbanDiplomas();
  });
  document.querySelectorAll("#fd-categorias .pill").forEach((pill) => {
    pill.addEventListener("click", () => {
      const id = pill.dataset.cat;
      const idx = FILTROS_DIPLOMAS.diploma.indexOf(id);
      if (idx === -1) FILTROS_DIPLOMAS.diploma.push(id);
      else FILTROS_DIPLOMAS.diploma.splice(idx, 1);
      renderFiltrosDiplomas(); renderKanbanDiplomas();
    });
  });
  document.getElementById("fd-clear").addEventListener("click", () => {
    FILTROS_DIPLOMAS = { eje: "TODOS", diploma: [], busqueda: "" };
    renderFiltrosDiplomas(); renderKanbanDiplomas();
  });
}

function renderKanbanDiplomas() {
  const items = getItemsDiplomasFiltrados();
  const board = document.getElementById("kanban-diplomas-board");
  if (!board) return;

  board.innerHTML = DIPLOMA_CATEGORIAS.map((cat) => {
    const itemsCat = items.filter((i) => i.diplomaCategoria.id === cat.id);
    return `
      <div class="kcol" style="--kcolor:${cat.color}">
        <div class="kcol-head">
          <span class="title">${cat.nombre}</span>
          <span class="count">${itemsCat.length}</span>
        </div>
        <div class="kcol-body">
          ${itemsCat.length ? itemsCat.map((item) => renderCardDiploma(item)).join("") : `<div class="empty-state">Sin items</div>`}
        </div>
      </div>
    `;
  }).join("");

  // El total de referencia para el % excluye "No corresponde": esos diplomas
  // nunca se iban a entregar, así que no deben contarse como "pendientes".
  const itemsQueCorresponden = items.filter((i) => i.diplomaCategoria.id !== "no_corresponde");
  const totalQueCorresponde = itemsQueCorresponden.length;
  const entregados = items.filter((i) => i.diplomaCategoria.id === "entregado").length;
  const pct = totalQueCorresponde ? Math.round((entregados / totalQueCorresponde) * 100) : 0;
  const resumenEl = document.getElementById("diplomas-resumen");
  if (resumenEl) {
    resumenEl.textContent = totalQueCorresponde
      ? `${entregados} de ${totalQueCorresponde} diplomas entregados (${pct}%) — sin contar los que no corresponden`
      : "No hay declaraciones aprobadas con diploma pendiente o entregado, con estos filtros.";
  }
}

function renderCardDiploma(item) {
  return `
    <div class="kcard" style="cursor:default;">
      <span class="title">${escapeHtml(item.titulo)}</span>
      <div class="meta">
        <span class="tag">${escapeHtml(item.eje)}</span>
      </div>
    </div>
  `;
}

/* =========================================================================
   EXPORTAR: Excel y PDF
   ========================================================================= */

function renderExportBarSeguimiento() {
  const el = document.getElementById("export-bar-seguimiento");
  if (!el) return;
  el.innerHTML = `
    <label class="export-checkbox">
      <input type="checkbox" id="exp-seg-filtrado" ${EXPORTAR_SOLO_FILTRADO_SEG ? "checked" : ""}>
      Exportar solo lo filtrado (si no, exporta todo)
    </label>
    <button class="export-btn excel" id="exp-seg-excel">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6"/></svg>
      Exportar a Excel
    </button>
    <button class="export-btn pdf" id="exp-seg-pdf">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6"/></svg>
      Exportar a PDF
    </button>
  `;
  document.getElementById("exp-seg-filtrado").addEventListener("change", (e) => {
    EXPORTAR_SOLO_FILTRADO_SEG = e.target.checked;
  });
  document.getElementById("exp-seg-excel").addEventListener("click", () => exportarSeguimientoExcel());
  document.getElementById("exp-seg-pdf").addEventListener("click", () => exportarSeguimientoPDF());
}

function renderExportBarDiplomas() {
  const el = document.getElementById("export-bar-diplomas");
  if (!el) return;
  el.innerHTML = `
    <label class="export-checkbox">
      <input type="checkbox" id="exp-dip-filtrado" ${EXPORTAR_SOLO_FILTRADO_DIP ? "checked" : ""}>
      Exportar solo lo filtrado (si no, exporta todo)
    </label>
    <button class="export-btn excel" id="exp-dip-excel">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6"/></svg>
      Exportar a Excel
    </button>
    <button class="export-btn pdf" id="exp-dip-pdf">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6"/></svg>
      Exportar a PDF
    </button>
  `;
  document.getElementById("exp-dip-filtrado").addEventListener("change", (e) => {
    EXPORTAR_SOLO_FILTRADO_DIP = e.target.checked;
  });
  document.getElementById("exp-dip-excel").addEventListener("click", () => exportarDiplomasExcel());
  document.getElementById("exp-dip-pdf").addEventListener("click", () => exportarDiplomasPDF());
}

// Devuelve { headers, rows } listos para exportar de la vista Seguimiento,
// respetando si hay una sola hoja seleccionada o "todas" (igual que renderTabla).
function getDatosExportSeguimiento() {
  const items = EXPORTAR_SOLO_FILTRADO_SEG ? getItemsFiltrados() : getAllItems();
  if (FILTROS.hoja !== "TODAS" && RAW_DATA[FILTROS.hoja]) {
    const headers = RAW_DATA[FILTROS.hoja].headers;
    const rows = items.map((item) => headers.map((h) => item.row[h] ?? ""));
    return { headers: [...headers, "Etapa"], rows: items.map((item, idx) => [...rows[idx], item.etapa.nombre]) };
  }
  const headers = ["Hoja", "Proyecto / Declaración", "Eje de trabajo", "Año", "Etapa"];
  const rows = items.map((item) => [
    HOJAS_CONFIG[item.sheet].etiqueta, item.titulo, item.eje, item.anio, item.etapa.nombre,
  ]);
  return { headers, rows };
}

function getDatosExportDiplomas() {
  const items = EXPORTAR_SOLO_FILTRADO_DIP ? getItemsDiplomasFiltrados() : getAllItemsDiplomas();
  const headers = ["Declaración", "Eje de trabajo", "Estado del diploma"];
  const rows = items.map((item) => [item.titulo, item.eje, item.diplomaCategoria.nombre]);
  return { headers, rows };
}

// Genera y descarga un .xlsx en el navegador usando SheetJS (cargado vía CDN).
function descargarExcel(headers, rows, nombreArchivo) {
  if (typeof XLSX === "undefined") {
    showToast("No se pudo cargar la librería de Excel. Revisá tu conexión a internet.", true);
    return;
  }
  const data = [headers, ...rows];
  const ws = XLSX.utils.aoa_to_sheet(data);
  ws["!cols"] = headers.map(() => ({ wch: 28 }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Datos");
  XLSX.writeFile(wb, nombreArchivo);
  showToast("Excel descargado ✓");
}

function exportarSeguimientoExcel() {
  const { headers, rows } = getDatosExportSeguimiento();
  if (!rows.length) { showToast("No hay datos para exportar con estos filtros.", true); return; }
  descargarExcel(headers, rows, "seguimiento_proyectos.xlsx");
}

function exportarDiplomasExcel() {
  const { headers, rows } = getDatosExportDiplomas();
  if (!rows.length) { showToast("No hay datos para exportar con estos filtros.", true); return; }
  descargarExcel(headers, rows, "diplomas.xlsx");
}

// Arma una tabla HTML dentro de #print-area y dispara el diálogo de impresión
// del navegador (la persona elige "Guardar como PDF" ahí).
function imprimirComoPDF(titulo, headers, rows) {
  const printArea = document.getElementById("print-area");
  const fecha = new Date().toLocaleString("es-AR");
  printArea.innerHTML = `
    <h1>${escapeHtml(titulo)}</h1>
    <div class="print-meta">Generado el ${fecha} · ${rows.length} resultado${rows.length === 1 ? "" : "s"}</div>
    <table>
      <thead><tr>${headers.map((h) => `<th>${escapeHtml(h)}</th>`).join("")}</tr></thead>
      <tbody>${rows.map((r) => `<tr>${r.map((c) => `<td>${escapeHtml(c)}</td>`).join("")}</tr>`).join("")}</tbody>
    </table>
  `;
  window.print();
}

function exportarSeguimientoPDF() {
  const { headers, rows } = getDatosExportSeguimiento();
  if (!rows.length) { showToast("No hay datos para exportar con estos filtros.", true); return; }
  imprimirComoPDF("Seguimiento Proyectos Emmanuel Ferrario", headers, rows);
}

function exportarDiplomasPDF() {
  const { headers, rows } = getDatosExportDiplomas();
  if (!rows.length) { showToast("No hay datos para exportar con estos filtros.", true); return; }
  imprimirComoPDF("Entrega de diplomas — Declaraciones de Interés aprobadas", headers, rows);
}

/* =========================================================================
   RENDER ALL
   ========================================================================= */

function renderAll() {
  renderSelectorVista();
  if (VISTA_ACTIVA === "seguimiento") {
    renderOverview();
    renderFiltros();
    renderKanban();
    renderSheetTabs();
    renderTabla();
    renderExportBarSeguimiento();
  } else {
    renderFiltrosDiplomas();
    renderKanbanDiplomas();
    renderExportBarDiplomas();
  }
}

function render() {
  renderAll();
}

/* =========================================================================
   INIT
   ========================================================================= */

cargarDatos(false);
reiniciarAutoRefresh();
