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
};

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
let FILTROS = {
  hoja: "TODAS",
  eje: "TODOS",
  etapa: "TODAS",
  busqueda: "",
};
let autoRefreshTimer = null;

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
      items.push({
        sheet: sheetName,
        etapa: getEtapa(row),
        titulo: getTitulo(row, sheetName),
        eje: getEje(row),
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
  if (FILTROS.etapa !== "TODAS") items = items.filter((i) => i.etapa.id === FILTROS.etapa);
  if (FILTROS.busqueda.trim() !== "") {
    const q = normaliza(FILTROS.busqueda);
    items = items.filter((i) => normaliza(i.titulo).includes(q));
  }
  return items;
}

function getEjesUnicos() {
  const set = new Set();
  getAllItems().forEach((i) => set.add(i.eje));
  return Array.from(set).filter(Boolean).sort();
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
    <div class="filter-group" style="min-width:280px;">
      <label>Etapa</label>
      <div class="filter-pills" id="f-etapas">
        <span class="pill ${FILTROS.etapa==='TODAS'?'active':''}" data-etapa="TODAS">Todas</span>
        ${ETAPAS.map((e) => `<span class="pill ${FILTROS.etapa===e.id?'active':''}" data-etapa="${e.id}" style="${FILTROS.etapa===e.id?`background:${e.color};border-color:${e.color};`:''}">${e.nombre}</span>`).join("")}
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
      FILTROS.etapa = pill.dataset.etapa;
      renderFiltros(); renderOverview(); renderKanban(); renderTabla();
    });
  });
  document.getElementById("f-clear").addEventListener("click", () => {
    FILTROS = { hoja: "TODAS", eje: "TODOS", etapa: "TODAS", busqueda: "" };
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
   RENDER ALL
   ========================================================================= */

function renderAll() {
  renderOverview();
  renderFiltros();
  renderKanban();
  renderSheetTabs();
  renderTabla();
}

function render() {
  renderAll();
}

/* =========================================================================
   INIT
   ========================================================================= */

cargarDatos(false);
reiniciarAutoRefresh();
