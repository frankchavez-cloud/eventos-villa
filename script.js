
function buildOperativoHTML(events){
 return events.sort((a,b)=>a.hora.localeCompare(b.hora)).map(e=>`
 <div class="operativo-card">
 <h3>📍 ${e.ubicacion||"Sin ubicación"}</h3>
 <div><strong>Hora:</strong> ${e.hora}</div>
 <div><strong>Evento:</strong> ${e.evento}</div>
 <div><strong>Requerimientos:</strong></div>
 <ul>${String(e.requerimientos||"").split(/[\n;]/).filter(x=>x.trim()).map(x=>`<li>${x}</li>`).join("")}</ul>
 </div>`).join("");
}
function showModalTab(tab){
 document.getElementById("eventScroll").classList.toggle("hidden",tab!=="eventos");
 document.getElementById("operativoScroll").classList.toggle("hidden",tab!=="operativo");
}
const PASSWORD = "villa2026";
let eventosBase = [];
let currentDate = new Date(2026, 5, 1);
let currentView = "month";

const feriadosPeru2026 = {
  "2026-01-01": "Año Nuevo",
  "2026-04-02": "Jueves Santo",
  "2026-04-03": "Viernes Santo",
  "2026-05-01": "Día del Trabajo",
  "2026-06-07": "Batalla de Arica",
  "2026-06-29": "San Pedro y San Pablo",
  "2026-07-23": "Día de la Fuerza Aérea del Perú",
  "2026-07-28": "Fiestas Patrias",
  "2026-07-29": "Fiestas Patrias",
  "2026-08-06": "Batalla de Junín",
  "2026-08-30": "Santa Rosa de Lima",
  "2026-10-08": "Combate de Angamos",
  "2026-11-01": "Todos los Santos",
  "2026-12-08": "Inmaculada Concepción",
  "2026-12-09": "Batalla de Ayacucho",
  "2026-12-25": "Navidad"
};

function login() {
  const pass = document.getElementById("passwordInput").value;
  if (pass === PASSWORD) {
    document.getElementById("login").classList.add("hidden");
    document.getElementById("app").classList.remove("hidden");
    cargarExcel();
  } else {
    alert("Contraseña incorrecta");
  }
}
function logout() {
  document.getElementById("app").classList.add("hidden");
  document.getElementById("login").classList.remove("hidden");
}

async function cargarExcel() {
  const status = document.getElementById("statusMessage");
  try {
    const response = await fetch("PLANTILLA CONTROL EVENTOS CAMPUS VILLA.xlsx?t=" + Date.now());
    if (!response.ok) throw new Error("No se pudo leer PLANTILLA CONTROL EVENTOS CAMPUS VILLA.xlsx");
    console.log("URL Excel:", response.url);
console.log("STATUS:", response.status);

    const arrayBuffer = await response.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: "array", cellDates: true });

    const sheetName = workbook.SheetNames.includes("EVENTOS") ? "EVENTOS" : workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    const rows = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false });

    eventosBase = rows.map(mapRow).filter(e => e.fecha && e.evento);

    initFilters();
    status.textContent = `Eventos cargados desde Excel: ${eventosBase.length}`;
    render();
  } catch (err) {
    status.textContent = "Error: no se pudo cargar el Excel. Verifica que 'PLANTILLA CONTROL EVENTOS CAMPUS VILLA.xlsx' esté subido junto a index.html.";
    console.error(err);
  }
}

function parseCSV(text) {
  text = text.replace(/^\uFEFF/, "");
  const rows = [];
  let row = [], field = "", inQuotes = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i], n = text[i + 1];
    if (c === '"' && inQuotes && n === '"') { field += '"'; i++; }
    else if (c === '"') inQuotes = !inQuotes;
    else if ((c === "," || c === ";") && !inQuotes) { row.push(field); field = ""; }
    else if ((c === "\n" || c === "\r") && !inQuotes) {
      if (c === "\r" && n === "\n") i++;
      row.push(field); field = "";
      if (row.some(v => v.trim() !== "")) rows.push(row);
      row = [];
    } else field += c;
  }
  if (field || row.length) { row.push(field); if (row.some(v => v.trim() !== "")) rows.push(row); }
  const headers = rows.shift().map(h => h.trim());
  return rows.map(r => Object.fromEntries(headers.map((h, i) => [h, (r[i] || "").trim()])));
}

function mapRow(r) {
  return {
    fecha: normalizeDate(r["Fecha"]),
    hora: normalizeTime(r["Hora"]),
    area: r["Área Solicitante"] || r["Area Solicitante"] || "",
    ot: r["N° OT"] || r["N OT"] || r["Numero OT"] || "",
    evento: r["Evento / Actividad"] || "",
    ubicacion: r["Ubicación"] || r["Ubicacion"] || "",
    requerimientos: r["Requerimientos"] || "",
    prioridad: normalizePriority(r["Prioridad"] || ""),
    categoria: r["Categoría"] || r["Categoria"] || "",
    estado: r["Estado"] || ""
  };
}

function normalizePriority(p) {
  const s = String(p || "").trim();
  const n = normalize(s);
  if (n.includes("alta")) return "Alta";
  if (n.includes("medio") || n.includes("media")) return "Media";
  if (n.includes("baja") || n.includes("bajo")) return "Baja";
  return s;
}
function normalizeDate(value) {
  if (value === null || value === undefined || value === "") return "";

  if (value instanceof Date && !isNaN(value)) {
    const y = value.getFullYear();
    const m = String(value.getMonth() + 1).padStart(2, "0");
    const d = String(value.getDate()).padStart(2, "0");
    return `${y}-${m}-${d}`;
  }

  if (typeof value === "number") {
    const date = new Date(Math.round((value - 25569) * 86400 * 1000));
    const y = date.getUTCFullYear();
    const m = String(date.getUTCMonth() + 1).padStart(2, "0");
    const d = String(date.getUTCDate()).padStart(2, "0");
    return `${y}-${m}-${d}`;
  }

  let s = String(value || "").trim();
  if (!s) return "";

  let matchIso = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
  if (matchIso) {
    const y = matchIso[1];
    const m = String(matchIso[2]).padStart(2, "0");
    const d = String(matchIso[3]).padStart(2, "0");
    return `${y}-${m}-${d}`;
  }

  const parts = s.split(/[\/\-]/);

if (parts.length === 3) {

    // YYYY-MM-DD
    if (parts[0].length === 4) {
        return `${parts[0]}-${parts[1].padStart(2,"0")}-${parts[2].padStart(2,"0")}`;
    }

    // DD/MM/YY o DD/MM/YYYY
    let year = parts[2];

    if (year.length === 2) {
        year = "20" + year;
    }

    return `${year}-${parts[1].padStart(2,"0")}-${parts[0].padStart(2,"0")}`;
}

  return s;
}
function normalizeTime(value) {
  if (value === null || value === undefined || value === "") return "";

  if (value instanceof Date && !isNaN(value)) {
    const h = String(value.getHours()).padStart(2, "0");
    const m = String(value.getMinutes()).padStart(2, "0");
    return `${h}:${m}`;
  }

  const s = String(value || "").trim();
  if (!s) return "";

  const match = s.match(/(\d{1,2})[:;](\d{2})/);
  if (match) return `${match[1].padStart(2,"0")}:${match[2]}`;

  return s;
}
function updateClock() {
  const now = new Date();
  document.getElementById("currentDate").textContent = now.toLocaleDateString("es-PE", { weekday: "long", day: "2-digit", month: "long", year: "numeric" });
  document.getElementById("currentTime").textContent = now.toLocaleTimeString("es-PE", { hour: "2-digit", minute: "2-digit", second: "2-digit" });
}
setInterval(updateClock, 1000);
window.addEventListener("DOMContentLoaded", updateClock);

function goToday() {
  const today = new Date();
  currentDate = new Date(today.getFullYear(), today.getMonth(), 1);
  render();
}
function changeMonth(delta) { currentDate.setMonth(currentDate.getMonth() + delta); render(); }
function setView(view) {
  currentView = view;
  document.getElementById("btnMonth").classList.toggle("active", view === "month");
  document.getElementById("btnWeek").classList.toggle("active", view === "week");
  render();
}
function toISODate(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth()+1).padStart(2,"0");
  const d = String(date.getDate()).padStart(2,"0");
  return `${y}-${m}-${d}`;
}
function parseDate(iso) { return new Date(iso + "T00:00:00"); }
function normalize(text) { return String(text || "").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, ""); }
function priorityClass(priority) {
  const p = normalize(priority);
  if (p.includes("alta")) return "alta";
  if (p.includes("media")) return "media";
  if (p.includes("baja")) return "baja";
  return "";
}
function filteredEvents() {
  const search = normalize(document.getElementById("searchInput")?.value || "");
  const area = document.getElementById("areaFilter")?.value || "";
  const ubicacion = document.getElementById("ubicacionFilter")?.value || "";
  const estado = document.getElementById("estadoFilter")?.value || "";
  const categoria = document.getElementById("categoriaFilter")?.value || "";
  return eventosBase.filter(e => {
    const blob = normalize(`${e.ot} ${e.evento} ${e.requerimientos} ${e.area} ${e.estado} ${e.categoria} ${e.ubicacion}`);
    return (!search || blob.includes(search)) && (!area || e.area === area) && (!ubicacion || e.ubicacion === ubicacion) && (!estado || e.estado === estado) && (!categoria || e.categoria === categoria);
  });
}
function initFilters() {
  fillSelect("areaFilter", [...new Set(eventosBase.map(e => e.area).filter(Boolean))].sort());
  fillSelect("ubicacionFilter", [...new Set(eventosBase.map(e => e.ubicacion).filter(Boolean))].sort());
  fillSelect("estadoFilter", [...new Set(eventosBase.map(e => e.estado).filter(Boolean))].sort());
  fillSelect("categoriaFilter", [...new Set(eventosBase.map(e => e.categoria).filter(Boolean))].sort());
}
function fillSelect(id, values) {
  const select = document.getElementById(id);
  const firstText = select.options[0].textContent;
  select.innerHTML = `<option value="">${firstText}</option>`;
  values.forEach(v => {
    const option = document.createElement("option");
    option.value = v; option.textContent = v; select.appendChild(option);
  });
}
function render() {
  updateTitle();
  if (currentView === "month") renderMonth(); else renderWeek();
  updateKPIs();
}
function updateTitle() {
  document.getElementById("monthTitle").textContent = currentDate.toLocaleDateString("es-PE", { month: "long", year: "numeric" });
}
function renderMonth() {
  const calendar = document.getElementById("calendar");
  calendar.innerHTML = "";
  ["LUN","MAR","MIÉ","JUE","VIE","SÁB","DOM"].forEach(d => {
    const div = document.createElement("div"); div.className = "day-name"; div.textContent = d; calendar.appendChild(div);
  });
  const year = currentDate.getFullYear(), month = currentDate.getMonth();
  const first = new Date(year, month, 1);
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  let startOffset = first.getDay() - 1; if (startOffset < 0) startOffset = 6;
  for (let i=0; i<startOffset; i++) { const empty = document.createElement("div"); empty.className = "day empty"; calendar.appendChild(empty); }
  for (let day=1; day<=daysInMonth; day++) renderDayCell(calendar, new Date(year, month, day));
}
function renderWeek() {
  const calendar = document.getElementById("calendar");
  calendar.innerHTML = "";
  const today = new Date(currentDate);
  const day = today.getDay() || 7;
  const monday = new Date(today); monday.setDate(today.getDate() - day + 1);
  ["LUN","MAR","MIÉ","JUE","VIE","SÁB","DOM"].forEach(d => {
    const div = document.createElement("div"); div.className = "day-name"; div.textContent = d; calendar.appendChild(div);
  });
  for (let i=0; i<7; i++) { const date = new Date(monday); date.setDate(monday.getDate()+i); renderDayCell(calendar, date, true); }
}
function renderDayCell(calendar, date, isWeek=false) {
  const iso = toISODate(date);
  const events = filteredEvents().filter(e => e.fecha === iso);
  const holiday = feriadosPeru2026[iso];
  const todayIso = toISODate(new Date());
  const div = document.createElement("div");
  div.className = "day";
  if (holiday) div.classList.add("holiday");
  if (iso === todayIso) div.classList.add("today");
  const topPriority = events.some(e => priorityClass(e.prioridad)==="alta") ? "alta" : events.some(e => priorityClass(e.prioridad)==="media") ? "media" : events.some(e => priorityClass(e.prioridad)==="baja") ? "baja" : "";
  if (topPriority) div.classList.add(`priority-${topPriority}`);
  div.onclick = () => openModal(iso);
  const mainArea = events[0]?.area || "";
  const mainLocation = events.find(e => e.ubicacion)?.ubicacion || "";
  div.innerHTML = `
    <div class="day-number">${date.getDate()}</div>
    ${isWeek ? `<div class="area-label">${date.toLocaleDateString("es-PE", { month: "long" })}</div>` : ""}
    ${holiday ? `<div class="holiday-label">🇵🇪 ${holiday}</div>` : ""}
    <div class="event-summary"><span class="event-chip">${events.length} EV</span></div>
    ${mainLocation ? `<div class="location-label">📍 ${mainLocation}</div>` : ""}
    ${mainArea ? `<div class="area-label">🏢 ${mainArea}${events.length > 1 ? " +" + (events.length-1) : ""}</div>` : ""}
    <div class="mini-dots">${events.slice(0,10).map(e => `<i class="mini-dot ${priorityClass(e.prioridad)}"></i>`).join("")}</div>`;
  calendar.appendChild(div);
}
function monthEvents() {
  const year = currentDate.getFullYear(), month = currentDate.getMonth();
  return filteredEvents().filter(e => { const d = parseDate(e.fecha); return d.getFullYear() === year && d.getMonth() === month; });
}
function updateKPIs() {
  const events = monthEvents();
  const year = currentDate.getFullYear(), month = currentDate.getMonth();
  const holidayCount = Object.keys(feriadosPeru2026).filter(iso => { const d = parseDate(iso); return d.getFullYear() === year && d.getMonth() === month; }).length;
  document.getElementById("kpiTotal").textContent = events.length;
  const hoyIso = toISODate(new Date());
  const hoyEventos = events.filter(e => e.fecha === hoyIso);
  document.getElementById("kpiDone").textContent = hoyEventos.length; // HOY
  const manana = new Date();
  manana.setDate(manana.getDate()+1);
  const mananaIso = toISODate(manana);
  document.getElementById("kpiPending").textContent = events.filter(e => e.fecha === mananaIso).length;
  const ahora = new Date();
  let proximoTxt = "0";
  const futuros = hoyEventos.filter(e => e.hora).map(e=>{
    const [h,m]=String(e.hora).split(":");
    const d=new Date(); d.setHours(+h||0,+m||0,0,0);
    return {e,d};
  }).filter(x=>x.d>ahora).sort((a,b)=>a.d-b.d);
  if(futuros.length){
    const mins=Math.floor((futuros[0].d-ahora)/60000);
    proximoTxt=`${Math.floor(mins/60)}h ${mins%60}m`;
  }
  document.getElementById("kpiHigh").textContent = proximoTxt;
  document.getElementById("kpiHoliday").textContent = holidayCount;
}

function splitRequirements(text) {
  return String(text || "")
    .replace(/Requerimiento:/gi, "")
    .replace(/Layout:/gi, "Layout:")
    .replace(/Instalación:/gi, "Instalación:")
    .replace(/Retiro:/gi, "Retiro:")
    .split(/[\/;\n]+/)
    .map(x => x.trim())
    .filter(Boolean);
}

function reqIcon(text) {
  const t = normalize(text);
  if (t.includes("energia") || t.includes("luz") || t.includes("tomacorriente") || t.includes("punto")) return "🔌";
  if (t.includes("mesa") || t.includes("silla") || t.includes("mobiliario") || t.includes("butaca")) return "🪑";
  if (t.includes("limpio") || t.includes("limpieza")) return "🧹";
  if (t.includes("audio") || t.includes("micro") || t.includes("parlante")) return "🎤";
  if (t.includes("banner") || t.includes("señal") || t.includes("senal")) return "🏷️";
  if (t.includes("retiro")) return "🚚";
  return "✓";
}

function eventStatusLabel(e, iso) {
  const today = toISODate(new Date());
  if (iso !== today) return "";
  if (!e.hora) return "";
  const [h,m] = String(e.hora).split(":");
  const start = new Date();
  start.setHours(+h || 0, +m || 0, 0, 0);
  const now = new Date();
  const diff = Math.round((start - now) / 60000);
  if (diff < -60) return `<span class="status-pill done">Finalizado</span>`;
  if (diff <= 0) return `<span class="status-pill live">En curso</span>`;
  if (diff <= 120) return `<span class="status-pill next">Próximo</span>`;
  return "";
}

function dayExecutiveSummary(events, iso) {
  const sedes = new Set(events.map(e => e.ubicacion).filter(Boolean)).size;
  const altas = events.filter(e => priorityClass(e.prioridad) === "alta").length;
  const total = events.length;
  const label = feriadosPeru2026[iso] ? `🇵🇪 ${feriadosPeru2026[iso]} · ` : "";
  return `
    <span>${label}${total} evento(s) registrado(s)</span>
    <span class="modal-metrics">
      <b>${altas}</b> alta prioridad · <b>${sedes}</b> ubicación(es)
    </span>
    <span class="day-progress"><i style="width:${total ? 100 : 0}%"></i></span>
  `;
}

function eventCardHTML(e, index = 0, total = 1, iso = "") {
  const reqs = splitRequirements(e.requerimientos);
  const status = eventStatusLabel(e, iso);
  return `<article class="event-card ${priorityClass(e.prioridad)}" data-event-index="${index + 1}">
      <div class="timeline-dot"></div>
      <div class="event-topline">
        <div class="event-time">${e.hora || "--:--"}</div>
        <div class="event-count">${index + 1} / ${total}</div>
      </div>
      <div class="event-title">${e.evento}</div>
      ${status}
      <div class="badges">
        <span class="badge area">🏢 Área: ${e.area || "-"}</span>
        <span class="badge sede">📍 Sede: ${e.ubicacion || "Sin ubicación"}</span>
        <span class="badge ot">📧 OT/SOL: ${e.ot || "Sin OT"}</span>
        <span class="badge prioridad">🔥 Prioridad: ${e.prioridad || "-"}</span>
      </div>
      <div class="requirements">
        <strong>Requerimientos / atención:</strong>
        <div class="req-chip-list">
          ${reqs.length ? reqs.map(x => `<span class="req-chip">${reqIcon(x)} ${x}</span>`).join("") : `<span class="req-chip">✓ Sin requerimientos registrados</span>`}
        </div>
      </div>
    </article>`;
}
function openModal(iso) {
  const modal = document.getElementById("eventModal");
  const scroll = document.getElementById("eventScroll");
  const events = filteredEvents().filter(e => e.fecha === iso).sort((a,b) => a.hora.localeCompare(b.hora));
  const date = parseDate(iso);

  document.getElementById("modalDate").textContent =
    date.toLocaleDateString("es-PE", { weekday:"long", day:"2-digit", month:"long", year:"numeric" }).toUpperCase();

  document.getElementById("modalInfo").innerHTML = dayExecutiveSummary(events, iso);

  scroll.innerHTML = events.length
    ? events.map((e,i) => eventCardHTML(e, i, events.length, iso)).join("")
    : `<div class="no-events">No hay eventos registrados para este día.</div>`;

  document.getElementById("operativoScroll").innerHTML = buildOperativoHTML(events);
  showModalTab("eventos");
  modal.classList.remove("hidden");
  scroll.scrollTop = 0;

  setTimeout(() => {
    updateFocusedCard();
  }, 220);
}
function closeModal() { document.getElementById("eventModal").classList.add("hidden"); }


function openKpiModal(type) {
  const list = document.getElementById("kpiList");
  const title = document.getElementById("kpiModalTitle");
  const info = document.getElementById("kpiModalInfo");
  const events = monthEvents().sort((a,b)=>a.fecha.localeCompare(b.fecha)||a.hora.localeCompare(b.hora));

 function operativoRows(rows, groupByDate = false){

  const ordered = [...rows].sort(
    (a,b) => a.fecha.localeCompare(b.fecha) || a.hora.localeCompare(b.hora)
  );

  let lastDate = "";
  let sequence = 0;

  const content = ordered.map(e => {
    sequence += 1;

    let dateHeader = "";
    if (groupByDate && e.fecha !== lastDate) {
      lastDate = e.fecha;
      const dateObj = parseDate(e.fecha);
      const dateLabel = dateObj.toLocaleDateString("es-PE", {
        weekday: "long",
        day: "2-digit",
        month: "long",
        year: "numeric"
      }).toUpperCase();

      dateHeader = `
        <div class="agenda-fecha-grupo">
          <span class="agenda-fecha-icon">📅</span>
          <span>${dateLabel}</span>
        </div>`;
    }

    const reqs = String(e.requerimientos || "")
      .split(/[\/;\n]+/)
      .map(x => x.trim())
      .filter(Boolean);

    return `
      ${dateHeader}
      <div class="agenda-item agenda-item-uniforme">
        <div class="agenda-secuencia">${sequence}</div>

        <div class="agenda-hora">
          ${e.hora || "--:--"}
        </div>

        <div class="agenda-evento">
          <div class="agenda-titulo">
            ${e.evento}
          </div>

          <div class="agenda-ambiente">
            📍 ${e.ubicacion || "-"}
          </div>

          <div class="agenda-req">
            ${reqs.length
              ? reqs.map(x => `<span>${x}</span>`).join("")
              : `<span>Sin requerimientos registrados</span>`}
          </div>
        </div>
      </div>`;
  }).join("");

  return `<div class="agenda-resumen">${content}</div>`;
}
  let rows=[];

if(type==="eventos"){
   title.textContent="TOTAL EVENTOS";
   rows=events;
}

else if(type==="hoy"){
   title.textContent="EVENTOS DE HOY";
   rows=events.filter(e=>e.fecha===toISODate(new Date()));
}

else if(type==="manana"){
   title.textContent="EVENTOS DE MAÑANA";
   const manana = new Date();
   manana.setDate(manana.getDate()+1);
   rows = events.filter(e=>e.fecha===toISODate(manana));
}

else if(type==="feriados"){
   title.textContent="FERIADOS PERÚ";
   info.textContent="Feriados del mes seleccionado";
   list.innerHTML = buildFeriadosHTML();
   document.getElementById("kpiModal").classList.remove("hidden");
   return;
}

else if(type==="proximo"){

   title.textContent="PRÓXIMO EVENTO";

   const ahora = new Date();

   rows = events.filter(e=>{

      if(e.fecha !== toISODate(new Date())) return false;

      const [h,m] = String(e.hora||"00:00").split(":");

      const d = new Date();
      d.setHours(+h||0,+m||0,0,0);

      return d > ahora;

   }).sort((a,b)=>a.hora.localeCompare(b.hora));

   rows = rows.slice(0,1);
}
  else { rows=events; }

  info.textContent = type === "eventos" ? `${rows.length} evento(s) ordenados por fecha y hora` : `${rows.length} registro(s)`;
  list.innerHTML = rows.length ? operativoRows(rows, type === "eventos") : `<div class="no-events">No hay eventos.</div>`;
  document.getElementById("kpiModal").classList.remove("hidden");
}
function closeKpiModal() { document.getElementById("kpiModal").classList.add("hidden"); }

function updateFocusedCard() {
  const container = document.getElementById("eventScroll");
  if (!container || container.classList.contains("hidden")) return;

  const cards = container.querySelectorAll(".event-card");
  if (!cards.length) return;

  cards.forEach(card => card.classList.remove("focused"));

  if (container.scrollTop + container.clientHeight >= container.scrollHeight - 24) {
    cards[cards.length - 1].classList.add("focused");
    return;
  }

  const containerTop = container.getBoundingClientRect().top;
  const focusLine = containerTop + Math.min(260, container.clientHeight * 0.38);

  let closest = null;
  let best = Infinity;

  cards.forEach(card => {
    const rect = card.getBoundingClientRect();
    const distance = Math.abs(rect.top - focusLine);

    if (rect.bottom > containerTop + 60 && distance < best) {
      best = distance;
      closest = card;
    }
  });

  (closest || cards[0]).classList.add("focused");
}

document.getElementById("eventScroll").addEventListener("scroll", updateFocusedCard);
document.addEventListener("keydown", e => { if (e.key === "Escape") { closeModal(); closeKpiModal(); } });
