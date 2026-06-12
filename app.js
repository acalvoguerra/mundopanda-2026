const EXCEL_FILE = "ranking.xlsx?v=" + Date.now();
const preferredSheetNames = ["CLAS", "WEB_CLASIFICACION", "Clasificacion", "CLASIFICACION", "Clasificación"];
const rankingEl = document.getElementById("ranking");
const podiumEl = document.getElementById("podium");
const statsEl = document.getElementById("stats");
const searchEl = document.getElementById("search");
let players = [];

function normalize(value) {
  return String(value ?? "").trim();
}

function normalizeKey(value) {
  return normalize(value)
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ");
}

function number(value) {
  const raw = String(value ?? "0").replace(",", ".").replace(/[^0-9.\-]/g, "");
  const n = Number(raw || 0);
  return Number.isFinite(n) ? n : 0;
}

function findHeaderRow(matrix) {
  for (let r = 0; r < Math.min(matrix.length, 20); r++) {
    const row = matrix[r].map(normalizeKey);
    const hasJugador = row.some(c => c.includes("jugador") || c.includes("participante") || c.includes("nombre"));
    const hasPuntos = row.some(c => c.includes("puntos totales") || c === "puntos" || c.includes("pts") || c.includes("total"));
    if (hasJugador && hasPuntos) return r;
  }
  return -1;
}

function columnIndex(headers, candidates) {
  const normalized = headers.map(normalizeKey);
  for (const candidate of candidates.map(normalizeKey)) {
    const exact = normalized.findIndex(h => h === candidate);
    if (exact >= 0) return exact;
  }
  for (const candidate of candidates.map(normalizeKey)) {
    const partial = normalized.findIndex(h => h.includes(candidate));
    if (partial >= 0) return partial;
  }
  return -1;
}

function extractPlayersFromSheet(sheet) {
  const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", blankrows: false });
  const headerRowIndex = findHeaderRow(matrix);
  if (headerRowIndex < 0) throw new Error("No encuentro la fila de cabecera con Jugador y Puntos Totales.");

  const headers = matrix[headerRowIndex];
  const idxPos = columnIndex(headers, ["Pos", "Posicion", "Posición", "Puesto"]);
  const idxName = columnIndex(headers, ["Jugador", "Participante", "Nombre", "Usuario"]);
  const idxPoints = columnIndex(headers, ["Puntos Totales", "Puntos", "Total", "Pts"]);
  const idxGrupos = columnIndex(headers, ["F. Grupos", "Fase Grupos"]);
  const idxPosGrupos = columnIndex(headers, ["Pos. Grupos", "Pos Grupos"]);
  const idxEquipos16 = columnIndex(headers, ["Equipos 1/16"]);
  const idxPartidos16 = columnIndex(headers, ["Partidos 1/16"]);
  const idxEquipos8 = columnIndex(headers, ["Equipos 1/8"]);
  const idxPartidos8 = columnIndex(headers, ["Partidos 1/8"]);
  const idxEquipos4 = columnIndex(headers, ["Equipos 1/4"]);
  const idxPartidos4 = columnIndex(headers, ["Partidos 1/4"]);
  const idxEquipos2 = columnIndex(headers, ["Equipos 1/2"]);
  const idxPartidos2 = columnIndex(headers, ["Partidos 1/2"]);
  const idxEquipos34 = columnIndex(headers, ["Equipos 3-4"]);
  const idxEquiposFinal = columnIndex(headers, ["Equipos Final"]);
  const idxPartido34 = columnIndex(headers, ["Partido 3-4"]);
  const idxPartidoFinal = columnIndex(headers, ["Partido Final"]);
  const idxHonor = columnIndex(headers, ["Cuadro de Honor"]);

  if (idxName < 0 || idxPoints < 0) {
    throw new Error("No encuentro columnas de Jugador y Puntos Totales.");
  }

  return matrix.slice(headerRowIndex + 1)
    .map((row, i) => ({
      pos: idxPos >= 0 ? number(row[idxPos]) : i + 1,
      name: normalize(row[idxName]),
      points: number(row[idxPoints]),
      grupos: idxGrupos >= 0 ? number(row[idxGrupos]) : 0,
      posGrupos: idxPosGrupos >= 0 ? number(row[idxPosGrupos]) : 0,
      equipos16: idxEquipos16 >= 0 ? number(row[idxEquipos16]) : 0,
      partidos16: idxPartidos16 >= 0 ? number(row[idxPartidos16]) : 0,
      equipos8: idxEquipos8 >= 0 ? number(row[idxEquipos8]) : 0,
      partidos8: idxPartidos8 >= 0 ? number(row[idxPartidos8]) : 0,
      equipos4: idxEquipos4 >= 0 ? number(row[idxEquipos4]) : 0,
      partidos4: idxPartidos4 >= 0 ? number(row[idxPartidos4]) : 0,
      equipos2: idxEquipos2 >= 0 ? number(row[idxEquipos2]) : 0,
      partidos2: idxPartidos2 >= 0 ? number(row[idxPartidos2]) : 0,
      equipos34: idxEquipos34 >= 0 ? number(row[idxEquipos34]) : 0,
      equiposFinal: idxEquiposFinal >= 0 ? number(row[idxEquiposFinal]) : 0,
      partido34: idxPartido34 >= 0 ? number(row[idxPartido34]) : 0,
      partidoFinal: idxPartidoFinal >= 0 ? number(row[idxPartidoFinal]) : 0,
      honor: idxHonor >= 0 ? number(row[idxHonor]) : 0,
    }))
    .filter(p => p.name && !normalizeKey(p.name).includes("total"))
    .sort((a, b) => b.points - a.points || a.name.localeCompare(b.name, "es"))
    .map((p, i, arr) => ({ ...p, pos: i > 0 && p.points === arr[i - 1].points ? arr[i - 1].pos : i + 1 }));
}

async function loadExcel() {
  try {
    const res = await fetch(EXCEL_FILE, { cache: "no-store" });
    if (!res.ok) throw new Error("No se encontro ranking.xlsx");
    const buffer = await res.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array" });
    const sheetName = preferredSheetNames.find(n => workbook.SheetNames.includes(n)) || workbook.SheetNames[0];
    players = extractPlayersFromSheet(workbook.Sheets[sheetName]);
    render(players);
  } catch (error) {
    rankingEl.innerHTML = `<div class="empty">No se pudo cargar el Excel: ${error.message}</div>`;
    console.error(error);
  }
}

function medal(pos) {
  if (pos === 1) return "🥇";
  if (pos === 2) return "🥈";
  if (pos === 3) return "🥉";
  return `#${pos}`;
}

function initials(name) {
  return name.split(/\s+/).slice(0, 2).map(w => w[0]).join("").toUpperCase();
}

function render(data) {
  renderStats(data);
  renderPodium(data.slice(0, 3));
  renderRanking(data);
}

function renderStats(data) {
  const leader = data[0];
  const total = data.length;
  const maxPoints = leader?.points ?? 0;
  statsEl.innerHTML = `
    <article><span>${total}</span><small>Participantes</small></article>
    <article><span>${maxPoints}</span><small>Puntos lider</small></article>
    <article><span>${leader ? leader.name : "-"}</span><small>Lider actual</small></article>
  `;
}

function renderPodium(top) {
  podiumEl.innerHTML = top.map((p, idx) => `
    <article class="podium-card rank-${idx + 1}">
      <div class="medal">${medal(p.pos)}</div>
      <div class="avatar">${initials(p.name)}</div>
      <h3>${p.name}</h3>
      <strong>${p.points} pts</strong>
    </article>
  `).join("");
}

function renderRanking(data) {
  rankingEl.innerHTML = data.map(p => `
    <article class="player-card">
      <div class="player-main">
        <div class="position">${medal(p.pos)}</div>
        <div class="avatar small">${initials(p.name)}</div>
        <div><h3>${p.name}</h3><p>${p.points} puntos</p></div>
      </div>
      <div class="phase-grid">
        <span><b>${p.grupos}</b><small>F. grupos</small></span>
        <span><b>${p.posGrupos}</b><small>Pos. grupos</small></span>
        <span><b>${p.equipos16 + p.partidos16}</b><small>1/16</small></span>
        <span><b>${p.equipos8 + p.partidos8}</b><small>1/8</small></span>
        <span><b>${p.equipos4 + p.partidos4}</b><small>1/4</small></span>
        <span><b>${p.equipos2 + p.partidos2}</b><small>1/2</small></span>
        <span><b>${p.equiposFinal + p.partidoFinal}</b><small>Final</small></span>
      </div>
    </article>
  `).join("");
}

searchEl.addEventListener("input", e => {
  const q = e.target.value.toLowerCase();
  renderRanking(players.filter(p => p.name.toLowerCase().includes(q)));
});

loadExcel();
