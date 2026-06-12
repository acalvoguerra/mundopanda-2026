const EXCEL_FILE = "ranking.xlsx?v=" + Date.now();
const preferredSheetNames = ["WEB_CLASIFICACION", "Clasificacion", "CLASIFICACION", "Clasificación"];
const rankingEl = document.getElementById("ranking");
const podiumEl = document.getElementById("podium");
const statsEl = document.getElementById("stats");
const searchEl = document.getElementById("search");
let players = [];

function normalize(value) {
  return String(value ?? "").trim();
}

function number(value) {
  const n = Number(String(value ?? "0").replace(",", "."));
  return Number.isFinite(n) ? n : 0;
}

function pick(row, names) {
  const keys = Object.keys(row);
  const found = keys.find(k => names.some(n => k.toLowerCase().includes(n.toLowerCase())));
  return found ? row[found] : "";
}

function mapRows(rows) {
  return rows.map((row, i) => ({
    pos: number(pick(row, ["pos", "posición", "puesto"])) || i + 1,
    name: normalize(pick(row, ["jugador", "participante", "nombre", "usuario"])) || `Jugador ${i + 1}`,
    points: number(pick(row, ["puntos", "total", "pts"])),
    grupos: number(pick(row, ["grupos", "fase grupos", "liguilla"])),
    dieciseis: number(pick(row, ["1/16", "dieciseis", "dieciseisavos"])),
    octavos: number(pick(row, ["1/8", "octavos"])),
    cuartos: number(pick(row, ["1/4", "cuartos"])),
    semis: number(pick(row, ["1/2", "semifinal", "semis"])),
    final: number(pick(row, ["final"])),
  })).filter(p => p.name && p.points >= 0)
    .sort((a,b) => b.points - a.points || a.name.localeCompare(b.name))
    .map((p, i, arr) => ({ ...p, pos: i > 0 && p.points === arr[i-1].points ? arr[i-1].pos : i + 1 }));
}

async function loadExcel() {
  try {
    const res = await fetch(EXCEL_FILE, { cache: "no-store" });
    if (!res.ok) throw new Error("No se encontro ranking.xlsx");
    const buffer = await res.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array" });
    const sheetName = preferredSheetNames.find(n => workbook.SheetNames.includes(n)) || workbook.SheetNames[0];
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: "" });
    players = mapRows(rows);
    render(players);
  } catch (error) {
    rankingEl.innerHTML = `<div class="empty">No se pudo cargar el Excel. Sube un fichero desde admin.html.</div>`;
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
  return name.split(/\s+/).slice(0,2).map(w => w[0]).join("").toUpperCase();
}

function render(data) {
  renderStats(data);
  renderPodium(data.slice(0,3));
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
        <span><b>${p.grupos}</b><small>Grupos</small></span>
        <span><b>${p.dieciseis}</b><small>1/16</small></span>
        <span><b>${p.octavos}</b><small>1/8</small></span>
        <span><b>${p.cuartos}</b><small>1/4</small></span>
        <span><b>${p.semis}</b><small>1/2</small></span>
        <span><b>${p.final}</b><small>Final</small></span>
      </div>
    </article>
  `).join("");
}

searchEl.addEventListener("input", e => {
  const q = e.target.value.toLowerCase();
  renderRanking(players.filter(p => p.name.toLowerCase().includes(q)));
});

loadExcel();
