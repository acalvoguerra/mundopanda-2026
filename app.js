const cfg = window.MUNDOPANDA_CONFIG;
const statusEl = document.getElementById("status");
const bodyEl = document.getElementById("rankingBody");
const cardsEl = document.getElementById("mobileCards");
const podiumEl = document.getElementById("podium");
const summaryEl = document.getElementById("summaryGrid");
const searchEl = document.getElementById("search");

let players = [];

function n(value) {
  const parsed = Number(String(value ?? "").replace(",", "."));
  return Number.isFinite(parsed) ? parsed : 0;
}

function clean(value) {
  return String(value ?? "").replace(/\s+/g, " ").trim();
}

function initial(name) {
  return clean(name).slice(0, 1).toUpperCase() || "?";
}

function isPlaceholderName(name) {
  const value = clean(name).toLowerCase();
  return !value ||
    value.includes("pegar valores") ||
    value.includes("nombre j") ||
    value === "jugador" ||
    value === "participante";
}

function rankBadge(pos) {
  const rank = Number(pos);
  if (rank === 1) return "🥇";
  if (rank === 2) return "🥈";
  if (rank === 3) return "🥉";
  return String(pos || "-");
}

function rankText(pos) {
  const rank = Number(pos);
  return rank ? `${rank}º` : "-";
}

function parseClas(workbook) {
  const sheet = workbook.Sheets[cfg.SHEET_NAME] || workbook.Sheets["CLAS"] || workbook.Sheets[workbook.SheetNames[0]];
  if (!sheet) throw new Error("No encuentro la hoja CLAS en el Excel.");

  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: false });

  // Formato real de CLAS:
  // fila 4: B=Pos, C=Jugador, D=Puntos Totales
  // datos desde fila 5. En arrays JS: B=1, C=2, D=3.
  const data = [];
  for (let i = 4; i < rows.length; i++) {
    const row = rows[i] || [];
    const jugador = clean(row[2]);
    const puntos = n(row[3]);

    // La hoja tiene filas de plantilla tipo "Pegar Valores Nombre J19".
    // No son participantes reales y no deben contarse ni mostrarse.
    if (isPlaceholderName(jugador)) continue;

    data.push({
      excelOrder: n(row[0]) || i,
      pos: n(row[1]),
      jugador,
      puntos,
      grupos: n(row[4]),
      posGrupos: n(row[5]),
      equipos116: n(row[6]),
      partidos116: n(row[7]),
      equipos18: n(row[8]),
      partidos18: n(row[9]),
      equipos14: n(row[10]),
      partidos14: n(row[11]),
      equipos12: n(row[12]),
      partidos12: n(row[13]),
      equipos3_4: n(row[14]),
      equiposFinal: n(row[15]),
      partido3_4: n(row[16]),
      partidoFinal: n(row[17]),
      cuadroHonor: n(row[18])
    });
  }

  if (!data.length) {
    throw new Error("No he podido leer jugadores reales desde CLAS. Revisa filas desde la 5 y columnas B, C y D.");
  }

  return data.sort((a, b) =>
    b.puntos - a.puntos ||
    a.pos - b.pos ||
    a.excelOrder - b.excelOrder ||
    a.jugador.localeCompare(b.jugador)
  );
}

function renderSummary(data) {
  const leader = data[0];
  const total = data.length;
  const maxPoints = leader?.puntos ?? 0;
  const avg = total ? Math.round(data.reduce((acc, p) => acc + p.puntos, 0) / total) : 0;
  const leaders = data.filter(p => p.puntos === maxPoints).map(p => p.jugador).join(" · ");
  summaryEl.innerHTML = `
    <article><span>Líder${leaders.includes(" · ") ? "es" : ""}</span><strong>${leaders || "-"}</strong></article>
    <article><span>Puntuación líder</span><strong>${maxPoints}</strong></article>
    <article><span>Participantes</span><strong>${total}</strong></article>
    <article><span>Media puntos</span><strong>${avg}</strong></article>
  `;
}

function renderPodium(data) {
  const top = data.filter(p => Number(p.pos) <= 3).slice(0, 6);
  podiumEl.innerHTML = top.map((p) => `
    <article class="podium-card podium-card--rank-${p.pos}">
      <div class="medal">${rankBadge(p.pos)}</div>
      <div class="avatar">${initial(p.jugador)}</div>
      <h3>${p.jugador}</h3>
      <strong>${p.puntos} pts</strong>
      <span>Posición ${rankText(p.pos)}</span>
    </article>
  `).join("");
}

function phaseFinal(p) {
  return p.equiposFinal + p.partidoFinal;
}

function renderRows(data) {
  bodyEl.innerHTML = data.map((p) => `
    <tr>
      <td class="pos">${rankBadge(p.pos)}</td>
      <td><div class="player"><span>${initial(p.jugador)}</span>${p.jugador}</div></td>
      <td class="points">${p.puntos}</td>
      <td>${p.grupos}</td>
      <td>${p.equipos116 + p.partidos116}</td>
      <td>${p.equipos18 + p.partidos18}</td>
      <td>${p.equipos14 + p.partidos14}</td>
      <td>${p.equipos12 + p.partidos12}</td>
      <td>${phaseFinal(p)}</td>
    </tr>
  `).join("");

  cardsEl.innerHTML = data.map((p) => `
    <article class="mobile-card">
      <div class="mobile-card__rank">${rankBadge(p.pos)}<small>${rankText(p.pos)}</small></div>
      <div class="mobile-card__main">
        <div class="mobile-card__name"><span>${initial(p.jugador)}</span><h3>${p.jugador}</h3></div>
        <div class="mini-stats">
          <span>Grupos <b>${p.grupos}</b></span>
          <span>1/16 <b>${p.equipos116 + p.partidos116}</b></span>
          <span>1/8 <b>${p.equipos18 + p.partidos18}</b></span>
          <span>Final <b>${phaseFinal(p)}</b></span>
        </div>
      </div>
      <strong>${p.puntos}<small>pts</small></strong>
    </article>
  `).join("");
}

function applyFilter() {
  const q = clean(searchEl.value).toLowerCase();
  const filtered = q ? players.filter(p => p.jugador.toLowerCase().includes(q)) : players;
  renderRows(filtered);
}

async function loadRanking() {
  try {
    const url = `${cfg.EXCEL_FILE}?v=${Date.now()}`;
    const response = await fetch(url, { cache: "no-store" });
    if (!response.ok) throw new Error(`No puedo descargar ${cfg.EXCEL_FILE}. HTTP ${response.status}`);
    const buffer = await response.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array" });
    players = parseClas(workbook);
    renderSummary(players);
    renderPodium(players);
    renderRows(players);
    const updated = new Date().toLocaleString("es-ES", { dateStyle: "medium", timeStyle: "short" });
    statusEl.textContent = `Actualizado · ${updated}`;
  } catch (error) {
    console.error(error);
    statusEl.textContent = error.message;
    statusEl.classList.add("status-pill--error");
  }
}

searchEl.addEventListener("input", applyFilter);
loadRanking();
