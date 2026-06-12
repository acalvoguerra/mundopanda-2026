(() => {
  const cfg = window.MUNDOPANDA_CONFIG;
  const statusEl = document.getElementById("status");
  const bodyEl = document.getElementById("rankingBody");
  const cardsEl = document.getElementById("mobileCards");
  const podiumEl = document.getElementById("podium");
  const summaryEl = document.getElementById("summaryGrid");
  const searchEl = document.getElementById("search");
  let players = [];

  const n = value => {
    const parsed = Number(String(value ?? "").replace(",", "."));
    return Number.isFinite(parsed) ? parsed : 0;
  };
  const clean = value => String(value ?? "").trim();
  const esc = value => clean(value).replace(/[&<>'"]/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;","'":"&#39;","\"":"&quot;"}[c]));
  const initial = name => clean(name).slice(0, 1).toUpperCase() || "?";
  const medal = (pos, index) => index === 0 ? "🥇" : index === 1 ? "🥈" : index === 2 ? "🥉" : pos;

  function parseClas(workbook) {
    const sheet = workbook.Sheets[cfg.SHEET_NAME] || workbook.Sheets["CLAS"] || workbook.Sheets[workbook.SheetNames[0]];
    if (!sheet) throw new Error("No encuentro la hoja CLAS en el Excel.");
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: false });
    const data = [];

    // Formato real de CLAS: fila 4 encabezados, datos desde fila 5.
    // B=Pos, C=Jugador, D=Puntos Totales, E=F.Grupos...
    for (let i = 4; i < rows.length; i++) {
      const row = rows[i] || [];
      const jugador = clean(row[2]);
      if (!jugador || /jugador|participante/i.test(jugador)) continue;
      data.push({
        pos: n(row[1]),
        jugador,
        puntos: n(row[3]),
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
    if (!data.length) throw new Error("No he podido leer jugadores desde CLAS. Deben estar desde la fila 5 y columnas B, C y D.");
    return data.sort((a, b) => b.puntos - a.puntos || a.pos - b.pos || a.jugador.localeCompare(b.jugador));
  }

  function renderSummary(data) {
    const leader = data[0];
    const total = data.length;
    const maxPoints = leader?.puntos ?? 0;
    const avg = total ? Math.round(data.reduce((acc, p) => acc + p.puntos, 0) / total) : 0;
    const perseguidor = data[1] ? Math.max(0, maxPoints - data[1].puntos) : 0;
    summaryEl.innerHTML = `
      <article><span>Líder</span><strong>${esc(leader?.jugador ?? "-")}</strong></article>
      <article><span>Puntos líder</span><strong>${maxPoints}</strong></article>
      <article><span>Participantes</span><strong>${total}</strong></article>
      <article><span>Ventaja</span><strong>${perseguidor}</strong></article>
      <article><span>Media</span><strong>${avg}</strong></article>
    `;
  }

  function renderPodium(data) {
    const top = data.slice(0, 3);
    podiumEl.innerHTML = top.map((p, i) => `
      <article class="podium-card podium-card--${i + 1}">
        <div class="medal">${medal(p.pos, i)}</div>
        <div class="avatar">${initial(p.jugador)}</div>
        <h3>${esc(p.jugador)}</h3>
        <strong>${p.puntos} pts</strong>
        <span>Posición ${p.pos || i + 1}</span>
      </article>
    `).join("");
  }

  function renderRows(data) {
    bodyEl.innerHTML = data.map((p, i) => `
      <tr>
        <td class="pos">${medal(p.pos, i)}</td>
        <td><div class="player"><span>${initial(p.jugador)}</span>${esc(p.jugador)}</div></td>
        <td class="points">${p.puntos}</td>
        <td>${p.grupos}</td>
        <td>${p.equipos116 + p.partidos116}</td>
        <td>${p.equipos18 + p.partidos18}</td>
        <td>${p.equipos14 + p.partidos14}</td>
        <td>${p.equipos12 + p.partidos12}</td>
        <td>${p.equiposFinal + p.partidoFinal}</td>
      </tr>
    `).join("");

    cardsEl.innerHTML = data.map((p, i) => `
      <article class="mobile-card">
        <div class="mobile-card__rank">${medal(p.pos, i)}</div>
        <div>
          <h3>${esc(p.jugador)}</h3>
          <p>Grupos ${p.grupos} · Eliminatorias ${p.equipos116 + p.partidos116 + p.equipos18 + p.partidos18 + p.equipos14 + p.partidos14 + p.equipos12 + p.partidos12 + p.equiposFinal + p.partidoFinal}</p>
        </div>
        <strong>${p.puntos} pts</strong>
      </article>
    `).join("");
  }

  function applyFilter() {
    const q = clean(searchEl.value).toLowerCase();
    renderRows(q ? players.filter(p => p.jugador.toLowerCase().includes(q)) : players);
  }

  async function loadRanking() {
    try {
      const response = await fetch(`${cfg.EXCEL_FILE}?v=${Date.now()}`, { cache: "no-store" });
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
})();
