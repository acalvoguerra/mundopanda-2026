const dropzone = document.getElementById("dropzone");
const fileInput = document.getElementById("fileInput");
const uploadBtn = document.getElementById("uploadBtn");
const statusEl = document.getElementById("status");
const lockedEl = document.getElementById("locked");
let selectedFile = null;
let uploadKey = "";

function getUploadKey() {
  const hash = new URLSearchParams(location.hash.replace(/^#/, ""));
  const keyFromUrl = hash.get("key");
  if (keyFromUrl) {
    sessionStorage.setItem("mundopanda_upload_key", keyFromUrl);
    history.replaceState(null, "", location.pathname);
    return keyFromUrl;
  }
  return sessionStorage.getItem("mundopanda_upload_key") || "";
}

function setStatus(text, type = "") {
  statusEl.textContent = text;
  statusEl.className = `status ${type}`;
}

function lockIfNeeded() {
  uploadKey = getUploadKey();
  const locked = !uploadKey;
  lockedEl.hidden = !locked;
  dropzone.classList.toggle("disabled", locked);
  uploadBtn.disabled = true;
  if (locked) setStatus("Subida bloqueada.", "error");
}

async function validateExcel(file) {
  if (!file.name.match(/\.(xlsx|xlsm|xls)$/i)) throw new Error("El fichero debe ser Excel.");
  if (file.size > 8 * 1024 * 1024) throw new Error("El Excel es demasiado grande. Maximo 8 MB.");

  const header = new Uint8Array(await file.slice(0, 4).arrayBuffer());
  const isZipExcel = header[0] === 0x50 && header[1] === 0x4b;
  if (!isZipExcel && file.name.match(/\.(xlsx|xlsm)$/i)) throw new Error("El fichero no parece un Excel valido.");

  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array" });
  const allowed = window.MUNDOPANDA_ALLOWED_SHEETS || [];
  const sheetName = allowed.find(n => workbook.SheetNames.includes(n)) || workbook.SheetNames[0];
  const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: "" });
  if (!rows.length) throw new Error("La hoja de clasificación está vacía.");

  const keys = Object.keys(rows[0]).map(k => k.toLowerCase());
  const hasName = keys.some(k => ["jugador", "participante", "nombre", "usuario"].some(x => k.includes(x)));
  const hasPoints = keys.some(k => ["puntos", "pts", "total"].some(x => k.includes(x)));
  if (!hasName || !hasPoints) throw new Error("No encuentro columnas de jugador/participante y puntos/total.");

  return { sheetName, rows: rows.length };
}

async function setFile(file) {
  if (!file || !uploadKey) return;
  uploadBtn.disabled = true;
  setStatus("Validando Excel...", "working");
  try {
    const info = await validateExcel(file);
    selectedFile = file;
    uploadBtn.disabled = false;
    setStatus(`Excel valido: ${file.name} · hoja ${info.sheetName} · ${info.rows} filas`, "ready");
  } catch (error) {
    selectedFile = null;
    setStatus(error.message, "error");
  }
}

fileInput.addEventListener("change", e => setFile(e.target.files[0]));
["dragenter", "dragover"].forEach(evt => dropzone.addEventListener(evt, e => {
  e.preventDefault();
  if (!uploadKey) return;
  dropzone.classList.add("dragging");
}));
["dragleave", "drop"].forEach(evt => dropzone.addEventListener(evt, e => {
  e.preventDefault();
  dropzone.classList.remove("dragging");
}));
dropzone.addEventListener("drop", e => setFile(e.dataTransfer.files[0]));

uploadBtn.addEventListener("click", async () => {
  if (!selectedFile || !uploadKey) return;
  if (!window.MUNDOPANDA_UPLOAD_ENDPOINT || window.MUNDOPANDA_UPLOAD_ENDPOINT.includes("TU-WORKER")) {
    setStatus("Falta configurar admin-config.js con la URL del Worker.", "error");
    return;
  }
  uploadBtn.disabled = true;
  setStatus("Subiendo Excel actualizado...", "working");

  const form = new FormData();
  form.append("file", selectedFile, "ranking.xlsx");

  try {
    const response = await fetch(window.MUNDOPANDA_UPLOAD_ENDPOINT, {
      method: "POST",
      headers: { "X-Upload-Key": uploadKey },
      body: form,
    });
    const result = await response.json().catch(() => ({}));
    if (!response.ok) throw new Error(result.error || "Error al subir el fichero");
    setStatus("Clasificacion actualizada correctamente.", "success");
  } catch (error) {
    setStatus(error.message, "error");
    uploadBtn.disabled = false;
  }
});

lockIfNeeded();
