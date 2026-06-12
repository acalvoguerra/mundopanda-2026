const cfg = window.MUNDOPANDA_CONFIG;
const fileInput = document.getElementById("fileInput");
const dropzone = document.getElementById("dropzone");
const uploadBtn = document.getElementById("uploadBtn");
const logEl = document.getElementById("log");
const authState = document.getElementById("authState");
let selectedFile = null;

function log(message) {
  logEl.textContent = message;
}

function getAdminKey() {
  const hash = new URLSearchParams(location.hash.replace(/^#/, ""));
  return hash.get("key") || "";
}

const adminKey = getAdminKey();
if (!adminKey) {
  authState.textContent = "Acceso no válido. Usa el enlace privado con #key=";
  authState.classList.add("status-pill--error");
  dropzone.classList.add("is-disabled");
} else {
  authState.textContent = "Acceso privado validado en navegador";
}

function setFile(file) {
  if (!adminKey) return;
  if (!file) return;
  if (!/\.(xlsx|xlsm|xls)$/i.test(file.name)) {
    selectedFile = null;
    uploadBtn.disabled = true;
    log("El fichero debe ser Excel: .xlsx, .xlsm o .xls");
    return;
  }
  selectedFile = file;
  uploadBtn.disabled = false;
  dropzone.querySelector("strong").textContent = file.name;
  log(`Listo para subir: ${file.name}`);
}

fileInput.addEventListener("change", e => setFile(e.target.files[0]));

["dragenter", "dragover"].forEach(eventName => {
  dropzone.addEventListener(eventName, e => {
    e.preventDefault();
    dropzone.classList.add("is-dragover");
  });
});

["dragleave", "drop"].forEach(eventName => {
  dropzone.addEventListener(eventName, e => {
    e.preventDefault();
    dropzone.classList.remove("is-dragover");
  });
});

dropzone.addEventListener("drop", e => setFile(e.dataTransfer.files[0]));

uploadBtn.addEventListener("click", async () => {
  if (!selectedFile || !adminKey) return;

  uploadBtn.disabled = true;
  log("Subiendo Excel actualizado...");

  try {
    const form = new FormData();
    form.append("file", selectedFile, cfg.EXCEL_FILE);
    form.append("filename", cfg.EXCEL_FILE);
    form.append("sheetName", cfg.SHEET_NAME);

    const response = await fetch(cfg.WORKER_URL, {
      method: "POST",
      headers: { "X-Admin-Key": adminKey },
      body: form
    });

    const text = await response.text();
    let payload;
    try { payload = JSON.parse(text); } catch { payload = { message: text }; }

    if (!response.ok) {
      throw new Error(payload.error || payload.message || `Error HTTP ${response.status}`);
    }

    log("Clasificación actualizada correctamente. GitHub Pages puede tardar 1-2 minutos en reflejar el cambio.");
  } catch (error) {
    console.error(error);
    log(`Error: ${error.message || "Failed to fetch"}`);
  } finally {
    uploadBtn.disabled = false;
  }
});
