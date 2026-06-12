(() => {
  const cfg = window.MUNDOPANDA_CONFIG;
  const fileInput = document.getElementById("fileInput");
  const dropzone = document.getElementById("dropzone");
  const uploadBtn = document.getElementById("uploadBtn");
  const logEl = document.getElementById("log");
  const authState = document.getElementById("authState");
  const fileNameEl = document.getElementById("fileName");
  const dropTitle = document.getElementById("dropTitle");
  let selectedFile = null;

  function log(message, type = "") {
    logEl.textContent = message;
    logEl.dataset.type = type;
  }

  function getAdminKey() {
    const hash = new URLSearchParams(location.hash.replace(/^#/, ""));
    return hash.get("key") || "";
  }

  const adminKey = getAdminKey();
  if (!adminKey) {
    authState.textContent = "Acceso no autorizado";
    authState.classList.add("status-pill--error");
    dropzone.classList.add("is-disabled");
    fileInput.disabled = true;
    log("Usa el enlace privado con #key= para acceder a la subida.", "error");
  } else {
    authState.textContent = "Acceso privado activo";
    log("Esperando fichero...");
  }

  function setFile(file) {
    if (!adminKey) return;
    if (!file) return;

    if (!/\.(xlsx|xlsm|xls)$/i.test(file.name)) {
      selectedFile = null;
      uploadBtn.disabled = true;
      dropTitle.textContent = "Fichero no válido";
      fileNameEl.textContent = "Selecciona un Excel .xlsx, .xlsm o .xls";
      log("El fichero debe ser Excel: .xlsx, .xlsm o .xls", "error");
      return;
    }

    selectedFile = file;
    uploadBtn.disabled = false;
    dropzone.classList.add("has-file");
    dropTitle.textContent = "Excel preparado";
    fileNameEl.textContent = `${file.name} · ${(file.size / 1024 / 1024).toFixed(2)} MB`;
    log(`Listo para subir: ${file.name}`);
  }

  dropzone.addEventListener("click", (event) => {
    if (!adminKey) return;
    if (event.target !== fileInput) fileInput.click();
  });

  dropzone.addEventListener("keydown", (event) => {
    if ((event.key === "Enter" || event.key === " ") && adminKey) {
      event.preventDefault();
      fileInput.click();
    }
  });

  fileInput.addEventListener("change", event => setFile(event.target.files?.[0]));

  ["dragenter", "dragover"].forEach(eventName => {
    dropzone.addEventListener(eventName, event => {
      event.preventDefault();
      if (!adminKey) return;
      dropzone.classList.add("is-dragover");
    });
  });

  ["dragleave", "drop"].forEach(eventName => {
    dropzone.addEventListener(eventName, event => {
      event.preventDefault();
      dropzone.classList.remove("is-dragover");
    });
  });

  dropzone.addEventListener("drop", event => setFile(event.dataTransfer.files?.[0]));

  uploadBtn.addEventListener("click", async () => {
    if (!selectedFile || !adminKey) {
      log("Selecciona primero un fichero Excel.", "error");
      return;
    }

    uploadBtn.disabled = true;
    uploadBtn.textContent = "Subiendo...";
    log("Subiendo Excel actualizado a GitHub...");

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

      log("Clasificación actualizada correctamente. GitHub Pages puede tardar 1-2 minutos en refrescar.", "ok");
      dropTitle.textContent = "Actualización completada";
    } catch (error) {
      console.error(error);
      log(`Error: ${error.message || "Failed to fetch"}`, "error");
    } finally {
      uploadBtn.disabled = false;
      uploadBtn.textContent = "Actualizar clasificación";
    }
  });
})();
