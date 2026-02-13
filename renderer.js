const form = document.getElementById('form');
const pickBtn = document.getElementById('pick');
const fileInput = document.getElementById('file');
const fileDropTarget = document.getElementById('fileDropTarget');
const submitBtn = document.getElementById('submitBtn');
const statusEl = document.getElementById('status');
const requiredFields = ['file', 'centro', 'nome', 'codice', 'peso', 'spessore'];

pickBtn.addEventListener('click', async () => {
  const file = await window.api.selectFile();
  if (file) fileInput.value = file;
});

if (fileDropTarget) {
  const preventDefault = (e) => {
    e.preventDefault();
    e.stopPropagation();
  };

  ['dragenter', 'dragover'].forEach((eventName) => {
    fileDropTarget.addEventListener(eventName, (e) => {
      preventDefault(e);
      fileDropTarget.classList.add('is-dragover');
    });
  });

  ['dragleave', 'dragend'].forEach((eventName) => {
    fileDropTarget.addEventListener(eventName, (e) => {
      preventDefault(e);
      fileDropTarget.classList.remove('is-dragover');
    });
  });

  fileDropTarget.addEventListener('drop', async (e) => {
    preventDefault(e);
    fileDropTarget.classList.remove('is-dragover');
    const droppedFile = e.dataTransfer?.files?.[0];
    if (!droppedFile) return;

    const droppedPath = await resolveDroppedFilePath(e.dataTransfer, droppedFile);
    const droppedName = String(droppedFile.name || '');
    const isXlsx = droppedPath.toLowerCase().endsWith('.xlsx') || droppedName.toLowerCase().endsWith('.xlsx');
    if (!isXlsx) {
      setStatus('Puoi trascinare solo file .xlsx', true);
      return;
    }

    if (!droppedPath) {
      setStatus('Impossibile leggere il percorso completo del file trascinato.', true);
      return;
    }

    fileInput.value = droppedPath;
    setStatus(`File selezionato: ${fileInput.value}`);
  });
}

async function resolveDroppedFilePath(dataTransfer, file) {
  let directPath = '';
  if (window.api?.getPathForFile) {
    try {
      directPath = String(window.api.getPathForFile(file) || '').trim();
    } catch {
      directPath = '';
    }
  }

  if (!directPath) {
    directPath = String(file?.path || '').trim();
  }

  if (directPath) return directPath;

  const uriList = String(dataTransfer?.getData('text/uri-list') || '').trim();
  const uriPath = uriToWindowsPath(uriList);
  if (uriPath) return uriPath;

  const plainText = String(dataTransfer?.getData('text/plain') || '').trim();
  const plainPath = uriToWindowsPath(plainText) || plainText;
  return plainPath && looksLikeAbsolutePath(plainPath) ? plainPath : '';
}

function uriToWindowsPath(value) {
  if (!value) return '';
  const firstLine = value
    .split(/\r?\n/)
    .map((line) => line.trim())
    .find((line) => line && !line.startsWith('#'));
  if (!firstLine) return '';
  if (!firstLine.toLowerCase().startsWith('file://')) return '';

  let decoded = '';
  try {
    decoded = decodeURIComponent(firstLine);
  } catch {
    return '';
  }

  // Windows file URI: file:///C:/path/file.xlsx
  return decoded.replace(/^file:\/\/\/?/i, '').replace(/\//g, '\\');
}

function looksLikeAbsolutePath(p) {
  return /^[a-zA-Z]:\\/.test(String(p || ''));
}

form.addEventListener('submit', async (e) => {
  e.preventDefault();
  setSubmitLoading(true);
  const missing = requiredFields.find((id) => !String(document.getElementById(id).value || '').trim());
  if (missing) {
    setSubmitLoading(false);
    setStatus('Compila tutti i campi obbligatori prima di elaborare.', true);
    document.getElementById(missing).focus();
    return;
  }
  setStatus('Elaborazione in corso...');
  try {
    const plichiA4 = Number(document.querySelector('input[name="plichiA4"]:checked')?.value) || 3;
    const outPath = await window.api.processFile({
      filePath: fileInput.value,
      centroImpostazione: document.getElementById('centro').value.trim(),
      nomeLavorazione: document.getElementById('nome').value.trim(),
      codiceLavorazione: document.getElementById('codice').value.trim(),
      pesoInvio: Number(document.getElementById('peso').value) || null,
      spessoreInvio: Number(document.getElementById('spessore').value) || null,
      plichiPerFoglioA4: plichiA4,
    });
    setStatus(`File generato: ${outPath}`);
  } catch (err) {
    console.error(err);
    setStatus(`Errore: ${err.message}`, true);
  } finally {
    setSubmitLoading(false);
  }
});

function setStatus(text, isError = false) {
  statusEl.textContent = text;
  statusEl.style.background = isError ? '#fef2f2' : '#ecfeff';
  statusEl.style.borderColor = isError ? '#fca5a5' : '#22d3ee';
}

function setSubmitLoading(isLoading) {
  if (!submitBtn) return;
  submitBtn.disabled = isLoading;
  submitBtn.innerHTML = isLoading
    ? '<span class="spinner-border spinner-border-sm me-2" role="status" aria-hidden="true"></span>Elaborazione...'
    : 'Elabora';
}
