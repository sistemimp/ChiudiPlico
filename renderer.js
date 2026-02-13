const form = document.getElementById('form');
const pickBtn = document.getElementById('pick');
const fileInput = document.getElementById('file');
const statusEl = document.getElementById('status');
const requiredFields = ['file', 'centro', 'nome', 'codice', 'peso', 'spessore'];

pickBtn.addEventListener('click', async () => {
  const file = await window.api.selectFile();
  if (file) fileInput.value = file;
});

form.addEventListener('submit', async (e) => {
  e.preventDefault();
  const missing = requiredFields.find((id) => !String(document.getElementById(id).value || '').trim());
  if (missing) {
    setStatus('Compila tutti i campi obbligatori prima di elaborare.', true);
    document.getElementById(missing).focus();
    return;
  }
  setStatus('Elaborazione in corso...');
  try {
    const outPath = await window.api.processFile({
      filePath: fileInput.value,
      centroImpostazione: document.getElementById('centro').value.trim(),
      nomeLavorazione: document.getElementById('nome').value.trim(),
      codiceLavorazione: document.getElementById('codice').value.trim(),
      pesoInvio: Number(document.getElementById('peso').value) || null,
      spessoreInvio: Number(document.getElementById('spessore').value) || null,
    });
    setStatus(`File generato: ${outPath}`);
  } catch (err) {
    console.error(err);
    setStatus(`Errore: ${err.message}`, true);
  }
});

function setStatus(text, isError = false) {
  statusEl.textContent = text;
  statusEl.style.background = isError ? '#fef2f2' : '#ecfeff';
  statusEl.style.borderColor = isError ? '#fca5a5' : '#22d3ee';
}
