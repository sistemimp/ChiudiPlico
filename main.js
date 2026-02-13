const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const ExcelJS = require('exceljs');
const fs = require('fs');

const END_ARROW = '\u2192';
const UNKNOWN_BACINO = 'MIX BACINI';

// Load bacino ranges pre-parsed from PDF
const bacinoRanges = JSON.parse(
  fs.readFileSync(path.join(__dirname, 'bacini.json'), 'utf-8')
);

const COLUMN_ORDER = [
  'id',
  'nome',
  'indirizzo',
  'cap',
  'citta',
  'provincia',
  'spare1',
  'spare2',
  'spare3',
  'spare4',
  'spare5',
];

function createWindow() {
  const win = new BrowserWindow({
    width: 1100,
    height: 760,
    icon: path.join(__dirname, 'img', 'logo.ico'),
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
    },
  });
  win.loadFile('index.html');
}

app.whenReady().then(() => {
  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});

ipcMain.handle('select-file', async () => {
  const res = await dialog.showOpenDialog({
    filters: [{ name: 'Excel', extensions: ['xlsx'] }],
    properties: ['openFile'],
  });
  if (res.canceled || res.filePaths.length === 0) return null;
  return res.filePaths[0];
});

ipcMain.handle('process-file', async (_evt, payload) => {
  const {
    filePath,
    centroImpostazione,
    nomeLavorazione,
    codiceLavorazione,
    pesoInvio, // grammi
    spessoreInvio, // millimetri
    plichiPerFoglioA4,
  } = payload || {};
  if (!filePath) throw new Error('Nessun file selezionato');
  const plichiPerFoglio = [2, 3].includes(Number(plichiPerFoglioA4))
    ? Number(plichiPerFoglioA4)
    : 3;

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  const inputSheet = workbook.worksheets[0];

  // Normalize headers
  const headerMap = {};
  inputSheet.getRow(1).eachCell((cell, colNumber) => {
    const key = normalizeText(readCellValue(cell.value)).toLowerCase();
    if (key) headerMap[key] = colNumber;
  });

  const records = [];
  inputSheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    const rec = {};
    COLUMN_ORDER.forEach((k) => {
      const col = headerMap[k];
      const raw = col ? readCellValue(row.getCell(col).value) : '';
      rec[k] = k === 'cap' ? normalizeCap(raw) : normalizeText(raw);
    });
    // Skip empty lines
    if (rec.id || rec.nome || rec.cap) {
      rec._bacino = findBacino(rec.cap) || UNKNOWN_BACINO;
      records.push(rec);
    }
  });

  // Ordine richiesto: bacino, provincia, cap, citta
  const sortKey = (r) => [
    bacinoSortKey(r._bacino || UNKNOWN_BACINO),
    r.provincia || '',
    r.cap || '',
    r.citta || '',
  ];
  records.sort((a, b) => {
    const ka = sortKey(a);
    const kb = sortKey(b);
    for (let i = 0; i < ka.length; i++) {
      const cmp = compareSortValue(ka[i], kb[i]);
      if (cmp !== 0) return cmp;
    }
    return 0;
  });
  // progressivo etichetta dopo il riordino
  records.forEach((r, idx) => (r._seq = idx + 1));

  const maxByWeight =
    pesoInvio && Number(pesoInvio) > 0
      ? Math.floor(10000 / Number(pesoInvio)) // 10 kg plico
      : Number.POSITIVE_INFINITY;
  const maxByThickness =
    spessoreInvio && Number(spessoreInvio) > 0
      ? Math.floor(120 / Number(spessoreInvio)) // 12 cm plico
      : Number.POSITIVE_INFINITY;
  const maxPerPlico = Math.max(1, Math.min(maxByWeight, maxByThickness) || 1);
  const MIN_PIECES = 5;

  // Progressive counter across all plichi
  let plicoCounter = 1;
  let plichi = [];

  // Più cicli su criteri "fini" prima di usare il bacino.
  const levels = [
    { name: 'cap+citta', key: (r) => (r.cap ? r.cap + '|' + r.citta : null) },
    { name: 'cap', key: (r) => r.cap || null },
    { name: 'provincia', key: (r) => r.provincia || null },
    {
      name: 'provincia+citta_capoluogo',
      key: (r) =>
        r.provincia &&
        r.citta &&
        normalize(r.citta) === normalize(r.provincia)
          ? r.provincia
          : null,
    },
    { name: 'cap+citta', key: (r) => (r.cap ? r.cap + '|' + r.citta : null) },
    { name: 'cap', key: (r) => r.cap || null },
    { name: 'provincia', key: (r) => r.provincia || null },
    {
      name: 'provincia+citta_capoluogo',
      key: (r) =>
        r.provincia &&
        r.citta &&
        normalize(r.citta) === normalize(r.provincia)
          ? r.provincia
          : null,
    },
    { name: 'bacino', key: (r) => findBacino(r.cap) || null },
    { name: 'mix', key: () => 'MIX' },
  ];

  let remaining = records;
  levels.forEach((level, idx) => {
    const isLast = idx === levels.length - 1;
    const buckets = new Map();
    remaining.forEach((rec) => {
      const k = level.key(rec);
      if (!k) return;
      if (!buckets.has(k)) buckets.set(k, []);
      buckets.get(k).push(rec);
    });

    const nextRemaining = [];

    buckets.forEach((arr) => {
      // Evita plichi con progressivi non contigui nel foglio riordinato.
      const runs = splitByConsecutiveSeq(arr);
      runs.forEach((run) => {
        const { chunks, leftovers } = splitGroup(
          run,
          maxPerPlico,
          MIN_PIECES,
          isLast
        );
        nextRemaining.push(...leftovers);

        chunks.forEach((chunk) => {
          const first = chunk[0];
          const bacino = buildBacinoLabel(chunk);
          const destinazione = buildDestinazione(
            level.name,
            first.cap,
            first.citta,
            first.provincia,
            bacino
          );
          const start = chunk[0]._seq;
          const end = chunk[chunk.length - 1]._seq;
          const barcode = buildBarcode(codiceLavorazione, first.cap, start, end);

          chunk.forEach((rec, idx) => {
            rec._plico = plicoCounter;
            rec._fine = idx === chunk.length - 1 ? END_ARROW : '';
          });

          plichi.push({
            plico: plicoCounter,
            groupType: level.name,
            items: chunk,
            cap: first.cap,
            provincia: first.provincia,
            citta: first.citta,
            bacino,
            destinazione,
            numeroInvii: chunk.length,
            pesoPlico:
              pesoInvio && Number(pesoInvio) > 0
                ? Number(pesoInvio) * chunk.length
                : '',
            altezzaPlico:
              spessoreInvio && Number(spessoreInvio) > 0
                ? Number(spessoreInvio) * chunk.length
                : '',
            start,
            end,
            barcode,
          });
          plicoCounter += 1;
        });
      });
    });

    remaining
      .filter((r) => !level.key(r))
      .forEach((r) => nextRemaining.push(r));

    remaining = nextRemaining;
  });

  const sheet2 = new ExcelJS.Workbook().addWorksheet('Etichettatura Plichi');
  sheet2.addRow([
    'Centro impostazione',
    'Bacino',
    'Destinazione',
    'Progressivo plico',
    'Numero invii',
    'Peso plico (g)',
    'Altezza plico (mm)',
    'Progressivo iniziale',
    'Progressivo finale',
    'Nome lavorazione',
    'Codice lavorazione',
    'Barcode',
  ]);

  // Ordina i plichi con la stessa logica dei dati
  plichi.sort((a, b) => {
    const ka = [bacinoSortKey(a.bacino || UNKNOWN_BACINO), a.provincia || '', a.cap || '', a.citta || ''];
    const kb = [bacinoSortKey(b.bacino || UNKNOWN_BACINO), b.provincia || '', b.cap || '', b.citta || ''];
    for (let i = 0; i < ka.length; i++) {
      const cmp = compareSortValue(ka[i], kb[i]);
      if (cmp !== 0) return cmp;
    }
    return 0;
  });

  // Se un plico di tipo bacino contiene province diverse, preferiamo separarlo per provincia.
  // Questo rende etichettatura e range piu' coerenti con l'aggregazione richiesta per provincia.
  let splitIdx = 0;
  while (splitIdx < plichi.length) {
    const p = plichi[splitIdx];
    if (!p || p.groupType !== 'bacino') {
      splitIdx += 1;
      continue;
    }

    const byProv = new Map();
    p.items.forEach((rec) => {
      const prov = rec.provincia || '';
      if (!byProv.has(prov)) byProv.set(prov, []);
      byProv.get(prov).push(rec);
    });

    if (byProv.size <= 1) {
      splitIdx += 1;
      continue;
    }

    const splitPlichi = Array.from(byProv.entries()).map(([prov, items]) => {
      items.sort((a, b) => Number(a._seq) - Number(b._seq));
      const first = items[0];
      const last = items[items.length - 1];
      const bacino = buildBacinoLabel(items);
      return {
        plico: 0,
        groupType: 'provincia',
        items,
        cap: first.cap,
        provincia: prov,
        citta: first.citta,
        bacino,
        destinazione: buildDestinazione('provincia', first.cap, first.citta, prov, bacino),
        numeroInvii: items.length,
        pesoPlico:
          pesoInvio && Number(pesoInvio) > 0
            ? Number(pesoInvio) * items.length
            : '',
        altezzaPlico:
          spessoreInvio && Number(spessoreInvio) > 0
            ? Number(spessoreInvio) * items.length
            : '',
        start: first._seq,
        end: last._seq,
        barcode: buildBarcode(codiceLavorazione, first.cap, first._seq, last._seq),
      };
    });

    plichi.splice(splitIdx, 1, ...splitPlichi);
    splitIdx += splitPlichi.length;
  }
  const softMaxPerPlico = Math.max(3, Math.ceil(maxPerPlico * 1.35));

  const getPlicoZoneKey = (p) =>
    [p.bacino || UNKNOWN_BACINO, p.provincia || '', p.cap || '', p.citta || ''].join('|');

  const refreshPlicoMeta = (p) => {
    if (!p.items.length) return;
    const first = p.items[0];
    const last = p.items[p.items.length - 1];
    p.cap = first.cap;
    p.provincia = first.provincia;
    p.citta = first.citta;
    p.bacino = buildBacinoLabel(p.items);
    p.destinazione = resolveDestinazioneForGroup(
      p.groupType || 'cap',
      first.cap,
      first.citta,
      first.provincia,
      p.bacino
    );
    p.numeroInvii = p.items.length;
    p.start = first._seq;
    p.end = last._seq;
    p.barcode = buildBarcode(codiceLavorazione, first.cap, p.start, p.end);
    p.pesoPlico =
      pesoInvio && Number(pesoInvio) > 0
        ? Number(pesoInvio) * p.items.length
        : '';
    p.altezzaPlico =
      spessoreInvio && Number(spessoreInvio) > 0
        ? Number(spessoreInvio) * p.items.length
        : '';
  };

  // Evita plichi da 1-2 pezzi: accorpa al precedente/successivo della stessa zona.
  let merged = true;
  while (merged) {
    merged = false;
    for (let i = 0; i < plichi.length; i++) {
      const curr = plichi[i];
      if (!curr || curr.items.length > 2) continue;
      const zone = getPlicoZoneKey(curr);

      const prev = i > 0 ? plichi[i - 1] : null;
      const next = i < plichi.length - 1 ? plichi[i + 1] : null;

      const canMergePrev =
        prev &&
        getPlicoZoneKey(prev) === zone &&
        prev.items.length + curr.items.length <= softMaxPerPlico;
      const canMergeNext =
        next &&
        getPlicoZoneKey(next) === zone &&
        next.items.length + curr.items.length <= softMaxPerPlico;

      if (canMergePrev) {
        prev.items.push(...curr.items);
        refreshPlicoMeta(prev);
        plichi.splice(i, 1);
        merged = true;
        break;
      }

      if (canMergeNext) {
        next.items.unshift(...curr.items);
        refreshPlicoMeta(next);
        plichi.splice(i, 1);
        merged = true;
        break;
      }

      // Fallback: se resta un plico da 1/2 pezzi, accorpa comunque sulla stessa zona
      // anche oltre il limite "soft".
      if (prev && getPlicoZoneKey(prev) === zone) {
        prev.items.push(...curr.items);
        refreshPlicoMeta(prev);
        plichi.splice(i, 1);
        merged = true;
        break;
      }

      if (next && getPlicoZoneKey(next) === zone) {
        next.items.unshift(...curr.items);
        refreshPlicoMeta(next);
        plichi.splice(i, 1);
        merged = true;
        break;
      }
    }
  }

  // Consolidamento plichi singoli (1 pezzo) per provincia, se adiacenti.
  const getProvKey = (p) => p.provincia || '';

  let k = 0;
  while (k < plichi.length) {
    const curr = plichi[k];
    if (!curr || curr.items.length !== 1 || !getProvKey(curr)) {
      k += 1;
      continue;
    }

    const prov = getProvKey(curr);
    const collect = [curr];
    let j = k + 1;

    while (j < plichi.length) {
      const pj = plichi[j];
      if (!pj || getProvKey(pj) !== prov) break;
      if (pj.items.length === 1) {
        collect.push(pj);
        j += 1;
        continue;
      }
      break;
    }

    if (collect.length <= 1) {
      k += 1;
      continue;
    }

    const items = collect
      .flatMap((x) => x.items)
      .sort((a, b) => Number(a._seq) - Number(b._seq));
    const bacino = buildBacinoLabel(items);
    const start = items[0]._seq;
    const end = items[items.length - 1]._seq;

    const aggregated = {
      plico: 0,
      groupType: 'provincia',
      items,
      cap: '',
      provincia: prov,
      citta: '',
      bacino,
      destinazione: buildProvinciaBacinoDestinazione(prov, bacino),
      numeroInvii: items.length,
      pesoPlico:
        pesoInvio && Number(pesoInvio) > 0
          ? Number(pesoInvio) * items.length
          : '',
      altezzaPlico:
        spessoreInvio && Number(spessoreInvio) > 0
          ? Number(spessoreInvio) * items.length
          : '',
      start,
      end,
      barcode: buildBarcode(codiceLavorazione, 'MIX', start, end),
    };

    plichi.splice(k, collect.length, aggregated);
    k += 1;
  }
  // Consolidamento plichi piccoli su blocchi adiacenti Provincia/Bacino.
  // Evita accorpamenti "globali" che creano intervalli Progressivo incoerenti.
  const getProvBacKey = (p) => [p.bacino || UNKNOWN_BACINO, p.provincia || ''].join('|');

  let i = 0;
  while (i < plichi.length) {
    const curr = plichi[i];
    if (!curr || curr.items.length >= MIN_PIECES) {
      i += 1;
      continue;
    }

    const zone = getProvBacKey(curr);
    let j = i;
    let total = 0;
    const collect = [];

    // Raccoglie solo plichi piccoli adiacenti della stessa provincia/bacino.
    while (j < plichi.length) {
      const pj = plichi[j];
      if (!pj || getProvBacKey(pj) !== zone || pj.items.length >= MIN_PIECES) break;
      collect.push(pj);
      total += pj.items.length;
      j += 1;
    }

    if (collect.length <= 1) {
      i += 1;
      continue;
    }

    const [bacino, provincia] = zone.split('|');
    const items = collect
      .flatMap((x) => x.items)
      .sort((a, b) => Number(a._seq) - Number(b._seq));

    const start = items[0]._seq;
    const end = items[items.length - 1]._seq;

    const aggregated = {
      plico: 0,
      groupType: 'provincia',
      items,
      cap: '',
      provincia,
      citta: '',
      bacino,
      destinazione: buildProvinciaBacinoDestinazione(provincia, bacino),
      numeroInvii: items.length,
      pesoPlico:
        pesoInvio && Number(pesoInvio) > 0
          ? Number(pesoInvio) * items.length
          : '',
      altezzaPlico:
        spessoreInvio && Number(spessoreInvio) > 0
          ? Number(spessoreInvio) * items.length
          : '',
      start,
      end,
      barcode: buildBarcode(codiceLavorazione, 'MIX', start, end),
    };

    // Sostituisce il blocco adiacente con un solo plico provincia/bacino.
    plichi.splice(i, collect.length, aggregated);
    i += 1;
  }

  // Fase di recupero: prima del consolidamento finale, prova a convertire plichi bacino
  // in plichi provincia quando esiste una corrispondenza provinciale aggregabile.
  const byBacinoProvincia = new Map();
  plichi.forEach((p, idx) => {
    if (!p || p.groupType !== 'bacino') return;
    const prov = p.provincia || '';
    if (!prov) return;
    const key = [p.bacino || UNKNOWN_BACINO, prov].join('|');
    if (!byBacinoProvincia.has(key)) byBacinoProvincia.set(key, []);
    byBacinoProvincia.get(key).push(idx);
  });

  if (byBacinoProvincia.size > 0) {
    const toRemove = new Set();
    const toAdd = [];

    byBacinoProvincia.forEach((idxs, key) => {
      if (idxs.length < 2) return;
      const [bacino, provincia] = key.split('|');
      const items = idxs
        .flatMap((i) => plichi[i].items)
        .sort((a, b) => Number(a._seq) - Number(b._seq));
      if (items.length < MIN_PIECES) return;

      const first = items[0];
      const last = items[items.length - 1];
      toAdd.push({
        plico: 0,
        groupType: 'provincia',
        items,
        cap: first.cap,
        provincia,
        citta: '',
        bacino,
        destinazione: buildProvinciaBacinoDestinazione(provincia, bacino),
        numeroInvii: items.length,
        pesoPlico:
          pesoInvio && Number(pesoInvio) > 0
            ? Number(pesoInvio) * items.length
            : '',
        altezzaPlico:
          spessoreInvio && Number(spessoreInvio) > 0
            ? Number(spessoreInvio) * items.length
            : '',
        start: first._seq,
        end: last._seq,
        barcode: buildBarcode(codiceLavorazione, first.cap, first._seq, last._seq),
      });
      idxs.forEach((i) => toRemove.add(i));
    });

    if (toRemove.size > 0) {
      plichi = plichi.filter((_, idx) => !toRemove.has(idx));
      plichi.push(...toAdd);
    }
  }
  // Consolidamento forte: per ciascun Bacino+Provincia unisce i plichi da 1/2 pezzi.
  // Serve a ridurre micro-plichi residui non adiacenti ma della stessa area provinciale.
  plichi.sort((a, b) => {
    const ka = [
      bacinoSortKey(a.bacino || UNKNOWN_BACINO),
      a.provincia || '',
      a.cap || '',
      a.citta || '',
    ];
    const kb = [
      bacinoSortKey(b.bacino || UNKNOWN_BACINO),
      b.provincia || '',
      b.cap || '',
      b.citta || '',
    ];
    for (let n = 0; n < ka.length; n++) {
      const cmp = compareSortValue(ka[n], kb[n]);
      if (cmp !== 0) return cmp;
    }
    return 0;
  });

  const mergedByProvince = [];
  let z = 0;
  while (z < plichi.length) {
    const curr = plichi[z];
    const zoneBacino = curr?.bacino || UNKNOWN_BACINO;
    const zoneProvincia = curr?.provincia || '';

    let w = z;
    const zoneItems = [];
    while (w < plichi.length) {
      const pw = plichi[w];
      if (!pw) break;
      if ((pw.bacino || UNKNOWN_BACINO) !== zoneBacino) break;
      if ((pw.provincia || '') !== zoneProvincia) break;
      zoneItems.push(pw);
      w += 1;
    }

    const tiny = zoneItems.filter((p) => p.items.length <= 2);
    const normal = zoneItems.filter((p) => p.items.length > 2);

    if (tiny.length >= 2) {
      const items = tiny
        .flatMap((p) => p.items)
        .sort((a, b) => Number(a._seq) - Number(b._seq));
      const first = items[0];
      const last = items[items.length - 1];
      const bacino = buildBacinoLabel(items);
      const aggregated = {
        plico: 0,
        groupType: 'provincia',
        items,
        cap: first?.cap || '',
        provincia: zoneProvincia,
        citta: '',
        bacino,
        destinazione: buildProvinciaBacinoDestinazione(zoneProvincia, bacino),
        numeroInvii: items.length,
        pesoPlico:
          pesoInvio && Number(pesoInvio) > 0
            ? Number(pesoInvio) * items.length
            : '',
        altezzaPlico:
          spessoreInvio && Number(spessoreInvio) > 0
            ? Number(spessoreInvio) * items.length
            : '',
        start: first?._seq || '',
        end: last?._seq || '',
        barcode: buildBarcode(codiceLavorazione, 'MIX', first?._seq || '', last?._seq || ''),
      };
      mergedByProvince.push(...normal, aggregated);
    } else {
      mergedByProvince.push(...zoneItems);
    }

    z = w;
  }
  plichi = mergedByProvince;

  // Pass finale obbligatorio: nessun plico deve contenere province miste.
  // Se accade (es. residue da livello bacino/mix), splitta per provincia.
  const splitMultiProvincePlichi = () => {
    let changed = false;
    const out = [];

    plichi.forEach((p) => {
      if (!p || !Array.isArray(p.items) || p.items.length === 0) {
        out.push(p);
        return;
      }

      const byProv = new Map();
      p.items.forEach((rec) => {
        const prov = rec.provincia || '';
        if (!byProv.has(prov)) byProv.set(prov, []);
        byProv.get(prov).push(rec);
      });

      if (byProv.size <= 1) {
        out.push(p);
        return;
      }

      changed = true;
      byProv.forEach((items, prov) => {
        items.sort((a, b) => Number(a._seq) - Number(b._seq));
        const first = items[0];
        const last = items[items.length - 1];
        const bacino = buildBacinoLabel(items);

        out.push({
          plico: 0,
          groupType: 'provincia',
          items,
          cap: first.cap,
          provincia: prov,
          citta: '',
          bacino,
          destinazione: buildProvinciaBacinoDestinazione(prov, bacino),
          numeroInvii: items.length,
          pesoPlico:
            pesoInvio && Number(pesoInvio) > 0
              ? Number(pesoInvio) * items.length
              : '',
          altezzaPlico:
            spessoreInvio && Number(spessoreInvio) > 0
              ? Number(spessoreInvio) * items.length
              : '',
          start: first._seq,
          end: last._seq,
          barcode: buildBarcode(codiceLavorazione, first.cap, first._seq, last._seq),
        });
      });
    });

    if (changed) plichi = out;
  };

  splitMultiProvincePlichi();
  // Terza fase: per plichi bacino con destinazione uguale al bacino,
  // tenta prima raggruppamento per comune, poi per provincia (senza CAP in destinazione).
  const makePlicoFromItems = (items, groupType, bacinoOverride = '') => {
    if (!items || items.length === 0) return null;
    items.sort((a, b) => Number(a._seq) - Number(b._seq));
    const first = items[0];
    const last = items[items.length - 1];
    const bacino = bacinoOverride || buildBacinoLabel(items);
    const provincia = first.provincia || '';
    const citta = first.citta || '';
    return {
      plico: 0,
      groupType,
      items,
      cap: first.cap,
      provincia,
      citta,
      bacino,
      destinazione: resolveDestinazioneForGroup(groupType, first.cap, citta, provincia, bacino),
      numeroInvii: items.length,
      pesoPlico:
        pesoInvio && Number(pesoInvio) > 0
          ? Number(pesoInvio) * items.length
          : '',
      altezzaPlico:
        spessoreInvio && Number(spessoreInvio) > 0
          ? Number(spessoreInvio) * items.length
          : '',
      start: first._seq,
      end: last._seq,
      barcode: buildBarcode(codiceLavorazione, first.cap, first._seq, last._seq),
    };
  };

  const isBacinoDest = (p) =>
    p &&
    normalizeText(p.groupType || '') === 'BACINO' &&
    normalizeText(p.destinazione || '') === normalizeText(p.bacino || '');

  const keep = [];
  const candidatesByBacino = new Map();
  plichi.forEach((p) => {
    if (!isBacinoDest(p)) {
      keep.push(p);
      return;
    }
    const b = normalizeText(p.bacino || UNKNOWN_BACINO) || UNKNOWN_BACINO;
    if (!candidatesByBacino.has(b)) candidatesByBacino.set(b, []);
    candidatesByBacino.get(b).push(...p.items);
  });

  const rebuilt = [];
  candidatesByBacino.forEach((items, bacino) => {
    const byComune = new Map();
    const poolProvincia = [];

    items.forEach((rec) => {
      const comune = normalizeText(rec.citta || '');
      if (comune) {
        if (!byComune.has(comune)) byComune.set(comune, []);
        byComune.get(comune).push(rec);
      } else {
        poolProvincia.push(rec);
      }
    });

    Array.from(byComune.keys())
      .sort((a, b) => String(a).localeCompare(String(b)))
      .forEach((comuneKey) => {
        const arr = byComune.get(comuneKey) || [];
        arr.sort((a, b) => Number(a._seq) - Number(b._seq));
        const { chunks, leftovers } = splitGroup(arr, maxPerPlico, MIN_PIECES, false);
        chunks.forEach((chunk) => {
          const plico = makePlicoFromItems(chunk, 'comune', bacino);
          if (plico) rebuilt.push(plico);
        });
        poolProvincia.push(...leftovers);
      });

    const byProv = new Map();
    poolProvincia.forEach((rec) => {
      const prov = normalizeText(rec.provincia || '');
      if (!prov) return;
      if (!byProv.has(prov)) byProv.set(prov, []);
      byProv.get(prov).push(rec);
    });

    const leftoverBacino = [];
    Array.from(byProv.keys())
      .sort((a, b) => String(a).localeCompare(String(b)))
      .forEach((provKey) => {
        const arr = byProv.get(provKey) || [];
        arr.sort((a, b) => Number(a._seq) - Number(b._seq));
        const { chunks, leftovers } = splitGroup(arr, maxPerPlico, MIN_PIECES, false);
        chunks.forEach((chunk) => {
          const plico = makePlicoFromItems(chunk, 'provincia', bacino);
          if (plico) rebuilt.push(plico);
        });
        leftoverBacino.push(...leftovers);
      });

    if (leftoverBacino.length) {
      const plico = makePlicoFromItems(leftoverBacino, 'bacino', bacino);
      if (plico) rebuilt.push(plico);
    }
  });

  plichi = [...keep, ...rebuilt];

  // Fase aggiuntiva: se la destinazione coincide con il bacino,
  // prova sempre il raggruppamento per provincia prima di lasciare BACINO puro.
  const promoteBacinoDestToProvincia = () => {
    const out = [];

    plichi.forEach((p) => {
      const bacino = p?.bacino || UNKNOWN_BACINO;
      const isPureBacinoDest =
        p &&
        normalizeText(p.destinazione || '') === normalizeText(bacino || '') &&
        normalizeText(bacino || '') !== normalizeText(UNKNOWN_BACINO);

      if (!isPureBacinoDest || !Array.isArray(p.items) || p.items.length === 0) {
        out.push(p);
        return;
      }

      const byProv = new Map();
      const noProv = [];

      p.items.forEach((rec) => {
        const prov = normalizeText(rec.provincia || '');
        if (!prov) {
          noProv.push(rec);
          return;
        }
        if (!byProv.has(prov)) byProv.set(prov, []);
        byProv.get(prov).push(rec);
      });

      if (byProv.size === 0) {
        out.push(p);
        return;
      }

      Array.from(byProv.keys())
        .sort((a, b) => compareSortValue(a, b))
        .forEach((prov) => {
          const arr = byProv.get(prov) || [];
          arr.sort((a, b) => Number(a._seq) - Number(b._seq));
          splitByMaxSize(arr, maxPerPlico).forEach((chunk) => {
            const first = chunk[0];
            const last = chunk[chunk.length - 1];
            out.push({
              plico: 0,
              groupType: 'provincia',
              items: chunk,
              cap: first.cap,
              provincia: prov,
              citta: '',
              bacino,
              destinazione: buildProvinciaBacinoDestinazione(prov, bacino),
              numeroInvii: chunk.length,
              pesoPlico:
                pesoInvio && Number(pesoInvio) > 0
                  ? Number(pesoInvio) * chunk.length
                  : '',
              altezzaPlico:
                spessoreInvio && Number(spessoreInvio) > 0
                  ? Number(spessoreInvio) * chunk.length
                  : '',
              start: first._seq,
              end: last._seq,
              barcode: buildBarcode(codiceLavorazione, first.cap, first._seq, last._seq),
            });
          });
        });

      if (noProv.length) {
        const fallback = makePlicoFromItems(noProv, 'bacino', bacino);
        if (fallback) out.push(fallback);
      }
    });

    plichi = out;
  };

  promoteBacinoDestToProvincia();

  const compareFinalKeys = (a, b) => {
    const ka = [
      bacinoSortKey(a.bacino || UNKNOWN_BACINO),
      normalizeText(a.destinazione || ''),
      a.cap || '',
      a.citta || '',
    ];
    const kb = [
      bacinoSortKey(b.bacino || UNKNOWN_BACINO),
      normalizeText(b.destinazione || ''),
      b.cap || '',
      b.citta || '',
    ];
    for (let n = 0; n < ka.length; n++) {
      const cmp = compareSortValue(ka[n], kb[n]);
      if (cmp !== 0) return cmp;
    }
    return 0;
  };

  // Primo allineamento destinazione per record (serve al sort finale records)
  plichi.forEach((p) => {
    p.items.forEach((rec) => {
      rec._destinazione = p.destinazione || '';
      rec._bacino = p.bacino || rec._bacino || UNKNOWN_BACINO;
    });
  });

  // Riordino finale richiesto per dati riordinati: bacino, destinazione, cap, citta
  records.sort((a, b) => {
    const ka = [
      bacinoSortKey(a._bacino || UNKNOWN_BACINO),
      normalizeText(a._destinazione || ''),
      a.cap || '',
      a.citta || '',
    ];
    const kb = [
      bacinoSortKey(b._bacino || UNKNOWN_BACINO),
      normalizeText(b._destinazione || ''),
      b.cap || '',
      b.citta || '',
    ];
    for (let n = 0; n < ka.length; n++) {
      const cmp = compareSortValue(ka[n], kb[n]);
      if (cmp !== 0) return cmp;
    }
    return 0;
  });
  records.forEach((r, idx) => (r._seq = idx + 1));

  // Ricalcolo metadati plico sui nuovi progressivi
  plichi.forEach((p) => {
    p.items.sort((x, y) => Number(x._seq) - Number(y._seq));
    refreshPlicoMeta(p);
  });

  // Riordino finale etichettatura omogeneo al foglio dati
  plichi.sort(compareFinalKeys);

  // Rinumeriamo i plichi secondo l'ordinamento finale e riallineiamo i flag
  plichi.forEach((p, idx) => {
    const newNum = idx + 1;
    p.plico = newNum;
    p.items.forEach((rec, i) => {
      rec._plico = newNum;
      rec._fine = i === p.items.length - 1 ? END_ARROW : '';
      rec._destinazione = p.destinazione || '';
      rec._bacino = p.bacino || rec._bacino || UNKNOWN_BACINO;
    });
  });

  // Allineamento finale: ordina i record con i metadati plico definitivi.
  records.sort((a, b) => {
    const ka = [
      bacinoSortKey(a._bacino || UNKNOWN_BACINO),
      normalizeText(a._destinazione || ''),
      a.cap || '',
      a.citta || '',
    ];
    const kb = [
      bacinoSortKey(b._bacino || UNKNOWN_BACINO),
      normalizeText(b._destinazione || ''),
      b.cap || '',
      b.citta || '',
    ];
    for (let n = 0; n < ka.length; n++) {
      const cmp = compareSortValue(ka[n], kb[n]);
      if (cmp !== 0) return cmp;
    }
    return 0;
  });
  records.forEach((r, idx) => (r._seq = idx + 1));

  // Calcolo etichette come ultimo step: ricostruisce i plichi dai dati gia' riordinati.
  const inferGroupTypeFromDest = (dest = '', bacino = '') => {
    const d = normalizeText(dest || '');
    const b = normalizeText(bacino || UNKNOWN_BACINO);
    if (!d) return 'mix';
    if (d === b) return 'bacino';
    if (/ PROVINCIA$/.test(d)) return 'provincia';
    if (/^\d{5} [A-Z]/.test(d)) return 'cap+citta';
    if (/^\d{5}$/.test(d)) return 'cap';
    return 'comune';
  };

  const buildFinalPlichi = () => {
    const out = [];
    let idx = 0;

    while (idx < records.length) {
      const first = records[idx];
      const bacino = first._bacino || UNKNOWN_BACINO;
      const destinazione = first._destinazione || '';
      const groupType = inferGroupTypeFromDest(destinazione, bacino);

      const groupItems = [];
      while (idx < records.length) {
        const rec = records[idx];
        const sameBacino =
          normalizeText(rec._bacino || UNKNOWN_BACINO) === normalizeText(bacino);
        const sameDest =
          normalizeText(rec._destinazione || '') === normalizeText(destinazione);
        if (!sameBacino || !sameDest) break;
        groupItems.push(rec);
        idx += 1;
      }

      splitForFinalPlichi(groupItems, maxPerPlico, MIN_PIECES).forEach((chunk) => {
        if (!chunk.length) return;
        chunk.sort((a, b) => Number(a._seq) - Number(b._seq));
        const firstItem = chunk[0];
        const lastItem = chunk[chunk.length - 1];
        out.push({
          plico: 0,
          groupType,
          items: chunk,
          cap: firstItem.cap,
          provincia: firstItem.provincia,
          citta: firstItem.citta,
          bacino,
          destinazione,
          numeroInvii: chunk.length,
          pesoPlico:
            pesoInvio && Number(pesoInvio) > 0
              ? Number(pesoInvio) * chunk.length
              : '',
          altezzaPlico:
            spessoreInvio && Number(spessoreInvio) > 0
              ? Number(spessoreInvio) * chunk.length
              : '',
          start: firstItem._seq,
          end: lastItem._seq,
          barcode: buildBarcode(codiceLavorazione, firstItem.cap, firstItem._seq, lastItem._seq),
        });
      });
    }

    return out;
  };

  plichi = buildFinalPlichi();

  plichi.forEach((p, idx) => {
    const newNum = idx + 1;
    p.plico = newNum;
    p.items.sort((x, y) => Number(x._seq) - Number(y._seq));
    p.items.forEach((rec, i) => {
      rec._plico = newNum;
      rec._fine = i === p.items.length - 1 ? END_ARROW : '';
      rec._destinazione = p.destinazione || rec._destinazione || '';
      rec._bacino = p.bacino || rec._bacino || UNKNOWN_BACINO;
    });
    refreshPlicoMeta(p);
  });

  const etichetteRows = plichi.map((p) => [
    centroImpostazione || '',
    p.bacino || '',
    p.destinazione,
    p.plico,
    p.numeroInvii,
    p.pesoPlico,
    p.altezzaPlico,
    p.start,
    p.end,
    nomeLavorazione || '',
    codiceLavorazione || '',
    p.barcode,
  ]);

  etichetteRows.forEach((row) => sheet2.addRow(row));

  // Foglio per stampa A4:
  // ordine intercalato in 2 o 3 colonne per ottenere mazzetti progressivi dopo il taglio.
  const sheet3 = sheet2.workbook.addWorksheet('X_Stampa_Etichette');
  sheet3.addRow([
    'Centro impostazione',
    'Bacino',
    'Destinazione',
    'Progressivo plico',
    'Numero invii',
    'Peso plico (g)',
    'Altezza plico (mm)',
    'Progressivo iniziale',
    'Progressivo finale',
    'Nome lavorazione',
    'Codice lavorazione',
    'Barcode',
  ]);
  createA4PrintOrder(etichetteRows, plichiPerFoglio).forEach((row) => sheet3.addRow(row));

  // Rebuild sheet 1 in desired order + plico info AFTER renumber
  let output = sheet2.workbook;
  const sheet1 = output.addWorksheet('Dati Riordinati', { position: 1 });
  sheet1.addRow([...COLUMN_ORDER, 'Bacino', 'Destinazione', 'Progressivo', 'Plico', 'Fine Plico']);
  records.forEach((rec) => {
    sheet1.addRow([
      ...COLUMN_ORDER.map((k) => rec[k] ?? ''),
      rec._bacino || UNKNOWN_BACINO,
      rec._destinazione || '',
      rec._seq || '',
      rec._plico || '',
      rec._fine || '',
    ]);
  });

  const outPath = makeOutputPath(filePath);
  await output.xlsx.writeFile(outPath);
  return outPath;
});

function makeOutputPath(inputPath) {
  const dir = path.dirname(inputPath);
  const base = path.basename(inputPath, path.extname(inputPath));
  return path.join(dir, `${base}-plico.xlsx`);
}

function findBacino(cap) {
  if (!cap) return '';
  const capNum = Number(cap);
  const match = bacinoRanges.find(
    (r) => capNum >= Number(r.start) && capNum <= Number(r.end)
  );
  return match ? match.bacino : '';
}

function buildBacinoLabel(items = []) {
  const uniq = Array.from(
    new Set(
      items
        .map((r) => normalizeText(findBacino(r.cap)))
        .filter((x) => x)
    )
  );
  if (uniq.length === 0) return UNKNOWN_BACINO;
  if (uniq.length > 1) return 'MIX BACINI';
  return uniq[0];
}

function buildDestinazione(groupType, cap = '', citta = '', provincia = '', bacino = '') {
  const capNorm = normalizeCap(cap);
  const cittaNorm = normalizeText(citta);
  const provNorm = normalizeText(provincia);

  switch (groupType) {
    case 'cap+citta':
      return `${capNorm} ${cittaNorm}`.trim();
    case 'cap':
      return `${capNorm}`.trim();
    case 'provincia':
      return `${provNorm} PROVINCIA`.trim();
    case 'provincia+citta_capoluogo':
      return `${provNorm} CITPROV`.trim();
    case 'bacino':
      return `${normalizeText(bacino)}`.trim();
    case 'mix':
      return normalizeText(bacino) || 'MIX BACINI';
    default:
      return `${capNorm} ${cittaNorm || provNorm}`.trim();
  }
}

function buildBarcode(codice, cap, start, end) {
  const code = codice ? String(codice) : 'LAV';
  return `PL-${code}-${cap || 'MIX'}-${String(start).padStart(6, '0')}-${String(
    end
  ).padStart(6, '0')}`;
}

function readCellValue(value) {
  if (value == null) return '';
  if (typeof value === 'object') {
    if (Array.isArray(value.richText)) {
      return value.richText.map((x) => x.text || '').join('');
    }
    if (value.text != null) return String(value.text);
    if (value.result != null) return String(value.result);
  }
  return String(value);
}

function normalizeText(value) {
  return readCellValue(value)
    .replace(/[\r\n\t]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toUpperCase();
}

function normalizeCap(value) {
  const raw = readCellValue(value).replace(/\D/g, '');
  if (!raw) return '';
  return raw.slice(-5).padStart(5, '0');
}
function resolveDestinazioneForGroup(groupType, cap = '', citta = '', provincia = '', bacino = '') {
  const type = normalizeText(groupType || '');
  if (type === 'PROVINCIA') {
    return buildProvinciaBacinoDestinazione(provincia, bacino);
  }
  return buildDestinazione(groupType, cap, citta, provincia, bacino);
}

function buildProvinciaBacinoDestinazione(provincia = '', bacino = '') {
  const prov = normalizeText(provincia || '').trim();
  if (prov) return `${prov} PROVINCIA`;
  const bac = normalizeText(bacino || '').trim();
  if (bac) return `${bac}`;
  return UNKNOWN_BACINO;
}
function bacinoSortKey(bacino = '') {
  const val = normalizeText(bacino || UNKNOWN_BACINO);
  if (val === UNKNOWN_BACINO) return val;
  return val;
}
function compareSortValue(a, b) {
  const av = normalizeForSort(a);
  const bv = normalizeForSort(b);
  return av.localeCompare(bv, 'it', { numeric: true, sensitivity: 'base' });
}
function normalizeForSort(value) {
  return normalizeText(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
}

function normalize(str = '') {
  return String(str).toLowerCase().replace(/\s+/g, '').trim();
}

function splitByConsecutiveSeq(arr) {
  if (!arr.length) return [];
  const runs = [];
  let current = [arr[0]];
  for (let i = 1; i < arr.length; i++) {
    const prev = Number(arr[i - 1]._seq || 0);
    const curr = Number(arr[i]._seq || 0);
    if (curr === prev + 1) {
      current.push(arr[i]);
    } else {
      runs.push(current);
      current = [arr[i]];
    }
  }
  runs.push(current);
  return runs;
}
function splitByMaxSize(arr, maxSize) {
  if (!Array.isArray(arr) || arr.length === 0) return [];
  const size = Math.max(1, Number(maxSize) || 1);
  const out = [];
  for (let i = 0; i < arr.length; i += size) {
    out.push(arr.slice(i, i + size));
  }
  return out;
}

function splitForFinalPlichi(arr, maxSize, minSize) {
  if (!Array.isArray(arr) || arr.length === 0) return [];
  const size = Math.max(1, Number(maxSize) || 1);
  const min = Math.max(1, Number(minSize) || 1);

  const chunks = splitByMaxSize(arr, size);
  if (!chunks.length) return [];

  // Evita ultimo frammento troppo piccolo spostando elementi dal chunk precedente.
  if (chunks.length > 1) {
    let last = chunks[chunks.length - 1];
    while (last.length < min) {
      const prev = chunks[chunks.length - 2];
      if (!prev || prev.length <= min) break;
      last.unshift(prev.pop());
    }

    // Se resta micro-chunk (1/2), accorpa comunque al precedente.
    if (last.length <= 2 && chunks.length > 1) {
      const prev = chunks[chunks.length - 2];
      prev.push(...last);
      chunks.pop();
    }
  }

  return chunks.filter((c) => c.length > 0);
}


function createA4PrintOrder(rows, plichiPerFoglio = 3) {
  const total = Array.isArray(rows) ? rows.length : 0;
  if (!total) return [];

  const slots = [2, 3].includes(Number(plichiPerFoglio))
    ? Number(plichiPerFoglio)
    : 3;
  const block = Math.ceil(total / slots);
  const ordered = [];

  for (let i = 0; i < block; i++) {
    for (let section = 0; section < slots; section++) {
      const idx = i + section * block;
      if (idx < total) ordered.push(rows[idx]);
    }
  }

  return ordered;
}

function splitGroup(arr, maxSize, minSize, acceptSmall) {
  if (arr.length < minSize && !acceptSmall) return { chunks: [], leftovers: arr };
  const chunks = [];
  let i = 0;
  while (i < arr.length) {
    const end = Math.min(i + maxSize, arr.length);
    let chunk = arr.slice(i, end);
    if (chunk.length < minSize) {
      if (chunks.length === 0) {
        if (!acceptSmall) return { chunks: [], leftovers: arr };
        chunks.push(chunk);
      } else {
        chunks[chunks.length - 1].push(...chunk);
      }
      break;
    }
    chunks.push(chunk);
    i = end;
  }
  return { chunks, leftovers: [] };
}

