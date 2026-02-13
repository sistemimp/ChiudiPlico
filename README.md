# chiudi-plico

App Electron per generare plichi ed etichette da file XLSX.

## Requisiti

- Node.js 18+ (consigliato LTS)
- npm

## Installazione

```bash
npm install
```

## Avvio

```bash
npm start
```

## Funzionamento completo (passo passo)

1. L'utente avvia l'app e seleziona un file `.xlsx` con il pulsante `Scegli file`.
2. L'utente compila i campi obbligatori:
   - centro di impostazione
   - nome lavorazione
   - codice lavorazione
   - peso invio (grammi)
   - spessore invio (millimetri)
3. Al click su `Elabora`, il renderer valida che tutti i campi siano compilati; se manca qualcosa, blocca il processo.
4. Il processo main legge il primo foglio Excel e mappa le colonne attese:
   `id`, `nome`, `indirizzo`, `cap`, `citta`, `provincia`, `spare1..spare5`.
5. I dati vengono normalizzati (maiuscolo, spazi puliti, CAP a 5 cifre) e vengono scartate le righe vuote.
6. Per ogni record viene calcolato il `Bacino` usando `bacini.json` (se non trovato: `MIX BACINI`).
7. I record vengono ordinati e viene assegnato un progressivo interno.
8. Il sistema calcola la capienza massima plico in base a:
   - limite peso: `10000 / pesoInvio` (10 kg)
   - limite spessore: `120 / spessoreInvio` (12 cm)
   - valore effettivo: minimo tra i due.
9. Parte la costruzione dei plichi con livelli progressivi di raggruppamento:
   - `CAP + CITTA`
   - `CAP`
   - `PROVINCIA`
   - `PROVINCIA + CITTA capoluogo`
   - `BACINO`
   - fallback `MIX`
10. Il motore applica più passaggi di consolidamento per ridurre micro-plichi (1-2 invii), mantenere sequenze coerenti e migliorare la destinazione finale (preferendo provincia/comune quando possibile).
11. Per ogni plico vengono calcolati:
   - progressivo plico
   - numero invii
   - peso e altezza plico
   - progressivo iniziale/finale
   - barcode (`PL-<codice>-<cap>-<start>-<end>`)
   - destinazione finale
12. Viene generato il file output nella stessa cartella dell'input, con nome:
   `<nomefile>-plico.xlsx`.
13. L'output contiene 3 fogli:
   - `Dati Riordinati`: dati originali riordinati con colonne aggiuntive (`Bacino`, `Destinazione`, `Progressivo`, `Plico`, `Fine Plico`).
   - `Etichettatura Plichi`: una riga per plico con tutti i dati di etichetta.
   - `X_Stampa_Etichette`: stesso contenuto etichette in ordine 3-up per stampa/taglio.
14. L'app mostra a video il percorso completo del file generato.

## Struttura principale

- `main.js`: processo principale Electron
- `preload.js`: bridge sicuro tra main e renderer
- `renderer.js`: logica UI lato renderer
- `index.html`: interfaccia applicativa
- `run-electron.js`: bootstrap avvio Electron
- `bacini.json`: dati di supporto

## Dipendenze principali

- `electron`
- `exceljs`
