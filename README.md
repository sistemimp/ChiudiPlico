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
