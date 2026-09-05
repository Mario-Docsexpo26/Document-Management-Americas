#!/usr/bin/env node
/**
 * build-dashboard.js
 *
 * Reads data/bitacora.xlsx, replicates EXACTLY the same parsing rules the
 * dashboard's own "Actualizar datos (subir Excel)" button uses in-browser
 * (parseWorkbookToData + normalizeVesselNames from index.html), and writes
 * the resulting RAW_DATA array back into index.html in place.
 *
 * Run by the GitHub Action on every push that touches data/*.xlsx — see
 * .github/workflows/update-dashboard.yml. Can also be run locally:
 *   npm install xlsx@0.18.5
 *   node scripts/build-dashboard.js
 */
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

const REPO_ROOT = path.resolve(__dirname, '..');
const EXCEL_PATH = path.join(REPO_ROOT, 'data', 'bitacora.xlsx');
const HTML_PATH = path.join(REPO_ROOT, 'index.html');

// ---------------- same constants as index.html ----------------
const COL = {
  key: 0, wk: 1, fechaCarga: 2, wkArribo: 3, mercado: 4, pod: 5, origen: 6, exportador: 7,
  claseExp: 8, especie: 9, manejo: 10, po: 11, contenedor: 12, consignee: 13, cliente: 14,
  line: 15, nave: 16, etdReal: 17, etaReal: 18, ol: 19, etaDoc: 20, etdDoc: 21, ataDoc: 22,
  obtencion: 23, anticipacion: 24, isfStatus: 34, fullDocument: 35, claseExportador: 38,
  statusDocumental: 39,
};
const VIA_LABELS = { SEA: 'Marítimo', AIR: 'Aéreo', TRK: 'Terrestre' };

function excelSerialToDate(v) {
  if (v instanceof Date && !isNaN(v)) return v;
  if (typeof v === 'number') return new Date(Math.round((v - 25569) * 86400 * 1000));
  return null;
}
function toISO(d) {
  return (d instanceof Date && !isNaN(d)) ? d.toISOString().slice(0, 10) : null;
}

function isMaritime(d) { return !d.claseExp || d.claseExp === 'SEA'; }

// ---------------- parseWorkbookToData (mirrors index.html) ----------------
function parseWorkbookToData(wb) {
  const sheetName =
    wb.SheetNames.find(n => n.trim() === 'Bitacora') ||
    wb.SheetNames.find(n => n.toLowerCase().replace(/[^a-z]/g, '').includes('bitacora'));
  if (!sheetName) throw new Error('No se encontró la pestaña "Bitacora" en el archivo.');

  const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, raw: true, defval: null });
  const headerRow = rows[0] || [];

  function findHeaderCol(patterns) {
    for (let i = 0; i < headerRow.length; i++) {
      const h = String(headerRow[i] || '')
        .toLowerCase()
        .normalize('NFD')
        .replace(/[̀-ͯ]/g, '')
        .replace(/[^a-z0-9]/g, '');
      if (patterns.some(p => h.includes(p))) return i;
    }
    return -1;
  }

  // Same explicit decision as in the dashboard: "Status Container" (legacy) is never
  // used, only "Container Status" is read. See business-rules.md #2 in the docexpo-specialist
  // skill for the full reasoning if this ever needs revisiting.
  const COL_CONTAINER_STATUS = findHeaderCol(['containerstatus']);
  const COL_RELEASE_DATE = findHeaderCol(['releasedate', 'fecharelease', 'fechaliberacion', 'fechadeliberacion']);
  const COL_COMMENTS = findHeaderCol(['commentsstatuscontainer', 'comentariosstatuscontainer', 'commentscontainerstatus']);

  const body = rows.slice(1).filter(r => r && r[COL.key]);
  const out = [];

  body.forEach(r => {
    const etdReal = excelSerialToDate(r[COL.etdReal]);
    const etaReal = excelSerialToDate(r[COL.etaReal]);
    const etaDoc = excelSerialToDate(r[COL.etaDoc]);
    const etdDoc = excelSerialToDate(r[COL.etdDoc]);
    const ataDoc = excelSerialToDate(r[COL.ataDoc]);
    const obtRaw = r[COL.obtencion], antRaw = r[COL.anticipacion];
    const statusContainerRaw = COL_CONTAINER_STATUS >= 0 ? r[COL_CONTAINER_STATUS] : null;
    const releaseDateRaw = COL_RELEASE_DATE >= 0 ? excelSerialToDate(r[COL_RELEASE_DATE]) : null;
    const commentsRaw = COL_COMMENTS >= 0 ? r[COL_COMMENTS] : null;
    const claseExpRaw = r[COL.claseExp] ? String(r[COL.claseExp]).trim().toUpperCase() : null;

    out.push({
      key: r[COL.key], wk: r[COL.wk], wkArribo: r[COL.wkArribo], mercado: r[COL.mercado], pod: r[COL.pod],
      origen: r[COL.origen], exportador: r[COL.exportador],
      claseExp: claseExpRaw, via: VIA_LABELS[claseExpRaw] || (claseExpRaw || 'Sin vía'),
      claseExportador: r[COL.claseExportador] ? String(r[COL.claseExportador]).trim() : null,
      especie: r[COL.especie], manejo: r[COL.manejo], po: r[COL.po], contenedor: r[COL.contenedor],
      consignee: (r[COL.consignee] == null || r[COL.consignee] === '') ? 'Sin consignee' : r[COL.consignee],
      cliente: (r[COL.cliente] == null || r[COL.cliente] === '') ? 'Sin cliente' : r[COL.cliente],
      linea: r[COL.line], nave: r[COL.nave] ? String(r[COL.nave]).trim() : null,
      etdReal: toISO(etdReal), etaReal: toISO(etaReal), etaDoc: toISO(etaDoc), etdDoc: toISO(etdDoc), ataDoc: toISO(ataDoc),
      obtencion: typeof obtRaw === 'number' ? obtRaw : null,
      anticipacion: typeof antRaw === 'number' ? antRaw : null,
      fullDocument: r[COL.fullDocument],
      statusDocumental: r[COL.statusDocumental] ? String(r[COL.statusDocumental]).trim() : '',
      isfStatus: (r[COL.isfStatus] == null || r[COL.isfStatus] === '') ? null : String(r[COL.isfStatus]).trim(),
      statusContainer: (statusContainerRaw == null || statusContainerRaw === '') ? null : String(statusContainerRaw).trim(),
      releaseDate: toISO(releaseDateRaw),
      comments: (commentsRaw == null || String(commentsRaw).trim() === '') ? null : String(commentsRaw).trim(),
      onTime: (etaDoc instanceof Date && etaReal instanceof Date && !isNaN(etaDoc) && !isNaN(etaReal)) ? (etaDoc < etaReal) : false,
    });
  });

  return out;
}

// ---------------- normalizeVesselNames (mirrors index.html) ----------------
function levenshtein(a, b) {
  const m = a.length, n = b.length;
  if (m === 0) return n;
  if (n === 0) return m;
  const dp = new Array(n + 1);
  for (let j = 0; j <= n; j++) dp[j] = j;
  for (let i = 1; i <= m; i++) {
    let prev = dp[0];
    dp[0] = i;
    for (let j = 1; j <= n; j++) {
      const tmp = dp[j];
      dp[j] = a[i - 1] === b[j - 1] ? prev : 1 + Math.min(prev, dp[j], dp[j - 1]);
      prev = tmp;
    }
  }
  return dp[n];
}

function normalizeVesselNames(data) {
  function coreTokens(nave) {
    const upper = String(nave).toUpperCase().trim().replace(/\s+/g, ' ');
    return upper.split(' ').filter(tok => tok && !/\d/.test(tok));
  }
  function tokensMatch(tokA, tokB) {
    const setA = new Set(tokA), setB = new Set(tokB);
    const isSubset = (small, big) => [...small].every(t => big.has(t));
    if (setA.size && setB.size && (isSubset(setA, setB) || isSubset(setB, setA))) return true;
    for (const a of tokA) {
      for (const b of tokB) {
        const dist = levenshtein(a, b);
        if (dist <= 1) return true;
        if (dist <= 2 && a.length >= 6 && b.length >= 6) return true;
      }
    }
    return false;
  }

  const groups = new Map();
  data.forEach(d => {
    if (!d.nave) return;
    const key = d.etaReal;
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(d);
  });

  groups.forEach(records => {
    const rawCounts = new Map();
    records.forEach(d => rawCounts.set(d.nave, (rawCounts.get(d.nave) || 0) + 1));
    const rawValues = [...rawCounts.keys()];
    const cores = rawValues.map(v => coreTokens(v));
    const tokenCounts = cores.map(c => c.length);

    const parent = rawValues.map((_, i) => i);
    function find(i) { while (parent[i] !== i) { parent[i] = parent[parent[i]]; i = parent[i]; } return i; }
    function union(i, j) { const ri = find(i), rj = find(j); if (ri !== rj) parent[ri] = rj; }
    for (let i = 0; i < rawValues.length; i++) {
      for (let j = i + 1; j < rawValues.length; j++) {
        if (tokensMatch(cores[i], cores[j])) union(i, j);
      }
    }

    const clusterMap = new Map();
    rawValues.forEach((_, i) => {
      const root = find(i);
      if (!clusterMap.has(root)) clusterMap.set(root, []);
      clusterMap.get(root).push(i);
    });

    clusterMap.forEach(indices => {
      if (indices.length < 2) return;
      const clusterTotal = indices.reduce((sum, idx) => sum + rawCounts.get(rawValues[idx]), 0);

      let canonicalIdx = indices[0];
      indices.forEach(idx => {
        const curTok = tokenCounts[idx], bestTok = tokenCounts[canonicalIdx];
        const curCount = rawCounts.get(rawValues[idx]), bestCount = rawCounts.get(rawValues[canonicalIdx]);
        if (curTok > bestTok) {
          canonicalIdx = idx;
        } else if (curTok === bestTok) {
          if (curCount > bestCount || (curCount === bestCount && rawValues[idx].length > rawValues[canonicalIdx].length)) {
            canonicalIdx = idx;
          }
        }
      });

      const winnerTok = tokenCounts[canonicalIdx];
      const winnerCount = rawCounts.get(rawValues[canonicalIdx]);
      const strictlyMoreTokensThanAllOthers = indices.every(idx => idx === canonicalIdx || tokenCounts[idx] < winnerTok);
      const clearMajority = clusterTotal > 0 && (winnerCount / clusterTotal) >= 0.7;
      if (!strictlyMoreTokensThanAllOthers && !clearMajority) return;

      const canonical = rawValues[canonicalIdx];
      const rawStringsInCluster = new Set(indices.map(idx => rawValues[idx]));
      records.forEach(d => {
        if (rawStringsInCluster.has(d.nave) && d.nave !== canonical) {
          d.naveOriginal = d.nave;
          d.nave = canonical;
        }
      });
    });
  });
}

// ---------------- main ----------------
function main() {
  if (!fs.existsSync(EXCEL_PATH)) {
    console.error(`No se encontró ${EXCEL_PATH}. Sube la Bitácora como data/bitacora.xlsx antes de correr esto.`);
    process.exit(1);
  }
  const buf = fs.readFileSync(EXCEL_PATH);
  const wb = XLSX.read(buf, { type: 'buffer', cellDates: true });

  const data = parseWorkbookToData(wb);
  if (!data.length) {
    console.error('El Excel no arrojó ningún registro (revisa que la pestaña se llame "Bitacora" y tenga datos con Key en la primera columna).');
    process.exit(1);
  }
  normalizeVesselNames(data);

  const html = fs.readFileSync(HTML_PATH, 'utf8');
  const startMarker = 'let RAW_DATA = [';
  const endAnchor = '\n\n\n/* ---------------- ICONS';

  const startIdx = html.indexOf(startMarker);
  if (startIdx === -1) throw new Error('No se encontró "let RAW_DATA = [" en index.html — revisa que el archivo no se haya editado a mano de forma que rompa este marcador.');
  const anchorIdx = html.indexOf(endAnchor, startIdx);
  if (anchorIdx === -1) throw new Error('No se encontró el marcador de cierre ("/* ---------------- ICONS") después de RAW_DATA en index.html.');

  // The array's closing "];" sits immediately before the anchor.
  const closeIdx = html.lastIndexOf('];', anchorIdx);
  if (closeIdx === -1 || closeIdx < startIdx) throw new Error('No se pudo ubicar el cierre "];" del arreglo RAW_DATA.');

  const before = html.slice(0, startIdx);
  const after = html.slice(closeIdx + 1); // keep the ';' from '];'
  const newArrayLiteral = 'let RAW_DATA = ' + JSON.stringify(data) + ';';

  const newHtml = before + newArrayLiteral + after;
  fs.writeFileSync(HTML_PATH, newHtml, 'utf8');

  console.log(`OK — ${data.length} registros escritos en index.html desde data/bitacora.xlsx.`);
}

main();
