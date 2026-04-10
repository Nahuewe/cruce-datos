/* eslint-disable react-hooks/exhaustive-deps */
import { useState, useMemo, useCallback, useRef, useEffect } from "react";
import * as XLSX from "xlsx";

/* ─── Web Worker ─────────────────────────────────────────────────────────────── */
const WORKER_CODE = `
const normalizar = (v) =>
  String(v ?? "").toLowerCase()
    .normalize("NFD").replace(/[\\u0300-\\u036f]/g, "")
    .replace(/\\s+/g, " ").trim();

const levenshtein = (a, b) => {
  if (!a.length) return b.length;
  if (!b.length) return a.length;
  const prev = Array.from({length: b.length + 1}, (_, i) => i);
  const curr = new Array(b.length + 1);
  for (let i = 1; i <= a.length; i++) {
    curr[0] = i;
    for (let j = 1; j <= b.length; j++) {
      curr[j] = a[i-1] === b[j-1]
        ? prev[j-1]
        : 1 + Math.min(prev[j-1], curr[j-1], prev[j]);
    }
    prev.set ? prev.set(curr) : curr.forEach((v,k) => prev[k] = v);
  }
  return prev[b.length];
};

const similitud = (a, b) => {
  const na = normalizar(a), nb = normalizar(b);
  if (na === nb) return 1;
  const maxLen = Math.max(na.length, nb.length);
  return maxLen ? 1 - levenshtein(na, nb) / maxLen : 1;
};

self.onmessage = function(e) {
  const { base, comparar, columnasClaves, umbral, indice, nombreArchivo } = e.data;
  const resultados = [];
  const coincidenciasUsadas = new Set();
  const total = base.length + comparar.length;
  let procesados = 0;
  const BATCH = 50;

  const procesarLote = (desde) => {
    const hasta = Math.min(desde + BATCH, base.length);
    for (let bi = desde; bi < hasta; bi++) {
      const filaBase = base[bi];
      const claveBase = columnasClaves.map(c => String(filaBase[c] ?? "")).join("|");
      let mejorCoincidencia = null, mejorPuntaje = 0, mejorIdx = -1;

      for (let ci = 0; ci < comparar.length; ci++) {
        const claveComp = columnasClaves.map(c => String(comparar[ci][c] ?? "")).join("|");
        const puntaje = similitud(claveBase, claveComp);
        if (puntaje > mejorPuntaje) { mejorPuntaje = puntaje; mejorCoincidencia = comparar[ci]; mejorIdx = ci; }
      }

      if (mejorCoincidencia && mejorPuntaje >= umbral) {
        coincidenciasUsadas.add(mejorIdx);
        const todasColumnas = Array.from(new Set([...Object.keys(filaBase), ...Object.keys(mejorCoincidencia)]));
        const columnasMerge = {}, diferencias = {};
        for (const col of todasColumnas) {
          const valBase = filaBase[col] ?? "", valComp = mejorCoincidencia[col] ?? "";
          columnasMerge[col] = { base: valBase, comparar: valComp };
          diferencias[col] = normalizar(String(valBase)) !== normalizar(String(valComp));
        }
        resultados.push({ columnasMerge, diferencias, puntajeCoincidencia: mejorPuntaje, estado: "coincide" });
      } else {
        const columnasMerge = {}, diferencias = {};
        for (const col of Object.keys(filaBase)) { columnasMerge[col] = { base: filaBase[col] ?? "", comparar: "" }; diferencias[col] = false; }
        resultados.push({ columnasMerge, diferencias, puntajeCoincidencia: 0, estado: "solo_base" });
      }
      procesados++;
    }

    const progreso = Math.round((procesados / total) * 100);
    self.postMessage({ tipo: "progreso", indice, progreso, procesados, total });

    if (hasta < base.length) {
      setTimeout(() => procesarLote(hasta), 0);
    } else {
      for (let ci = 0; ci < comparar.length; ci++) {
        if (coincidenciasUsadas.has(ci)) continue;
        const filaComp = comparar[ci];
        const columnasMerge = {}, diferencias = {};
        for (const col of Object.keys(filaComp)) { columnasMerge[col] = { base: "", comparar: filaComp[col] ?? "" }; diferencias[col] = false; }
        resultados.push({ columnasMerge, diferencias, puntajeCoincidencia: 0, estado: "solo_comparar" });
      }
      self.postMessage({ tipo: "listo", indice, nombreArchivo, resultados });
    }
  };

  procesarLote(0);
};
`;

function crearWorker() {
  const blob = new Blob([WORKER_CODE], { type: "application/javascript" });
  return new Worker(URL.createObjectURL(blob));
}

/* ─── SVG Icons ─────────────────────────────────────────────────────────────── */
const IcoUpload = () => (
  <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8">
    <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" /><polyline points="17 8 12 3 7 8" />
    <line x1="12" y1="3" x2="12" y2="15" />
  </svg>
);
const IcoFile = () => (
  <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8">
    <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
    <polyline points="14 2 14 8 20 8" />
  </svg>
);
const IcoPlus = () => (
  <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
    <line x1="12" y1="5" x2="12" y2="19" /><line x1="5" y1="12" x2="19" y2="12" />
  </svg>
);
const IcoClose = () => (
  <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
    <line x1="18" y1="6" x2="6" y2="18" /><line x1="6" y1="6" x2="18" y2="18" />
  </svg>
);
const IcoDownload = () => (
  <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
    <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
    <polyline points="7 10 12 15 17 10" />
    <line x1="12" y1="15" x2="12" y2="3" />
  </svg>
);
const IcoReset = () => (
  <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
    <polyline points="1 4 1 10 7 10" />
    <path d="M3.51 15a9 9 0 1 0 .49-4.5" />
  </svg>
);
const IcoSearch = () => (
  <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
    <circle cx="11" cy="11" r="8" /><line x1="21" y1="21" x2="16.65" y2="16.65" />
  </svg>
);
const IcoVote = () => (
  <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8">
    <path d="M9 11l3 3L22 4" /><path d="M21 12v7a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11" />
  </svg>
);
const IcoShield = () => (
  <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8">
    <path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z" />
  </svg>
);
const IcoCheck = () => (
  <svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
    <polyline points="20 6 9 17 4 12" />
  </svg>
);

/* ─── Styles ──────────────────────────────────────────────────────────────────── */
const GLOBAL_STYLES = `
  @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500;600&family=IBM+Plex+Sans:wght@400;500;600;700&display=swap');

  * { box-sizing: border-box; }
  body { margin: 0; }

  .comparador-root {
    --c-bg: #F4F3EF;
    --c-surface: #FFFFFF;
    --c-border: #D6D1C8;
    --c-border-strong: #A89F92;
    --c-text: #1A1814;
    --c-text-muted: #6B6560;
    --c-text-light: #9B9590;
    --c-accent: #1B3A6B;
    --c-accent-light: #E8EDF5;
    --c-accent-hover: #142D55;
    --c-success: #1A5C3A;
    --c-success-bg: #EBF5EE;
    --c-success-border: #B8DECA;
    --c-warning: #7A4F00;
    --c-warning-bg: #FFF8EC;
    --c-warning-border: #F0D99A;
    --c-danger: #7A1C1C;
    --c-danger-bg: #FDECEC;
    --c-danger-border: #F0B8B8;
    --c-info: #0F4D7A;
    --c-info-bg: #E8F4FD;
    --c-info-border: #AAD5F5;
    --font-sans: 'IBM Plex Sans', system-ui, sans-serif;
    --font-mono: 'IBM Plex Mono', monospace;
    min-height: 100vh;
    background: var(--c-bg);
    font-family: var(--font-sans);
    color: var(--c-text);
  }

  .comparador-root ::-webkit-scrollbar { width: 6px; height: 6px; }
  .comparador-root ::-webkit-scrollbar-track { background: var(--c-bg); }
  .comparador-root ::-webkit-scrollbar-thumb { background: var(--c-border-strong); border-radius: 3px; }

  .comparador-root input[type=range] {
    -webkit-appearance: none; height: 3px; border-radius: 2px;
    outline: none; background: var(--c-border);
  }
  .comparador-root input[type=range]::-webkit-slider-thumb {
    -webkit-appearance: none; width: 15px; height: 15px; border-radius: 50%;
    background: var(--c-accent); cursor: pointer; border: 2px solid white;
    box-shadow: 0 1px 3px rgba(0,0,0,0.3);
  }

  @keyframes spin { to { transform: rotate(360deg); } }
  @keyframes pulse-dot {
    0%, 100% { opacity: 0.3; transform: scale(0.8); }
    50% { opacity: 1; transform: scale(1.1); }
  }
  @keyframes fadeIn { from { opacity: 0; transform: translateY(8px); } to { opacity: 1; transform: translateY(0); } }

  .fade-in { animation: fadeIn 0.3s ease forwards; }

  .card { background: var(--c-surface); border: 1px solid var(--c-border); border-radius: 4px; overflow: hidden; }
  .card-header { padding: 12px 20px; border-bottom: 1px solid var(--c-border); background: #FAFAF8; display: flex; align-items: center; gap: 10px; }

  .step-badge { width: 22px; height: 22px; border-radius: 2px; display: flex; align-items: center; justify-content: center; font-size: 10px; font-weight: 700; font-family: var(--font-mono); background: var(--c-accent); color: white; flex-shrink: 0; letter-spacing: 0; }
  .step-badge.done { background: var(--c-success); }

  .card-title { font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.08em; color: var(--c-text-muted); font-family: var(--font-mono); }

  .btn-primary { background: var(--c-accent); color: white; border: none; padding: 9px 18px; border-radius: 3px; font-size: 12px; font-weight: 600; cursor: pointer; font-family: var(--font-mono); letter-spacing: 0.04em; text-transform: uppercase; transition: background 0.15s, transform 0.1s; display: flex; align-items: center; gap: 7px; }
  .btn-primary:hover:not(:disabled) { background: var(--c-accent-hover); }
  .btn-primary:active:not(:disabled) { transform: scale(0.98); }
  .btn-primary:disabled { opacity: 0.4; cursor: not-allowed; }

  .btn-secondary { background: transparent; color: var(--c-text-muted); border: 1px solid var(--c-border); padding: 8px 16px; border-radius: 3px; font-size: 11px; font-weight: 600; cursor: pointer; font-family: var(--font-mono); letter-spacing: 0.03em; text-transform: uppercase; transition: all 0.15s; display: flex; align-items: center; gap: 7px; }
  .btn-secondary:hover { border-color: var(--c-border-strong); color: var(--c-text); background: #F4F3EF; }

  .badge { font-size: 9px; font-weight: 600; padding: 2px 7px; border-radius: 2px; font-family: var(--font-mono); text-transform: uppercase; letter-spacing: 0.06em; border: 1px solid; }
  .badge-success { background: var(--c-success-bg); color: var(--c-success); border-color: var(--c-success-border); }
  .badge-danger  { background: var(--c-danger-bg);  color: var(--c-danger);  border-color: var(--c-danger-border); }
  .badge-info    { background: var(--c-info-bg);    color: var(--c-info);    border-color: var(--c-info-border); }
  .badge-warning { background: var(--c-warning-bg); color: var(--c-warning); border-color: var(--c-warning-border); }
  .badge-neutral { background: #F0EDE8; color: var(--c-text-muted); border-color: var(--c-border); }

  .drop-zone { border: 1.5px dashed var(--c-border); border-radius: 3px; padding: 24px 16px; cursor: pointer; text-align: center; transition: all 0.2s; user-select: none; min-height: 120px; display: flex; flex-direction: column; align-items: center; justify-content: center; gap: 8px; background: #FAFAF8; }
  .drop-zone:hover { border-color: var(--c-accent); background: var(--c-accent-light); }
  .drop-zone.loaded { border-style: solid; border-color: var(--c-accent); background: var(--c-accent-light); }
  .drop-zone.drag { border-color: var(--c-accent); background: var(--c-accent-light); transform: scale(1.01); }

  .data-table { width: 100%; border-collapse: collapse; font-size: 11px; font-family: var(--font-mono); }
  .data-table th { background: #F4F3EF; padding: 9px 12px; text-align: left; font-size: 9px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.08em; color: var(--c-text-muted); border-bottom: 2px solid var(--c-border); white-space: nowrap; position: sticky; top: 0; z-index: 1; }
  .data-table td { padding: 8px 12px; border-bottom: 1px solid #F0EDE8; white-space: nowrap; vertical-align: top; }
  .data-table tr.clickable { cursor: pointer; transition: background 0.1s; }
  .data-table tr.clickable:hover { background: #FAFAF8; }
  .data-table tr.expanded { background: var(--c-accent-light); }

  .stat-box { text-align: center; padding: 16px 12px; }
  .stat-num { font-size: 28px; font-weight: 700; font-family: var(--font-mono); line-height: 1; }
  .stat-lbl { font-size: 9px; font-weight: 600; text-transform: uppercase; letter-spacing: 0.07em; color: var(--c-text-muted); margin-top: 5px; font-family: var(--font-mono); }

  .dot { width: 6px; height: 6px; border-radius: 50%; background: var(--c-accent); display: inline-block; }

  .official-stripe { height: 4px; background: linear-gradient(90deg, var(--c-accent) 0%, var(--c-accent) 60%, #D4A82A 60%, #D4A82A 100%); }

  .filter-pill { padding: 5px 12px; border-radius: 2px; border: 1px solid var(--c-border); font-size: 10px; font-weight: 600; cursor: pointer; transition: all 0.15s; font-family: var(--font-mono); text-transform: uppercase; letter-spacing: 0.04em; background: transparent; color: var(--c-text-muted); }
  .filter-pill:hover { border-color: var(--c-accent); color: var(--c-accent); }
  .filter-pill.active { background: var(--c-accent); border-color: var(--c-accent); color: white; }
  .filter-pill.active-warn { background: var(--c-warning-bg); border-color: var(--c-warning-border); color: var(--c-warning); }

  .col-chip { padding: 5px 11px; border-radius: 2px; border: 1px solid var(--c-border); font-size: 10px; font-weight: 500; cursor: pointer; transition: all 0.15s; font-family: var(--font-mono); background: white; color: var(--c-text-muted); display: flex; align-items: center; gap: 5px; }
  .col-chip:hover { border-color: var(--c-accent); color: var(--c-accent); }
  .col-chip.active { background: var(--c-accent); border-color: var(--c-accent); color: white; }

  .prog-bar-track { height: 4px; background: #E8E5DF; border-radius: 2px; overflow: hidden; }
  .prog-bar-fill { height: 100%; background: var(--c-accent); border-radius: 2px; transition: width 0.3s ease; }

  .circ-progress { transform: rotate(-90deg); }
`;

/* ─── File colors ─────────────────────────────────────────────────────────────── */
const FILE_COLORS = {
  0: { accent: "#1B3A6B", bg: "var(--c-accent-light)", label: "ARCHIVO BASE" },
  1: { accent: "#1A5C3A", bg: "var(--c-success-bg)", label: "COMPARAR A" },
  2: { accent: "#7A4F00", bg: "var(--c-warning-bg)", label: "COMPARAR B" },
};

/* ─── Upload Zone ─────────────────────────────────────────────────────────────── */
function ZonaCarga({ indice, datos, alCargar, alQuitar }) {
  const refInput = useRef();
  const [drag, setDrag] = useState(false);
  const fc = FILE_COLORS[indice];

  const leerArchivo = (archivo) => {
    if (!archivo) return;
    const lector = new FileReader();
    lector.onload = (e) => {
      const wb = XLSX.read(e.target.result, { type: "binary" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json(ws, { defval: "" });
      alCargar(indice, { json, nombre: archivo.name, filas: json.length, columnas: Object.keys(json[0] || {}).length });
    };
    lector.readAsBinaryString(archivo);
  };

  return (
    <div className="relative">
      {datos && alQuitar && (
        <button
          onClick={() => alQuitar(indice)}
          className="absolute -top-2 -right-2 z-10 w-5 h-5 rounded-full bg-white flex items-center justify-center cursor-pointer transition-all"
          style={{ border: "1.5px solid var(--c-border)", color: "var(--c-danger)" }}
        >
          <IcoClose />
        </button>
      )}
      <div
        onDragOver={(e) => { e.preventDefault(); setDrag(true); }}
        onDragLeave={() => setDrag(false)}
        onDrop={(e) => { e.preventDefault(); setDrag(false); leerArchivo(e.dataTransfer.files[0]); }}
        onClick={() => refInput.current?.click()}
        className={`drop-zone${datos ? " loaded" : ""}${drag ? " drag" : ""}`}
        style={datos ? { borderColor: fc.accent, background: fc.bg } : {}}
      >
        <input ref={refInput} type="file" accept=".xlsx,.xls,.csv" className="hidden"
          onChange={(e) => leerArchivo(e.target.files[0])} />

        <div style={{ color: datos ? fc.accent : "var(--c-text-light)" }}>
          {datos ? <IcoFile /> : <IcoUpload />}
        </div>

        <div className="font-mono text-sm font-bold uppercase tracking-widest"
          style={{ color: datos ? fc.accent : "var(--c-text-muted)" }}>
          {fc.label}
        </div>

        {datos ? (
          <>
            <div className="text-base font-semibold max-w-40 overflow-hidden text-ellipsis whitespace-nowrap"
              style={{ color: "var(--c-text)" }}>
              {datos.nombre}
            </div>
            <div className="badge badge-neutral">
              {datos.filas.toLocaleString()} reg · {datos.columnas} col
            </div>
          </>
        ) : (
          <div className="text-sm" style={{ color: "var(--c-text-light)" }}>
            Arrastrar o hacer clic · xlsx, xls, csv
          </div>
        )}
      </div>
    </div>
  );
}

/* ─── Loading Screen ─────────────────────────────────────────────────────────── */
const PROCESSING_MSGS = [
  "Normalizando registros...",
  "Aplicando distancia de Levenshtein...",
  "Calculando similitud entre registros...",
  "Cruzando padrones...",
  "Detectando discrepancias...",
  "Consolidando resultados...",
  "Verificando integridad...",
];

function PantallaLoading({ progresos, archivos }) {
  const [msgIdx, setMsgIdx] = useState(0);
  useEffect(() => {
    const t = setInterval(() => setMsgIdx(i => (i + 1) % PROCESSING_MSGS.length), 2000);
    return () => clearInterval(t);
  }, []);

  const globalProg = progresos.length
    ? Math.round(progresos.reduce((s, p) => s + p, 0) / progresos.length) : 0;

  const R = 36, CIRC = 2 * Math.PI * R;

  return (
    <div className="card fade-in">
      <div className="official-stripe" />
      <div className="flex flex-col items-center gap-5 px-7 py-8">

        {/* Circular progress */}
        <div className="relative w-24 h-24">
          <svg width="96" height="96" viewBox="0 0 96 96" className="circ-progress">
            <circle cx="48" cy="48" r={R} fill="none" stroke="#E8E5DF" strokeWidth="5" />
            <circle cx="48" cy="48" r={R} fill="none" stroke="var(--c-accent)" strokeWidth="5"
              strokeLinecap="round"
              strokeDasharray={CIRC}
              strokeDashoffset={CIRC * (1 - globalProg / 100)}
              style={{ transition: "stroke-dashoffset 0.4s ease" }}
            />
          </svg>
          <div className="absolute inset-0 flex flex-col items-center justify-center">
            <span className="font-mono text-lg font-bold leading-none" style={{ color: "var(--c-accent)" }}>
              {globalProg}%
            </span>
            <span className="font-mono text-xs mt-0.5 uppercase" style={{ color: "var(--c-text-muted)" }}>
              procesado
            </span>
          </div>
        </div>

        {/* Message */}
        <div className="text-center">
          <div className="text-sm font-semibold mb-1" style={{ color: "var(--c-text)" }}>
            {PROCESSING_MSGS[msgIdx]}
          </div>
          <div className="font-mono text-sm" style={{ color: "var(--c-text-muted)" }}>
            Procesando {archivos.filter(Boolean).length} archivos en background
          </div>
        </div>

        {/* Per-file bars */}
        <div className="w-full max-w-sm flex flex-col gap-3">
          {progresos.map((prog, i) => {
            const fc = FILE_COLORS[i + 1];
            const arch = archivos[i + 1];
            return (
              <div key={i}>
                <div className="flex justify-between mb-1">
                  <span className="font-mono text-sm uppercase tracking-wider" style={{ color: "var(--c-text-muted)" }}>
                    Base ↔ {arch?.nombre ?? `Archivo ${i + 2}`}
                  </span>
                  <span className="font-mono text-sm" style={{ color: "var(--c-text-muted)" }}>{prog}%</span>
                </div>
                <div className="prog-bar-track">
                  <div className="prog-bar-fill" style={{ width: `${prog}%`, background: fc.accent }} />
                </div>
              </div>
            );
          })}
        </div>

        {/* Animated dots */}
        <div className="flex gap-1.5">
          {[0, 1, 2].map(i => (
            <div key={i} className="dot"
              style={{ animation: `pulse-dot 1.4s ease-in-out ${i * 0.22}s infinite` }} />
          ))}
        </div>
      </div>
    </div>
  );
}

/* ─── Status Badge ───────────────────────────────────────────────────────────── */
function BadgeEstado({ estado }) {
  const cfg = {
    coincide: { cls: "badge-success", txt: "Coincide" },
    solo_base: { cls: "badge-danger", txt: "Solo base" },
    solo_comparar: { cls: "badge-info", txt: "Solo comp." },
  };
  const { cls, txt } = cfg[estado] || {};
  return <span className={`badge ${cls}`}>{txt}</span>;
}

/* ─── Step Card ──────────────────────────────────────────────────────────────── */
function PasoCard({ numero, titulo, listo, activo, extra, children }) {
  return (
    <div className="card">
      {activo && <div className="official-stripe" />}
      <div className="card-header">
        <span className={`step-badge${listo ? " done" : ""}`}>
          {listo ? <IcoCheck /> : numero}
        </span>
        <span className="card-title">{titulo}</span>
        {activo && !listo && (
          <div className="flex gap-0.5 ml-1">
            {[0, 1, 2].map(i => (
              <div key={i} className="dot" style={{
                width: 4, height: 4,
                animation: `pulse-dot 1.2s ease-in-out ${i * 0.2}s infinite`,
              }} />
            ))}
          </div>
        )}
        {extra && <div className="ml-auto">{extra}</div>}
      </div>
      <div className="p-5">{children}</div>
    </div>
  );
}

/* ─── Main Component ─────────────────────────────────────────────────────────── */
export default function ComparadorExcel() {
  const [archivos, setArchivos] = useState([null, null]);
  const [columnasClaves, setColumnasClaves] = useState([]);
  const [umbral, setUmbral] = useState(82);
  const [resultados, setResultados] = useState(null);
  const [procesando, setProcesando] = useState(false);
  const [progresos, setProgresos] = useState([]);
  const [filtroEstado, setFiltroEstado] = useState("todos");
  const [filtroDiferencias, setFiltroDiferencias] = useState(false);
  const [busqueda, setBusqueda] = useState("");
  const [pagina, setPagina] = useState(1);
  const [filaExpandida, setFilaExpandida] = useState(null);
  const [tabActiva, setTabActiva] = useState(0);
  const workersRef = useRef([]);
  const FILAS_POR_PAG = 25;

  const todasColumnas = useMemo(() => {
    const cols = new Set();
    archivos.forEach(d => d?.json?.forEach(f => Object.keys(f).forEach(k => cols.add(k))));
    return Array.from(cols);
  }, [archivos]);

  const cargarArchivo = useCallback((idx, datos) => {
    setArchivos(prev => { const n = [...prev]; n[idx] = datos; return n; });
    setResultados(null); setColumnasClaves([]);
  }, []);

  const quitarArchivo = useCallback((idx) => {
    setArchivos(prev => { const n = [...prev]; n.splice(idx, 1); return n; });
    setResultados(null);
  }, []);

  const ejecutarComparacion = useCallback(() => {
    const validos = archivos.filter(Boolean);
    if (validos.length < 2 || columnasClaves.length === 0) return;

    workersRef.current.forEach(w => w.terminate());
    workersRef.current = [];

    const comparaciones = archivos.slice(1).filter(Boolean);
    setProgresos(comparaciones.map(() => 0));
    setProcesando(true);
    setResultados(null);

    const resultadosParciales = new Array(comparaciones.length).fill(null);
    let listos = 0;

    comparaciones.forEach((archivo, i) => {
      const worker = crearWorker();
      workersRef.current.push(worker);

      worker.onmessage = (e) => {
        const msg = e.data;
        if (msg.tipo === "progreso") {
          setProgresos(prev => { const n = [...prev]; n[i] = msg.progreso; return n; });
        } else if (msg.tipo === "listo") {
          resultadosParciales[i] = { indice: i + 1, nombreArchivo: msg.nombreArchivo, filas: msg.resultados };
          listos++;
          worker.terminate();
          if (listos === comparaciones.length) {
            setResultados(resultadosParciales.filter(Boolean));
            setTabActiva(0); setPagina(1);
            setProcesando(false);
          }
        }
      };

      worker.postMessage({
        base: archivos[0].json, comparar: archivo.json,
        columnasClaves, umbral: umbral / 100,
        indice: i + 1, nombreArchivo: archivo.nombre,
      });
    });
  }, [archivos, columnasClaves, umbral]);

  const resultadoActual = resultados?.[tabActiva]?.filas ?? [];

  const estadisticas = useMemo(() => {
    if (!resultados) return null;
    return resultados.map(r => ({
      nombreArchivo: r.nombreArchivo, total: r.filas.length,
      coinciden: r.filas.filter(f => f.estado === "coincide").length,
      conDiferencias: r.filas.filter(f => f.estado === "coincide" && Object.values(f.diferencias).some(Boolean)).length,
      soloBase: r.filas.filter(f => f.estado === "solo_base").length,
      soloComparar: r.filas.filter(f => f.estado === "solo_comparar").length,
    }));
  }, [resultados]);

  const filasFiltradas = useMemo(() => resultadoActual.filter(f => {
    if (filtroEstado !== "todos" && f.estado !== filtroEstado) return false;
    if (filtroDiferencias && !Object.values(f.diferencias).some(Boolean)) return false;
    if (busqueda) {
      const b = busqueda.toLowerCase();
      return Object.values(f.columnasMerge).some(v =>
        String(v.base).toLowerCase().includes(b) || String(v.comparar).toLowerCase().includes(b));
    }
    return true;
  }), [resultadoActual, filtroEstado, filtroDiferencias, busqueda]);

  const filasPaginadas = useMemo(() => {
    const ini = (pagina - 1) * FILAS_POR_PAG;
    return filasFiltradas.slice(ini, ini + FILAS_POR_PAG);
  }, [filasFiltradas, pagina]);

  const totalPaginas = Math.ceil(filasFiltradas.length / FILAS_POR_PAG);

  const exportar = () => {
    if (!resultados) return;
    resultados.forEach(r => {
      if (!r.filas.length) return;
      const cols = Object.keys(r.filas[0].columnasMerge);
      const filas = r.filas.map(f => {
        const fila = { _estado: f.estado, _similitud: Math.round(f.puntajeCoincidencia * 100) + "%" };
        cols.forEach(c => {
          fila[`${c}_base`] = f.columnasMerge[c]?.base ?? "";
          fila[`${c}_comparar`] = f.columnasMerge[c]?.comparar ?? "";
          fila[`${c}_diferente`] = f.diferencias[c] ? "SÍ" : "";
        });
        return fila;
      });
      const hoja = XLSX.utils.json_to_sheet(filas);
      const libro = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(libro, hoja, "Comparación");
      XLSX.writeFile(libro, `cruce_padron_${r.nombreArchivo}`);
    });
  };

  const puedoComparar = archivos.filter(Boolean).length >= 2 && columnasClaves.length > 0;
  const now = new Date();
  const fechaStr = now.toLocaleDateString("es-AR", { day: "2-digit", month: "long", year: "numeric" });

  return (
    <div className="comparador-root">
      <style>{GLOBAL_STYLES}</style>

      {/* ── Top accent bar ── */}
      <div className="h-1.25" style={{ background: "var(--c-accent)" }} />

      {/* ── Header ── */}
      <div className="border-b" style={{ background: "var(--c-surface)", borderColor: "var(--c-border)" }}>
        <div className="max-w-375 mx-auto px-6">
          <div className="flex items-center justify-between py-3.5 gap-4 flex-wrap">

            {/* Logo + title */}
            <div className="flex items-center gap-3.5">
              <div className="w-10 h-10 rounded flex items-center justify-center text-white shrink-0"
                style={{ background: "var(--c-accent)" }}>
                <IcoVote />
              </div>
              <div>
                <div className="text-[15px] font-bold leading-tight" style={{ color: "var(--c-text)" }}>
                  Sistema de Cruzamiento de Datos Electorales
                </div>
                <div className="font-mono text-xs uppercase tracking-wider mt-0.5" style={{ color: "var(--c-text-muted)" }}>
                  Verificación y comparación de padrones
                </div>
              </div>
            </div>

            {/* Meta info */}
            <div className="flex items-center gap-4 flex-wrap">
              <div className="flex items-center gap-1.5 font-mono text-sm" style={{ color: "var(--c-text-muted)" }}>
                <span style={{ color: "var(--c-accent)" }}><IcoShield /></span>
                Uso oficial interno
              </div>
              <div className="font-mono text-sm px-2.5 py-1 rounded border"
                style={{ color: "var(--c-text-muted)", background: "#F4F3EF", borderColor: "var(--c-border)" }}>
                {fechaStr}
              </div>
              {resultados && (
                <button className="btn-secondary" onClick={exportar} style={{ fontSize: 13, padding: "6px 12px" }}>
                  <IcoDownload /> Exportar
                </button>
              )}
            </div>
          </div>
        </div>
      </div>

      {/* ── Gold stripe ── */}
      <div className="h-0.5" style={{ background: "linear-gradient(90deg, var(--c-accent) 0%, #D4A82A 100%)" }} />

      {/* ── Main content ── */}
      <div className="max-w-375 mx-auto px-6 pt-6 pb-12">

        {/* Instruction banner */}
        <div
          className="rounded p-4 mb-5 flex gap-3 items-start border"
          style={{ background: "var(--c-accent-light)", borderColor: "#C8D7ED" }}
        >
          <div className="shrink-0 mt-0.5" style={{ color: "var(--c-accent)" }}>
            <IcoShield />
          </div>

          <div>
            <div
              className="text-[15px] font-bold mb-1"
              style={{ color: "var(--c-accent)" }}
            >
              Instrucciones de uso
            </div>

            <ul className="text-[15px] leading-relaxed list-decimal pl-5" style={{ color: "var(--c-text-muted)" }}>
              <li>Cargue los archivos de padrón en formato Excel o CSV.</li>
              <li>Seleccione las columnas identificadoras de cada registro (DNI, nombre, etc.).</li>
              <li>Ajuste el umbral de similitud y ejecute el cruzamiento.</li>
              <li>Revise las discrepancias detectadas y exporte el informe.</li>
            </ul>
          </div>
        </div>

        <div className="flex flex-col gap-4">

          {/* ── Paso 1: Archivos ── */}
          <PasoCard numero="1" titulo="Carga de archivos"
            listo={archivos.filter(Boolean).length >= 2}
            activo={archivos.filter(Boolean).length < 2}
            extra={
              <span className="font-mono text-sm" style={{ color: "var(--c-text-muted)" }}>
                {archivos.filter(Boolean).length}/{archivos.length} archivos cargados
              </span>
            }>
            <div
              className="grid gap-3.5"
              style={{ gridTemplateColumns: archivos.length === 3 ? "repeat(3, 1fr)" : `repeat(${Math.min(archivos.length + 1, 3)}, 1fr)` }}
            >
              {archivos.map((datos, i) => (
                <ZonaCarga key={i} indice={i} datos={datos} alCargar={cargarArchivo}
                  alQuitar={i >= 2 ? quitarArchivo : null} />
              ))}
              {archivos.length < 3 && (
                <button
                  onClick={() => setArchivos(prev => [...prev, null])}
                  className="transition-all min-h-30"
                  style={{
                    border: "1.5px dashed var(--c-border)", borderRadius: 3,
                    padding: "24px 16px", cursor: "pointer", background: "transparent",
                    display: "flex", flexDirection: "column", alignItems: "center",
                    justifyContent: "center", gap: 8, color: "var(--c-text-light)",
                    fontFamily: "var(--font-sans)",
                  }}
                  onMouseOver={e => { e.currentTarget.style.borderColor = "var(--c-accent)"; e.currentTarget.style.color = "var(--c-accent)"; }}
                  onMouseOut={e => { e.currentTarget.style.borderColor = "var(--c-border)"; e.currentTarget.style.color = "var(--c-text-light)"; }}
                >
                  <div className="w-7.5 h-7.5 rounded-full flex items-center justify-center"
                    style={{ border: "1.5px solid currentColor" }}>
                    <IcoPlus />
                  </div>
                  <span className="text-sm font-semibold uppercase tracking-wider font-mono">
                    Agregar tercer archivo
                  </span>
                </button>
              )}
            </div>
          </PasoCard>

          {/* ── Paso 2: Columnas ── */}
          <PasoCard numero="2" titulo="Columnas identificadoras"
            listo={columnasClaves.length > 0}
            activo={archivos.filter(Boolean).length >= 2 && columnasClaves.length === 0}
            extra={columnasClaves.length > 0 && (
              <span className="badge badge-success">{columnasClaves.length} col. selec.</span>
            )}>
            {todasColumnas.length > 0 ? (
              <>
                <p className="text-[15px] mb-3 mt-0 leading-relaxed" style={{ color: "var(--c-text-muted)" }}>
                  Seleccione las columnas que identifican de forma unívoca cada registro (DNI, número de elector, nombre completo, etc.).
                </p>
                <div className="flex flex-wrap gap-1.5 mb-5">
                  {todasColumnas.map(col => {
                    const activa = columnasClaves.includes(col);
                    return (
                      <button key={col}
                        onClick={() => setColumnasClaves(prev => activa ? prev.filter(x => x !== col) : [...prev, col])}
                        className={`col-chip${activa ? " active" : ""}`}>
                        {activa && <IcoCheck />}
                        {col}
                      </button>
                    );
                  })}
                </div>

                {/* Similarity threshold */}
                <div className="rounded border p-3.5" style={{ background: "#FAFAF8", borderColor: "var(--c-border)" }}>
                  <div className="text-xs font-bold uppercase tracking-wider font-mono mb-2.5"
                    style={{ color: "var(--c-text-muted)" }}>
                    Umbral de similitud para coincidencia
                  </div>
                  <div className="flex items-center gap-4 flex-wrap">
                    <input type="range" min={0} max={100} value={umbral}
                      onChange={e => setUmbral(Number(e.target.value))}
                      className="flex-1 min-w-30" />
                    <div className="flex items-baseline gap-1">
                      <span className="font-mono text-xl font-bold" style={{ color: "var(--c-accent)" }}>
                        {umbral}%
                      </span>
                    </div>
                    <span className={`badge ${umbral >= 95 ? "badge-success" : umbral >= 80 ? "badge-warning" : "badge-danger"}`}>
                      {umbral >= 95 ? "Coincidencia exacta" : umbral >= 80 ? "Coincidencia moderada" : "Coincidencia amplia"}
                    </span>
                  </div>
                </div>
              </>
            ) : (
              <p className="text-[15px] m-0" style={{ color: "var(--c-text-muted)" }}>
                Cargue al menos 2 archivos para visualizar las columnas disponibles.
              </p>
            )}
          </PasoCard>

          {/* ── Paso 3: Ejecutar ── */}
          <PasoCard numero="3" titulo="Ejecución del cruzamiento"
            listo={!!resultados}
            activo={archivos.filter(Boolean).length >= 2 && columnasClaves.length > 0}>
            <div className="flex items-center gap-2.5 flex-wrap">
              <button className="btn-primary" onClick={ejecutarComparacion} disabled={!puedoComparar || procesando}>
                {procesando ? (
                  <>
                    <div className="w-3.25 h-3.25 rounded-full border-2 border-white/30 border-t-white"
                      style={{ animation: "spin .7s linear infinite" }} />
                    Procesando...
                  </>
                ) : (
                  <><IcoVote /> Iniciar cruzamiento</>
                )}
              </button>

              {(resultados || procesando) && (
                <button className="btn-secondary" onClick={() => {
                  workersRef.current.forEach(w => w.terminate()); workersRef.current = [];
                  setResultados(null); setArchivos([null, null]); setColumnasClaves([]); setProcesando(false);
                }}>
                  <IcoReset /> {procesando ? "Cancelar" : "Reiniciar"}
                </button>
              )}

              {!puedoComparar && !procesando && (
                <span className="text-sm font-mono" style={{ color: "var(--c-text-light)" }}>
                  {archivos.filter(Boolean).length < 2 ? "↑ Cargue al menos 2 archivos. " : ""}
                  {columnasClaves.length === 0 ? "↑ Seleccione columnas identificadoras." : ""}
                </span>
              )}
            </div>

            {/* Time estimate */}
            {puedoComparar && !procesando && !resultados && (() => {
              const nBase = archivos[0]?.filas ?? 0;
              const nComp = archivos[1]?.filas ?? 0;
              const ops = nBase * nComp;
              const segs = Math.round(ops / 80000);
              if (segs < 2) return null;
              return (
                <div className="mt-2.5 text-sm font-mono" style={{ color: "var(--c-text-muted)" }}>
                  Tiempo estimado: ~{segs < 60 ? `${segs}s` : `${Math.round(segs / 60)}min`} · {nBase.toLocaleString()} × {nComp.toLocaleString()} comparaciones
                </div>
              );
            })()}
          </PasoCard>

          {/* ── Loading ── */}
          {procesando && <PantallaLoading progresos={progresos} archivos={archivos} />}

          {/* ── Results ── */}
          {resultados && !procesando && (
            <div className="fade-in flex flex-col gap-4">

              {/* Statistics */}
              {estadisticas.map((est, i) => (
                <div key={i} className="card">
                  <div className="card-header">
                    <div className="w-2 h-2 rounded-full shrink-0"
                      style={{ background: FILE_COLORS[i + 1]?.accent || "var(--c-accent)" }} />
                    <span className="card-title">Informe: Base ↔ {est.nombreArchivo}</span>
                    <div className="ml-auto flex items-center gap-2">
                      <div className="prog-bar-track" style={{ width: 80 }}>
                        <div className="prog-bar-fill" style={{ width: `${Math.round(est.coinciden / est.total * 100)}%` }} />
                      </div>
                      <span className="font-mono text-sm" style={{ color: "var(--c-text-muted)" }}>
                        {Math.round(est.coinciden / est.total * 100)}% coincidencia
                      </span>
                    </div>
                  </div>
                  <div className="grid grid-cols-5">
                    {[
                      ["Total registros", est.total, "var(--c-text)"],
                      ["Coincidencias", est.coinciden, "var(--c-success)"],
                      ["Con discrepancias", est.conDiferencias, "#7A4F00"],
                      ["Solo en base", est.soloBase, "var(--c-danger)"],
                      ["Solo en comparación", est.soloComparar, "var(--c-info)"],
                    ].map(([lbl, val, color], j) => (
                      <div key={j} className="stat-box"
                        style={{ borderLeft: j > 0 ? "1px solid var(--c-border)" : "none" }}>
                        <div className="stat-num" style={{ color }}>{val.toLocaleString()}</div>
                        <div className="stat-lbl">{lbl}</div>
                      </div>
                    ))}
                  </div>
                </div>
              ))}

              {/* Alert messages */}
              {estadisticas[0]?.coinciden === 0 && (
                <div className="rounded border p-3 text-[15px] font-mono"
                  style={{ background: "var(--c-warning-bg)", borderColor: "var(--c-warning-border)", color: "var(--c-warning)" }}>
                  ⚠ No se encontraron coincidencias. Verifique las columnas identificadoras y el umbral.
                </div>
              )}
              {estadisticas[0]?.conDiferencias > 0 && (
                <div className="rounded border p-3 text-[15px] font-mono"
                  style={{ background: "var(--c-info-bg)", borderColor: "var(--c-info-border)", color: "var(--c-info)" }}>
                  ◈ Se detectaron {estadisticas[0].conDiferencias.toLocaleString()} registros con discrepancias en sus datos.
                </div>
              )}

              {/* Tabs (multiple files) */}
              {resultados.length > 1 && (
                <div className="flex gap-1 border-b" style={{ borderColor: "var(--c-border)" }}>
                  {resultados.map((r, i) => (
                    <button key={i} onClick={() => { setTabActiva(i); setPagina(1); setFilaExpandida(null); }}
                      className="px-4 py-2 font-mono text-sm font-semibold uppercase tracking-wider cursor-pointer rounded-t transition-all"
                      style={{
                        border: "1px solid var(--c-border)",
                        borderBottom: tabActiva === i ? "2px solid var(--c-accent)" : "1px solid transparent",
                        background: tabActiva === i ? "var(--c-surface)" : "transparent",
                        color: tabActiva === i ? "var(--c-accent)" : "var(--c-text-muted)",
                      }}>
                      Base ↔ {r.nombreArchivo}
                    </button>
                  ))}
                </div>
              )}

              {/* Data table */}
              <div className="card">
                <div className="card-header">
                  <span className="step-badge done"><IcoCheck /></span>
                  <span className="card-title">Resultados del cruzamiento</span>
                  <span className="ml-auto font-mono text-sm px-2 py-0.5 rounded border"
                    style={{ color: "var(--c-text-muted)", background: "#F4F3EF", borderColor: "var(--c-border)" }}>
                    {filasFiltradas.length.toLocaleString()} registros
                  </span>
                </div>

                {/* Filters */}
                <div className="px-4 py-2.5 border-b flex flex-wrap gap-1.5 items-center"
                  style={{ background: "#FAFAF8", borderColor: "var(--c-border)" }}>
                  {[["todos", "Todos"], ["coincide", "Coinciden"], ["solo_base", "Solo base"], ["solo_comparar", "Solo comp."]].map(([val, lbl]) => (
                    <button key={val} className={`filter-pill${filtroEstado === val ? " active" : ""}`}
                      onClick={() => { setFiltroEstado(val); setPagina(1); }}>
                      {lbl}
                    </button>
                  ))}
                  <div className="w-px h-4.5 mx-0.5" style={{ background: "var(--c-border)" }} />
                  <button onClick={() => { setFiltroDiferencias(v => !v); setPagina(1); }}
                    className={`filter-pill${filtroDiferencias ? " active-warn" : ""}`}>
                    ◈ Solo discrepancias
                  </button>
                  <div className="relative flex-1 min-w-40 ml-1">
                    <div className="absolute left-2 top-1/2 -translate-y-1/2 pointer-events-none"
                      style={{ color: "var(--c-text-light)" }}>
                      <IcoSearch />
                    </div>
                    <input value={busqueda} onChange={e => { setBusqueda(e.target.value); setPagina(1); }}
                      placeholder="Buscar en resultados..."
                      className="w-full rounded border pl-7 pr-2.5 py-1.5 text-sm font-mono outline-none transition-colors"
                      style={{ background: "white", borderColor: "var(--c-border)", color: "var(--c-text)" }}
                      onFocus={e => e.target.style.borderColor = "var(--c-accent)"}
                      onBlur={e => e.target.style.borderColor = "var(--c-border)"}
                    />
                  </div>
                </div>

                {filasPaginadas.length === 0 ? (
                  <div className="py-12 text-center font-mono text-[15px]" style={{ color: "var(--c-text-light)" }}>
                    Sin resultados con los filtros aplicados.
                  </div>
                ) : (
                  <div className="overflow-x-auto">
                    <table className="data-table">
                      <thead>
                        <tr>
                          <th style={{ width: 28 }} />
                          <th>Estado</th>
                          <th>Similitud</th>
                          {filasPaginadas[0] && Object.keys(filasPaginadas[0].columnasMerge).map(col => (
                            <th key={col}>{col}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {filasPaginadas.map((fila, ri) => {
                          const idxGlobal = (pagina - 1) * FILAS_POR_PAG + ri;
                          const expandida = filaExpandida === idxGlobal;
                          const tieneDiff = Object.values(fila.diferencias).some(Boolean);
                          return (
                            <>
                              <tr key={ri} onClick={() => setFilaExpandida(expandida ? null : idxGlobal)}
                                className={`clickable${expandida ? " expanded" : ""}`}
                                style={tieneDiff && !expandida ? { background: "var(--c-warning-bg)" } : {}}>
                                <td className="text-center text-sm" style={{ color: "var(--c-text-muted)" }}>
                                  {expandida ? "▼" : "▶"}
                                </td>
                                <td><BadgeEstado estado={fila.estado} /></td>
                                <td>
                                  {fila.estado === "coincide" ? (
                                    <span className={`badge ${fila.puntajeCoincidencia >= 0.95 ? "badge-success" : fila.puntajeCoincidencia >= 0.82 ? "badge-warning" : "badge-danger"}`}>
                                      {Math.round(fila.puntajeCoincidencia * 100)}%
                                    </span>
                                  ) : <span style={{ color: "var(--c-text-light)" }}>—</span>}
                                </td>
                                {Object.keys(fila.columnasMerge).map((col, ci) => {
                                  const celda = fila.columnasMerge[col];
                                  const diff = fila.diferencias[col];
                                  return (
                                    <td key={ci} style={diff ? { background: "rgba(255,220,100,0.12)" } : {}}>
                                      {diff ? (
                                        <div className="flex flex-col gap-0.5">
                                          <span className="text-sm opacity-80 line-through" style={{ color: "var(--c-danger)" }}>
                                            {String(celda.base)}
                                          </span>
                                          <span className="font-semibold" style={{ color: "var(--c-success)" }}>
                                            {String(celda.comparar)}
                                          </span>
                                        </div>
                                      ) : (
                                        <span style={{ color: "var(--c-text)" }}>
                                          {String(celda.base || celda.comparar)}
                                        </span>
                                      )}
                                    </td>
                                  );
                                })}
                              </tr>

                              {expandida && (
                                <tr key={`exp-${ri}`}>
                                  <td colSpan={Object.keys(fila.columnasMerge).length + 3}
                                    className="px-4 py-3.5" style={{ background: "#EEF3FA" }}>
                                    <div className="text-sm font-bold uppercase tracking-wider font-mono mb-2.5"
                                      style={{ color: "var(--c-accent)" }}>
                                      ◈ Detalle del registro
                                    </div>
                                    <div className="grid gap-2"
                                      style={{ gridTemplateColumns: "repeat(auto-fill, minmax(300px, 1fr))" }}>
                                      {Object.keys(fila.columnasMerge).map(col => {
                                        const celda = fila.columnasMerge[col];
                                        const diff = fila.diferencias[col];
                                        return (
                                          <div key={col} className="rounded border px-3 py-2.5"
                                            style={{
                                              background: diff ? "var(--c-warning-bg)" : "white",
                                              borderColor: diff ? "var(--c-warning-border)" : "var(--c-border)",
                                            }}>
                                            <div className="text-xs font-bold uppercase tracking-wider font-mono mb-1.5"
                                              style={{ color: diff ? "var(--c-warning)" : "var(--c-text-muted)" }}>
                                              {col}
                                            </div>
                                            <div className="flex flex-col gap-0.5">
                                              <div className="text-xs font-mono">
                                                <span className="text-xs font-bold mr-1" style={{ color: "var(--c-text-muted)" }}>BASE</span>
                                                <span style={{ color: diff ? "var(--c-danger)" : "var(--c-text)" }}>
                                                  {String(celda.base) || "—"}
                                                </span>
                                              </div>
                                              {fila.estado === "coincide" && (
                                                <div className="text-xs font-mono">
                                                  <span className="text-xs font-bold mr-1" style={{ color: "var(--c-text-muted)" }}>COMP</span>
                                                  <span style={{ color: diff ? "var(--c-success)" : "var(--c-text)", fontWeight: diff ? 600 : 400 }}>
                                                    {String(celda.comparar) || "—"}
                                                  </span>
                                                </div>
                                              )}
                                            </div>
                                          </div>
                                        );
                                      })}
                                    </div>
                                  </td>
                                </tr>
                              )}
                            </>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                )}

                {/* Pagination */}
                {totalPaginas > 1 && (
                  <div className="flex items-center justify-center gap-1.5 px-4 py-3 border-t"
                    style={{ borderColor: "var(--c-border)", background: "#FAFAF8" }}>
                    {[["«", 1], ["‹", pagina - 1]].map(([lbl, p]) => (
                      <button key={lbl} onClick={() => setPagina(p)} disabled={pagina === 1}
                        className="btn-secondary" style={{ padding: "4px 10px", fontSize: 15, minWidth: 32 }}>
                        {lbl}
                      </button>
                    ))}
                    <span className="font-mono text-xs px-2" style={{ color: "var(--c-text-muted)" }}>
                      Pág. {pagina} de {totalPaginas} · {filasFiltradas.length.toLocaleString()} registros
                    </span>
                    {[["›", pagina + 1], ["»", totalPaginas]].map(([lbl, p]) => (
                      <button key={lbl} onClick={() => setPagina(p)} disabled={pagina === totalPaginas}
                        className="btn-secondary" style={{ padding: "4px 10px", fontSize: 15, minWidth: 32 }}>
                        {lbl}
                      </button>
                    ))}
                  </div>
                )}
              </div>
            </div>
          )}
        </div>

        {/* Footer */}
        <div className="mt-8 pt-4 border-t flex justify-between flex-wrap gap-2"
          style={{ borderColor: "var(--c-border)" }}>
          <span className="font-mono text-sm uppercase tracking-wider" style={{ color: "var(--c-text-light)" }}>
            Sistema de Cruzamiento Electoral · Procesamiento local · Sin envío de datos externos
          </span>
          <span className="font-mono text-sm" style={{ color: "var(--c-text-light)" }}>
            Algoritmo: distancia de Levenshtein normalizada
          </span>
        </div>
      </div>
    </div>
  );
}
