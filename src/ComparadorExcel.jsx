/* eslint-disable react-hooks/exhaustive-deps */
import { useState, useMemo, useCallback, useRef } from "react";
import * as XLSX from "xlsx";

/* ─── Motor de comparación ─────────────────────────────────────────────────── */

const normalizar = (v) =>
    String(v ?? "")
        .toLowerCase()
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .replace(/\s+/g, " ")
        .trim();

const levenshtein = (a, b) => {
    if (!a.length) return b.length;
    if (!b.length) return a.length;
    const m = Array.from({ length: a.length + 1 }, (_, i) =>
        Array.from({ length: b.length + 1 }, (_, j) => (i === 0 ? j : j === 0 ? i : 0))
    );
    for (let i = 1; i <= a.length; i++)
        for (let j = 1; j <= b.length; j++)
            m[i][j] =
                a[i - 1] === b[j - 1]
                    ? m[i - 1][j - 1]
                    : 1 + Math.min(m[i - 1][j - 1], m[i][j - 1], m[i - 1][j]);
    return m[a.length][b.length];
};

const similitud = (a, b) => {
    const na = normalizar(a), nb = normalizar(b);
    if (na === nb) return 1;
    const maxLen = Math.max(na.length, nb.length);
    return maxLen ? 1 - levenshtein(na, nb) / maxLen : 1;
};

function compararConBase(base, comparar, columnasClaves, umbral = 0.82) {
    const resultados = [];
    const coincidenciasUsadas = new Set();

    for (const filaBase of base) {
        const claveBase = columnasClaves.map((c) => String(filaBase[c] ?? "")).join("|");
        let mejorCoincidencia = null, mejorPuntaje = 0, mejorIdx = -1;
        comparar.forEach((filaComp, idx) => {
            const claveComp = columnasClaves.map((c) => String(filaComp[c] ?? "")).join("|");
            const puntaje = similitud(claveBase, claveComp);
            if (puntaje > mejorPuntaje) { mejorPuntaje = puntaje; mejorCoincidencia = filaComp; mejorIdx = idx; }
        });

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
    }

    for (let idx = 0; idx < comparar.length; idx++) {
        if (coincidenciasUsadas.has(idx)) continue;
        const filaComp = comparar[idx];
        const columnasMerge = {}, diferencias = {};
        for (const col of Object.keys(filaComp)) { columnasMerge[col] = { base: "", comparar: filaComp[col] ?? "" }; diferencias[col] = false; }
        resultados.push({ columnasMerge, diferencias, puntajeCoincidencia: 0, estado: "solo_comparar" });
    }
    return resultados;
}

/* ─── Iconos ───────────────────────────────────────────────────────────────── */
const IcoSubir = () => <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" /><polyline points="17 8 12 3 7 8" /><line x1="12" y1="3" x2="12" y2="15" /></svg>;
const IcoArchivo = () => <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" /><polyline points="14 2 14 8 20 8" /></svg>;
const IcoMas = () => <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"><line x1="12" y1="5" x2="12" y2="19" /><line x1="5" y1="12" x2="19" y2="12" /></svg>;
const IcoX = () => <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"><line x1="18" y1="6" x2="6" y2="18" /><line x1="6" y1="6" x2="18" y2="18" /></svg>;
const IcoPulse = () => <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><polyline points="22 12 18 12 15 21 9 3 6 12 2 12" /></svg>;
const IcoDescargar = () => <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" /><polyline points="7 10 12 15 17 10" /><line x1="12" y1="15" x2="12" y2="3" /></svg>;
const IcoReset = () => <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><polyline points="1 4 1 10 7 10" /><path d="M3.51 15a9 9 0 1 0 .49-4.5" /></svg>;

/* ─── Zona de carga ────────────────────────────────────────────────────────── */
const CONFIG_COLORES = {
    base: { borde: "border-emerald-500/50", bg: "bg-emerald-500/5", icono: "text-emerald-400", etiqueta: "text-emerald-400", badge: "bg-emerald-500/10 text-emerald-300 border-emerald-500/30", acento: "#4ade80" },
    segundo: { borde: "border-sky-400/50", bg: "bg-sky-400/5", icono: "text-sky-400", etiqueta: "text-sky-400", badge: "bg-sky-500/10 text-sky-300 border-sky-400/30", acento: "#38bdf8" },
    tercero: { borde: "border-violet-400/50", bg: "bg-violet-400/5", icono: "text-violet-400", etiqueta: "text-violet-400", badge: "bg-violet-500/10 text-violet-300 border-violet-400/30", acento: "#a78bfa" },
};

function ZonaCarga({ indice, datos, alCargar, alQuitar, etiqueta, colorClave }) {
    const refInput = useRef();
    const [arrastrando, setArrastrando] = useState(false);
    const c = CONFIG_COLORES[colorClave];

    const procesarArchivo = (archivo) => {
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
                <button onClick={() => alQuitar(indice)}
                    className="absolute -top-2 -right-2 z-10 w-5 h-5 rounded-full bg-zinc-700 hover:bg-rose-500 border border-zinc-600 flex items-center justify-center text-zinc-300 hover:text-white transition-all">
                    <IcoX />
                </button>
            )}
            <div
                onDragOver={(e) => { e.preventDefault(); setArrastrando(true); }}
                onDragLeave={() => setArrastrando(false)}
                onDrop={(e) => { e.preventDefault(); setArrastrando(false); procesarArchivo(e.dataTransfer.files[0]); }}
                onClick={() => refInput.current?.click()}
                className={`rounded-xl border-2 border-dashed p-5 cursor-pointer text-center transition-all duration-200 select-none min-h-32.5 flex flex-col items-center justify-center gap-2
          ${datos ? `${c.borde} ${c.bg}` : "border-zinc-700 bg-zinc-800/40 hover:border-zinc-600"}
          ${arrastrando ? "scale-[1.02]" : ""}`}
            >
                <input ref={refInput} type="file" accept=".xlsx,.xls,.csv" className="hidden"
                    onChange={(e) => procesarArchivo(e.target.files[0])} />
                <div className={datos ? c.icono : "text-zinc-500"}>{datos ? <IcoArchivo /> : <IcoSubir />}</div>
                <div className={`font-display font-bold text-sm ${datos ? c.etiqueta : "text-zinc-300"}`}>{etiqueta}</div>
                {datos ? (
                    <>
                        <div className="text-[10px] text-zinc-400 truncate max-w-45">{datos.nombre}</div>
                        <span className={`text-[10px] px-2 py-0.5 rounded-full border font-mono ${c.badge}`}>
                            {datos.filas} filas · {datos.columnas} col.
                        </span>
                    </>
                ) : (
                    <div className="text-[10px] text-zinc-600">Arrastrar o hacer clic · xlsx, xls, csv</div>
                )}
            </div>
        </div>
    );
}

/* ─── Insignia de estado ───────────────────────────────────────────────────── */
function InsigniaEstado({ estado }) {
    const mapa = {
        coincide: "bg-emerald-500/10 text-emerald-400 border-emerald-500/30 · Coincide",
        solo_base: "bg-rose-500/10 text-rose-400 border-rose-400/30 · Solo base",
        solo_comparar: "bg-sky-500/10 text-sky-400 border-sky-400/30 · Solo comp.",
    };
    const partes = (mapa[estado] || "").split(" · ");
    return <span className={`text-[9px] font-bold uppercase tracking-wider px-2 py-0.5 rounded-full border font-mono ${partes[0]}`}>{partes[1]}</span>;
}

/* ─── Paso wrapper ─────────────────────────────────────────────────────────── */
function TarjetaPaso({ numero, titulo, listo, extra, children }) {
    return (
        <div className="rounded-2xl border border-zinc-800 bg-zinc-900 overflow-hidden">
            <div className="flex items-center gap-3 px-5 py-3.5 border-b border-zinc-800 bg-zinc-800/40">
                <span className={`w-5 h-5 rounded-full text-[10px] font-bold flex items-center justify-center font-display transition-colors ${listo ? "bg-emerald-400 text-black" : "bg-zinc-700 text-zinc-400"}`}>
                    {listo ? "✓" : numero}
                </span>
                <span className="font-display font-bold text-sm uppercase tracking-widest text-zinc-200">{titulo}</span>
                {extra && <div className="ml-auto">{extra}</div>}
            </div>
            <div className="p-5">{children}</div>
        </div>
    );
}

/* ─── Componente principal ─────────────────────────────────────────────────── */
export default function ComparadorExcel() {
    const [archivos, setArchivos] = useState([null, null]);
    const [columnasClaves, setColumnasClaves] = useState([]);
    const [umbral, setUmbral] = useState(82);
    const [resultados, setResultados] = useState(null);
    const [procesando, setProcesando] = useState(false);
    const [filtroEstado, setFiltroEstado] = useState("todos");
    const [filtroDiferencias, setFiltroDiferencias] = useState(false);
    const [busqueda, setBusqueda] = useState("");
    const [pagina, setPagina] = useState(1);
    const [filaExpandida, setFilaExpandida] = useState(null);
    const [tabActiva, setTabActiva] = useState(0);
    const FILAS_POR_PAGINA = 25;

    const claveColores = ["base", "segundo", "tercero"];
    const etiquetasArchivo = ["Archivo base", "Comparar con A", "Comparar con B"];

    const todasColumnas = useMemo(() => {
        const cols = new Set();
        archivos.forEach(d => d?.json?.forEach(f => Object.keys(f).forEach(k => cols.add(k))));
        return Array.from(cols);
    }, [archivos]);

    const cargarArchivo = useCallback((indice, datos) => {
        setArchivos(prev => { const n = [...prev]; n[indice] = datos; return n; });
        setResultados(null); setColumnasClaves([]);
    }, []);

    const quitarArchivo = useCallback((indice) => {
        setArchivos(prev => { const n = [...prev]; n.splice(indice, 1); return n; });
        setResultados(null);
    }, []);

    const agregarArchivo = () => {
        if (archivos.length < 3) setArchivos(prev => [...prev, null]);
    };

    const ejecutarComparacion = useCallback(() => {
        const validos = archivos.filter(Boolean);
        if (validos.length < 2 || columnasClaves.length === 0) return;
        setProcesando(true);
        setTimeout(() => {
            const todosResultados = [];
            for (let i = 1; i < archivos.length; i++) {
                if (!archivos[i]) continue;
                const filas = compararConBase(archivos[0].json, archivos[i].json, columnasClaves, umbral / 100);
                todosResultados.push({ indice: i, nombreArchivo: archivos[i].nombre, filas });
            }
            setResultados(todosResultados);
            setTabActiva(0); setPagina(1); setProcesando(false);
        }, 60);
    }, [archivos, columnasClaves, umbral]);

    const resultadoActual = resultados?.[tabActiva]?.filas ?? [];

    const estadisticas = useMemo(() => {
        if (!resultados) return null;
        return resultados.map(r => ({
            nombreArchivo: r.nombreArchivo,
            total: r.filas.length,
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
            return Object.values(f.columnasMerge).some(v => String(v.base).toLowerCase().includes(b) || String(v.comparar).toLowerCase().includes(b));
        }
        return true;
    }), [resultadoActual, filtroEstado, filtroDiferencias, busqueda]);

    const filasPaginadas = useMemo(() => {
        const inicio = (pagina - 1) * FILAS_POR_PAGINA;
        return filasFiltradas.slice(inicio, inicio + FILAS_POR_PAGINA);
    }, [filasFiltradas, pagina]);

    const totalPaginas = Math.ceil(filasFiltradas.length / FILAS_POR_PAGINA);

    const exportarResultado = () => {
        if (!resultados) return;
        resultados.forEach(r => {
            if (!r.filas.length) return;
            const columnas = Object.keys(r.filas[0].columnasMerge);
            const filas = r.filas.map(f => {
                const fila = { _estado: f.estado, _similitud: Math.round(f.puntajeCoincidencia * 100) + "%" };
                columnas.forEach(c => {
                    fila[`${c}_base`] = f.columnasMerge[c]?.base ?? "";
                    fila[`${c}_comparar`] = f.columnasMerge[c]?.comparar ?? "";
                    fila[`${c}_diferente`] = f.diferencias[c] ? "SÍ" : "";
                });
                return fila;
            });
            const hoja = XLSX.utils.json_to_sheet(filas);
            const libro = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(libro, hoja, "Comparación");
            XLSX.writeFile(libro, `comparacion_${r.nombreArchivo}`);
        });
    };

    const puedoComparar = archivos.filter(Boolean).length >= 2 && columnasClaves.length > 0;
    const cantArchivos = archivos.length;

    return (
        <div className="min-h-screen bg-zinc-950 text-zinc-100" style={{ fontFamily: "'DM Mono', monospace" }}>
            <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Mono:wght@400;500&family=Bricolage+Grotesque:wght@600;700;800&display=swap');
        .font-display { font-family: 'Bricolage Grotesque', sans-serif; }
        input[type=range]{-webkit-appearance:none;height:4px;border-radius:2px;outline:none;background:#3f3f46}
        input[type=range]::-webkit-slider-thumb{-webkit-appearance:none;width:16px;height:16px;border-radius:50%;background:#4ade80;cursor:pointer;border:2px solid #09090b}
        .spin{animation:spin .7s linear infinite}
        @keyframes spin{to{transform:rotate(360deg)}}
        ::-webkit-scrollbar{width:5px;height:5px}
        ::-webkit-scrollbar-track{background:#18181b}
        ::-webkit-scrollbar-thumb{background:#3f3f46;border-radius:3px}
      `}</style>

            <div className="max-w-6xl mx-auto px-4 py-10 pb-24 space-y-5">

                {/* Encabezado */}
                <div className="flex items-center justify-between flex-wrap gap-4 mb-6">
                    <div className="flex items-center gap-3">
                        <div className="w-10 h-10 rounded-xl bg-linear-to-br from-emerald-400 to-emerald-600 flex items-center justify-center text-black shrink-0">
                            <IcoPulse />
                        </div>
                        <div>
                            <h1 className="font-display text-xl font-extrabold tracking-tight">Comparador Excel</h1>
                            <p className="text-[11px] text-zinc-500">Cruzamiento inteligente · hasta 3 archivos</p>
                        </div>
                    </div>
                    {resultados && (
                        <button onClick={exportarResultado}
                            className="flex items-center gap-2 px-4 py-2 rounded-lg border border-zinc-700 bg-zinc-800 hover:border-zinc-500 text-zinc-300 hover:text-white text-xs transition-all cursor-pointer">
                            <IcoDescargar /> Exportar resultados
                        </button>
                    )}
                </div>

                {/* ── Paso 1 ── */}
                <TarjetaPaso numero="1" titulo="Cargar archivos"
                    listo={archivos.filter(Boolean).length >= 2}
                    extra={<span className="text-[10px] text-zinc-500 font-mono">{archivos.filter(Boolean).length}/{cantArchivos} archivos</span>}>
                    <div className={`grid gap-4 ${cantArchivos === 3 ? "grid-cols-1 sm:grid-cols-3" : "grid-cols-1 sm:grid-cols-2"}`}>
                        {archivos.map((datos, i) => (
                            <ZonaCarga key={i} indice={i} datos={datos} alCargar={cargarArchivo}
                                alQuitar={i >= 2 ? quitarArchivo : null}
                                etiqueta={etiquetasArchivo[i] ?? `Archivo ${i + 1}`}
                                colorClave={claveColores[i]} />
                        ))}
                        {cantArchivos < 3 && (
                            <button onClick={agregarArchivo}
                                className="rounded-xl border-2 border-dashed border-zinc-700 hover:border-violet-400/60 hover:bg-violet-400/5 p-5 flex flex-col items-center justify-center gap-2 cursor-pointer text-zinc-600 hover:text-violet-400 transition-all min-h-32.5">
                                <div className="w-7 h-7 rounded-full border-2 border-current flex items-center justify-center">
                                    <IcoMas />
                                </div>
                                <span className="font-display font-bold text-xs">Agregar tercer archivo</span>
                                <span className="text-[10px] text-zinc-700">Comparación adicional</span>
                            </button>
                        )}
                    </div>
                </TarjetaPaso>

                {/* ── Paso 2 ── */}
                <TarjetaPaso numero="2" titulo="Columnas clave para cruzar"
                    listo={columnasClaves.length > 0}
                    extra={columnasClaves.length > 0 && (
                        <span className="text-[10px] bg-emerald-500/10 text-emerald-400 border border-emerald-500/30 px-2 py-0.5 rounded-full font-mono">
                            {columnasClaves.length} seleccionada{columnasClaves.length !== 1 ? "s" : ""}
                        </span>
                    )}>
                    {todasColumnas.length > 0 ? (
                        <>
                            <p className="text-[11px] text-zinc-500 mb-3">Elegí las columnas que identifican de forma única cada registro.</p>
                            <div className="flex flex-wrap gap-2">
                                {todasColumnas.map(col => {
                                    const activa = columnasClaves.includes(col);
                                    return (
                                        <button key={col} onClick={() => setColumnasClaves(prev => activa ? prev.filter(x => x !== col) : [...prev, col])}
                                            className={`cursor-pointer px-3 py-1.5 rounded-lg text-[11px] border transition-all
                        ${activa ? "bg-emerald-500/15 border-emerald-500/50 text-emerald-300" : "bg-zinc-800 border-zinc-700 text-zinc-400 hover:border-zinc-500 hover:text-zinc-200"}`}>
                                            {activa && <span className="mr-1 text-emerald-400">✓</span>}
                                            {col}
                                        </button>
                                    );
                                })}
                            </div>
                            <div className="mt-5 flex items-center gap-4 flex-wrap">
                                <span className="text-[11px] text-zinc-500 shrink-0">Umbral de similitud</span>
                                <input type="range" min={50} max={100} value={umbral} onChange={e => setUmbral(Number(e.target.value))} className="flex-1 min-w-25" />
                                <span className="text-sm font-bold text-emerald-400 w-10 text-center">{umbral}%</span>
                                <span className="text-[10px] text-zinc-600 hidden sm:block">
                                    {umbral >= 95 ? "Exacto" : umbral >= 80 ? "Fuzzy moderado" : "Fuzzy amplio"}
                                </span>
                            </div>
                        </>
                    ) : (
                        <p className="text-[11px] text-zinc-500">Cargá al menos 2 archivos para ver las columnas disponibles.</p>
                    )}
                </TarjetaPaso>

                {/* ── Paso 3 ── */}
                <TarjetaPaso numero="3" titulo="Ejecutar comparación" listo={!!resultados}>
                    <div className="flex gap-3 flex-wrap items-center">
                        <button onClick={ejecutarComparacion} disabled={!puedoComparar || procesando}
                            className="flex items-center gap-2 px-5 py-2.5 rounded-lg bg-emerald-400 hover:bg-emerald-300 active:scale-95 text-black text-xs font-display font-bold uppercase tracking-wider transition-all disabled:opacity-30 disabled:cursor-not-allowed cursor-pointer">
                            {procesando
                                ? <><span className="spin w-3.5 h-3.5 border-2 border-black/30 border-t-black rounded-full" />Procesando...</>
                                : <><IcoPulse />Comparar archivos</>}
                        </button>
                        {resultados && (
                            <button onClick={() => { setResultados(null); setArchivos([null, null]); setColumnasClaves([]); }}
                                className="cursor-pointer flex items-center gap-2 px-4 py-2.5 rounded-lg border border-zinc-700 bg-zinc-800 hover:border-zinc-500 text-zinc-300 text-xs transition-all">
                                <IcoReset />Reiniciar
                            </button>
                        )}
                        {!puedoComparar && !procesando && (
                            <span className="text-[10px] text-zinc-600 font-mono">
                                {archivos.filter(Boolean).length < 2 ? "• Cargá al menos 2 archivos. " : ""}
                                {columnasClaves.length === 0 ? "• Seleccioná columnas clave." : ""}
                            </span>
                        )}
                    </div>
                </TarjetaPaso>

                {/* ── Spinner ── */}
                {procesando && (
                    <div className="rounded-2xl border border-zinc-800 bg-zinc-900 py-14 flex items-center justify-center gap-3 text-zinc-400 text-sm">
                        <span className="spin w-5 h-5 border-2 border-zinc-700 border-t-emerald-400 rounded-full" />
                        Analizando y cruzando datos...
                    </div>
                )}

                {/* ── Resultados ── */}
                {resultados && !procesando && (
                    <>
                        {/* Estadísticas */}
                        <div className="grid gap-4">
                            {estadisticas.map((est, i) => (
                                <div key={i} className="rounded-2xl border border-zinc-800 bg-zinc-900 overflow-hidden">
                                    <div className="flex items-center gap-2 px-5 py-3 border-b border-zinc-800 bg-zinc-800/40">
                                        <span className="w-2 h-2 rounded-full bg-sky-400" />
                                        <span className="font-display font-bold text-xs uppercase tracking-widest text-zinc-300">
                                            Base vs {est.nombreArchivo}
                                        </span>
                                    </div>
                                    <div className="grid grid-cols-2 sm:grid-cols-5 divide-zinc-800" style={{ borderLeft: "none" }}>
                                        {[
                                            ["Total filas", est.total, "text-zinc-100"],
                                            ["Coincidencias", est.coinciden, "text-emerald-400"],
                                            ["Con diferencias", est.conDiferencias, "text-amber-400"],
                                            ["Solo en base", est.soloBase, "text-rose-400"],
                                            ["Solo en comparar", est.soloComparar, "text-sky-400"],
                                        ].map(([lbl, val, cls], j) => (
                                            <div key={j} className={`p-4 text-center ${j > 0 ? "border-l border-zinc-800" : ""}`}>
                                                <div className={`font-display text-3xl font-extrabold leading-none ${cls}`}>{val}</div>
                                                <div className="text-[9px] text-zinc-500 mt-1.5 uppercase tracking-wider">{lbl}</div>
                                            </div>
                                        ))}
                                    </div>
                                </div>
                            ))}
                        </div>

                        {/* Tabs si hay múltiples comparaciones */}
                        {resultados.length > 1 && (
                            <div className="flex gap-2">
                                {resultados.map((r, i) => (
                                    <button key={i} onClick={() => { setTabActiva(i); setPagina(1); setFilaExpandida(null); }}
                                        className={`cursor-pointer px-3 py-1.5 rounded-lg text-[11px] border transition-all font-mono ${tabActiva === i ? "bg-zinc-700 border-zinc-500 text-white" : "bg-zinc-800 border-zinc-700 text-zinc-400 hover:text-zinc-200"
                                            }`}>
                                        Base vs {r.nombreArchivo}
                                    </button>
                                ))}
                            </div>
                        )}

                        {/* Tabla */}
                        <div className="rounded-2xl border border-zinc-800 bg-zinc-900 overflow-hidden">
                            <div className="flex items-center gap-3 px-5 py-3.5 border-b border-zinc-800 bg-zinc-800/40">
                                <span className="w-5 h-5 rounded-full bg-emerald-400 text-black text-[10px] font-bold flex items-center justify-center font-display">✓</span>
                                <span className="font-display font-bold text-sm uppercase tracking-widest text-zinc-200">Resultados</span>
                                <span className="ml-auto text-[10px] bg-zinc-800 border border-zinc-700 text-zinc-400 px-2 py-0.5 rounded-full font-mono">
                                    {filasFiltradas.length} filas
                                </span>
                            </div>

                            {/* Filtros */}
                            <div className="px-5 py-3 border-b border-zinc-800 flex flex-wrap gap-2 items-center">
                                {[["todos", "Todos"], ["coincide", "Coinciden"], ["solo_base", "Solo base"], ["solo_comparar", "Solo comp."]].map(([val, lbl]) => (
                                    <button key={val} onClick={() => { setFiltroEstado(val); setPagina(1); }}
                                        className={`cursor-pointer px-3 py-1 rounded-md text-[11px] border transition-all font-mono ${filtroEstado === val ? "bg-zinc-700 border-zinc-500 text-white" : "bg-zinc-800 border-zinc-700 text-zinc-400 hover:text-zinc-200"}`}>
                                        {lbl}
                                    </button>
                                ))}
                                <div className="w-px h-4 bg-zinc-700 mx-0.5" />
                                <button onClick={() => { setFiltroDiferencias(v => !v); setPagina(1); }}
                                    className={`cursor-pointer px-3 py-1 rounded-md text-[11px] border transition-all font-mono ${filtroDiferencias ? "bg-amber-500/10 border-amber-400/40 text-amber-300" : "bg-zinc-800 border-zinc-700 text-zinc-400 hover:text-zinc-200"}`}>
                                    ⚡ Solo diferencias
                                </button>
                                <div className="w-px h-4 bg-zinc-700 mx-0.5" />
                                <input value={busqueda} onChange={e => { setBusqueda(e.target.value); setPagina(1); }}
                                    placeholder="Buscar en resultados..."
                                    className="flex-1 min-w-35 bg-zinc-800 border border-zinc-700 focus:border-zinc-500 rounded-lg px-3 py-1 text-[11px] text-zinc-200 placeholder-zinc-600 outline-none transition-colors font-mono" />
                            </div>

                            {/* Tabla de datos */}
                            {filasPaginadas.length === 0 ? (
                                <div className="py-16 text-center text-zinc-600 text-sm">Sin resultados con los filtros actuales.</div>
                            ) : (
                                <div className="overflow-x-auto">
                                    <table className="w-full text-[11px]">
                                        <thead>
                                            <tr className="bg-zinc-800/50 border-b border-zinc-800">
                                                <th className="w-7 px-3 py-2.5" />
                                                <th className="px-3 py-2.5 text-left font-display text-[9px] uppercase tracking-widest text-zinc-500 whitespace-nowrap">Estado</th>
                                                <th className="px-3 py-2.5 text-left font-display text-[9px] uppercase tracking-widest text-zinc-500 whitespace-nowrap">Similitud</th>
                                                {filasPaginadas[0] && Object.keys(filasPaginadas[0].columnasMerge).map(col => (
                                                    <th key={col} className="px-3 py-2.5 text-left font-display text-[9px] uppercase tracking-widest text-zinc-500 whitespace-nowrap">{col}</th>
                                                ))}
                                            </tr>
                                        </thead>
                                        <tbody className="divide-y divide-zinc-800/50">
                                            {filasPaginadas.map((fila, ri) => {
                                                const idxGlobal = (pagina - 1) * FILAS_POR_PAGINA + ri;
                                                const expandida = filaExpandida === idxGlobal;
                                                return (
                                                    <>
                                                        <tr key={ri} onClick={() => setFilaExpandida(expandida ? null : idxGlobal)}
                                                            className={`cursor-pointer transition-colors ${expandida ? "bg-emerald-500/5" : "hover:bg-zinc-800/30"}`}>
                                                            <td className="px-3 py-2.5 text-zinc-600 text-[9px] text-center">{expandida ? "▼" : "▶"}</td>
                                                            <td className="px-3 py-2.5 whitespace-nowrap"><InsigniaEstado estado={fila.estado} /></td>
                                                            <td className="px-3 py-2.5 whitespace-nowrap">
                                                                {fila.estado === "coincide" ? (
                                                                    <span className={`text-[9px] px-2 py-0.5 rounded-full border font-mono ${fila.puntajeCoincidencia >= 0.95 ? "bg-emerald-500/10 text-emerald-400 border-emerald-500/30" : fila.puntajeCoincidencia >= 0.82 ? "bg-amber-500/10 text-amber-400 border-amber-400/30" : "bg-rose-500/10 text-rose-400 border-rose-400/30"}`}>
                                                                        {Math.round(fila.puntajeCoincidencia * 100)}%
                                                                    </span>
                                                                ) : <span className="text-zinc-600">—</span>}
                                                            </td>
                                                            {Object.keys(fila.columnasMerge).map((col, ci) => {
                                                                const celda = fila.columnasMerge[col];
                                                                const diferente = fila.diferencias[col];
                                                                return (
                                                                    <td key={ci} className={`px-3 py-2.5 whitespace-nowrap ${diferente ? "bg-amber-500/5" : ""}`}>
                                                                        {diferente ? (
                                                                            <div className="flex flex-col gap-0.5">
                                                                                <span className="text-rose-400 line-through text-[9px] opacity-75">{String(celda.base)}</span>
                                                                                <span className="text-emerald-400 text-[10px] font-medium">{String(celda.comparar)}</span>
                                                                            </div>
                                                                        ) : (
                                                                            <span className="text-zinc-300">{String(celda.base || celda.comparar)}</span>
                                                                        )}
                                                                    </td>
                                                                );
                                                            })}
                                                        </tr>
                                                        {expandida && (
                                                            <tr key={`exp-${ri}`}>
                                                                <td colSpan={Object.keys(fila.columnasMerge).length + 3} className="bg-zinc-800/20 px-5 py-4">
                                                                    <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 gap-2.5">
                                                                        {Object.keys(fila.columnasMerge).map(col => {
                                                                            const celda = fila.columnasMerge[col];
                                                                            const diferente = fila.diferencias[col];
                                                                            return (
                                                                                <div key={col} className={`rounded-lg p-3 border ${diferente ? "border-amber-500/30 bg-amber-500/5" : "border-zinc-700/50 bg-zinc-800/40"}`}>
                                                                                    <div className={`text-[8px] uppercase tracking-wider mb-1.5 font-display font-bold ${diferente ? "text-amber-400" : "text-zinc-500"}`}>{col}</div>
                                                                                    <div className="space-y-1">
                                                                                        <div className="text-[10px]"><span className="text-zinc-600 mr-1 font-bold text-[8px]">BASE</span><span className={diferente ? "text-rose-300" : "text-zinc-300"}>{String(celda.base) || "—"}</span></div>
                                                                                        {fila.estado === "coincide" && (
                                                                                            <div className="text-[10px]"><span className="text-zinc-600 mr-1 font-bold text-[8px]">COMP</span><span className={diferente ? "text-emerald-300 font-medium" : "text-zinc-300"}>{String(celda.comparar) || "—"}</span></div>
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

                            {/* Paginación */}
                            {totalPaginas > 1 && (
                                <div className="flex items-center justify-center gap-2 px-5 py-4 border-t border-zinc-800">
                                    {[["«", 1], ["‹", pagina - 1]].map(([lbl, p]) => (
                                        <button key={lbl} onClick={() => setPagina(p)} disabled={pagina === 1}
                                            className="px-2.5 py-1 rounded-md border border-zinc-700 bg-zinc-800 text-zinc-400 hover:text-zinc-200 hover:border-zinc-500 text-xs disabled:opacity-30 disabled:cursor-not-allowed transition-all cursor-pointer">
                                            {lbl}
                                        </button>
                                    ))}
                                    <span className="text-[10px] text-zinc-500 font-mono px-2">Página {pagina} de {totalPaginas} · {filasFiltradas.length} filas</span>
                                    {[["›", pagina + 1], ["»", totalPaginas]].map(([lbl, p]) => (
                                        <button key={lbl} onClick={() => setPagina(p)} disabled={pagina === totalPaginas}
                                            className="px-2.5 py-1 rounded-md border border-zinc-700 bg-zinc-800 text-zinc-400 hover:text-zinc-200 hover:border-zinc-500 text-xs disabled:opacity-30 disabled:cursor-not-allowed transition-all cursor-pointer">
                                            {lbl}
                                        </button>
                                    ))}
                                </div>
                            )}
                        </div>
                    </>
                )}

            </div>
        </div>
    );
}
