"use client";

import { useEffect, useState, useRef } from "react";
import { db } from "@/lib/firebaseClient";
import {
  collection,
  doc,
  onSnapshot,
  orderBy,
  query,
  limit,
  where,
  getDocs,
} from "firebase/firestore";
import { generarReportePorArea, generarReporteConsolidadoGlobal } from "@/lib/exportarReporte";
import { generarReporteExcelPorArea } from "@/lib/exportarExcel";
import { LOGO_PUNO_BASE64 } from "@/lib/logoPuno";
import {
  getAuth,
  GoogleAuthProvider,
  signInWithPopup,
  onAuthStateChanged,
} from "firebase/auth";

const auth = getAuth(db.app);

const COLLAPSE_STORAGE_KEY = "acocollo_i2_grupos_colapsados";

const ESTADO_COLOR = {
  completa: "#e5b80b",
  incompleta: "#e67e22",
  vacia: "#c0392b",
};

const AREA_COLORS = [
  "#e5b80b",
  "#e67e22",
  "#c0392b",
  "#457b9d",
  "#1d3557",
  "#f1c40f",
  "#e5b80b",
];

function colorForArea(area) {
  let hash = 0;
  for (let i = 0; i < area.length; i++) hash = area.charCodeAt(i) + ((hash << 5) - hash);
  return AREA_COLORS[Math.abs(hash) % AREA_COLORS.length];
}

const ESTADO_OPTIONS = [
  { value: "pendientes", label: "Pendientes", color: "#e5b80b" },
  { value: "incompleta", label: "Incompletas", color: "#e67e22" },
  { value: "vacia", label: "Vacías", color: "#c0392b" },
  { value: "completa", label: "Completas", color: "#f1c40f" },
  { value: "todas", label: "Todas", color: "#457b9d" },
];

const EVENTO_LABEL = {
  archivo_subido: "subió",
  archivo_reemplazado: "reemplazó",
  archivo_borrado: "borró",
  carpeta_creada: "creó la carpeta",
  carpeta_borrada: "borró la carpeta",
  carpeta_movida: "movió la carpeta",
  carpeta_marcada_completa: "marcó como completa",
  carpeta_marcada_incompleta: "marcó como incompleta",
  carpeta_desmarcada: "revirtió la marca de",
};

const EVENTO_COLOR = {
  archivo_subido: "#e5b80b",
  archivo_reemplazado: "#e67e22",
  archivo_borrado: "#c0392b",
  carpeta_creada: "#264653",
  carpeta_borrada: "#c0392b",
  carpeta_movida: "#e67e22",
  carpeta_marcada_completa: "#e5b80b",
  carpeta_marcada_incompleta: "#e67e22",
  carpeta_desmarcada: "#457b9d",
};

const EVENTO_ICONO = {
  archivo_subido: "↑",
  archivo_reemplazado: "⟲",
  archivo_borrado: "✕",
  carpeta_creada: "+",
  carpeta_borrada: "✕",
  carpeta_movida: "⇄",
  carpeta_marcada_completa: "✓",
  carpeta_marcada_incompleta: "⚠",
  carpeta_desmarcada: "↺",
};

function fechaLimaISO(fecha) {
  const partes = new Intl.DateTimeFormat("en-CA", {
    timeZone: "America/Lima",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).formatToParts(fecha);
  const obj = {};
  for (const p of partes) obj[p.type] = p.value;
  return `${obj.year}-${obj.month}-${obj.day}`;
}

function formatearFechaLarga(fechaEntrada) {
  const fecha = typeof fechaEntrada === "string" ? new Date(fechaEntrada + "T12:00:00") : fechaEntrada;
  if (!fecha || isNaN(fecha.getTime())) return String(fechaEntrada);
  const texto = fecha.toLocaleDateString("es-PE", { weekday: "long", day: "numeric", month: "long", year: "numeric" });
  return texto.charAt(0).toUpperCase() + texto.slice(1);
}

function tiempoRelativo(date) {
  if (!date) return "...";
  const diffMs = Date.now() - date.getTime();
  const diffMin = Math.floor(diffMs / 60000);
  if (diffMin < 1) return "justo ahora";
  if (diffMin < 60) return `hace ${diffMin} min`;
  const diffH = Math.floor(diffMin / 60);
  if (diffH < 24) return `hace ${diffH} h`;
  const diffD = Math.floor(diffH / 24);
  if (diffD < 7) return `hace ${diffD} d`;
  return date.toLocaleDateString("es-PE");
}

export default function Dashboard() {
  const [resumen, setResumen] = useState(null);
  const [carpetas, setCarpetas] = useState([]);
  const [eventos, setEventos] = useState([]);
  const [filtroArea, setFiltroArea] = useState("Todas");
  const [filtroEstado, setFiltroEstado] = useState("pendientes");
  const [tipoExportacion, setTipoExportacion] = useState("todas");
  const [sincronizando, setSincronizando] = useState(false);
  const [mensajeSync, setMensajeSync] = useState(null);
  const [busqueda, setBusqueda] = useState("");
  const [busquedaMarcadas, setBusquedaMarcadas] = useState(""); // 🔍 Buscador interactivo en modal de marcadas
  const [colapsados, setColapsados] = useState({});
  const [colapsoListo, setColapsoListo] = useState(false);
  const [exportandoArea, setExportandoArea] = useState(null);
  const [exportandoExcelArea, setExportandoExcelArea] = useState(null);
  const [exportandoGlobal, setExportandoGlobal] = useState(false);
  const [modoPresentacion, setModoPresentacion] = useState(false);
  const [historial, setHistorial] = useState([]);
  const [actividadPorDia, setActividadPorDia] = useState({});
  const [marcandoId, setMarcandoId] = useState(null);
  const [mostrarMarcadas, setMostrarMarcadas] = useState(false);
  const [usuarioGoogle, setUsuarioGoogle] = useState(null);
  const [rangoDiasHeatmap, setRangoDiasHeatmap] = useState(84);

  useEffect(() => {
    const unsub = onAuthStateChanged(auth, (user) => {
      setUsuarioGoogle(user ? { email: user.email, displayName: user.displayName } : null);
    });
    return () => unsub();
  }, []);

  useEffect(() => {
    try {
      const guardado = localStorage.getItem(COLLAPSE_STORAGE_KEY);
      if (guardado) setColapsados(JSON.parse(guardado));
    } catch {}
    setColapsoListo(true);
  }, []);

  useEffect(() => {
    if (!colapsoListo) return;
    try {
      localStorage.setItem(COLLAPSE_STORAGE_KEY, JSON.stringify(colapsados));
    } catch {}
  }, [colapsados, colapsoListo]);

  function toggleGrupo(key) {
    setColapsados((prev) => ({ ...prev, [key]: !prev[key] }));
  }

  function filtrarCarpetasParaExportar(listaCarpetas, tipoForzado) {
    const filtroUsado = tipoForzado || tipoExportacion;
    if (filtroUsado === "completas") {
      return listaCarpetas.filter((c) => c.estado === "completa");
    }
    if (filtroUsado === "incompletas") {
      return listaCarpetas.filter((c) => c.estado === "incompleta");
    }
    if (filtroUsado === "vacias") {
      return listaCarpetas.filter((c) => c.estado === "vacia");
    }
    if (filtroUsado === "incompletas_vacias") {
      return listaCarpetas.filter((c) => c.estado === "incompleta" || c.estado === "vacia");
    }
    return listaCarpetas;
  }

  function handleExportarArea(areaNombre, carpetasDelArea, tipoForzado) {
    setExportandoArea(areaNombre);
    try {
      const usuarioFirma = usuarioGoogle?.email || usuarioGoogle?.displayName || "Sistema Acocollo I-2";
      const listaFiltrada = filtrarCarpetasParaExportar(carpetasDelArea, tipoForzado);
      generarReportePorArea(areaNombre, listaFiltrada, { usuarioFirma });
    } finally {
      setExportandoArea(null);
    }
  }

  async function handleExportarGlobal(tipoForzado) {
    setExportandoGlobal(true);
    try {
      const usuarioFirma = usuarioGoogle?.email || usuarioGoogle?.displayName || "Sistema Acocollo I-2";
      const listaFiltrada = filtrarCarpetasParaExportar(carpetas, tipoForzado);
      generarReporteConsolidadoGlobal(listaFiltrada, { usuarioFirma });
    } catch (err) {
      alert(`No se pudo generar el reporte consolidado: ${err.message}`);
    } finally {
      setExportandoGlobal(false);
    }
  }

  async function handleExportarExcelArea(areaNombre, carpetasDelArea, tipoForzado) {
    setExportandoExcelArea(areaNombre);
    try {
      const timeoutPromise = new Promise((_, reject) =>
        setTimeout(() => reject(new Error("Tiempo de espera agotado al generar el Excel")), 10000)
      );
      const listaFiltrada = filtrarCarpetasParaExportar(carpetasDelArea, tipoForzado);
      await Promise.race([
        generarReporteExcelPorArea(areaNombre, listaFiltrada),
        timeoutPromise,
      ]);
    } catch (err) {
      alert(`No se pudo generar el Excel: ${err.message}`);
    } finally {
      setExportandoExcelArea(null);
    }
  }

  // estado: "completa" | "incompleta" | null (null = revertir al cálculo automático)
  async function handleMarcarCompleta(folderId, estado, folderName, folderRuta) {
    let user = auth.currentUser;
    if (!user) {
      try {
        const cred = await signInWithPopup(auth, new GoogleAuthProvider());
        user = cred.user;
      } catch (err) {
        alert(`Necesitas iniciar sesión con Google para actualizar carpetas. ${err.message || ""}`);
        return;
      }
    }

    let motivo = "";
    if (estado === "completa") {
      motivo = window.prompt("¿Por qué se marca como completa manualmente?", "");
      if (motivo === null) return;
    } else if (estado === "incompleta") {
      motivo = window.prompt("¿Por qué se marca como incompleta manualmente?", "");
      if (motivo === null) return;
    } else {
      motivo = window.prompt("¿Por qué se revierte esta marca manual? (opcional)", "");
      if (motivo === null) return;
    }

    setMarcandoId(folderId);
    try {
      const idToken = await user.getIdToken();
      const res = await fetch("/api/marcar-completo", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ folderId, estado, motivo, idToken, folderName, folderRuta }),
      });
      if (!res.ok) {
        const data = await res.json().catch(() => ({}));
        alert(`No se pudo actualizar: ${data.error || res.statusText}`);
      }
    } catch (err) {
      alert(`Error de conexión: ${err.message}`);
    } finally {
      setMarcandoId(null);
    }
  }

  async function handleSync() {
    setSincronizando(true);
    setMensajeSync(null);
    try {
      const res = await fetch("/api/manual-sync", { method: "POST" });
      const data = await res.json();
      if (data.ok) {
        setMensajeSync({ tipo: "ok", texto: `Listo — ${data.eventos} eventos nuevos detectados` });
      } else {
        setMensajeSync({ tipo: "error", texto: `Error: ${data.error || "desconocido"}` });
      }
    } catch (err) {
      setMensajeSync({ tipo: "error", texto: "Error de conexión al sincronizar" });
    } finally {
      setSincronizando(false);
    }
  }

  useEffect(() => {
    const unsubResumen = onSnapshot(doc(db, "_meta", "resumen"), (snap) => {
      if (snap.exists()) setResumen(snap.data());
    });

    const unsubCarpetas = onSnapshot(collection(db, "carpetas"), (snap) => {
      setCarpetas(snap.docs.map((d) => ({ id: d.id, ...d.data() })));
    });

    const eventosQuery = query(collection(db, "eventos"), orderBy("timestamp", "desc"), limit(50));
    const unsubEventos = onSnapshot(eventosQuery, (snap) => {
      setEventos(snap.docs.map((d) => ({ id: d.id, ...d.data() })));
    });

    const unsubActividadPorDia = onSnapshot(doc(db, "_meta", "actividadPorDia"), (snap) => {
      setActividadPorDia(snap.exists() ? snap.data() : {});
    });

    const historialQuery = query(collection(db, "historial"), orderBy("fecha", "asc"), limit(90));
    const unsubHistorial = onSnapshot(historialQuery, (snap) => {
      setHistorial(snap.docs.map((d) => ({ id: d.id, ...d.data() })));
    });

    return () => {
      unsubResumen();
      unsubCarpetas();
      unsubEventos();
      unsubActividadPorDia();
      unsubHistorial();
    };
  }, []);

  const pct = resumen && resumen.totalFinales
    ? Math.round((resumen.completas / resumen.totalFinales) * 100)
    : 0;

  const areas = Array.from(new Set(carpetas.map((c) => c.area || "Sin área"))).sort();

  const carpetasPorArea = {};
  for (const c of carpetas) {
    const a = c.area || "Sin área";
    if (!carpetasPorArea[a]) carpetasPorArea[a] = [];
    carpetasPorArea[a].push(c);
  }

  const carpetasForzadas = carpetas
    .filter((c) => c.forzada)
    .sort((a, b) => new Date(b.marcadoEn || 0) - new Date(a.marcadoEn || 0));

  // Filtrado en tiempo real de carpetas marcadas en el modal
  const carpetasForzadasFiltradas = busquedaMarcadas.trim()
    ? carpetasForzadas.filter((c) =>
        (c.nombre || "").toLowerCase().includes(busquedaMarcadas.trim().toLowerCase()) ||
        (c.ruta || "").toLowerCase().includes(busquedaMarcadas.trim().toLowerCase())
      )
    : carpetasForzadas;

  const areaStats = {};
  const especialidadPorArea = {};
  for (const c of carpetas) {
    const a = c.area || "Sin área";
    if (!areaStats[a])
      areaStats[a] = { total: 0, completas: 0, incompletas: 0, vacias: 0, archivosNecesarios: 0, archivosCompletados: 0 };
    areaStats[a].total++;
    if (c.estado === "completa") areaStats[a].completas++;
    if (c.estado === "incompleta") areaStats[a].incompletas++;
    if (c.estado === "vacia") areaStats[a].vacias++;
    areaStats[a].archivosNecesarios += c.archivosNecesarios || 0;
    areaStats[a].archivosCompletados += c.archivosCompletados || 0;

    const partesRuta = (c.ruta || c.nombre || "").split(" / ").filter(Boolean);
    const especialidad = partesRuta.length > 1 ? partesRuta[1] : "(raíz)";
    if (!especialidadPorArea[a]) especialidadPorArea[a] = {};
    if (!especialidadPorArea[a][especialidad])
      especialidadPorArea[a][especialidad] = { total: 0, completas: 0, incompletas: 0, vacias: 0, archivosNecesarios: 0, archivosCompletados: 0 };
    
    especialidadPorArea[a][especialidad].total++;
    if (c.estado === "completa") especialidadPorArea[a][especialidad].completas++;
    if (c.estado === "incompleta") especialidadPorArea[a][especialidad].incompletas++;
    if (c.estado === "vacia") especialidadPorArea[a][especialidad].vacias++;
    especialidadPorArea[a][especialidad].archivosNecesarios += c.archivosNecesarios || 0;
    especialidadPorArea[a][especialidad].archivosCompletados += c.archivosCompletados || 0;
  }

  let listaBase = carpetas;
  if (filtroEstado === "pendientes") {
    listaBase = carpetas.filter((c) => c.estado !== "completa");
  } else if (filtroEstado !== "todas") {
    listaBase = carpetas.filter((c) => c.estado === filtroEstado);
  }

  let visibles = listaBase.sort((a, b) =>
    (a.ruta || "").localeCompare(b.ruta || "", undefined, { numeric: true, sensitivity: "base" })
  );

  if (filtroArea !== "Todas") {
    visibles = visibles.filter((c) => (c.area || "Sin área") === filtroArea);
  }

  if (busqueda.trim()) {
    const q = busqueda.trim().toLowerCase();
    visibles = visibles.filter((c) =>
      (c.nombre || "").toLowerCase().includes(q) || (c.ruta || "").toLowerCase().includes(q)
    );
  }

  const ESTADO_FILTRO_LABEL = {
    pendientes: "Pendientes (incompletas + vacías)",
    incompleta: "Solo incompletas",
    vacia: "Solo vacías",
    completa: "Solo completas",
    todas: "Todas las carpetas",
  };
  const areaLabel = filtroArea !== "Todas" ? ` · ${filtroArea}` : "";

  return (
    <div className="acocollo-fondo-animado" style={{ minHeight: "100vh", width: "100%", paddingBottom: 60 }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

        .acocollo-fondo-animado, .acocollo-fondo-animado * {
          font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
        }
        .acocollo-fondo-animado {
          background: linear-gradient(
            -45deg,
            #0c1015,
            #141c24,
            #1f2d3d,
            #2b3e55,
            #1f2d3d,
            #141c24,
            #0c1015
          );
          background-size: 500% 500%;
          animation: acocolloGradiente 5s ease infinite;
        }
        @keyframes acocolloGradiente {
          0%   { background-position: 0% 50%; }
          50%  { background-position: 100% 50%; }
          100% { background-position: 0% 50%; }
        }
        @media (prefers-reduced-motion: reduce) {
          .acocollo-fondo-animado { animation: none; }
        }
        .acocollo-header-sticky {
          position: sticky;
          top: 0;
          z-index: 40;
          backdrop-filter: blur(12px);
          background: rgba(12,16,21,.98);
          border-bottom: 3.5px solid #e5b80b;
          box-shadow: 0 6px 25px rgba(0,0,0,.7);
        }
        .acocollo-fade-in {
          animation: acocolloFadeIn .32s cubic-bezier(.16,1,.3,1) both;
        }
        @keyframes acocolloFadeIn {
          from { opacity: 0; transform: translateY(10px); }
          to   { opacity: 1; transform: translateY(0); }
        }
        .acocollo-stagger-item {
          animation: acocolloStaggerAnim .35s cubic-bezier(.16,1,.3,1) both;
        }
        @keyframes acocolloStaggerAnim {
          from { opacity: 0; transform: translateY(12px); }
          to   { opacity: 1; transform: translateY(0); }
        }
        .acocollo-barra-avance {
          animation: acocolloRayas 0.8s linear infinite;
        }
        @keyframes acocolloRayas {
          from { background-position: 0 0, 0 0; }
          to   { background-position: 36px 0, 0 0; }
        }
        .acocollo-fondo-animado button:not(:disabled) {
          transition: transform .18s cubic-bezier(.2,.8,.2,1), filter .18s ease, box-shadow .18s ease;
        }
        .acocollo-fondo-animado button:not(:disabled):hover {
          transform: translateY(-2px) scale(1.02);
          filter: brightness(1.18);
          box-shadow: 0 6px 20px rgba(229,184,11,.3);
        }
        .acocollo-fondo-animado button:not(:disabled):active {
          transform: translateY(0) scale(0.97);
          filter: brightness(0.95);
        }
        .acocollo-tarjeta-viva {
          animation: acocolloTarjetaEntrada .5s cubic-bezier(.25,.9,.35,1.25) both;
          transition: transform .2s ease, box-shadow .2s ease, border-color .2s ease;
        }
        .acocollo-tarjeta-viva:hover {
          transform: translateY(-3px);
          border-color: #e5b80b !important;
          box-shadow: 0 8px 25px rgba(229,184,11,.25) !important;
        }
        @keyframes acocolloTarjetaEntrada {
          from { opacity: 0; transform: translateY(10px) scale(.96); }
          to   { opacity: 1; transform: translateY(0) scale(1); }
        }
        .acocollo-celda-heatmap {
          animation: acocolloCeldaEntrada .4s ease both;
        }
        @keyframes acocolloCeldaEntrada {
          from { opacity: 0; transform: scale(.4); }
          to   { opacity: 1; transform: scale(1); }
        }
        .acocollo-celda-heatmap:hover {
          transform: scale(1.35);
          transition: transform .12s ease;
          box-shadow: 0 0 12px rgba(229,184,11,1);
          z-index: 70;
        }
        .acocollo-celda-hoy {
          position: relative;
        }
        .acocollo-celda-hoy::after {
          content: "";
          position: absolute;
          inset: 0;
          border-radius: inherit;
          border: 2px solid #e5b80b;
          animation: acocolloHoyPulso 1.4s ease-out infinite;
          pointer-events: none;
        }
        @keyframes acocolloHoyPulso {
          0%   { transform: scale(1); opacity: 1; }
          100% { transform: scale(2.2); opacity: 0; }
        }
        .acocollo-modo-transicion {
          animation: acocolloModoEntrada .35s cubic-bezier(.2,.85,.35,1.15) both;
        }
        @keyframes acocolloModoEntrada {
          from { opacity: 0; transform: scale(.985); }
          to   { opacity: 1; transform: scale(1); }
        }
        .acocollo-barra-flotante {
          animation: acocolloFlotarIn .3s cubic-bezier(.16,1,.3,1) both;
        }
        @keyframes acocolloFlotarIn {
          from { opacity: 0; transform: translate(-50%, 20px); }
          to   { opacity: 1; transform: translate(-50%, 0); }
        }
      `}</style>

      {/* ENCABEZADO SUPERIOR FIJO */}
      <div className="acocollo-header-sticky">
        <div
          style={{
            maxWidth: modoPresentacion ? "100%" : 1500,
            margin: "0 auto",
            padding: modoPresentacion ? "18px 48px" : "16px 28px",
            display: "flex",
            justifyContent: "space-between",
            alignItems: "center",
            flexWrap: "wrap",
            gap: 16,
            color: "#ffffff",
          }}
        >
          <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
            <img
              src={LOGO_PUNO_BASE64}
              alt="Escudo Gobierno Regional de Puno"
              style={{ width: modoPresentacion ? 52 : 42, height: modoPresentacion ? 58 : 48, flexShrink: 0 }}
            />
            <div>
              <h1 style={{ fontSize: modoPresentacion ? 36 : 22, marginBottom: 2, fontWeight: 800, letterSpacing: -0.3, color: "#e5b80b", textShadow: "0 0 18px rgba(229,184,11,.7)" }}>
                Expediente Técnico — C.S. ACOCOLLO I-2
              </h1>
              <p style={{ color: "#a8dadc", marginTop: 0, marginBottom: 2, fontSize: 13 }}>
                Estado en tiempo real de la carga de documentación
              </p>
              {resumen?.ultimaSync?.toDate && (
                <p style={{ color: "#e5b80b", fontSize: 11, marginTop: 0, fontWeight: 700 }}>
                  Última sincronización: {tiempoRelativo(resumen.ultimaSync.toDate())}
                  {usuarioGoogle && <span style={{ marginLeft: 8, color: "#457b9d" }}>· {usuarioGoogle.email}</span>}
                </p>
              )}
            </div>
          </div>

          {/* PANEL DE BOTONES DE EXPORTACIÓN POR FILTROS */}
          <div
            style={{
              background: "linear-gradient(135deg, #1f2d3d, #141c24)",
              border: "2px solid #e5b80b",
              borderRadius: 14,
              padding: "10px 16px",
              display: "flex",
              alignItems: "center",
              gap: 10,
              boxShadow: "0 0 20px rgba(229,184,11,0.3)",
              flexWrap: "wrap",
            }}
          >
            <span style={{ fontSize: 12, fontWeight: 800, color: "#e5b80b" }}>
              ⚙️ EXPORTAR FILTRO:
            </span>
            <div style={{ display: "flex", gap: 6 }}>
              {[
                { id: "todas", label: "📂 Todas", color: "#457b9d" },
                { id: "completas", label: "✅ Completas", color: "#2a9d8f" },
                { id: "incompletas", label: "⚠️ Incompletas", color: "#e67e22" },
                { id: "vacias", label: "❌ Vacías", color: "#c0392b" },
                { id: "incompletas_vacias", label: "🚨 Inc. + Vacías", color: "#e5b80b" },
              ].map((f) => (
                <button
                  key={f.id}
                  onClick={() => setTipoExportacion(f.id)}
                  style={{
                    fontSize: 11,
                    fontWeight: 700,
                    padding: "6px 10px",
                    borderRadius: 16,
                    border: `1.5px solid ${tipoExportacion === f.id ? f.color : "#457b9d"}`,
                    background: tipoExportacion === f.id ? f.color + "44" : "#0c1015",
                    color: tipoExportacion === f.id ? f.color : "#ffffff",
                    cursor: "pointer",
                  }}
                >
                  {f.label} {tipoExportacion === f.id ? "✓" : ""}
                </button>
              ))}
            </div>
            <button
              onClick={() => handleExportarGlobal(tipoExportacion)}
              disabled={exportandoGlobal || carpetas.length === 0}
              style={{
                fontSize: 11.5,
                fontWeight: 800,
                padding: "8px 14px",
                borderRadius: 10,
                border: "2px solid #e5b80b",
                background: "#e5b80b",
                color: "#0c1015",
                cursor: "pointer",
                display: "flex",
                alignItems: "center",
                gap: 6,
                whiteSpace: "nowrap",
              }}
              title="Exportar Reporte Global con el filtro seleccionado"
            >
              <span>📑</span> {exportandoGlobal ? "Generando..." : `PDF Global (${tipoExportacion.toUpperCase()})`}
            </button>
          </div>

          <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
            {carpetasForzadas.length > 0 && (
              <button
                onClick={() => setMostrarMarcadas(true)}
                style={{
                  fontSize: 12,
                  fontWeight: 700,
                  padding: "10px 14px",
                  borderRadius: 12,
                  border: "2px solid #e5b80b",
                  background: "#141c24",
                  color: "#e5b80b",
                  cursor: "pointer",
                  display: "flex",
                  alignItems: "center",
                  gap: 6,
                }}
              >
                <span>✓</span> Marcadas ({carpetasForzadas.length})
              </button>
            )}

            <button
              onClick={() => setModoPresentacion((v) => !v)}
              style={{
                fontSize: 12,
                fontWeight: 700,
                padding: "10px 14px",
                borderRadius: 12,
                border: modoPresentacion ? "2px solid #e5b80b" : "1.5px solid #457b9d",
                background: modoPresentacion ? "#e5b80b33" : "#141c24",
                color: modoPresentacion ? "#e5b80b" : "#ffffff",
                cursor: "pointer",
              }}
            >
              {modoPresentacion ? "Salir pres." : "🖥 Presentación"}
            </button>

            <button
              onClick={handleSync}
              disabled={sincronizando}
              style={{
                fontSize: 13,
                fontWeight: 800,
                padding: "10px 18px",
                borderRadius: 12,
                border: "2px solid #e5b80b",
                background: sincronizando ? "#141c24" : "linear-gradient(90deg,#e5b80b55,#1f2d3d55)",
                color: sincronizando ? "#a8dadc" : "#e5b80b",
                cursor: sincronizando ? "not-allowed" : "pointer",
              }}
            >
              {sincronizando ? "⟳ Sinc..." : "⟳ Sincronizar"}
            </button>
          </div>
        </div>
      </div>

      <div style={{ maxWidth: modoPresentacion ? "100%" : 1500, margin: "0 auto", padding: modoPresentacion ? "24px 48px 36px" : "24px 28px 32px", color: "#ffffff" }}>

        {/* Barra de progreso */}
        <div
          style={{
            marginBottom: 20,
            background: "#141c24",
            borderRadius: 12,
            padding: "16px 18px",
            border: "2px solid #e5b80b66",
            boxShadow: "0 4px 20px rgba(229,184,11,.15)",
          }}
        >
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", marginBottom: 4 }}>
            <span style={{ fontSize: 16, fontWeight: 700, color: "#ffffff", letterSpacing: 0.5 }}>
              <span style={{ color: "#e5b80b" }}>»» </span>AVANCE POR CARPETAS
            </span>
            <strong style={{ fontSize: 30, color: "#e5b80b", textShadow: "0 0 18px rgba(229,184,11,.7)" }}>{pct}%</strong>
          </div>
          <div style={{ fontSize: 11, color: "#a8dadc", marginBottom: 8 }}>
            {resumen?.completas ?? "–"} de {resumen?.totalFinales ?? "–"} carpetas marcadas como completas
          </div>
          <div
            style={{
              height: 34,
              background: "#0c1015",
              borderRadius: 17,
              overflow: "hidden",
              boxShadow: "inset 0 2px 6px rgba(0,0,0,.6), 0 0 0 1px #e5b80b66",
            }}
          >
            <div
              className="acocollo-barra-avance"
              style={{
                width: `${pct}%`,
                height: "100%",
                backgroundImage:
                  "repeating-linear-gradient(45deg, rgba(255,255,255,.2) 0px, rgba(255,255,255,.2) 9px, transparent 9px, transparent 18px), linear-gradient(90deg,#e5b80b,#e67e22)",
                backgroundSize: "36px 36px, 100% 100%",
                transition: "width .4s ease",
                boxShadow: "0 0 25px rgba(229,184,11,.8)",
                borderRadius: 17,
              }}
            />
          </div>
        </div>

        {/* Barra de archivos */}
        <div
          style={{
            marginBottom: 32,
            background: "#141c24",
            borderRadius: 12,
            padding: "16px 18px",
            border: "2px solid #e5b80b66",
            boxShadow: "0 4px 20px rgba(229,184,11,.15)",
          }}
        >
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", marginBottom: 4 }}>
            <span style={{ fontSize: 16, fontWeight: 700, color: "#ffffff", letterSpacing: 0.5 }}>
              <span style={{ color: "#e5b80b" }}>»» </span>AVANCE POR ARCHIVOS <span style={{ fontSize: 11, color: "#a8dadc", fontWeight: 400 }}>(más preciso)</span>
            </span>
            <strong style={{ fontSize: 30, color: "#e5b80b", textShadow: "0 0 18px rgba(229,184,11,.7)" }}>
              {resumen?.pctArchivos ?? "–"}%
            </strong>
          </div>
          <div style={{ fontSize: 11, color: "#a8dadc", marginBottom: 8 }}>
            {resumen?.totalArchivosCompletados ?? "–"} de {resumen?.totalArchivosNecesarios ?? "–"} archivos que hacen falta, ya están subidos
          </div>
          <div
            style={{
              height: 22,
              background: "#0c1015",
              borderRadius: 11,
              overflow: "hidden",
              boxShadow: "inset 0 2px 6px rgba(0,0,0,.6), 0 0 0 1px #e5b80b66",
            }}
          >
            <div
              style={{
                width: `${resumen?.pctArchivos ?? 0}%`,
                height: "100%",
                background: "linear-gradient(90deg,#e5b80b,#e67e22)",
                transition: "width .4s ease",
                boxShadow: "0 0 18px rgba(229,184,11,.7)",
                borderRadius: 11,
              }}
            />
          </div>
        </div>

        {/* Contadores con desglose */}
        <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 12, marginBottom: 32 }}>
          <Card label="Carpetas finales" value={resumen?.totalFinales ?? "–"} color="#457b9d" grande={modoPresentacion} />
          <Card label="Completas" value={resumen?.completas ?? "–"} color="#e5b80b" grande={modoPresentacion} />
          <Card label="Incompletas" value={resumen?.incompletas ?? "–"} color="#e67e22" grande={modoPresentacion} />
          <Card label="Vacías" value={resumen?.vacias ?? "–"} color="#c0392b" grande={modoPresentacion} />
        </div>

        {/* Selector de Rango de Fechas para Actividad/Heatmap */}
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16, flexWrap: "wrap", gap: 10 }}>
          <div style={{ fontSize: 15, fontWeight: 700, color: "#ffffff" }}>
            Visualización de Actividad e Historial
          </div>
          <div style={{ display: "flex", gap: 6 }}>
            {[
              { label: "30 días", val: 30 },
              { label: "84 días", val: 84 },
              { label: "119 días", val: 119 },
            ].map((btn) => (
              <button
                key={btn.val}
                onClick={() => setRangoDiasHeatmap(btn.val)}
                style={{
                  fontSize: 11,
                  padding: "5px 12px",
                  borderRadius: 16,
                  border: `1.5px solid ${rangoDiasHeatmap === btn.val ? "#e5b80b" : "#457b9d"}`,
                  background: rangoDiasHeatmap === btn.val ? "#e5b80b33" : "#141c24",
                  color: rangoDiasHeatmap === btn.val ? "#e5b80b" : "#a8dadc",
                  fontWeight: 600,
                  cursor: "pointer",
                }}
              >
                {btn.label}
              </button>
            ))}
          </div>
        </div>

        {/* Sección de Tendencia de avance y Actividad */}
        <div style={{ display: "grid", gridTemplateColumns: "1.4fr 1fr", gap: 16, marginBottom: 32 }}>
          <TendenciaChart historial={historial} grande={modoPresentacion} actividadPorDia={actividadPorDia} />
          <ActividadHeatmap actividadPorDia={actividadPorDia} diasCustom={rangoDiasHeatmap} grande={modoPresentacion} onMarcarCompleta={handleMarcarCompleta} marcandoId={marcandoId} />
        </div>

        {modoPresentacion && (
          <div
            className="acocollo-fade-in acocollo-modo-transicion"
            style={{
              display: "grid",
              gridTemplateColumns: "repeat(auto-fit, minmax(280px, 1fr))",
              gap: 28,
              justifyItems: "center",
              marginTop: 8,
            }}
          >
            {areas.map((a) => {
              const stats = areaStats[a] || { total: 0, completas: 0, incompletas: 0, vacias: 0, archivosNecesarios: 0, archivosCompletados: 0 };
              const pctArea =
                stats.archivosNecesarios > 0
                  ? Math.round((stats.archivosCompletados / stats.archivosNecesarios) * 100)
                  : 0;
              return (
                <AreaMiniCard
                  key={a}
                  area={a}
                  pct={pctArea}
                  total={stats.total}
                  incompletas={stats.incompletas}
                  vacias={stats.vacias}
                  color={colorForArea(a)}
                  active={false}
                  onClick={() => {}}
                  tamano={280}
                />
              );
            })}
          </div>
        )}

        {modoPresentacion && areas.length > 0 && (
          <div className="acocollo-fade-in acocollo-modo-transicion" style={{ marginTop: 36 }}>
            <div style={{ fontSize: 16, fontWeight: 700, color: "#ffffff", marginBottom: 20, display: "flex", alignItems: "center", gap: 8 }}>
              <span style={{ color: "#e5b80b" }}>»» </span>AVANCE POR ESPECIALIDAD, POR ÁREA
            </div>
            {areas.map((a) => {
              const especialidadesDelArea = especialidadPorArea[a] || {};
              const nombresOrdenados = Object.keys(especialidadesDelArea).sort((x, y) =>
                x.localeCompare(y, undefined, { numeric: true, sensitivity: "base" })
              );
              if (nombresOrdenados.length === 0) return null;
              return (
                <div key={a} style={{ marginBottom: 28 }}>
                  <div
                    style={{
                      fontSize: 13,
                      fontWeight: 700,
                      color: "#ffffff",
                      textTransform: "uppercase",
                      letterSpacing: 0.5,
                      marginBottom: 12,
                      padding: "8px 12px",
                      display: "flex",
                      alignItems: "center",
                      gap: 8,
                      background: "#141c24",
                      borderRadius: 8,
                      borderBottom: `2.5px solid ${colorForArea(a)}`,
                    }}
                  >
                    <span
                      style={{
                        width: 10,
                        height: 10,
                        borderRadius: "50%",
                        background: colorForArea(a),
                        flexShrink: 0,
                        boxShadow: `0 0 8px ${colorForArea(a)}`,
                      }}
                    />
                    {a} <span style={{ color: "#a8dadc", fontWeight: 400, textTransform: "none" }}>({nombresOrdenados.length} especialidades)</span>
                  </div>
                  <div
                    style={{
                      display: "grid",
                      gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))",
                      gap: 18,
                    }}
                  >
                    {nombresOrdenados.map((esp, i) => {
                      const s = especialidadesDelArea[esp];
                      const pctEsp = s.archivosNecesarios > 0 ? Math.round((s.archivosCompletados / s.archivosNecesarios) * 100) : 0;
                      return (
                        <EspecialidadMiniCard
                          key={esp}
                          nombre={esp}
                          pct={pctEsp}
                          total={s.total}
                          incompletas={s.incompletas}
                          vacias={s.vacias}
                          delay={i * 30}
                        />
                      );
                    })}
                  </div>
                </div>
              );
            })}
          </div>
        )}

        {!modoPresentacion && (
        <div className="acocollo-modo-transicion" style={{ display: "grid", gridTemplateColumns: "1.3fr 1fr", gap: 24 }}>
          <div>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8, flexWrap: "wrap", gap: 8 }}>
              <h2 style={{ fontSize: 16, color: "#ffffff", margin: 0 }}>
                Carpetas — {ESTADO_FILTRO_LABEL[filtroEstado]}{areaLabel}
              </h2>
            </div>

            <input
              type="text"
              value={busqueda}
              onChange={(e) => setBusqueda(e.target.value)}
              placeholder="🔍 Buscar carpeta por nombre..."
              style={{
                width: "100%",
                boxSizing: "border-box",
                background: "#141c24",
                color: "#ffffff",
                border: "1.5px solid #457b9d",
                borderRadius: 8,
                padding: "9px 12px",
                fontSize: 13,
                marginBottom: 12,
                outline: "none",
              }}
            />

            <div style={{ display: "flex", gap: 10, marginBottom: 14, flexWrap: "wrap" }}>
              {areas.map((a) => (
                <div key={a} style={{ display: "flex", gap: 4 }}>
                  <button
                    onClick={() => handleExportarArea(a, carpetasPorArea[a] || [], tipoExportacion)}
                    disabled={exportandoArea === a}
                    style={{
                      fontSize: 11,
                      padding: "6px 12px",
                      borderRadius: "20px 0 0 20px",
                      border: "1.5px solid #457b9d",
                      background: "#141c24",
                      color: exportandoArea === a ? "#a8dadc" : "#ffffff",
                      fontWeight: 600,
                      cursor: exportandoArea === a ? "not-allowed" : "pointer",
                    }}
                    title={`Exportar reporte PDF de ${a}`}
                  >
                    📄 {exportandoArea === a ? "Generando..." : `PDF ${a}`}
                  </button>
                  <button
                    onClick={() => handleExportarExcelArea(a, carpetasPorArea[a] || [], tipoExportacion)}
                    disabled={exportandoExcelArea === a}
                    style={{
                      fontSize: 11,
                      padding: "6px 12px",
                      borderRadius: "0 20px 20px 0",
                      border: "1.5px solid #457b9d",
                      borderLeft: "none",
                      background: "#141c24",
                      color: exportandoExcelArea === a ? "#a8dadc" : "#e5b80b",
                      fontWeight: 600,
                      cursor: exportandoExcelArea === a ? "not-allowed" : "pointer",
                    }}
                    title={`Exportar reporte Excel de ${a}`}
                  >
                    📊 {exportandoExcelArea === a ? "Generando..." : "Excel"}
                  </button>
                </div>
              ))}
            </div>

            <div
              style={{
                display: "grid",
                gridTemplateColumns: "repeat(auto-fit, minmax(190px, 1fr))",
                gap: 8,
                marginBottom: 14,
              }}
            >
              <AreaMiniCard
                area="Todas"
                pct={resumen?.pctArchivos ?? 0}
                total={resumen?.totalFinales ?? 0}
                color="#e5b80b"
                active={filtroArea === "Todas"}
                onClick={() => setFiltroArea("Todas")}
              />
              {areas.map((a) => {
                const s = areaStats[a];
                if (!s) return null;
                const areaPct =
                  s.archivosNecesarios > 0 ? Math.round((s.archivosCompletados / s.archivosNecesarios) * 100) : 0;
                return (
                  <AreaMiniCard
                    key={a}
                    area={a}
                    pct={areaPct}
                    total={s.total}
                    color={colorForArea(a)}
                    active={filtroArea === a}
                    onClick={() => setFiltroArea(filtroArea === a ? "Todas" : a)}
                  />
                );
              })}
            </div>

            {filtroArea !== "Todas" && areaStats[filtroArea] && (
              <AreaProgressPanel area={filtroArea} stats={areaStats[filtroArea]} color={colorForArea(filtroArea)} />
            )}

            {/* BOTONES DE FILTRO Y CONTROL */}
            <div style={{ display: "flex", gap: 10, marginBottom: 16, flexWrap: "wrap", alignItems: "center" }}>
              {ESTADO_OPTIONS.map((opt) => (
                <button
                  key={opt.value}
                  onClick={() => setFiltroEstado(opt.value)}
                  style={chipStyle(filtroEstado === opt.value, opt.color)}
                >
                  {opt.label}
                </button>
              ))}
              <button
                onClick={() => setColapsados((prev) => ({ ...prev, __all: !prev.__all }))}
                style={{ ...chipStyle(false, "#a8dadc"), fontWeight: 700 }}
              >
                {colapsados.__all ? "▸ Expandir todo" : "▾ Colapsar todo"}
              </button>
            </div>

            <div
              key={`${filtroEstado}-${filtroArea}-${busqueda}`}
              className="acocollo-fade-in"
              style={{
                background: "#0c1015",
                borderRadius: 12,
                overflow: "hidden",
                border: "2px solid #e5b80b66",
                boxShadow: "0 6px 24px rgba(0,0,0,.6)",
              }}
            >
              {visibles.length === 0 && (
                <div
                  className="acocollo-fade-in"
                  style={{
                    display: "flex",
                    flexDirection: "column",
                    alignItems: "center",
                    justifyContent: "center",
                    textAlign: "center",
                    padding: "48px 24px",
                    gap: 8,
                  }}
                >
                  <div style={{ fontSize: 34, opacity: 0.7 }}>
                    {carpetas.length === 0 ? "⏳" : "🔍"}
                  </div>
                  <div style={{ fontSize: 15, fontWeight: 700, color: "#ffffff" }}>
                    {carpetas.length === 0 ? "Sin datos todavía" : "No hay carpetas que coincidan"}
                  </div>
                </div>
              )}
              {(() => {
                const grupos = {};
                const ordenGrupos = [];
                for (const c of visibles) {
                  const partes = (c.ruta || c.nombre || "").split(" / ").filter(Boolean);
                  const especialidad = partes.length > 1 ? partes[1] : "(raíz)";
                  const key = `${c.area || "Sin área"} / ${especialidad}`;
                  if (!grupos[key]) {
                    grupos[key] = { area: c.area || "Sin área", especialidad, items: [] };
                    ordenGrupos.push(key);
                  }
                  grupos[key].items.push(c);
                }
                ordenGrupos.sort((a, b) => a.localeCompare(b, undefined, { numeric: true, sensitivity: "base" }));

                return ordenGrupos.map((key, grupoIndex) => {
                  const g = grupos[key];
                  const pendientesGrupo = g.items.filter((c) => c.estado !== "completa").length;
                  const vaciasGrupo = g.items.filter((c) => c.estado === "vacia").length;
                  const tienePendientes = pendientesGrupo > 0;
                  const grupoColapsado = colapsados.__all ? !colapsados[key] : !!colapsados[key];
                  return (
                    <div 
                      key={key} 
                      className="acocollo-stagger-item"
                      style={{ animationDelay: `${grupoIndex * 35}ms` }}
                    >
                      <div
                        onClick={() => toggleGrupo(key)}
                        style={{
                          padding: "12px 16px 12px 14px",
                          background: "#141c24",
                          borderLeft: `5px solid ${vaciasGrupo > 0 ? "#c0392b" : tienePendientes ? "#e67e22" : "#e5b80b"}`,
                          borderTop: "1px solid #e5b80b44",
                          borderBottom: "1px solid #0c1015",
                          display: "flex",
                          alignItems: "baseline",
                          gap: 8,
                          cursor: "pointer",
                          userSelect: "none",
                        }}
                      >
                        <span style={{ fontSize: 13, color: "#a8dadc", transform: grupoColapsado ? "rotate(-90deg)" : "none", display: "inline-block", transition: "transform .15s ease" }}>
                          ▾
                        </span>
                        <span style={{ fontSize: 11.5, fontWeight: 800, color: "#a8dadc", textTransform: "uppercase", letterSpacing: 0.5 }}>
                          {g.area}
                        </span>
                        <span style={{ color: "#e5b80b", fontSize: 12 }}>›</span>
                        <span
                          style={{
                            fontSize: 15,
                            fontWeight: 800,
                            color: "#0c1015",
                            background: "#e5b80b",
                            padding: "3px 10px",
                            borderRadius: "6px",
                            boxShadow: "0 2px 8px rgba(229,184,11,.4)",
                            textShadow: "none",
                          }}
                        >
                          {g.especialidad}
                        </span>
                        <span style={{ display: "flex", alignItems: "center", gap: 8, marginLeft: "auto" }}>
                          <MiniDona completas={g.items.length - pendientesGrupo} total={g.items.length} />
                          {tienePendientes && (
                            <span
                              style={{
                                fontSize: 10.5,
                                padding: "2px 9px",
                                borderRadius: 20,
                                background: (vaciasGrupo > 0 ? "#c0392b" : "#e67e22") + "33",
                                color: vaciasGrupo > 0 ? "#ff6b6b" : "#e67e22",
                                fontWeight: 700,
                              }}
                            >
                              {pendientesGrupo} pendiente{pendientesGrupo !== 1 ? "s" : ""}
                            </span>
                          )}
                          <span style={{ fontSize: 10.5, color: "#a8dadc", fontWeight: 600 }}>
                            {g.items.length} carpeta{g.items.length !== 1 ? "s" : ""}
                          </span>
                        </span>
                      </div>
                      {!grupoColapsado && g.items.map((c, itemIndex) => {
                        const detalle = c.detalle || c.estado;
                        const driveUrl = `https://drive.google.com/drive/folders/${c.id}`;
                        const esCompleta = c.estado === "completa";
                        return (
                          <div
                            key={c.id}
                            className="acocollo-stagger-item"
                            onClick={() => window.open(driveUrl, "_blank", "noopener,noreferrer")}
                            style={{
                              padding: "12px 16px 12px 24px",
                              borderBottom: "1px solid #1a2533",
                              cursor: "pointer",
                              animationDelay: `${(grupoIndex * 35) + (itemIndex * 20)}ms`,
                              transition: "background .15s ease",
                            }}
                            onMouseEnter={(e) => (e.currentTarget.style.background = "#1b2736")}
                            onMouseLeave={(e) => (e.currentTarget.style.background = "transparent")}
                          >
                            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 8 }}>
                              <RutaJerarquica ruta={c.ruta} nombre={c.nombre} skipLevels={2} />
                              <span
                                style={{
                                  fontSize: 10,
                                  padding: "2px 8px",
                                  borderRadius: 20,
                                  background: (ESTADO_COLOR[c.estado] || "#e5b80b") + "33",
                                  color: ESTADO_COLOR[c.estado] || "#e5b80b",
                                  textTransform: "uppercase",
                                  fontWeight: 700,
                                  whiteSpace: "nowrap",
                                  flexShrink: 0,
                                }}
                              >
                                {c.estado}{c.forzada ? " · manual" : ""}
                              </span>
                            </div>
                            {c.forzada && (
                              <div
                                style={{
                                  marginTop: 6,
                                  padding: "6px 10px",
                                  background: "#e5b80b1c",
                                  border: "1.5px solid #e5b80b55",
                                  borderRadius: 8,
                                  fontSize: 11,
                                  color: "#a8dadc",
                                }}
                              >
                                ✓ Marcada por <strong style={{ color: "#e5b80b" }}>{c.marcadoPor || "alguien"}</strong>
                              </div>
                            )}
                            <div style={{ display: "flex", justifyContent: "flex-end", alignItems: "center", gap: 10, marginTop: 6 }}>
                              <span style={{ fontSize: 11, color: "#a8dadc" }}>{detalle}</span>
                              {/* 🔄 BOTÓN DINÁMICO BIDIRECCIONAL: MARCAR COMPLETA O INCOMPLETA */}
                              <button
                                onClick={(e) => {
                                  e.stopPropagation();
                                  handleMarcarCompleta(c.id, esCompleta ? "incompleta" : "completa", c.nombre, c.ruta);
                                }}
                                disabled={marcandoId === c.id}
                                style={{
                                  fontSize: 10,
                                  padding: "3px 10px",
                                  borderRadius: 20,
                                  border: esCompleta ? "1.5px solid #e67e22" : "1.5px solid #e5b80b",
                                  background: "transparent",
                                  color: marcandoId === c.id ? "#457b9d" : esCompleta ? "#e67e22" : "#ffffff",
                                  cursor: marcandoId === c.id ? "not-allowed" : "pointer",
                                  whiteSpace: "nowrap",
                                  flexShrink: 0,
                                  fontWeight: 700,
                                }}
                              >
                                {marcandoId === c.id ? "..." : esCompleta ? "⚠ Marcar incompleta" : "✓ Marcar completa"}
                              </button>
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  );
                });
              })()}
            </div>
          </div>

          <div>
            <h2 style={{ fontSize: 16, color: "#ffffff", marginBottom: 8 }}>Actividad reciente</h2>
            <div
              style={{
                background: "#0c1015",
                borderRadius: 12,
                maxHeight: 480,
                overflowY: "auto",
                border: "2px solid #e5b80b66",
                boxShadow: "0 6px 24px rgba(0,0,0,.6)",
              }}
            >
              {eventos.length === 0 && (
                <p style={{ padding: 16, color: "#a8dadc" }}>Sin eventos todavía.</p>
              )}
              {eventos.map((e) => {
                const color = EVENTO_COLOR[e.tipo] || "#e5b80b";
                const icono = EVENTO_ICONO[e.tipo] || "•";
                const fecha = e.timestamp?.toDate ? e.timestamp.toDate() : null;
                return (
                  <div
                    key={e.id}
                    style={{
                      display: "flex",
                      gap: 10,
                      padding: "10px 14px",
                      borderBottom: "1px solid #1a2533",
                    }}
                  >
                    <div
                      style={{
                        flexShrink: 0,
                        width: 26,
                        height: 26,
                        borderRadius: "50%",
                        background: color + "33",
                        border: `1.5px solid ${color}`,
                        color: color,
                        display: "flex",
                        alignItems: "center",
                        justifyContent: "center",
                        fontSize: 13,
                        fontWeight: 700,
                        marginTop: 1,
                      }}
                    >
                      {icono}
                    </div>
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <div style={{ fontSize: 13, color: "#ffffff" }}>
                        <strong>{e.usuario}</strong> <span style={{ color }}>{EVENTO_LABEL[e.tipo] || e.tipo}</span> <strong>{e.item}</strong>
                      </div>
                      <div style={{ fontSize: 10, color: "#a8dadc", marginTop: 2 }}>
                        {tiempoRelativo(fecha)}
                      </div>
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        </div>
        )}

      </div>

      {/* BARRA DE ACCIONES FLOTANTE INTELIGENTE */}
      <div
        className="acocollo-barra-flotante"
        style={{
          position: "fixed",
          bottom: 24,
          left: "50%",
          transform: "translateX(-50%)",
          zIndex: 50,
          background: "rgba(20, 28, 36, 0.92)",
          backdropFilter: "blur(16px)",
          border: "2.5px solid #e5b80b",
          borderRadius: 32,
          padding: "10px 22px",
          display: "flex",
          alignItems: "center",
          gap: 16,
          boxShadow: "0 10px 35px rgba(0,0,0,0.7), 0 0 20px rgba(229,184,11,0.25)",
        }}
      >
        <div style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13, fontWeight: 700, color: "#ffffff" }}>
          <span style={{ width: 8, height: 8, borderRadius: "50%", background: "#e5b80b", boxShadow: "0 0 8px #e5b80b" }} />
          <span>Filtro activo: <strong style={{ color: "#e5b80b" }}>{ESTADO_FILTRO_LABEL[filtroEstado]}</strong></span>
        </div>
        <div style={{ width: 1, height: 18, background: "#457b9d" }} />
        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
          <button
            onClick={() => window.scrollTo({ top: 0, behavior: "smooth" })}
            style={{
              background: "transparent",
              border: "1.5px solid #457b9d",
              color: "#a8dadc",
              fontSize: 12,
              fontWeight: 600,
              padding: "6px 14px",
              borderRadius: 20,
              cursor: "pointer",
            }}
          >
            ↑ Ir arriba
          </button>
          <button
            onClick={() => setColapsados((prev) => ({ ...prev, __all: !prev.__all }))}
            style={{
              background: "#e5b80b",
              border: "none",
              color: "#0c1015",
              fontSize: 12,
              fontWeight: 800,
              padding: "6px 14px",
              borderRadius: 20,
              cursor: "pointer",
            }}
          >
            {colapsados.__all ? "Expandir todo" : "Colapsar todo"}
          </button>
        </div>
      </div>

      {mostrarMarcadas && (
        <div
          onClick={() => { setMostrarMarcadas(false); setBusquedaMarcadas(""); }}
          style={{
            position: "fixed",
            inset: 0,
            background: "rgba(12,16,21,.88)",
            backdropFilter: "blur(4px)",
            zIndex: 100,
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            padding: 20,
          }}
        >
          <div
            onClick={(e) => e.stopPropagation()}
            className="acocollo-fade-in"
            style={{
              background: "#141c24",
              border: "2.5px solid #e5b80b",
              borderRadius: 16,
              width: "min(1100px, 100%)",
              maxHeight: "85vh",
              display: "flex",
              flexDirection: "column",
              boxShadow: "0 20px 60px rgba(0,0,0,.8)",
            }}
          >
            <div
              style={{
                display: "flex",
                justifyContent: "space-between",
                alignItems: "center",
                padding: "20px 26px",
                borderBottom: "1.5px solid #e5b80b66",
              }}
            >
              <div style={{ fontSize: 19, fontWeight: 700, color: "#ffffff" }}>
                ✓ Carpetas marcadas manualmente ({carpetasForzadas.length})
              </div>
              <button
                onClick={() => { setMostrarMarcadas(false); setBusquedaMarcadas(""); }}
                style={{
                  fontSize: 14,
                  padding: "7px 14px",
                  borderRadius: 8,
                  border: "1.5px solid #457b9d",
                  background: "transparent",
                  color: "#a8dadc",
                  cursor: "pointer",
                }}
              >
                ✕ Cerrar
              </button>
            </div>

            {/* 🔍 BUSCADOR DENTRO DEL MODAL DE CARPETAS MARCADAS */}
            <div style={{ padding: "16px 26px 0" }}>
              <input
                type="text"
                value={busquedaMarcadas}
                onChange={(e) => setBusquedaMarcadas(e.target.value)}
                placeholder="🔍 Buscar carpeta marcada por nombre o ruta..."
                style={{
                  width: "100%",
                  boxSizing: "border-box",
                  background: "#0c1015",
                  color: "#ffffff",
                  border: "1.5px solid #e5b80b",
                  borderRadius: 8,
                  padding: "10px 14px",
                  fontSize: 13,
                  outline: "none",
                }}
              />
            </div>

            <div style={{ overflowY: "auto", padding: "16px 26px 26px" }}>
              {carpetasForzadasFiltradas.length === 0 ? (
                <div style={{ color: "#a8dadc", fontSize: 15, padding: "24px 0", textAlign: "center" }}>
                  {carpetasForzadas.length === 0 ? "No hay ninguna carpeta marcada manualmente todavía." : "No se encontró ninguna carpeta con ese nombre o ruta."}
                </div>
              ) : (
                carpetasForzadasFiltradas.map((c) => (
                  <div
                    key={c.id}
                    style={{
                      padding: "16px 18px",
                      marginBottom: 12,
                      background: "#0c1015",
                      border: "1.5px solid #e5b80b66",
                      borderRadius: 10,
                    }}
                  >
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 12 }}>
                      <div style={{ minWidth: 0 }}>
                        <div style={{ fontWeight: 700, color: "#ffffff", fontSize: 16 }}>{c.nombre}</div>
                        <div style={{ fontSize: 13.5, color: "#a8dadc", marginTop: 3 }}>{c.ruta}</div>
                      </div>
                      <button
                        onClick={() => handleMarcarCompleta(c.id, null, c.nombre, c.ruta)}
                        disabled={marcandoId === c.id}
                        title="Quita la marca manual y deja que el próximo sync calcule el estado real"
                        style={{
                          flexShrink: 0,
                          fontSize: 12,
                          padding: "5px 12px",
                          borderRadius: 20,
                          border: "1.5px solid #457b9d",
                          background: "transparent",
                          color: marcandoId === c.id ? "#457b9d" : "#a8dadc",
                          cursor: marcandoId === c.id ? "not-allowed" : "pointer",
                          fontWeight: 700,
                        }}
                      >
                        {marcandoId === c.id ? "..." : "↺ Revertir marca"}
                      </button>
                    </div>
                  </div>
                ))
              )}
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

function RutaJerarquica({ ruta, nombre, skipLevels = 0 }) {
  let partes = (ruta || nombre || "").split(" / ").filter(Boolean);
  if (skipLevels > 0 && partes.length > skipLevels) {
    partes = partes.slice(skipLevels);
  }
  let mostrar = partes;

  return (
    <div style={{ display: "flex", flexWrap: "wrap", alignItems: "baseline", gap: 4, flex: 1, minWidth: 0 }}>
      {mostrar.map((p, i) => {
        const esUltimo = i === mostrar.length - 1;
        return (
          <span key={i} style={{ display: "inline-flex", alignItems: "baseline", gap: 4 }}>
            {i > 0 && <span style={{ color: "#457b9d", fontSize: 12, fontWeight: 700 }}>›</span>}
            <span
              style={{
                fontSize: esUltimo ? 14.5 : 12,
                fontWeight: esUltimo ? 800 : 600,
                color: esUltimo ? "#ffffff" : "#e5b80b",
                background: "transparent",
                padding: 0,
                borderRadius: 0,
                boxShadow: "none",
                textShadow: esUltimo ? "0 1px 2px rgba(0,0,0,.6)" : "none",
              }}
            >
              {p}
              {esUltimo && <span style={{ color: "#e5b80b", marginLeft: 4 }}>↗</span>}
            </span>
          </span>
        );
      })}
    </div>
  );
}

function EspecialidadMiniCard({ nombre, pct, total, incompletas = 0, vacias = 0, delay }) {
  const size = 110;
  const stroke = 8;
  const radius = (size - stroke) / 2;
  const circumference = 2 * Math.PI * radius;
  const offset = circumference - (pct / 100) * circumference;
  const color = pct >= 100 ? "#e5b80b" : pct >= 50 ? "#e67e22" : "#c0392b";

  return (
    <div
      className="acocollo-tarjeta-viva"
      style={{
        display: "flex",
        flexDirection: "column",
        alignItems: "center",
        gap: 8,
        padding: "18px 12px",
        borderRadius: 12,
        background: "#141c24",
        border: "1.5px solid #e5b80b66",
        animationDelay: `${delay}ms`,
        boxShadow: "0 4px 16px rgba(0,0,0,.4)",
      }}
    >
      <svg width={size} height={size}>
        <circle cx={size / 2} cy={size / 2} r={radius} fill="none" stroke="#0c1015" strokeWidth={stroke} />
        <circle
          cx={size / 2}
          cy={size / 2}
          r={radius}
          fill="none"
          stroke={color}
          strokeWidth={stroke}
          strokeLinecap="round"
          strokeDasharray={circumference}
          strokeDashoffset={offset}
          transform={`rotate(-90 ${size / 2} ${size / 2})`}
        />
        <text x="50%" y="50%" textAnchor="middle" dominantBaseline="central" fontSize="20" fontWeight="800" fill="#ffffff">
          {pct}%
        </text>
      </svg>
      <div style={{ fontSize: 13, fontWeight: 700, color: "#ffffff", textAlign: "center", lineHeight: 1.3, maxWidth: 150 }}>
        {nombre}
      </div>
      <div style={{ fontSize: 12, color: "#a8dadc", fontWeight: 700 }}>{total} carpetas</div>
      <div style={{ fontSize: 13, textAlign: "center", display: "flex", gap: 10, marginTop: 4 }}>
        <span style={{ color: "#e67e22", fontWeight: 700 }}>{incompletas} inc.</span>
        <span style={{ color: "#c0392b", fontWeight: 700 }}>{vacias} vacías</span>
      </div>
    </div>
  );
}

function AreaMiniCard({ area, pct, total, incompletas = 0, vacias = 0, color, active, onClick, tamano }) {
  const size = tamano || 128;
  const stroke = Math.max(9, Math.round(size * 0.06));
  const radius = (size - stroke) / 2;
  const circumference = 2 * Math.PI * radius;
  const offset = circumference - (pct / 100) * circumference;
  const fontPct = Math.round(size * 0.22);
  const fontLabel = Math.max(14, Math.round(size * 0.09));
  const fontCount = Math.max(12, Math.round(size * 0.07));

  return (
    <button
      onClick={onClick}
      style={{
        display: "flex",
        flexDirection: "column",
        alignItems: "center",
        gap: Math.round(size * 0.05),
        padding: "18px 14px",
        borderRadius: 14,
        border: `2.5px solid ${active ? color : "#e5b80b66"}`,
        background: active ? color + "25" : "#141c24",
        cursor: "pointer",
        transition: "all .15s ease",
        width: tamano ? "100%" : "auto",
        boxShadow: active ? `0 0 20px ${color}66` : "0 4px 14px rgba(0,0,0,.3)",
      }}
    >
      <svg width={size} height={size}>
        <circle cx={size / 2} cy={size / 2} r={radius} fill="none" stroke="#0c1015" strokeWidth={stroke} />
        <circle
          cx={size / 2}
          cy={size / 2}
          r={radius}
          fill="none"
          stroke={color}
          strokeWidth={stroke}
          strokeLinecap="round"
          strokeDasharray={circumference}
          strokeDashoffset={offset}
          transform={`rotate(-90 ${size / 2} ${size / 2})`}
        />
        <text x="50%" y="50%" textAnchor="middle" dominantBaseline="central" fontSize={fontPct} fontWeight="700" fill="#ffffff">
          {pct}%
        </text>
      </svg>
      <div style={{ fontSize: fontLabel, fontWeight: 700, color: active ? color : "#ffffff", textAlign: "center", marginTop: 4 }}>
        {area}
      </div>
      <div style={{ fontSize: fontCount, color: "#a8dadc", fontWeight: 700 }}>{total} carpetas</div>
      {tamano && (
        <div style={{ fontSize: 14, textAlign: "center", display: "flex", gap: 12, marginTop: 6 }}>
          <span style={{ color: "#e67e22", fontWeight: 700 }}>{incompletas} inc.</span>
          <span style={{ color: "#c0392b", fontWeight: 700 }}>{vacias} vacías</span>
        </div>
      )}
    </button>
  );
}

function AreaProgressPanel({ area, stats, color }) {
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 8, background: "#141c24", border: `2.5px solid #e5b80b`, borderRadius: 12, padding: "18px 22px", marginBottom: 14, boxShadow: `0 4px 20px rgba(229,184,11,.2)` }}>
      <div style={{ fontSize: 16, fontWeight: 800, color: "#e5b80b", textTransform: "uppercase", letterSpacing: 0.5 }}>
        {area}
      </div>
      <div style={{ fontSize: 20, fontWeight: 700, color: "#ffffff", lineHeight: 1.3 }}>
        <span style={{ color: "#e5b80b", fontWeight: 800 }}>{stats.completas}</span> completas de <strong style={{ color: "#ffffff" }}>{stats.total}</strong> carpetas
      </div>
      <div style={{ fontSize: 15, fontWeight: 600, color: "#a8dadc", display: "flex", gap: 12, flexWrap: "wrap", alignItems: "center" }}>
        <span>
          <strong style={{ color: "#e67e22", fontWeight: 700, fontSize: 17 }}>{stats.incompletas}</strong> incompletas
        </span>
        <span>·</span>
        <span>
          <strong style={{ color: "#c0392b", fontWeight: 700, fontSize: 17 }}>{stats.vacias}</strong> vacías
        </span>
      </div>
    </div>
  );
}

function MiniDona({ completas, total }) {
  const size = 22;
  const stroke = 4;
  const radius = (size - stroke) / 2;
  const circumference = 2 * Math.PI * radius;
  const pct = total > 0 ? completas / total : 0;
  const offset = circumference - pct * circumference;
  const color = pct >= 1 ? "#e5b80b" : pct > 0 ? "#e67e22" : "#c0392b";

  return (
    <svg width={size} height={size}>
      <circle cx={size / 2} cy={size / 2} r={radius} fill="none" stroke="#0c1015" strokeWidth={stroke} />
      <circle
        cx={size / 2}
        cy={size / 2}
        r={radius}
        fill="none"
        stroke={color}
        strokeWidth={stroke}
        strokeLinecap="round"
        strokeDasharray={circumference}
        strokeDashoffset={offset}
        transform={`rotate(-90 ${size / 2} ${size / 2})`}
      />
    </svg>
  );
}

function chipStyle(active, color) {
  return {
    fontSize: 13,
    padding: "8px 16px",
    borderRadius: 24,
    border: `1.5px solid ${active ? color : "#457b9d"}`,
    background: active ? color + "33" : "#141c24",
    color: active ? color : "#a8dadc",
    fontWeight: 600,
    cursor: "pointer",
    boxShadow: active ? `0 0 12px ${color}66` : "none",
  };
}

function useCountUp(target) {
  const [display, setDisplay] = useState(target);
  const prevRef = useRef(target);

  useEffect(() => {
    if (typeof target !== "number") {
      setDisplay(target);
      prevRef.current = target;
      return;
    }
    const from = typeof prevRef.current === "number" ? prevRef.current : target;
    const to = target;
    if (from === to) {
      setDisplay(to);
      return;
    }
    const duracion = 650;
    const inicio = performance.now();
    let raf;
    function tick(ahora) {
      const t = Math.min(1, (ahora - inicio) / duracion);
      const suavizado = 1 - Math.pow(1 - t, 3);
      setDisplay(Math.round(from + (to - from) * suavizado));
      if (t < 1) {
        raf = requestAnimationFrame(tick);
      } else {
        prevRef.current = to;
      }
    }
    raf = requestAnimationFrame(tick);
    return () => cancelAnimationFrame(raf);
  }, [target]);

  return display;
}

function Card({ label, value, color, grande }) {
  const valorAnimado = useCountUp(value);
  return (
    <div
      className="acocollo-tarjeta-viva"
      style={{
        background: "#141c24",
        borderRadius: 12,
        padding: grande ? "26px" : "18px",
        border: `1.5px solid ${color}55`,
        borderTop: `4px solid ${color}`,
        boxShadow: `0 0 22px ${color}33`,
      }}
    >
      <div style={{ fontSize: grande ? 44 : 28, fontWeight: 700, color: "#ffffff", textShadow: `0 0 14px ${color}66` }}>{valorAnimado}</div>
      <div style={{ fontSize: grande ? 15 : 12, color: "#a8dadc", letterSpacing: 0.3 }}>{label}</div>
    </div>
  );
}

function TendenciaChart({ historial, grande, actividadPorDia }) {
  const altoLinea = grande ? 420 : 280;
  const altoBarras = grande ? 100 : 70;
  const alto = altoLinea + altoBarras;
  
  const anchoPunto = grande ? 65 : 55;
  const paddingIzq = 60;
  const paddingDer = 40;
  const anchoMinimo = grande ? 960 : 720;
  const ancho = Math.max(anchoMinimo, paddingIzq + paddingDer + historial.length * anchoPunto);
  const paddingArriba = 24;

  return (
    <div
      style={{
        background: "#141c24",
        border: "2px solid #e5b80b66",
        borderRadius: 12,
        padding: grande ? "28px 32px" : "16px 18px",
        boxShadow: "0 4px 20px rgba(0,0,0,.5)",
      }}
    >
      <div style={{ fontSize: grande ? 20 : 15, fontWeight: 700, color: "#ffffff", marginBottom: 14, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
        <span>📈 Tendencia de avance {historial.length > 0 ? `(${historial.length} días)` : ""}</span>
        <span style={{ fontSize: 11, color: "#a8dadc", fontWeight: 400 }}>Desliza horizontalmente ↔</span>
      </div>

      {historial.length < 2 ? (
        <div style={{ fontSize: 12, color: "#a8dadc", padding: "20px 0" }}>
          Todavía no hay suficiente historial.
        </div>
      ) : (
        (() => {
          const puntos = historial.map((h, i) => {
            const x = paddingIzq + i * anchoPunto;
            const y = paddingArriba + altoLinea - paddingArriba - (h.pct / 100) * (altoLinea - paddingArriba * 2);
            const tiposDia = actividadPorDia?.[h.fecha] || {};
            const incidencias = Object.values(tiposDia).reduce((s, n) => s + n, 0);
            return { x, y, pct: h.pct, fecha: h.fecha, incidencias };
          });
          const pathLinea = puntos.map((p, i) => `${i === 0 ? "M" : "L"} ${p.x} ${p.y}`).join(" ");
          const pathArea =
            `M ${puntos[0].x} ${altoLinea - paddingArriba} ` +
            puntos.map((p) => `L ${p.x} ${p.y}`).join(" ") +
            ` L ${puntos[puntos.length - 1].x} ${altoLinea - paddingArriba} Z`;

          const maxIncidencias = Math.max(1, ...puntos.map((p) => p.incidencias));
          const yBaseBarras = altoLinea + altoBarras - 16;
          const pasoEtiqueta = 1;

          return (
            <div style={{ overflowX: "auto", overflowY: "hidden", width: "100%", paddingBottom: 8, scrollbarWidth: "thin", scrollbarColor: "#457b9d #141c24" }}>
              <svg viewBox={`0 0 ${ancho} ${alto}`} style={{ width: `${ancho}px`, height: `${alto}px`, display: "block" }}>
                {[0, 25, 50, 75, 100].map((v) => {
                  const y = paddingArriba + altoLinea - paddingArriba - (v / 100) * (altoLinea - paddingArriba * 2);
                  return (
                    <g key={v}>
                      <line x1={paddingIzq - 10} y1={y} x2={ancho - paddingDer} y2={y} stroke="#e5b80b33" strokeWidth="1" strokeDasharray="3,4" />
                      <text x={paddingIzq - 16} y={y + 4} textAnchor="end" fontSize={grande ? 14 : 12} fill="#a8dadc" fontWeight="600">
                        {v}%
                      </text>
                    </g>
                  );
                })}

                <path d={pathArea} fill="url(#tendenciaGradientGold)" opacity="0.45" />
                <path
                  d={pathLinea}
                  fill="none"
                  stroke="#e5b80b"
                  strokeWidth={grande ? "4.5" : "3.5"}
                  strokeLinecap="round"
                  strokeLinejoin="round"
                />
                {puntos.map((p, i) => (
                  <circle key={i} cx={p.x} cy={p.y} r={i === puntos.length - 1 ? (grande ? 7 : 5.5) : (grande ? 5 : 3.5)} fill="#e5b80b" />
                ))}

                <defs>
                  <linearGradient id="tendenciaGradientGold" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="0%" stopColor="#e5b80b" />
                    <stop offset="100%" stopColor="#141c24" stopOpacity="0" />
                  </linearGradient>
                </defs>

                <text x={puntos[puntos.length - 1].x} y={puntos[puntos.length - 1].y - 14} textAnchor="end" fontSize={grande ? 18 : 15} fontWeight="700" fill="#e5b80b">
                  {puntos[puntos.length - 1].pct}%
                </text>

                <line x1={paddingIzq - 10} y1={altoLinea + 8} x2={ancho - paddingDer} y2={altoLinea + 8} stroke="#e5b80b44" strokeWidth="1" />
                <text x={paddingIzq} y={altoLinea + 18} fontSize={grande ? 12 : 11} fill="#a8dadc" fontWeight="700">
                  INCIDENCIAS DEL DRIVE POR DÍA
                </text>
                {puntos.map((p, i) => {
                  const alturaBarrita = Math.max(3, (p.incidencias / maxIncidencias) * (altoBarras - 28));
                  return (
                    <rect
                      key={i}
                      x={p.x - 4}
                      y={yBaseBarras - alturaBarrita}
                      width="8"
                      height={alturaBarrita}
                      rx="2"
                      fill={p.incidencias > 0 ? "#e5b80b" : "#457b9d66"}
                    />
                  );
                })}
                {puntos.map((p, i) => {
                  if (i % pasoEtiqueta !== 0 && i !== puntos.length - 1) return null;
                  const fechaObj = new Date(p.fecha + "T12:00:00");
                  const etiqueta = isNaN(fechaObj.getTime())
                    ? p.fecha
                    : fechaObj.toLocaleDateString("es-PE", { day: "numeric", month: "short" });
                  return (
                    <text key={i} x={p.x} y={alto - 2} textAnchor="middle" fontSize={grande ? 13 : 11} fill="#ffffff" fontWeight="600">
                      {etiqueta}
                    </text>
                  );
                })}
              </svg>
            </div>
          );
        })()
      )}
    </div>
  );
}

function ActividadHeatmap({ actividadPorDia, diasCustom = 84, grande }) {
  const DIAS = diasCustom;
  const [tooltip, setTooltip] = useState(null);
  const [tooltipPos, setTooltipPos] = useState(null);
  const tooltipRef = useRef(null);
  const [diaSeleccionado, setDiaSeleccionado] = useState(null);
  const [eventosDelDia, setEventosDelDia] = useState(null);

  useEffect(() => {
    if (!tooltip) {
      setTooltipPos(null);
      return;
    }
    const el = tooltipRef.current;
    if (!el) return;
    const rect = el.getBoundingClientRect();
    let left = tooltip.anclaX - rect.width / 2;
    left = Math.max(8, Math.min(window.innerWidth - rect.width - 8, left));
    let top = tooltip.anclaY - rect.height - 10;
    if (top < 8) top = tooltip.anclaY + 18;
    setTooltipPos({ left, top });
  }, [tooltip]);

  const hoy = new Date();
  hoy.setHours(0, 0, 0, 0);

  const conteoPorDia = {};
  const conteoPorTipoTotal = {};
  for (const [fechaKey, tipos] of Object.entries(actividadPorDia || {})) {
    let totalDia = 0;
    for (const [tipo, cantidad] of Object.entries(tipos || {})) {
      totalDia += cantidad;
      conteoPorTipoTotal[tipo] = (conteoPorTipoTotal[tipo] || 0) + cantidad;
    }
    conteoPorDia[fechaKey] = totalDia;
  }

  const dias = [];
  for (let i = DIAS - 1; i >= 0; i--) {
    const d = new Date(hoy);
    d.setDate(d.getDate() - i);
    const key = fechaLimaISO(d);
    dias.push({ key, count: conteoPorDia[key] || 0, fecha: d });
  }

  function intensidad(count) {
    if (count === 0) return "#0c1015";
    if (count >= 11) return "#e5b80b";
    if (count >= 4) return "#e67e22";
    return "#c0392b";
  }

  const semanas = [];
  for (let i = 0; i < dias.length; i += 7) {
    semanas.push(dias.slice(i, i + 7));
  }

  const celda = grande ? 30 : 17;
  const gap = grande ? 7 : 4;
  const tiposOrdenados = Object.keys(conteoPorTipoTotal).sort((a, b) => conteoPorTipoTotal[b] - conteoPorTipoTotal[a]);

  async function abrirDetalleDia(d) {
    setDiaSeleccionado(d);
    setEventosDelDia(null);
    try {
      const inicioUTC = new Date(`${d.key}T05:00:00.000Z`);
      const finUTC = new Date(inicioUTC.getTime() + 24 * 60 * 60 * 1000);
      const q = query(
        collection(db, "eventos"),
        where("timestamp", ">=", inicioUTC),
        where("timestamp", "<", finUTC),
        orderBy("timestamp", "desc")
      );
      const snap = await getDocs(q);
      setEventosDelDia(snap.docs.map((doc) => ({ id: doc.id, ...doc.data() })));
    } catch {
      setEventosDelDia([]);
    }
  }

  function mostrarTooltip(e, texto) {
    const rect = e.currentTarget.getBoundingClientRect();
    setTooltip({ anclaX: rect.left + rect.width / 2, anclaY: rect.top, texto });
  }

  return (
    <div
      style={{
        background: "#141c24",
        border: "2px solid #e5b80b66",
        borderRadius: 12,
        padding: grande ? "22px 26px" : "16px 18px",
        overflowX: "auto",
        boxShadow: "0 4px 20px rgba(0,0,0,.5)",
      }}
    >
      <div style={{ fontSize: grande ? 18 : 14, fontWeight: 700, color: "#ffffff", marginBottom: grande ? 18 : 10 }}>
        🔥 Actividad ({DIAS} días)
      </div>
      <div style={{ display: "flex", gap: grande ? 40 : 20, alignItems: "flex-start", flexWrap: "wrap" }}>
        <div style={{ display: "flex", gap: gap }}>
          {semanas.map((semana, si) => (
            <div key={si} style={{ display: "flex", flexDirection: "column", gap: gap }}>
              {semana.map((d, di) => {
                const esHoy = d.key === fechaLimaISO(new Date());
                const textoTooltip = `${formatearFechaLarga(d.fecha)}${esHoy ? " (hoy)" : ""} — ${d.count} evento${d.count !== 1 ? "s" : ""}`;
                return (
                  <div
                    key={d.key}
                    className={`acocollo-celda-heatmap${esHoy ? " acocollo-celda-hoy" : ""}`}
                    onMouseEnter={(e) => mostrarTooltip(e, textoTooltip)}
                    onMouseLeave={() => setTooltip(null)}
                    onClick={() => abrirDetalleDia(d)}
                    style={{
                      width: celda,
                      height: celda,
                      borderRadius: grande ? 5 : 3,
                      background: intensidad(d.count),
                      animationDelay: `${(si * 7 + di) * 4}ms`,
                      cursor: "pointer",
                    }}
                  />
                );
              })}
            </div>
          ))}
        </div>

        <div style={{ display: "flex", flexDirection: "column", gap: 16, minWidth: 240 }}>
          {tiposOrdenados.length > 0 && (
            <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
              <div style={{ fontSize: 13, fontWeight: 700, color: "#a8dadc", textTransform: "uppercase", letterSpacing: 0.4 }}>
                Resumen del período
              </div>
              {tiposOrdenados.map((tipo) => (
                <div key={tipo} style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 15 }}>
                  <span
                    style={{
                      width: 28,
                      height: 28,
                      borderRadius: "50%",
                      background: (EVENTO_COLOR[tipo] || "#e5b80b") + "33",
                      border: `1.5px solid ${EVENTO_COLOR[tipo] || "#e5b80b"}`,
                      color: EVENTO_COLOR[tipo] || "#e5b80b",
                      display: "flex",
                      alignItems: "center",
                      justifyContent: "center",
                      fontSize: 13,
                      flexShrink: 0,
                    }}
                  >
                    {EVENTO_ICONO[tipo] || "•"}
                  </span>
                  <strong style={{ color: "#ffffff" }}>{conteoPorTipoTotal[tipo]}</strong>
                  <span style={{ color: "#a8dadc" }}>{EVENTO_LABEL[tipo] || tipo}</span>
                </div>
              ))}
            </div>
          )}
        </div>
      </div>

      {tooltip && (
        <div
          ref={tooltipRef}
          style={{
            position: "fixed",
            left: tooltipPos ? tooltipPos.left : tooltip.anclaX,
            top: tooltipPos ? tooltipPos.top : tooltip.anclaY,
            visibility: tooltipPos ? "visible" : "hidden",
            background: "#141c24",
            border: "2px solid #e5b80b",
            color: "#ffffff",
            padding: "8px 12px",
            borderRadius: 7,
            fontSize: 11.5,
            fontWeight: 600,
            width: 150,
            zIndex: 200,
            pointerEvents: "none",
          }}
        >
          {tooltip.texto}
        </div>
      )}

      {diaSeleccionado && (
        <div
          onClick={() => setDiaSeleccionado(null)}
          style={{
            position: "fixed",
            inset: 0,
            background: "rgba(12,16,21,.88)",
            backdropFilter: "blur(4px)",
            zIndex: 150,
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            padding: 20,
          }}
        >
          <div
            onClick={(e) => e.stopPropagation()}
            className="acocollo-fade-in"
            style={{
              background: "#141c24",
              border: "2.5px solid #e5b80b",
              borderRadius: 16,
              width: "min(750px, 100%)",
              maxHeight: "85vh",
              display: "flex",
              flexDirection: "column",
              boxShadow: "0 20px 60px rgba(0,0,0,.8)",
            }}
          >
            <div
              style={{
                display: "flex",
                justifyContent: "space-between",
                alignItems: "center",
                padding: "20px 26px",
                borderBottom: "1.5px solid #e5b80b66",
              }}
            >
              <div>
                <div style={{ fontSize: 18, fontWeight: 700, color: "#ffffff" }}>
                  Actividad del {formatearFechaLarga(diaSeleccionado.fecha)}
                </div>
                <div style={{ fontSize: 12, color: "#a8dadc", marginTop: 2 }}>
                  {diaSeleccionado.count} evento{diaSeleccionado.count !== 1 ? "s" : ""} registrado{diaSeleccionado.count !== 1 ? "s" : ""}
                </div>
              </div>
              <button
                onClick={() => setDiaSeleccionado(null)}
                style={{
                  fontSize: 14,
                  padding: "7px 14px",
                  borderRadius: 8,
                  border: "1.5px solid #457b9d",
                  background: "transparent",
                  color: "#a8dadc",
                  cursor: "pointer",
                }}
              >
                ✕ Cerrar
              </button>
            </div>
            <div style={{ overflowY: "auto", padding: "16px 26px 26px", flex: 1 }}>
              {eventosDelDia === null ? (
                <div style={{ textAlign: "center", padding: "40px 0", color: "#a8dadc" }}>
                  Cargando eventos del día...
                </div>
              ) : eventosDelDia.length === 0 ? (
                <div style={{ textAlign: "center", padding: "40px 0", color: "#a8dadc" }}>
                  No se encontraron eventos detallados para esta fecha.
                </div>
              ) : (
                <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                  {eventosDelDia.map((ev) => {
                    const color = EVENTO_COLOR[ev.tipo] || "#e5b80b";
                    const icono = EVENTO_ICONO[ev.tipo] || "•";
                    const fechaEv = ev.timestamp?.toDate ? ev.timestamp.toDate() : null;
                    const horaStr = fechaEv ? fechaEv.toLocaleTimeString("es-PE", { hour: "2-digit", minute: "2-digit", second: "2-digit" }) : "";
                    return (
                      <div
                        key={ev.id}
                        style={{
                          background: "#0c1015",
                          border: "1.5px solid #e5b80b66",
                          borderRadius: 10,
                          padding: "12px 16px",
                          display: "flex",
                          alignItems: "flex-start",
                          gap: 12,
                        }}
                      >
                        <div
                          style={{
                            flexShrink: 0,
                            width: 30,
                            height: 30,
                            borderRadius: "50%",
                            background: color + "33",
                            border: `1.5px solid ${color}`,
                            color: color,
                            display: "flex",
                            alignItems: "center",
                            justifyContent: "center",
                            fontSize: 14,
                            fontWeight: 700,
                          }}
                        >
                          {icono}
                        </div>
                        <div style={{ flex: 1, minWidth: 0 }}>
                          <div style={{ fontSize: 14, color: "#ffffff" }}>
                            <strong>{ev.usuario || "Usuario"}</strong> <span style={{ color }}>{EVENTO_LABEL[ev.tipo] || ev.tipo}</span> <strong style={{ color: "#e5b80b" }}>{ev.item}</strong>
                          </div>
                          {ev.ruta && (
                            <div style={{ fontSize: 12, color: "#a8dadc", marginTop: 3, wordBreak: "break-all" }}>
                              📁 {ev.ruta}
                            </div>
                          )}
                          <div style={{ fontSize: 11, color: "#e5b80b", marginTop: 6, display: "flex", justifyContent: "space-between" }}>
                            <span>{ev.tipo}</span>
                            <span>{horaStr}</span>
                          </div>
                        </div>
                      </div>
                    );
                  })}
                </div>
              )}
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
