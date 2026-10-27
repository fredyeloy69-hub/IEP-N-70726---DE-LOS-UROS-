"use client";

import { useEffect, useState, useRef } from "react";
import { db } from "../lib/firebaseClient";
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
import { generarReportePorArea, generarReporteConsolidadoGlobal, generarReporteActividadDiaPDF } from "../lib/exportarReporte";
import { generarReporteExcelPorArea, generarListaSeparadoresExcel, generarListaGeneralExcel, generarListaMaterialImpresion } from "../lib/exportarExcel";
import { LOGO_PUNO_BASE64 } from "../lib/logoPuno";
import { FOTO_UROS_BASE64 } from "../lib/fotoUros";
import {
  getAuth,
  GoogleAuthProvider,
  signInWithPopup,
  onAuthStateChanged,
} from "firebase/auth";

const auth = getAuth(db.app);

const COLLAPSE_STORAGE_KEY = "uros_grupos_colapsados";

const ESTADO_COLOR = {
  completa: "#2E9FD8",
  incompleta: "#e67e22",
  vacia: "#c0392b",
};

const AREA_COLORS = [
  "#2E9FD8",
  "#e67e22",
  "#c0392b",
  "#E2B659",
  "#E2B659",
  "#E2B659",
  "#2E9FD8",
];

function colorForArea(area) {
  let hash = 0;
  for (let i = 0; i < area.length; i++) hash = area.charCodeAt(i) + ((hash << 5) - hash);
  return AREA_COLORS[Math.abs(hash) % AREA_COLORS.length];
}

const ESTADO_OPTIONS = [
  { value: "pendientes", label: "Pendientes", color: "#2E9FD8" },
  { value: "incompleta", label: "Incompletas", color: "#e67e22" },
  { value: "vacia", label: "Vacías", color: "#c0392b" },
  { value: "completa", label: "Completas", color: "#E2B659" },
  { value: "todas", label: "Todas", color: "#E2B659" },
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
  archivo_subido: "#2E9FD8",
  archivo_reemplazado: "#e67e22",
  archivo_borrado: "#c0392b",
  carpeta_creada: "#E2B659",
  carpeta_borrada: "#c0392b",
  carpeta_movida: "#e67e22",
  carpeta_marcada_completa: "#2E9FD8",
  carpeta_marcada_incompleta: "#e67e22",
  carpeta_desmarcada: "#E2B659",
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

export default function Page() {
  const [resumen, setResumen] = useState(null);
  const [carpetas, setCarpetas] = useState([]);
  const [eventos, setEventos] = useState([]);
  const [filtroArea, setFiltroArea] = useState("Todas");
  const [filtroEstado, setFiltroEstado] = useState("pendientes");
  const [sincronizando, setSincronizando] = useState(false);
  const [mensajeSync, setMensajeSync] = useState(null);
  const [busqueda, setBusqueda] = useState("");
  const [colapsados, setColapsados] = useState({});
  const [colapsoListo, setColapsoListo] = useState(false);
  const [exportandoArea, setExportandoArea] = useState(null);
  const [exportandoExcelArea, setExportandoExcelArea] = useState(null);
  const [exportandoMaterial, setExportandoMaterial] = useState(null);
  const [menuMaterialAbierto, setMenuMaterialAbierto] = useState(false);
  const [menuSeparadoresAbierto, setMenuSeparadoresAbierto] = useState(false);
  const [exportandoSeparadores, setExportandoSeparadores] = useState(null);
  const [menuListaGeneralAbierto, setMenuListaGeneralAbierto] = useState(false);
  const [exportandoListaGeneral, setExportandoListaGeneral] = useState(null);
  const [exportandoGlobal, setExportandoGlobal] = useState(false);
  const [modoPresentacion, setModoPresentacion] = useState(false);
  const [slidePresentacion, setSlidePresentacion] = useState(0);
  const [carruselPausado, setCarruselPausado] = useState(false);
  const [historial, setHistorial] = useState([]);
  const [actividadPorDia, setActividadPorDia] = useState({});
  const [marcandoId, setMarcandoId] = useState(null);
  const [mostrarMarcadas, setMostrarMarcadas] = useState(false);
  const [usuarioGoogle, setUsuarioGoogle] = useState(null);
  const [rangoDiasHeatmap, setRangoDiasHeatmap] = useState(84);
  const [busquedaMarcadas, setBusquedaMarcadas] = useState("");
  const [tipoExportacion, setTipoExportacion] = useState("todas");

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
      const usuarioFirma = usuarioGoogle?.email || usuarioGoogle?.displayName || "Sistema Uros";
      const filtroUsado = tipoForzado || tipoExportacion;
      const listaFiltrada = filtrarCarpetasParaExportar(carpetasDelArea, filtroUsado);
      generarReportePorArea(areaNombre, listaFiltrada, { usuarioFirma, tipoFiltro: filtroUsado });
    } finally {
      setExportandoArea(null);
    }
  }

  async function handleExportarGlobal(tipoForzado) {
    setExportandoGlobal(true);
    try {
      const usuarioFirma = usuarioGoogle?.email || usuarioGoogle?.displayName || "Sistema Uros";
      const filtroUsado = tipoForzado || tipoExportacion;
      const listaFiltrada = filtrarCarpetasParaExportar(carpetas, filtroUsado);
      generarReporteConsolidadoGlobal(listaFiltrada, { usuarioFirma, tipoFiltro: filtroUsado });
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
      const filtroUsado = tipoForzado || tipoExportacion;
      const listaFiltrada = filtrarCarpetasParaExportar(carpetasDelArea, filtroUsado);
      await Promise.race([
        generarReporteExcelPorArea(areaNombre, listaFiltrada, filtroUsado),
        timeoutPromise,
      ]);
    } catch (err) {
      alert(`No se pudo generar el Excel: ${err.message}`);
    } finally {
      setExportandoExcelArea(null);
    }
  }

  // Lista de separadores: siempre usa TODAS las carpetas del área, sin
  // filtro por estado — es un índice de secciones, no un reporte de avance.
  async function handleExportarSeparadores(areaNombre, carpetasDelArea) {
    setExportandoSeparadores(areaNombre);
    setMenuSeparadoresAbierto(false);
    try {
      const timeoutPromise = new Promise((_, reject) =>
        setTimeout(() => reject(new Error("Tiempo de espera agotado al generar el Excel")), 10000)
      );
      await Promise.race([
        generarListaSeparadoresExcel(areaNombre, carpetasDelArea),
        timeoutPromise,
      ]);
    } catch (err) {
      alert(`No se pudo generar la lista de separadores: ${err.message}`);
    } finally {
      setExportandoSeparadores(null);
    }
  }

  // Lista general (índice general de la documentación, con columna de
  // archivadores): igual que separadores, usa TODAS las carpetas del área.
  async function handleExportarListaGeneral(areaNombre, carpetasDelArea) {
    setExportandoListaGeneral(areaNombre);
    setMenuListaGeneralAbierto(false);
    try {
      const timeoutPromise = new Promise((_, reject) =>
        setTimeout(() => reject(new Error("Tiempo de espera agotado al generar el Excel")), 10000)
      );
      await Promise.race([
        generarListaGeneralExcel(areaNombre, carpetasDelArea),
        timeoutPromise,
      ]);
    } catch (err) {
      alert(`No se pudo generar la lista general: ${err.message}`);
    } finally {
      setExportandoListaGeneral(null);
    }
  }

  // Material de impresión: cuenta hojas A4/A3/A2/A1/A0, siempre con TODAS
  // las carpetas del área (es un conteo de material, no de avance).
  async function handleExportarMaterialImpresion(areaNombre, carpetasDelArea) {
    setExportandoMaterial(areaNombre);
    setMenuMaterialAbierto(false);
    try {
      const timeoutPromise = new Promise((_, reject) =>
        setTimeout(() => reject(new Error("Tiempo de espera agotado al generar el Excel")), 10000)
      );
      await Promise.race([
        generarListaMaterialImpresion(areaNombre, carpetasDelArea),
        timeoutPromise,
      ]);
    } catch (err) {
      alert(`No se pudo generar el material de impresión: ${err.message}`);
    } finally {
      setExportandoMaterial(null);
    }
  }

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

  const SLIDES_PRESENTACION = ["resumen", "especialidades", "tendencia"];

  useEffect(() => {
    if (!modoPresentacion || carruselPausado) return;
    const intervalo = setInterval(() => {
      setSlidePresentacion((s) => (s + 1) % SLIDES_PRESENTACION.length);
    }, 15000);
    return () => clearInterval(intervalo);
  }, [modoPresentacion, carruselPausado]);

  useEffect(() => {
    if (!modoPresentacion) setSlidePresentacion(0);
  }, [modoPresentacion]);

  useEffect(() => {
    function alCambiarFullscreen() {
      if (!document.fullscreenElement) setModoPresentacion(false);
    }
    document.addEventListener("fullscreenchange", alCambiarFullscreen);
    return () => document.removeEventListener("fullscreenchange", alCambiarFullscreen);
  }, []);

  useEffect(() => {
    if (!modoPresentacion) return;
    function alPresionarTecla(e) {
      if (e.code === "Space" || e.code === "ArrowRight") {
        e.preventDefault();
        setSlidePresentacion((s) => (s + 1) % SLIDES_PRESENTACION.length);
      } else if (e.code === "ArrowLeft") {
        e.preventDefault();
        setSlidePresentacion((s) => (s - 1 + SLIDES_PRESENTACION.length) % SLIDES_PRESENTACION.length);
      }
    }
    window.addEventListener("keydown", alPresionarTecla);
    return () => window.removeEventListener("keydown", alPresionarTecla);
  }, [modoPresentacion]);

  const pct = resumen && resumen.totalFinales
    ? Math.round((resumen.completas / resumen.totalFinales) * 100)
    : 0;

  const deltaPct =
    historial.length >= 2 ? pct - historial[historial.length - 2].pct : null;

  const proyeccion = (() => {
    const DIAS_VENTANA = 14;
    if (historial.length < 2) return null;
    const base = historial[Math.max(0, historial.length - 1 - DIAS_VENTANA)];
    const actual = historial[historial.length - 1];
    const diasTranscurridos =
      (new Date(actual.fecha) - new Date(base.fecha)) / (1000 * 60 * 60 * 24);
    if (diasTranscurridos < 1) return null;
    const ritmoDiario = (actual.pct - base.pct) / diasTranscurridos;
    if (ritmoDiario <= 0.05) return { ritmoDiario, fecha: null };
    const diasRestantes = Math.ceil((100 - actual.pct) / ritmoDiario);
    const fechaProyectada = new Date(actual.fecha);
    fechaProyectada.setDate(fechaProyectada.getDate() + diasRestantes);
    return { ritmoDiario, fecha: fechaProyectada, diasRestantes };
  })();

  // Hook de conteo animado para el % del hero — debe llamarse siempre, sin
  // condicionales, aunque el hero solo se muestre en el slide 0.
  const pctArchivosAnimado = useCountUp(resumen?.pctArchivos ?? 0);

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

  const carpetasForzadasFiltradas = busquedaMarcadas.trim()
    ? carpetasForzadas.filter(
        (c) =>
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
    <div className="uros-fondo-animado" style={{ minHeight: "100vh", width: "100%", paddingBottom: 60 }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,600;9..144,700;9..144,800&family=Nunito:wght@400;600;700;800&display=swap');

        .uros-fondo-animado, .uros-fondo-animado * {
          font-family: 'Nunito', 'Segoe UI', sans-serif;
        }
        .uros-fondo-animado {
          background: linear-gradient(
            -45deg,
            #06182B,
            #0C2B4A,
            #144068,
            #1D5A8C,
            #144068,
            #0C2B4A,
            #06182B
          );
          background-size: 500% 500%;
          animation: urosGradiente 5s ease infinite;
        }
        @keyframes urosGradiente {
          0%   { background-position: 0% 50%; }
          50%  { background-position: 100% 50%; }
          100% { background-position: 0% 50%; }
        }
        @media (prefers-reduced-motion: reduce) {
          .uros-fondo-animado { animation: none; }
        }
        .uros-fondo-animado h1, .uros-fondo-animado h2 { font-family: 'Fraunces', Georgia, serif !important; }
        @keyframes urosOla { from { transform: translateX(0); } to { transform: translateX(-50%); } }
        .uros-olas-contenedor { position: absolute; left: 0; right: 0; bottom: 0; height: 40px; overflow: hidden; pointer-events: none; z-index: 0; }
        .uros-olas-sutil { position: absolute; left: 0; bottom: 0; width: 100%; overflow: hidden; }
        .uros-olas-sutil svg { width: 200%; display: block; }
        .uros-ola-s1 { height: 40px; }
        .uros-ola-s1 svg { height: 40px; animation: urosOla 16s linear infinite; }
        .uros-ola-s2 { height: 30px; }
        .uros-ola-s2 svg { height: 30px; animation: urosOla 10s linear infinite reverse; }
        @media (prefers-reduced-motion: reduce) { .uros-olas-contenedor * { animation: none !important; } }
        @keyframes urosPanoramica { 0%,100% { background-position: center 58%; } 50% { background-position: center 84%; } }
        @media (prefers-reduced-motion: reduce) { .uros-header-sticky { animation: none !important; } }
        .uros-header-sticky {
          position: sticky;
          top: 0;
          z-index: 40;
          backdrop-filter: blur(12px);
          background: #06182B;
          background-size: cover;
          background-repeat: no-repeat;
          border-bottom: 3px solid #8FD3F4;
          box-shadow: 0 6px 25px rgba(0,0,0,.6);
          animation: urosPanoramica 50s ease-in-out infinite;
        }
        .uros-fade-in {
          animation: urosFadeIn .32s cubic-bezier(.16,1,.3,1) both;
        }
        @keyframes urosFadeIn {
          from { opacity: 0; transform: translateY(10px); }
          to   { opacity: 1; transform: translateY(0); }
        }
        .uros-stagger-item {
          animation: urosStaggerAnim .35s cubic-bezier(.16,1,.3,1) both;
        }
        @keyframes urosStaggerAnim {
          from { opacity: 0; transform: translateY(12px); }
          to   { opacity: 1; transform: translateY(0); }
        }
        .uros-barra-avance {
          animation: urosRayas 0.8s linear infinite;
        }
        @keyframes urosRayas {
          from { background-position: 0 0, 0 0; }
          to   { background-position: 36px 0, 0 0; }
        }
        .uros-fondo-animado button:not(:disabled) {
          transition: transform .18s cubic-bezier(.2,.8,.2,1), filter .18s ease, box-shadow .18s ease;
        }
        .uros-fondo-animado button:not(:disabled):hover {
          transform: translateY(-2px) scale(1.02);
          filter: brightness(1.18);
          box-shadow: 0 6px 20px rgba(46,159,216,.3);
        }
        .uros-fondo-animado button:not(:disabled):active {
          transform: translateY(0) scale(0.97);
          filter: brightness(0.95);
        }
        .uros-tarjeta-viva {
          animation: urosTarjetaEntrada .5s cubic-bezier(.25,.9,.35,1.25) both;
          transition: transform .2s ease, box-shadow .2s ease, border-color .2s ease;
        }
        .uros-tarjeta-viva:hover {
          transform: translateY(-3px);
          border-color: #2E9FD8 !important;
          box-shadow: 0 8px 25px rgba(46,159,216,.25) !important;
        }
        @keyframes urosTarjetaEntrada {
          from { opacity: 0; transform: translateY(10px) scale(.96); }
          to   { opacity: 1; transform: translateY(0) scale(1); }
        }
        .uros-celda-heatmap {
          animation: urosCeldaEntrada .4s ease both;
        }
        @keyframes urosCeldaEntrada {
          from { opacity: 0; transform: scale(.4); }
          to   { opacity: 1; transform: scale(1); }
        }
        .uros-celda-heatmap:hover {
          transform: scale(1.35);
          transition: transform .12s ease;
          box-shadow: 0 0 12px rgba(46,159,216,1);
          z-index: 70;
        }
        .uros-celda-hoy {
          position: relative;
        }
        .uros-celda-hoy::after {
          content: "";
          position: absolute;
          inset: 0;
          border-radius: inherit;
          border: 2px solid #2E9FD8;
          animation: urosHoyPulso 1.4s ease-out infinite;
          pointer-events: none;
        }
        @keyframes urosHoyPulso {
          0%   { transform: scale(1); opacity: 1; }
          100% { transform: scale(2.2); opacity: 0; }
        }
        .uros-modo-transicion {
          animation: urosModoEntrada .35s cubic-bezier(.2,.85,.35,1.15) both;
        }
        @keyframes urosModoEntrada {
          from { opacity: 0; transform: scale(.985); }
          to   { opacity: 1; transform: scale(1); }
        }
        .uros-barra-flotante {
          animation: urosFlotarIn .3s cubic-bezier(.16,1,.3,1) both;
        }
        @keyframes urosFlotarIn {
          from { opacity: 0; transform: translate(-50%, 20px); }
          to   { opacity: 1; transform: translate(-50%, 0); }
        }
        .uros-hero-pulso {
          animation: urosHeroPulso 2.6s ease-in-out infinite;
        }
        @keyframes urosHeroPulso {
          0%, 100% { text-shadow: 0 0 40px rgba(46,159,216,.65); }
          50%      { text-shadow: 0 0 70px rgba(46,159,216,1), 0 0 110px rgba(46,159,216,.5); }
        }
        .uros-anillo-hero {
          position: absolute;
          inset: 0;
          margin: auto;
          border-radius: 50%;
          border: 1.5px solid #2E9FD855;
          animation: urosAnilloExpande 3.2s ease-out infinite;
          pointer-events: none;
        }
        @keyframes urosAnilloExpande {
          0%   { width: 60px; height: 60px; opacity: .9; }
          100% { width: 420px; height: 420px; opacity: 0; }
        }
        .uros-barra-brillo {
          position: relative;
          overflow: hidden;
        }
        .uros-barra-brillo::after {
          content: "";
          position: absolute;
          top: 0; bottom: 0; left: -60%;
          width: 45%;
          background: linear-gradient(90deg, transparent, rgba(255,255,255,.55), transparent);
          animation: urosBrilloBarra 2.6s ease-in-out infinite;
        }
        @keyframes urosBrilloBarra {
          0%   { left: -60%; }
          100% { left: 130%; }
        }
        .uros-ranking-item {
          animation: urosTarjetaEntrada .5s cubic-bezier(.25,.9,.35,1.25) both;
        }
        .uros-tarjeta-respira {
          animation: urosTarjetaEntrada .5s cubic-bezier(.25,.9,.35,1.25) both, urosRespira 3.4s ease-in-out infinite .5s;
        }
        @keyframes urosRespira {
          0%, 100% { box-shadow: 0 0 22px var(--glow, rgba(46,159,216,.3)); }
          50%      { box-shadow: 0 0 40px var(--glow, rgba(46,159,216,.55)); }
        }
        .uros-punto-pulso {
          animation: urosPuntoPulso 2s ease-out infinite 1.4s;
          transform-origin: center;
        }
        @keyframes urosPuntoPulso {
          0%   { r: 5.5; opacity: .9; }
          100% { r: 18; opacity: 0; }
        }
      `}</style>

      <div
        className="uros-header-sticky"
        style={{
          backgroundImage: `linear-gradient(90deg, rgba(3,12,24,.90) 0%, rgba(4,18,34,.72) 45%, rgba(6,30,56,.25) 100%), linear-gradient(180deg, rgba(3,12,24,.35) 0%, rgba(3,12,24,0) 40%, rgba(3,12,24,.30) 100%), url(${FOTO_UROS_BASE64})`,
        }}
      >
        <div className="uros-olas-contenedor" aria-hidden="true"><div className="uros-olas-sutil uros-ola-s1"><svg viewBox="0 0 2880 40" preserveAspectRatio="none"><path d="M0 17 Q120 1 240 17 T480 17 T720 17 T960 17 T1200 17 T1440 17 T1680 17 T1920 17 T2160 17 T2400 17 T2640 17 T2880 17 V40 H0 Z" fill="#2E9FD8" fillOpacity="0.3"/><path d="M0 17 Q120 1 240 17 T480 17 T720 17 T960 17 T1200 17 T1440 17 T1680 17 T1920 17 T2160 17 T2400 17 T2640 17 T2880 17" fill="none" stroke="#CFEFFF" strokeOpacity=".45" strokeWidth="3"/></svg></div><div className="uros-olas-sutil uros-ola-s2"><svg viewBox="0 0 2880 30" preserveAspectRatio="none"><path d="M0 13 Q90 1 180 13 T360 13 T540 13 T720 13 T900 13 T1080 13 T1260 13 T1440 13 T1620 13 T1800 13 T1980 13 T2160 13 T2340 13 T2520 13 T2700 13 T2880 13 V30 H0 Z" fill="#8FD3F4" fillOpacity="0.28"/><path d="M0 13 Q90 1 180 13 T360 13 T540 13 T720 13 T900 13 T1080 13 T1260 13 T1440 13 T1620 13 T1800 13 T1980 13 T2160 13 T2340 13 T2520 13 T2700 13 T2880 13" fill="none" stroke="#E8F7FF" strokeOpacity=".65" strokeWidth="3"/></svg></div></div>
        <div
          style={{
            maxWidth: modoPresentacion ? "100%" : 1500,
            margin: "0 auto",
            padding: modoPresentacion ? "18px 48px 30px" : "16px 28px 22px",
            position: "relative",
            zIndex: 1,
            display: "flex",
            justifyContent: "space-between",
            alignItems: "flex-start",
            flexWrap: "wrap",
            gap: 12,
            color: "#F4F8FC",
          }}
        >
          <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
            <img
              src={LOGO_PUNO_BASE64}
              alt="Escudo Gobierno Regional de Puno"
              style={{ width: modoPresentacion ? 104 : 80, height: modoPresentacion ? 116 : 90, flexShrink: 0 }}
            />
            <div>
              <h1 style={{ fontSize: modoPresentacion ? 36 : 24, marginBottom: 4, fontWeight: 800, letterSpacing: -0.3, color: "#F4F8FC", textShadow: "0 2px 6px rgba(0,0,0,.6)" }}>
                Proyecto Educativo en los Uros — Puno
              </h1>
              <p style={{ color: "#AFCBE3", marginTop: 0, marginBottom: 4, fontSize: modoPresentacion ? 16 : 14 }}>
                Estado en tiempo real de la carga de documentación
              </p>
              <p style={{ color: "#AFCBE3", marginTop: 0, marginBottom: 4, fontSize: modoPresentacion ? 12.5 : 10.5, maxWidth: 720, lineHeight: 1.35 }}>
                "PROYECTO EDUCATIVO EN LOS UROS - PUNO"
              </p>
              {resumen?.ultimaSync?.toDate && (
                <p style={{ color: "#2E9FD8", fontSize: 11, marginTop: 0, fontWeight: 700 }}>
                  Última sincronización: {tiempoRelativo(resumen.ultimaSync.toDate())}
                  {usuarioGoogle && <span style={{ marginLeft: 8, color: "#E2B659" }}>· {usuarioGoogle.email}</span>}
                </p>
              )}
            </div>
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
            {carpetasForzadas.length > 0 && (
              <button
                onClick={() => setMostrarMarcadas(true)}
                style={{
                  fontSize: 13,
                  fontWeight: 700,
                  padding: "14px 18px",
                  borderRadius: 14,
                  border: "2px solid #2E9FD8",
                  background: "#0C2B4A",
                  color: "#2E9FD8",
                  cursor: "pointer",
                  display: "flex",
                  alignItems: "center",
                  gap: 8,
                  whiteSpace: "nowrap",
                }}
              >
                <span style={{ fontSize: 16 }}>✓</span>
                Marcadas manualmente ({carpetasForzadas.length})
              </button>
            )}

            <button
              onClick={handleExportarGlobal}
              disabled={exportandoGlobal || carpetas.length === 0}
              style={{
                fontSize: 13,
                fontWeight: 700,
                padding: "14px 18px",
                borderRadius: 14,
                border: "2px solid #2E9FD8",
                background: "#0C2B4A",
                color: exportandoGlobal ? "#AFCBE3" : "#2E9FD8",
                cursor: exportandoGlobal || carpetas.length === 0 ? "not-allowed" : "pointer",
                display: "flex",
                alignItems: "center",
                gap: 8,
                whiteSpace: "nowrap",
              }}
              title="Generar PDF consolidado de todo el proyecto"
            >
              <span style={{ fontSize: 16 }}>📑</span>
              {exportandoGlobal ? "Generando Global..." : "Reporte Consolidado PDF"}
            </button>

            {/* BOTÓN ESPECIAL — Fase 2: Lista para separadores (dorado, distinto al resto) */}
            <div style={{ position: "relative" }}>
              <button
                onClick={() => setMenuSeparadoresAbierto((v) => !v)}
                style={{
                  fontSize: 13,
                  fontWeight: 800,
                  padding: "14px 18px",
                  borderRadius: 14,
                  border: "2px solid #7FB069",
                  background: "linear-gradient(135deg,#16300F,#0C2B4A)",
                  color: "#D5EBC8",
                  cursor: "pointer",
                  display: "flex",
                  alignItems: "center",
                  gap: 8,
                  whiteSpace: "nowrap",
                  boxShadow: "0 0 16px rgba(212,160,23,.35)",
                  letterSpacing: 0.2,
                }}
                title="Genera el índice numerado de secciones, listo para imprimir separadores físicos — Fase 2 del expediente"
              >
                <span style={{ fontSize: 16 }}>🗂️</span>
                EXPORTAR LISTA PARA SEPARADORES
                <span style={{ fontSize: 11, opacity: 0.8 }}>{menuSeparadoresAbierto ? "▲" : "▼"}</span>
              </button>

              {menuSeparadoresAbierto && (
                <div
                  style={{
                    position: "absolute",
                    top: "calc(100% + 6px)",
                    left: 0,
                    zIndex: 50,
                    background: "#0E1F0B",
                    border: "2px solid #7FB069",
                    borderRadius: 12,
                    minWidth: 260,
                    boxShadow: "0 8px 24px rgba(0,0,0,.5)",
                    overflow: "hidden",
                  }}
                >
                  <div style={{ padding: "9px 14px", fontSize: 11, fontWeight: 700, color: "#D5EBC8", borderBottom: "1px solid #7FB06955", textTransform: "uppercase", letterSpacing: 0.5 }}>
                    Elige la carpeta madre
                  </div>
                  {areas.map((a) => (
                    <button
                      key={a}
                      onClick={() => handleExportarSeparadores(a, carpetasPorArea[a] || [])}
                      disabled={exportandoSeparadores === a}
                      style={{
                        display: "block",
                        width: "100%",
                        textAlign: "left",
                        padding: "11px 14px",
                        fontSize: 13,
                        fontWeight: 600,
                        background: "transparent",
                        border: "none",
                        borderBottom: "1px solid #7FB06922",
                        color: exportandoSeparadores === a ? "#5C8A4A" : "#D5EBC8",
                        cursor: exportandoSeparadores === a ? "not-allowed" : "pointer",
                      }}
                      onMouseEnter={(e) => (e.currentTarget.style.background = "#7FB06922")}
                      onMouseLeave={(e) => (e.currentTarget.style.background = "transparent")}
                    >
                      📊 {exportandoSeparadores === a ? `Generando ${a}...` : a}
                    </button>
                  ))}
                </div>
              )}
            </div>

            {/* BOTÓN — Índice general de la documentación (azul, distinto al dorado de separadores) */}
            <div style={{ position: "relative" }}>
              <button
                onClick={() => setMenuListaGeneralAbierto((v) => !v)}
                style={{
                  fontSize: 13,
                  fontWeight: 800,
                  padding: "14px 18px",
                  borderRadius: 14,
                  border: "2px solid #C4456E",
                  background: "linear-gradient(135deg,#3A1424,#0C2B4A)",
                  color: "#F7C6D6",
                  cursor: "pointer",
                  display: "flex",
                  alignItems: "center",
                  gap: 8,
                  whiteSpace: "nowrap",
                  boxShadow: "0 0 16px rgba(63,127,191,.35)",
                  letterSpacing: 0.2,
                }}
                title="Genera el índice general de la documentación, con columna de N° de archivadores para completar a mano"
              >
                <span style={{ fontSize: 16 }}>📑</span>
                EXPORTAR LISTA GENERAL
                <span style={{ fontSize: 11, opacity: 0.8 }}>{menuListaGeneralAbierto ? "▲" : "▼"}</span>
              </button>

              {menuListaGeneralAbierto && (
                <div
                  style={{
                    position: "absolute",
                    top: "calc(100% + 6px)",
                    left: 0,
                    zIndex: 50,
                    background: "#2A0E18",
                    border: "2px solid #C4456E",
                    borderRadius: 12,
                    minWidth: 260,
                    boxShadow: "0 8px 24px rgba(0,0,0,.5)",
                    overflow: "hidden",
                  }}
                >
                  <div style={{ padding: "9px 14px", fontSize: 11, fontWeight: 700, color: "#F7C6D6", borderBottom: "1px solid #C4456E55", textTransform: "uppercase", letterSpacing: 0.5 }}>
                    Elige la carpeta madre
                  </div>
                  {areas.map((a) => (
                    <button
                      key={a}
                      onClick={() => handleExportarListaGeneral(a, carpetasPorArea[a] || [])}
                      disabled={exportandoListaGeneral === a}
                      style={{
                        display: "block",
                        width: "100%",
                        textAlign: "left",
                        padding: "11px 14px",
                        fontSize: 13,
                        fontWeight: 600,
                        background: "transparent",
                        border: "none",
                        borderBottom: "1px solid #C4456E22",
                        color: exportandoListaGeneral === a ? "#8A4A62" : "#F7C6D6",
                        cursor: exportandoListaGeneral === a ? "not-allowed" : "pointer",
                      }}
                      onMouseEnter={(e) => (e.currentTarget.style.background = "#C4456E22")}
                      onMouseLeave={(e) => (e.currentTarget.style.background = "transparent")}
                    >
                      📊 {exportandoListaGeneral === a ? `Generando ${a}...` : a}
                    </button>
                  ))}
                </div>
              )}
            </div>

            {/* Material de impresión: lista desplegable de las carpetas madre */}
            <div style={{ position: "relative" }}>
              <button
                onClick={() => setMenuMaterialAbierto((v) => !v)}
                style={{
                  fontSize: 13,
                  fontWeight: 800,
                  padding: "14px 18px",
                  borderRadius: 14,
                  border: "2px solid #E2B659",
                  background: "linear-gradient(135deg,#2B2008,#0C2B4A)",
                  color: "#F6E1A8",
                  cursor: "pointer",
                  display: "flex",
                  alignItems: "center",
                  gap: 8,
                  whiteSpace: "nowrap",
                  boxShadow: "0 0 16px rgba(226,182,89,.35)",
                  letterSpacing: 0.2,
                }}
                title="Material de impresión (hojas A4/A3/A2/A1/A0) por carpeta madre"
              >
                <span style={{ fontSize: 16 }}>🖨️</span>
                MATERIAL DE IMPRESIÓN
                <span style={{ fontSize: 11, opacity: 0.8 }}>{menuMaterialAbierto ? "▲" : "▼"}</span>
              </button>

              {menuMaterialAbierto && (
                <div
                  style={{
                    position: "absolute",
                    top: "calc(100% + 6px)",
                    left: 0,
                    zIndex: 50,
                    background: "#1E1606",
                    border: "2px solid #E2B659",
                    borderRadius: 12,
                    minWidth: 260,
                    boxShadow: "0 8px 24px rgba(0,0,0,.5)",
                    overflow: "hidden",
                  }}
                >
                  <div style={{ padding: "9px 14px", fontSize: 11, fontWeight: 700, color: "#F6E1A8", borderBottom: "1px solid #E2B65955", textTransform: "uppercase", letterSpacing: 0.5 }}>
                    Elige la carpeta madre
                  </div>
                  {areas.map((a) => (
                    <button
                      key={a}
                      onClick={() => handleExportarMaterialImpresion(a, carpetasPorArea[a] || [])}
                      disabled={exportandoMaterial === a}
                      style={{
                        display: "block",
                        width: "100%",
                        textAlign: "left",
                        padding: "11px 14px",
                        fontSize: 13,
                        fontWeight: 600,
                        background: "transparent",
                        border: "none",
                        borderBottom: "1px solid #E2B65922",
                        color: exportandoMaterial === a ? "#8A7327" : "#F6E1A8",
                        cursor: exportandoMaterial === a ? "not-allowed" : "pointer",
                      }}
                      onMouseEnter={(e) => (e.currentTarget.style.background = "#E2B65922")}
                      onMouseLeave={(e) => (e.currentTarget.style.background = "transparent")}
                    >
                      🖨️ {exportandoMaterial === a ? `Generando ${a}...` : a}
                    </button>
                  ))}
                </div>
              )}
            </div>

            <button
              onClick={() => {
                const entrando = !modoPresentacion;
                setModoPresentacion(entrando);
                try {
                  if (entrando && document.documentElement.requestFullscreen) {
                    document.documentElement.requestFullscreen().catch(() => {});
                  } else if (!entrando && document.fullscreenElement && document.exitFullscreen) {
                    document.exitFullscreen().catch(() => {});
                  }
                } catch {}
              }}
              style={{
                fontSize: 14,
                fontWeight: 700,
                padding: "14px 20px",
                borderRadius: 14,
                border: modoPresentacion ? "2.5px solid #2E9FD8" : "1.5px solid #E2B659",
                background: modoPresentacion ? "#2E9FD833" : "#0C2B4A",
                color: modoPresentacion ? "#2E9FD8" : "#F4F8FC",
                cursor: "pointer",
                display: "flex",
                alignItems: "center",
                gap: 8,
                whiteSpace: "nowrap",
              }}
            >
              <span style={{ fontSize: 18 }}>🖥</span>
              {modoPresentacion ? "Salir de presentación" : "Modo presentación"}
            </button>
            <div style={{ textAlign: "right" }}>
              <button
                onClick={handleSync}
                disabled={sincronizando}
                style={{
                  fontSize: 20,
                  fontWeight: 800,
                  padding: "22px 42px",
                  borderRadius: 16,
                  border: "2.5px solid #2E9FD8",
                  background: sincronizando ? "#0C2B4A" : "linear-gradient(90deg,#2E9FD855,#14406855)",
                  color: sincronizando ? "#AFCBE3" : "#2E9FD8",
                  cursor: sincronizando ? "not-allowed" : "pointer",
                  boxShadow: sincronizando ? "none" : "0 0 35px rgba(46,159,216,.6)",
                  letterSpacing: 0.3,
                }}
              >
                {sincronizando ? "⟳ Sincronizando..." : "⟳ Sincronizar ahora"}
              </button>
              {mensajeSync && (
                <p
                  style={{
                    fontSize: 11,
                    marginTop: 6,
                    color: mensajeSync.tipo === "ok" ? "#2E9FD8" : "#e76f51",
                  }}
                >
                  {mensajeSync.texto}
                </p>
              )}
            </div>
          </div>
        </div>
      </div>

      <div
        style={{
          maxWidth: modoPresentacion ? "100%" : 1500,
          margin: "0 auto",
          padding: modoPresentacion ? "6px 48px 20px" : "24px 28px 32px",
          color: "#F4F8FC",
          ...(modoPresentacion
            ? { minHeight: "calc(100vh - 165px)", display: "flex", flexDirection: "column", justifyContent: "center" }
            : {}),
        }}
      >

        {/* Hero de modo presentación: número gigante + contexto (delta y proyección) */}
        {modoPresentacion && slidePresentacion === 0 && (
          <div
            className="uros-fade-in"
            style={{
              textAlign: "center",
              marginBottom: 36,
              padding: "4px 10px 4px",
            }}
          >
            <div style={{ fontSize: 17, fontWeight: 700, color: "#AFCBE3", letterSpacing: 3, textTransform: "uppercase", marginBottom: 6 }}>
              Avance por archivos
            </div>
            <div style={{ position: "relative", display: "inline-block" }}>
              <div className="uros-anillo-hero" style={{ animationDelay: "0s" }} />
              <div className="uros-anillo-hero" style={{ animationDelay: "1.6s" }} />
              <div
                className="uros-hero-pulso"
                style={{
                  position: "relative",
                  fontSize: "clamp(90px, 15vw, 190px)",
                  fontWeight: 900,
                  lineHeight: 1,
                  color: "#2E9FD8",
                }}
              >
                {pctArchivosAnimado}%
              </div>
            </div>
            <div
              style={{
                display: "flex",
                gap: 16,
                justifyContent: "center",
                flexWrap: "wrap",
                marginTop: 20,
              }}
            >
              {deltaPct !== null && (
                <span
                  style={{
                    fontSize: 17,
                    fontWeight: 700,
                    padding: "10px 20px",
                    borderRadius: 24,
                    background: deltaPct > 0 ? "#2ECC7125" : deltaPct < 0 ? "#c0392b25" : "#0C2B4A",
                    color: deltaPct > 0 ? "#2ECC71" : deltaPct < 0 ? "#e57373" : "#AFCBE3",
                    border: `1.5px solid ${deltaPct > 0 ? "#2ECC7166" : deltaPct < 0 ? "#c0392b66" : "#2E9FD866"}`,
                  }}
                >
                  {deltaPct > 0 ? "↑" : deltaPct < 0 ? "↓" : "→"} {Math.abs(deltaPct)}% desde la última sincronización
                </span>
              )}
              {proyeccion?.fecha && (
                <span
                  style={{
                    fontSize: 17,
                    fontWeight: 700,
                    padding: "10px 20px",
                    borderRadius: 24,
                    background: "#0C2B4A",
                    color: "#F4F8FC",
                    border: "1.5px solid #E2B65988",
                  }}
                >
                  📅 A este ritmo, termina el{" "}
                  {proyeccion.fecha.toLocaleDateString("es-PE", { day: "2-digit", month: "long", year: "numeric" })}
                </span>
              )}
              {proyeccion && !proyeccion.fecha && (
                <span
                  style={{
                    fontSize: 17,
                    fontWeight: 700,
                    padding: "10px 20px",
                    borderRadius: 24,
                    background: "#0C2B4A",
                    color: "#AFCBE3",
                    border: "1.5px solid #c0392b66",
                  }}
                >
                  ⚠ Ritmo estancado en los últimos días
                </span>
              )}
            </div>
          </div>
        )}

        {(!modoPresentacion || slidePresentacion === 0) && (
        <>
        {/* Barra de progreso */}
        <div
          style={{
            marginBottom: 20,
            background: "#0C2B4A",
            borderRadius: 12,
            padding: "16px 18px",
            border: "2px solid #2E9FD866",
            boxShadow: "0 4px 20px rgba(46,159,216,.15)",
          }}
        >
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", marginBottom: 4 }}>
            <span style={{ fontSize: 16, fontWeight: 700, color: "#F4F8FC", letterSpacing: 0.5 }}>
              <span style={{ color: "#2E9FD8" }}>»» </span>AVANCE POR CARPETAS
            </span>
            <strong style={{ fontSize: 30, color: "#2E9FD8", textShadow: "0 0 18px rgba(46,159,216,.7)" }}>{pct}%</strong>
          </div>
          <div style={{ fontSize: 11, color: "#AFCBE3", marginBottom: 8 }}>
            {resumen?.completas ?? "–"} de {resumen?.totalFinales ?? "–"} carpetas marcadas como completas
          </div>
          <div
            style={{
              height: 34,
              background: "#06182B",
              borderRadius: 17,
              overflow: "hidden",
              boxShadow: "inset 0 2px 6px rgba(0,0,0,.6), 0 0 0 1px #2E9FD866",
            }}
          >
            <div
              className="uros-barra-avance"
              style={{
                width: `${pct}%`,
                height: "100%",
                backgroundImage:
                  "repeating-linear-gradient(45deg, rgba(255,255,255,.2) 0px, rgba(255,255,255,.2) 9px, transparent 9px, transparent 18px), linear-gradient(90deg,#2E9FD8,#e67e22)",
                backgroundSize: "36px 36px, 100% 100%",
                transition: "width .4s ease",
                boxShadow: "0 0 25px rgba(46,159,216,.8)",
                borderRadius: 17,
              }}
            />
          </div>
        </div>

        {/* Barra de archivos */}
        <div
          style={{
            marginBottom: 32,
            background: "#0C2B4A",
            borderRadius: 12,
            padding: "16px 18px",
            border: "2px solid #2E9FD866",
            boxShadow: "0 4px 20px rgba(46,159,216,.15)",
          }}
        >
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", marginBottom: 4 }}>
            <span style={{ fontSize: 16, fontWeight: 700, color: "#F4F8FC", letterSpacing: 0.5 }}>
              <span style={{ color: "#2E9FD8" }}>»» </span>AVANCE POR ARCHIVOS <span style={{ fontSize: 11, color: "#AFCBE3", fontWeight: 400 }}>(más preciso)</span>
            </span>
            <strong style={{ fontSize: 30, color: "#2E9FD8", textShadow: "0 0 18px rgba(46,159,216,.7)" }}>
              {resumen?.pctArchivos ?? "–"}%
            </strong>
          </div>
          <div style={{ fontSize: 11, color: "#AFCBE3", marginBottom: 8 }}>
            {resumen?.totalArchivosCompletados ?? "–"} de {resumen?.totalArchivosNecesarios ?? "–"} archivos que hacen falta, ya están subidos
          </div>
          <div
            style={{
              height: 22,
              background: "#06182B",
              borderRadius: 11,
              overflow: "hidden",
              boxShadow: "inset 0 2px 6px rgba(0,0,0,.6), 0 0 0 1px #2E9FD866",
            }}
          >
            <div
              style={{
                width: `${resumen?.pctArchivos ?? 0}%`,
                height: "100%",
                background: "linear-gradient(90deg,#2E9FD8,#e67e22)",
                transition: "width .4s ease",
                boxShadow: "0 0 18px rgba(46,159,216,.7)",
                borderRadius: 11,
              }}
            />
          </div>
        </div>

        {/* Contadores con desglose */}
        <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 12, marginBottom: 32 }}>
          <Card label="Carpetas finales" value={resumen?.totalFinales ?? "–"} color="#E2B659" grande={modoPresentacion} />
          <Card label="Completas" value={resumen?.completas ?? "–"} color="#2E9FD8" grande={modoPresentacion} />
          <Card label="Incompletas" value={resumen?.incompletas ?? "–"} color="#e67e22" grande={modoPresentacion} />
          <Card label="Vacías" value={resumen?.vacias ?? "–"} color="#c0392b" grande={modoPresentacion} />
        </div>
        </>
        )}

        {/* Ranking por área — reemplaza los círculos en modo presentación: de
            un vistazo se ve cuál área va más atrasada, sin comparar círculo
            por círculo. */}
        {modoPresentacion && slidePresentacion === 0 && areas.length > 0 && (
          <div className="uros-fade-in" style={{ marginBottom: 32, maxWidth: 1100, marginLeft: "auto", marginRight: "auto", width: "100%" }}>
            <div style={{ fontSize: 19, fontWeight: 700, color: "#F4F8FC", marginBottom: 22 }}>
              <span style={{ color: "#2E9FD8" }}>»» </span>RANKING DE AVANCE POR ÁREA
            </div>
            {areas
              .map((a) => {
                const stats = areaStats[a] || { total: 0, completas: 0, incompletas: 0, vacias: 0, archivosNecesarios: 0, archivosCompletados: 0 };
                const pctArea =
                  stats.archivosNecesarios > 0
                    ? Math.round((stats.archivosCompletados / stats.archivosNecesarios) * 100)
                    : 0;
                return { a, stats, pctArea };
              })
              .sort((x, y) => y.pctArea - x.pctArea)
              .map(({ a, stats, pctArea }, i) => (
                <div
                  key={a}
                  className="uros-ranking-item"
                  style={{ display: "flex", alignItems: "center", gap: 20, marginBottom: 22, animationDelay: `${i * 90}ms` }}
                >
                  <div style={{ width: 34, fontSize: 22, fontWeight: 800, color: "#AFCBE3", textAlign: "center", flexShrink: 0 }}>
                    {i + 1}
                  </div>
                  <div style={{ width: 250, fontSize: 19, fontWeight: 700, color: "#F4F8FC", flexShrink: 0 }}>
                    {a}
                    <div style={{ fontSize: 13, color: "#AFCBE3", fontWeight: 400 }}>
                      {stats.total} carpetas · {stats.incompletas} inc. · {stats.vacias} vacías
                    </div>
                  </div>
                  <div
                    className="uros-barra-brillo"
                    style={{ flex: 1, height: 48, background: "#06182B", borderRadius: 24, boxShadow: "inset 0 2px 6px rgba(0,0,0,.6)" }}
                  >
                    <div
                      style={{
                        width: `${pctArea}%`,
                        height: "100%",
                        background: `linear-gradient(90deg, ${colorForArea(a)}, ${colorForArea(a)}cc)`,
                        transition: "width 1s cubic-bezier(.16,1,.3,1)",
                        boxShadow: `0 0 20px ${colorForArea(a)}bb`,
                        borderRadius: 24,
                      }}
                    />
                  </div>
                  <div style={{ width: 80, fontSize: 26, fontWeight: 800, color: colorForArea(a), textAlign: "right", flexShrink: 0 }}>
                    <AnimatedPercent value={pctArea} />
                  </div>
                </div>
              ))}
          </div>
        )}

        {(!modoPresentacion || slidePresentacion === 2) && (
        <>
        {/* Selector de Rango de Fechas para Actividad/Heatmap */}
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16, flexWrap: "wrap", gap: 10 }}>
          <div style={{ fontSize: 15, fontWeight: 700, color: "#F4F8FC" }}>
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
                  border: `1px solid ${rangoDiasHeatmap === btn.val ? "#2E9FD8" : "#E2B659"}`,
                  background: rangoDiasHeatmap === btn.val ? "#2E9FD833" : "#0C2B4A",
                  color: rangoDiasHeatmap === btn.val ? "#2E9FD8" : "#AFCBE3",
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
        <div style={{ display: "grid", gridTemplateColumns: "1.4fr 1fr", gap: modoPresentacion ? 24 : 16, marginBottom: 32, alignItems: "stretch" }}>
          <TendenciaChart historial={historial} grande={modoPresentacion} actividadPorDia={actividadPorDia} />
          <ActividadHeatmap areasValidas={areas} actividadPorDia={actividadPorDia} diasCustom={rangoDiasHeatmap} grande={modoPresentacion} onMarcarCompleta={handleMarcarCompleta} marcandoId={marcandoId} />
        </div>
        </>
        )}

        {modoPresentacion && slidePresentacion === 1 && areas.length > 0 && (
          <div className="uros-fade-in uros-modo-transicion" style={{ marginTop: 8 }}>
            <div style={{ fontSize: 16, fontWeight: 700, color: "#F4F8FC", marginBottom: 20, display: "flex", alignItems: "center", gap: 8 }}>
              <span style={{ color: "#2E9FD8" }}>»» </span>AVANCE POR ESPECIALIDAD, POR ÁREA
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
                      color: "#F4F8FC",
                      textTransform: "uppercase",
                      letterSpacing: 0.5,
                      marginBottom: 12,
                      padding: "8px 12px",
                      display: "flex",
                      alignItems: "center",
                      gap: 8,
                      background: "#0C2B4A",
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
                    {a} <span style={{ color: "#AFCBE3", fontWeight: 400, textTransform: "none" }}>({nombresOrdenados.length} especialidades)</span>
                  </div>
                  <div
                    style={{
                      display: "grid",
                      gridTemplateColumns: `repeat(auto-fit, minmax(${modoPresentacion ? 230 : 200}px, 1fr))`,
                      gap: modoPresentacion ? 26 : 18,
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
                          delay={i * (modoPresentacion ? 60 : 30)}
                          grande={modoPresentacion}
                        />
                      );
                    })}
                  </div>
                </div>
              );
            })}
          </div>
        )}

        {modoPresentacion && (
          <div
            onMouseEnter={() => setCarruselPausado(true)}
            onMouseLeave={() => setCarruselPausado(false)}
            style={{
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              gap: 18,
              marginTop: 36,
              padding: "14px 0",
            }}
          >
            <button
              onClick={() =>
                setSlidePresentacion((s) => (s - 1 + SLIDES_PRESENTACION.length) % SLIDES_PRESENTACION.length)
              }
              style={{
                fontSize: 20,
                width: 40,
                height: 40,
                borderRadius: "50%",
                border: "1.5px solid #E2B65988",
                background: "#0C2B4A",
                color: "#F4F8FC",
                cursor: "pointer",
              }}
              title="Sección anterior"
            >
              ‹
            </button>
            <div style={{ display: "flex", gap: 10 }}>
              {SLIDES_PRESENTACION.map((s, i) => (
                <button
                  key={s}
                  onClick={() => setSlidePresentacion(i)}
                  title={{ resumen: "Resumen", especialidades: "Por especialidad", tendencia: "Tendencia y actividad" }[s]}
                  style={{
                    width: slidePresentacion === i ? 34 : 12,
                    height: 12,
                    borderRadius: 6,
                    border: "none",
                    background: slidePresentacion === i ? "#2E9FD8" : "#AFCBE366",
                    cursor: "pointer",
                    transition: "all .25s ease",
                    padding: 0,
                  }}
                />
              ))}
            </div>
            <button
              onClick={() => setSlidePresentacion((s) => (s + 1) % SLIDES_PRESENTACION.length)}
              style={{
                fontSize: 20,
                width: 40,
                height: 40,
                borderRadius: "50%",
                border: "1.5px solid #E2B65988",
                background: "#0C2B4A",
                color: "#F4F8FC",
                cursor: "pointer",
              }}
              title="Siguiente sección"
            >
              ›
            </button>
            {carruselPausado && (
              <span style={{ fontSize: 11, color: "#AFCBE3", marginLeft: 6 }}>⏸ en pausa</span>
            )}
          </div>
        )}

        {!modoPresentacion && (
        <div className="uros-modo-transicion" style={{ display: "grid", gridTemplateColumns: "1.3fr 1fr", gap: 24 }}>
          <div>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8, flexWrap: "wrap", gap: 8 }}>
              <h2 style={{ fontSize: 16, color: "#F4F8FC", margin: 0 }}>
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
                background: "#0C2B4A",
                color: "#F4F8FC",
                border: "1.5px solid #E2B659",
                borderRadius: 8,
                padding: "9px 12px",
                fontSize: 13,
                marginBottom: 12,
                outline: "none",
              }}
            />

            {/* FILTRO PARA EXPORTAR + BOTONES DE EXPORTAR POR ÁREA — juntos, mismo tamaño de texto */}
            <div
              style={{
                display: "flex",
                gap: 8,
                marginBottom: 10,
                flexWrap: "wrap",
                alignItems: "center",
                background: "#0C2B4A",
                border: "1.5px solid #2E9FD866",
                borderRadius: 10,
                padding: "10px 14px",
              }}
            >
              <span style={{ fontSize: 13, fontWeight: 700, color: "#AFCBE3", marginRight: 4 }}>
                ⚙️ EXPORTAR FILTRO:
              </span>
              {[
                { id: "todas", label: "📂 Todas", color: "#E2B659" },
                { id: "completas", label: "✅ Completas", color: "#2a9d8f" },
                { id: "incompletas", label: "⚠️ Incompletas", color: "#e67e22" },
                { id: "vacias", label: "❌ Vacías", color: "#c0392b" },
                { id: "incompletas_vacias", label: "🚨 Inc. + Vacías", color: "#2E9FD8" },
              ].map((f) => (
                <button
                  key={f.id}
                  onClick={() => setTipoExportacion(f.id)}
                  style={chipStyle(tipoExportacion === f.id, f.color)}
                >
                  {f.label}
                </button>
              ))}
            </div>

            <div style={{ display: "flex", gap: 10, marginBottom: 14, flexWrap: "wrap" }}>
              {areas.map((a) => (
                <div key={a} style={{ display: "flex", gap: 4 }}>
                  <button
                    onClick={() => handleExportarArea(a, carpetasPorArea[a] || [])}
                    disabled={exportandoArea === a}
                    style={{
                      fontSize: 13,
                      padding: "8px 16px",
                      borderRadius: "20px 0 0 20px",
                      border: "1.5px solid #E2B659",
                      background: "#0C2B4A",
                      color: exportandoArea === a ? "#AFCBE3" : "#F4F8FC",
                      fontWeight: 600,
                      cursor: exportandoArea === a ? "not-allowed" : "pointer",
                    }}
                    title={`Exportar reporte PDF de ${a}`}
                  >
                    📄 {exportandoArea === a ? "Generando..." : `PDF ${a}`}
                  </button>
                  <button
                    onClick={() => handleExportarExcelArea(a, carpetasPorArea[a] || [])}
                    disabled={exportandoExcelArea === a}
                    style={{
                      fontSize: 13,
                      padding: "8px 16px",
                      borderRadius: "0 20px 20px 0",
                      border: "1.5px solid #E2B659",
                      borderLeft: "none",
                      background: "#0C2B4A",
                      color: exportandoExcelArea === a ? "#AFCBE3" : "#2E9FD8",
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
                color="#2E9FD8"
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
                style={{ ...chipStyle(false, "#AFCBE3"), fontWeight: 700 }}
              >
                {colapsados.__all ? "▸ Expandir todo" : "▾ Colapsar todo"}
              </button>
            </div>

            <div
              key={`${filtroEstado}-${filtroArea}-${busqueda}`}
              className="uros-fade-in"
              style={{
                background: "#06182B",
                borderRadius: 12,
                overflow: "hidden",
                border: "2px solid #2E9FD866",
                boxShadow: "0 6px 24px rgba(0,0,0,.6)",
              }}
            >
              {visibles.length === 0 && (
                <div
                  className="uros-fade-in"
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
                  <div style={{ fontSize: 15, fontWeight: 700, color: "#F4F8FC" }}>
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
                      className="uros-stagger-item"
                      style={{ animationDelay: `${grupoIndex * 35}ms` }}
                    >
                      <div
                        onClick={() => toggleGrupo(key)}
                        style={{
                          padding: "12px 16px 12px 14px",
                          background: "#0C2B4A",
                          borderLeft: `5px solid ${vaciasGrupo > 0 ? "#c0392b" : tienePendientes ? "#e67e22" : "#2E9FD8"}`,
                          borderTop: "1px solid #2E9FD844",
                          borderBottom: "1px solid #06182B",
                          display: "flex",
                          alignItems: "baseline",
                          gap: 8,
                          cursor: "pointer",
                          userSelect: "none",
                        }}
                      >
                        <span style={{ fontSize: 13, color: "#AFCBE3", transform: grupoColapsado ? "rotate(-90deg)" : "none", display: "inline-block", transition: "transform .15s ease" }}>
                          ▾
                        </span>
                        <span style={{ fontSize: 11.5, fontWeight: 800, color: "#AFCBE3", textTransform: "uppercase", letterSpacing: 0.5 }}>
                          {g.area}
                        </span>
                        <span style={{ color: "#2E9FD8", fontSize: 12 }}>›</span>
                        <span
                          style={{
                            fontSize: 15,
                            fontWeight: 800,
                            color: "#06182B",
                            background: "#2E9FD8",
                            padding: "3px 10px",
                            borderRadius: "6px",
                            boxShadow: "0 2px 8px rgba(46,159,216,.4)",
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
                          <span style={{ fontSize: 10.5, color: "#AFCBE3", fontWeight: 600 }}>
                            {g.items.length} carpeta{g.items.length !== 1 ? "s" : ""}
                          </span>
                        </span>
                      </div>
                      {!grupoColapsado && g.items.map((c, itemIndex) => {
                        const detalle = c.detalle || c.estado;
                        const driveUrl = `https://drive.google.com/drive/folders/${c.id}`;
                        return (
                          <div
                            key={c.id}
                            className="uros-stagger-item"
                            onClick={() => window.open(driveUrl, "_blank", "noopener,noreferrer")}
                            style={{
                              padding: "12px 16px 12px 24px",
                              borderBottom: "1px solid #12385F",
                              cursor: "pointer",
                              animationDelay: `${(grupoIndex * 35) + (itemIndex * 20)}ms`,
                              transition: "background .15s ease",
                            }}
                            onMouseEnter={(e) => (e.currentTarget.style.background = "#0F3558")}
                            onMouseLeave={(e) => (e.currentTarget.style.background = "transparent")}
                          >
                            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 8 }}>
                              <RutaJerarquica ruta={c.ruta} nombre={c.nombre} skipLevels={2} />
                              <span
                                style={{
                                  fontSize: 10,
                                  padding: "2px 8px",
                                  borderRadius: 20,
                                  background: (ESTADO_COLOR[c.estado] || "#2E9FD8") + "33",
                                  color: ESTADO_COLOR[c.estado] || "#2E9FD8",
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
                                  background: "#2E9FD81c",
                                  border: "1.5px solid #2E9FD855",
                                  borderRadius: 8,
                                  fontSize: 11,
                                  color: "#AFCBE3",
                                }}
                              >
                                ✓ Marcada por <strong style={{ color: "#2E9FD8" }}>{c.marcadoPor || "alguien"}</strong>
                              </div>
                            )}
                            <div style={{ display: "flex", justifyContent: "flex-end", alignItems: "center", gap: 10, marginTop: 6 }}>
                              <span style={{ fontSize: 11, color: "#AFCBE3" }}>{detalle}</span>
                              <button
                                onClick={(e) => {
                                  e.stopPropagation();
                                  handleMarcarCompleta(c.id, c.estado === "completa" ? "incompleta" : "completa", c.nombre, c.ruta);
                                }}
                                disabled={marcandoId === c.id}
                                style={{
                                  fontSize: 10,
                                  padding: "3px 9px",
                                  borderRadius: 20,
                                  border: c.estado === "completa" ? "1.5px solid #e67e22" : "1.5px solid #2E9FD888",
                                  background: "transparent",
                                  color: marcandoId === c.id ? "#E2B659" : c.estado === "completa" ? "#e67e22" : "#F4F8FC",
                                  cursor: marcandoId === c.id ? "not-allowed" : "pointer",
                                  whiteSpace: "nowrap",
                                  flexShrink: 0,
                                }}
                              >
                                {marcandoId === c.id ? "..." : c.estado === "completa" ? "⚠ Marcar incompleta" : "✓ Marcar completa"}
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
            <h2 style={{ fontSize: 16, color: "#F4F8FC", marginBottom: 8 }}>Actividad reciente</h2>
            <div
              style={{
                background: "#06182B",
                borderRadius: 12,
                maxHeight: 480,
                overflowY: "auto",
                border: "2px solid #2E9FD866",
                boxShadow: "0 6px 24px rgba(0,0,0,.6)",
              }}
            >
              {eventos.length === 0 && (
                <p style={{ padding: 16, color: "#AFCBE3" }}>Sin eventos todavía.</p>
              )}
              {eventos.map((e) => {
                const color = EVENTO_COLOR[e.tipo] || "#2E9FD8";
                const icono = EVENTO_ICONO[e.tipo] || "•";
                const fecha = e.timestamp?.toDate ? e.timestamp.toDate() : null;
                return (
                  <div
                    key={e.id}
                    style={{
                      display: "flex",
                      gap: 10,
                      padding: "10px 14px",
                      borderBottom: "1px solid #12385F",
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
                      <div style={{ fontSize: 13, color: "#F4F8FC" }}>
                        <strong>{e.usuario}</strong> <span style={{ color }}>{EVENTO_LABEL[e.tipo] || e.tipo}</span> <strong>{e.item}</strong>
                      </div>
                      <div style={{ fontSize: 10, color: "#AFCBE3", marginTop: 2 }}>
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
        className="uros-barra-flotante"
        style={{
          position: "fixed",
          bottom: 24,
          left: "50%",
          transform: "translateX(-50%)",
          zIndex: 50,
          background: "rgba(20, 28, 36, 0.92)",
          backdropFilter: "blur(16px)",
          border: "2px solid #2E9FD8",
          borderRadius: 32,
          padding: "10px 22px",
          display: "flex",
          alignItems: "center",
          gap: 16,
          boxShadow: "0 10px 35px rgba(0,0,0,0.7), 0 0 20px rgba(46,159,216,0.25)",
        }}
      >
        <div style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13, fontWeight: 700, color: "#F4F8FC" }}>
          <span style={{ width: 8, height: 8, borderRadius: "50%", background: "#2E9FD8", boxShadow: "0 0 8px #2E9FD8" }} />
          <span>Filtro activo: <strong style={{ color: "#2E9FD8" }}>{ESTADO_FILTRO_LABEL[filtroEstado]}</strong></span>
        </div>
        <div style={{ width: 1, height: 18, background: "#E2B659" }} />
        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
          <button
            onClick={() => window.scrollTo({ top: 0, behavior: "smooth" })}
            style={{
              background: "transparent",
              border: "1.5px solid #E2B659",
              color: "#AFCBE3",
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
              background: "#2E9FD8",
              border: "none",
              color: "#06182B",
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
          onClick={() => setMostrarMarcadas(false)}
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
            className="uros-fade-in"
            style={{
              background: "#0C2B4A",
              border: "2.5px solid #2E9FD8",
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
                borderBottom: "1.5px solid #2E9FD866",
              }}
            >
              <div style={{ fontSize: 19, fontWeight: 700, color: "#F4F8FC" }}>
                ✓ Carpetas marcadas manualmente ({carpetasForzadas.length})
              </div>
              <button
                onClick={() => setMostrarMarcadas(false)}
                style={{
                  fontSize: 14,
                  padding: "7px 14px",
                  borderRadius: 8,
                  border: "1.5px solid #E2B659",
                  background: "transparent",
                  color: "#AFCBE3",
                  cursor: "pointer",
                }}
              >
                ✕ Cerrar
              </button>
            </div>
            <div style={{ overflowY: "auto", padding: "16px 26px 26px" }}>
              {carpetasForzadas.length > 0 && (
                <input
                  type="text"
                  value={busquedaMarcadas}
                  onChange={(e) => setBusquedaMarcadas(e.target.value)}
                  placeholder="🔍 Buscar carpeta por nombre o ruta..."
                  style={{
                    width: "100%",
                    boxSizing: "border-box",
                    background: "#06182B",
                    color: "#F4F8FC",
                    border: "1.5px solid #E2B659",
                    borderRadius: 8,
                    padding: "9px 12px",
                    fontSize: 13,
                    marginBottom: 14,
                    outline: "none",
                  }}
                />
              )}
              {carpetasForzadas.length === 0 ? (
                <div style={{ color: "#AFCBE3", fontSize: 15, padding: "24px 0" }}>
                  No hay ninguna carpeta marcada manualmente todavía.
                </div>
              ) : carpetasForzadasFiltradas.length === 0 ? (
                <div style={{ color: "#AFCBE3", fontSize: 15, padding: "24px 0" }}>
                  Ninguna carpeta marcada coincide con "{busquedaMarcadas}".
                </div>
              ) : (
                carpetasForzadasFiltradas.map((c) => (
                  <div
                    key={c.id}
                    style={{
                      padding: "16px 18px",
                      marginBottom: 12,
                      background: "#06182B",
                      border: "1.5px solid #2E9FD866",
                      borderRadius: 10,
                    }}
                  >
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 12 }}>
                      <div style={{ minWidth: 0 }}>
                        <div style={{ fontWeight: 700, color: "#F4F8FC", fontSize: 16 }}>{c.nombre}</div>
                        <div style={{ fontSize: 13.5, color: "#AFCBE3", marginTop: 3 }}>{c.ruta}</div>
                        <div style={{ fontSize: 11, color: "#2E9FD8", marginTop: 4, fontWeight: 700, textTransform: "uppercase" }}>
                          Forzada como: {c.estado || "completa"}
                        </div>
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
                          border: "1.5px solid #E2B659",
                          background: "transparent",
                          color: marcandoId === c.id ? "#E2B659" : "#AFCBE3",
                          cursor: marcandoId === c.id ? "not-allowed" : "pointer",
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
            {i > 0 && <span style={{ color: "#E2B659", fontSize: 12, fontWeight: 700 }}>›</span>}
            <span
              style={{
                fontSize: esUltimo ? 14.5 : 12,
                fontWeight: esUltimo ? 800 : 600,
                color: esUltimo ? "#F4F8FC" : "#E2B659",
                background: "transparent",
                padding: 0,
                borderRadius: 0,
                boxShadow: "none",
                textShadow: esUltimo ? "0 1px 2px rgba(0,0,0,.6)" : "none",
              }}
            >
              {p}
              {esUltimo && <span style={{ color: "#E2B659", marginLeft: 4 }}>↗</span>}
            </span>
          </span>
        );
      })}
    </div>
  );
}

function EspecialidadMiniCard({ nombre, pct, total, incompletas = 0, vacias = 0, delay, grande }) {
  const size = grande ? 150 : 110;
  const stroke = grande ? 11 : 8;
  const radius = (size - stroke) / 2;
  const circumference = 2 * Math.PI * radius;
  const color = pct >= 100 ? "#2E9FD8" : pct >= 50 ? "#e67e22" : "#c0392b";

  const [avanzado, setAvanzado] = useState(false);
  useEffect(() => {
    const t = setTimeout(() => setAvanzado(true), 80 + (delay || 0));
    return () => clearTimeout(t);
  }, [delay]);
  const offset = circumference - ((avanzado ? pct : 0) / 100) * circumference;

  return (
    <div
      className="uros-tarjeta-viva"
      style={{
        display: "flex",
        flexDirection: "column",
        alignItems: "center",
        gap: grande ? 12 : 8,
        padding: grande ? "26px 18px" : "18px 12px",
        borderRadius: 12,
        background: "#0C2B4A",
        border: "1.5px solid #2E9FD866",
        animationDelay: `${delay}ms`,
        boxShadow: "0 4px 16px rgba(0,0,0,.4)",
      }}
    >
      <svg width={size} height={size}>
        <circle cx={size / 2} cy={size / 2} r={radius} fill="none" stroke="#06182B" strokeWidth={stroke} />
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
          style={{ transition: "stroke-dashoffset 1.1s cubic-bezier(.16,1,.3,1)", filter: `drop-shadow(0 0 6px ${color}aa)` }}
          transform={`rotate(-90 ${size / 2} ${size / 2})`}
        />
        <text x="50%" y="50%" textAnchor="middle" dominantBaseline="central" fontSize={grande ? 28 : 20} fontWeight="800" fill="#F4F8FC">
          {pct}%
        </text>
      </svg>
      <div style={{ fontSize: grande ? 16 : 13, fontWeight: 700, color: "#F4F8FC", textAlign: "center", lineHeight: 1.3, maxWidth: grande ? 190 : 150 }}>
        {nombre}
      </div>
      <div style={{ fontSize: grande ? 14 : 12, color: "#AFCBE3", fontWeight: 700 }}>{total} carpetas</div>
      <div style={{ fontSize: grande ? 14 : 13, textAlign: "center", display: "flex", gap: 10, marginTop: 4 }}>
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
  const fontPct = Math.round(size * 0.22);
  const fontLabel = Math.max(14, Math.round(size * 0.09));
  const fontCount = Math.max(12, Math.round(size * 0.07));

  // El anillo arranca en 0% y se llena hasta su valor real apenas se monta,
  // igual que en modo presentación — para que se note como progreso animado.
  const [avanzado, setAvanzado] = useState(false);
  useEffect(() => {
    const t = setTimeout(() => setAvanzado(true), 100);
    return () => clearTimeout(t);
  }, []);
  const offset = circumference - ((avanzado ? pct : 0) / 100) * circumference;

  return (
    <button
      className="uros-tarjeta-viva"
      onClick={onClick}
      style={{
        display: "flex",
        flexDirection: "column",
        alignItems: "center",
        gap: Math.round(size * 0.05),
        padding: "18px 14px",
        borderRadius: 14,
        border: `2.5px solid ${active ? color : "#2E9FD866"}`,
        background: active ? color + "25" : "#0C2B4A",
        cursor: "pointer",
        transition: "all .15s ease",
        width: tamano ? "100%" : "auto",
        boxShadow: active ? `0 0 20px ${color}66` : "0 4px 14px rgba(0,0,0,.3)",
      }}
    >
      <svg width={size} height={size}>
        <circle cx={size / 2} cy={size / 2} r={radius} fill="none" stroke="#06182B" strokeWidth={stroke} />
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
          style={{ transition: "stroke-dashoffset 1.1s cubic-bezier(.16,1,.3,1)", filter: `drop-shadow(0 0 5px ${color}99)` }}
          transform={`rotate(-90 ${size / 2} ${size / 2})`}
        />
        {/* Arco chiquito que gira sin parar, para dar sensación de "actualizando en vivo" */}
        <circle
          cx={size / 2}
          cy={size / 2}
          r={radius}
          fill="none"
          stroke="#F4F8FC"
          strokeWidth={Math.max(2, stroke * 0.28)}
          strokeLinecap="round"
          strokeDasharray={`${circumference * 0.09} ${circumference * 0.91}`}
          opacity="0.75"
        >
          <animateTransform
            attributeName="transform"
            type="rotate"
            from={`0 ${size / 2} ${size / 2}`}
            to={`360 ${size / 2} ${size / 2}`}
            dur="1.3s"
            repeatCount="indefinite"
          />
        </circle>
        <text x="50%" y="50%" textAnchor="middle" dominantBaseline="central" fontSize={fontPct} fontWeight="700" fill="#F4F8FC">
          <AnimatedPercent value={pct} />
        </text>
      </svg>
      <div style={{ fontSize: fontLabel, fontWeight: 700, color: active ? color : "#F4F8FC", textAlign: "center", marginTop: 4 }}>
        {area}
      </div>
      <div style={{ fontSize: fontCount, color: "#AFCBE3", fontWeight: 700 }}>{total} carpetas</div>
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
    <div style={{ display: "flex", flexDirection: "column", gap: 8, background: "#0C2B4A", border: `2.5px solid #2E9FD8`, borderRadius: 12, padding: "18px 22px", marginBottom: 14, boxShadow: `0 4px 20px rgba(46,159,216,.2)` }}>
      <div style={{ fontSize: 16, fontWeight: 800, color: "#2E9FD8", textTransform: "uppercase", letterSpacing: 0.5 }}>
        {area}
      </div>
      <div style={{ fontSize: 20, fontWeight: 700, color: "#F4F8FC", lineHeight: 1.3 }}>
        <span style={{ color: "#2E9FD8", fontWeight: 800 }}>{stats.completas}</span> completas de <strong style={{ color: "#F4F8FC" }}>{stats.total}</strong> carpetas
      </div>
      <div style={{ fontSize: 15, fontWeight: 600, color: "#AFCBE3", display: "flex", gap: 12, flexWrap: "wrap", alignItems: "center" }}>
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
  const color = pct >= 1 ? "#2E9FD8" : pct > 0 ? "#e67e22" : "#c0392b";

  return (
    <svg width={size} height={size}>
      <circle cx={size / 2} cy={size / 2} r={radius} fill="none" stroke="#06182B" strokeWidth={stroke} />
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
    border: `1.5px solid ${active ? color : "#E2B659"}`,
    background: active ? color + "33" : "#0C2B4A",
    color: active ? color : "#AFCBE3",
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

// Envuelve useCountUp en un componente propio para poder usarlo dentro de
// un .map() (llamar hooks dentro de un loop directamente rompe las reglas
// de React; como componente aparte, cada instancia tiene su propio hook).
function AnimatedPercent({ value }) {
  const animado = useCountUp(value);
  return <>{animado}%</>;
}

function Card({ label, value, color, grande }) {
  const valorAnimado = useCountUp(value);
  return (
    <div
      className={grande ? "uros-tarjeta-viva uros-tarjeta-respira" : "uros-tarjeta-viva"}
      style={{
        background: "#0C2B4A",
        borderRadius: 12,
        padding: grande ? "26px" : "18px",
        border: `1.5px solid ${color}55`,
        borderTop: `4px solid ${color}`,
        boxShadow: `0 0 22px ${color}33`,
        ...(grande ? { "--glow": `${color}55` } : {}),
      }}
    >
      <div style={{ fontSize: grande ? 52 : 28, fontWeight: 700, color: "#F4F8FC", textShadow: `0 0 14px ${color}66` }}>{valorAnimado}</div>
      <div style={{ fontSize: grande ? 16 : 12, color: "#AFCBE3", letterSpacing: 0.3 }}>{label}</div>
    </div>
  );
}

function TendenciaChart({ historial, grande, actividadPorDia }) {
  const altoLinea = grande ? 540 : 280;
  const altoBarras = grande ? 130 : 70;
  const alto = altoLinea + altoBarras;
  
  const anchoPunto = grande ? 80 : 55;
  const paddingIzq = 60;
  const paddingDer = 40;
  const anchoMinimo = grande ? 1100 : 720;
  const ancho = Math.max(anchoMinimo, paddingIzq + paddingDer + historial.length * anchoPunto);
  const paddingArriba = 24;

  // La línea se "dibuja" de izquierda a derecha, el área aparece detrás, y
  // las barritas de incidencias crecen desde abajo — todo arranca apenas se
  // monta el gráfico, no es un dibujo estático.
  const [avanzado, setAvanzado] = useState(false);
  const [fluyendo, setFluyendo] = useState(false);
  useEffect(() => {
    const t1 = setTimeout(() => setAvanzado(true), 120);
    const t2 = setTimeout(() => setFluyendo(true), 1600);
    return () => {
      clearTimeout(t1);
      clearTimeout(t2);
    };
  }, []);

  return (
    <div
      style={{
        background: "#0C2B4A",
        border: "2px solid #2E9FD866",
        borderRadius: 12,
        padding: grande ? "28px 32px" : "16px 18px",
        boxShadow: "0 4px 20px rgba(0,0,0,.5)",
        minWidth: 0,
        overflow: "hidden",
      }}
    >
      <div style={{ fontSize: grande ? 20 : 15, fontWeight: 700, color: "#F4F8FC", marginBottom: 14, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
        <span>📈 Tendencia de avance {historial.length > 0 ? `(${historial.length} días)` : ""}</span>
        <span style={{ fontSize: 11, color: "#AFCBE3", fontWeight: 400 }}>Desliza horizontalmente ↔</span>
      </div>

      {historial.length < 2 ? (
        <div style={{ fontSize: 12, color: "#AFCBE3", padding: "20px 0" }}>
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
            <div style={{ overflowX: "auto", overflowY: "hidden", width: "100%", paddingBottom: 8, scrollbarWidth: "thin", scrollbarColor: "#E2B659 #0C2B4A" }}>
              <svg viewBox={`0 0 ${ancho} ${alto}`} style={{ width: `${ancho}px`, height: `${alto}px`, display: "block" }}>
                {[0, 25, 50, 75, 100].map((v) => {
                  const y = paddingArriba + altoLinea - paddingArriba - (v / 100) * (altoLinea - paddingArriba * 2);
                  return (
                    <g key={v}>
                      <line x1={paddingIzq - 10} y1={y} x2={ancho - paddingDer} y2={y} stroke="#2E9FD833" strokeWidth="1" strokeDasharray="3,4" />
                      <text x={paddingIzq - 16} y={y + 4} textAnchor="end" fontSize={grande ? 14 : 12} fill="#AFCBE3" fontWeight="600">
                        {v}%
                      </text>
                    </g>
                  );
                })}

                <path
                  d={pathArea}
                  fill="url(#tendenciaGradientGold)"
                  opacity={avanzado ? 0.45 : 0}
                  style={{ transition: "opacity 1.1s ease .3s" }}
                />
                <path
                  d={pathLinea}
                  fill="none"
                  stroke="#2E9FD8"
                  strokeWidth={grande ? "4.5" : "3.5"}
                  strokeLinecap="round"
                  strokeLinejoin="round"
                  pathLength="1"
                  strokeDasharray="1"
                  strokeDashoffset={avanzado ? 0 : 1}
                  style={{ transition: "stroke-dashoffset 1.4s cubic-bezier(.16,1,.3,1)", filter: "drop-shadow(0 0 6px rgba(46,159,216,.6))" }}
                />
                {puntos.map((p, i) => {
                  const esUltimo = i === puntos.length - 1;
                  return (
                    <circle
                      key={i}
                      cx={p.x}
                      cy={p.y}
                      r={esUltimo ? (grande ? 7 : 5.5) : (grande ? 5 : 3.5)}
                      fill="#2E9FD8"
                      opacity={avanzado ? 1 : 0}
                      style={{
                        transition: `opacity .4s ease ${0.15 + i * 0.03}s, r .4s ease`,
                        filter: esUltimo ? "drop-shadow(0 0 8px #2E9FD8)" : "none",
                      }}
                    />
                  );
                })}
                {avanzado && (
                  <circle
                    cx={puntos[puntos.length - 1].x}
                    cy={puntos[puntos.length - 1].y}
                    r={grande ? 7 : 5.5}
                    fill="none"
                    stroke="#2E9FD8"
                    strokeWidth="2"
                    className="uros-punto-pulso"
                  />
                )}
                {/* Punto brillante que recorre toda la línea sin parar */}
                {fluyendo && (
                  <circle r={grande ? 5.5 : 4} fill="#F4F8FC" style={{ filter: "drop-shadow(0 0 4px #F4F8FC)" }}>
                    <animateMotion dur="3.4s" repeatCount="indefinite" path={pathLinea} />
                  </circle>
                )}

                <defs>
                  <linearGradient id="tendenciaGradientGold" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="0%" stopColor="#2E9FD8" />
                    <stop offset="100%" stopColor="#0C2B4A" stopOpacity="0" />
                  </linearGradient>
                </defs>

                <text x={puntos[puntos.length - 1].x} y={puntos[puntos.length - 1].y - 14} textAnchor="end" fontSize={grande ? 18 : 15} fontWeight="700" fill="#2E9FD8">
                  {puntos[puntos.length - 1].pct}%
                </text>

                <line x1={paddingIzq - 10} y1={altoLinea + 8} x2={ancho - paddingDer} y2={altoLinea + 8} stroke="#2E9FD844" strokeWidth="1" />
                <text x={paddingIzq} y={altoLinea + 18} fontSize={grande ? 12 : 11} fill="#AFCBE3" fontWeight="700">
                  INCIDENCIAS DEL DRIVE POR DÍA
                </text>
                {puntos.map((p, i) => {
                  const alturaBarrita = Math.max(3, (p.incidencias / maxIncidencias) * (altoBarras - 28));
                  const alturaMostrada = avanzado ? alturaBarrita : 0;
                  return (
                    <rect
                      key={i}
                      x={p.x - 4}
                      y={yBaseBarras - alturaMostrada}
                      width="8"
                      height={alturaMostrada}
                      rx="2"
                      fill={p.incidencias > 0 ? "#2E9FD8" : "#E2B65966"}
                      style={{ transition: `height .6s ease ${0.2 + i * 0.02}s, y .6s ease ${0.2 + i * 0.02}s` }}
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
                    <text key={i} x={p.x} y={alto - 2} textAnchor="middle" fontSize={grande ? 13 : 11} fill="#F4F8FC" fontWeight="600">
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

function ActividadHeatmap({ actividadPorDia, diasCustom = 84, grande, areasValidas = [] }) {
  const DIAS = diasCustom;
  const [tooltip, setTooltip] = useState(null);
  const [tooltipPos, setTooltipPos] = useState(null);
  const tooltipRef = useRef(null);
  const [diaSeleccionado, setDiaSeleccionado] = useState(null);
  const [eventosDelDia, setEventosDelDia] = useState(null);
  const [exportandoPdfDia, setExportandoPdfDia] = useState(false);
  const [menuPdfDiaAbierto, setMenuPdfDiaAbierto] = useState(false);

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
    if (count === 0) return "#06182B";
    if (count >= 11) return "#2E9FD8";
    if (count >= 4) return "#e67e22";
    return "#c0392b";
  }

  const semanas = [];
  for (let i = 0; i < dias.length; i += 7) {
    semanas.push(dias.slice(i, i + 7));
  }

  const celda = grande ? 40 : 17;
  const gap = grande ? 10 : 4;
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

  // Carpeta madre (colegio) de un evento = primer tramo de su ruta
  function areaDeEvento(ev) {
    const r = ev.ruta || "";
    // se quita el ".lnk" que traen los accesos directos de Drive en el nombre
    return (r.split(" / ")[0] || "Sin área").replace(/\.lnk$/i, "").trim();
  }

  // areaNombre = null → todo el proyecto; si no, solo esa carpeta madre
  function exportarPdfDia(areaNombre = null) {
    if (!diaSeleccionado || !eventosDelDia || eventosDelDia.length === 0) return;
    const lista = areaNombre ? eventosDelDia.filter((ev) => areaDeEvento(ev) === areaNombre) : eventosDelDia;
    if (lista.length === 0) return;
    setMenuPdfDiaAbierto(false);
    setExportandoPdfDia(true);
    // pequeña pausa para que el botón muestre "Generando..." antes de armar el PDF
    setTimeout(() => {
      try {
        generarReporteActividadDiaPDF(formatearFechaLarga(diaSeleccionado.fecha), diaSeleccionado.key, lista, { areaNombre });
      } catch (err) {
        alert(`No se pudo generar el PDF: ${err.message}`);
      } finally {
        setExportandoPdfDia(false);
      }
    }, 60);
  }

  function mostrarTooltip(e, texto) {
    const rect = e.currentTarget.getBoundingClientRect();
    setTooltip({ anclaX: rect.left + rect.width / 2, anclaY: rect.top, texto });
  }

  return (
    <div
      style={{
        background: "#0C2B4A",
        border: "2px solid #2E9FD866",
        borderRadius: 12,
        padding: grande ? "22px 26px" : "16px 18px",
        overflowX: "auto",
        boxShadow: "0 4px 20px rgba(0,0,0,.5)",
      }}
    >
      <div style={{ fontSize: grande ? 18 : 14, fontWeight: 700, color: "#F4F8FC", marginBottom: grande ? 18 : 10 }}>
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
                    className={`uros-celda-heatmap${esHoy ? " uros-celda-hoy" : ""}`}
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
              <div style={{ fontSize: 13, fontWeight: 700, color: "#AFCBE3", textTransform: "uppercase", letterSpacing: 0.4 }}>
                Resumen del período
              </div>
              {tiposOrdenados.map((tipo) => (
                <div key={tipo} style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 15 }}>
                  <span
                    style={{
                      width: 28,
                      height: 28,
                      borderRadius: "50%",
                      background: (EVENTO_COLOR[tipo] || "#2E9FD8") + "33",
                      border: `1.5px solid ${EVENTO_COLOR[tipo] || "#2E9FD8"}`,
                      color: EVENTO_COLOR[tipo] || "#2E9FD8",
                      display: "flex",
                      alignItems: "center",
                      justifyContent: "center",
                      fontSize: 13,
                      flexShrink: 0,
                    }}
                  >
                    {EVENTO_ICONO[tipo] || "•"}
                  </span>
                  <strong style={{ color: "#F4F8FC" }}>{conteoPorTipoTotal[tipo]}</strong>
                  <span style={{ color: "#AFCBE3" }}>{EVENTO_LABEL[tipo] || tipo}</span>
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
            background: "#0C2B4A",
            border: "2px solid #2E9FD8",
            color: "#F4F8FC",
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
            className="uros-fade-in"
            style={{
              background: "#0C2B4A",
              border: "2.5px solid #2E9FD8",
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
                borderBottom: "1.5px solid #2E9FD866",
              }}
            >
              <div>
                <div style={{ fontSize: 18, fontWeight: 700, color: "#F4F8FC" }}>
                  Actividad del {formatearFechaLarga(diaSeleccionado.fecha)}
                </div>
                <div style={{ fontSize: 12, color: "#AFCBE3", marginTop: 2 }}>
                  {diaSeleccionado.count} evento{diaSeleccionado.count !== 1 ? "s" : ""} registrado{diaSeleccionado.count !== 1 ? "s" : ""}
                </div>
              </div>
              <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap", justifyContent: "flex-end" }}>
                <div style={{ position: "relative" }}>
                  <button
                    onClick={() => setMenuPdfDiaAbierto((v) => !v)}
                    disabled={exportandoPdfDia || !eventosDelDia || eventosDelDia.length === 0}
                    style={{
                      fontSize: 14,
                      fontWeight: 700,
                      padding: "7px 14px",
                      borderRadius: 8,
                      border: "1.5px solid #8FD3F4",
                      background: "#0F3558",
                      color: "#CFEFFF",
                      cursor: exportandoPdfDia || !eventosDelDia || eventosDelDia.length === 0 ? "not-allowed" : "pointer",
                      opacity: !eventosDelDia || eventosDelDia.length === 0 ? 0.5 : 1,
                    }}
                    title="Exportar este registro del día a PDF (quién, a qué hora, en qué carpeta y tipo de archivo)"
                  >
                    {exportandoPdfDia ? "Generando PDF..." : `📄 Exportar PDF ${menuPdfDiaAbierto ? "▲" : "▼"}`}
                  </button>
                  {menuPdfDiaAbierto && eventosDelDia && eventosDelDia.length > 0 && (
                    <div
                      style={{
                        position: "absolute",
                        top: "calc(100% + 6px)",
                        right: 0,
                        zIndex: 60,
                        background: "#06182B",
                        border: "2px solid #8FD3F4",
                        borderRadius: 12,
                        minWidth: 300,
                        boxShadow: "0 8px 24px rgba(0,0,0,.6)",
                        overflow: "hidden",
                      }}
                    >
                      <div style={{ padding: "9px 14px", fontSize: 11, fontWeight: 700, color: "#CFEFFF", borderBottom: "1px solid #8FD3F455", textTransform: "uppercase", letterSpacing: 0.5 }}>
                        Elige qué exportar
                      </div>
                      {(() => {
                        const conteo = {};
                        for (const ev of eventosDelDia) {
                          const a = areaDeEvento(ev);
                          conteo[a] = (conteo[a] || 0) + 1;
                        }
                        // Solo las carpetas madre reales del proyecto (los colegios), no cada especialidad
                        const opciones = [{ nombre: null, texto: "📁 Todo el proyecto", n: eventosDelDia.length }].concat(
                          areasValidas.map((a) => ({ nombre: a, texto: `🏫 ${a}`, n: conteo[a] || 0 }))
                        );
                        return opciones.map((o) => (
                          <button
                            key={o.nombre || "todo"}
                            onClick={() => o.n > 0 && exportarPdfDia(o.nombre)}
                            style={{
                              display: "flex",
                              justifyContent: "space-between",
                              gap: 12,
                              width: "100%",
                              textAlign: "left",
                              padding: "11px 14px",
                              fontSize: 13,
                              fontWeight: 600,
                              background: "transparent",
                              border: "none",
                              borderBottom: "1px solid #8FD3F422",
                              color: "#CFEFFF",
                              opacity: o.n > 0 ? 1 : 0.45,
                              cursor: o.n > 0 ? "pointer" : "not-allowed",
                            }}
                            onMouseEnter={(e) => (e.currentTarget.style.background = "#8FD3F422")}
                            onMouseLeave={(e) => (e.currentTarget.style.background = "transparent")}
                          >
                            <span>{o.texto}</span>
                            <span style={{ opacity: 0.7 }}>{o.n}</span>
                          </button>
                        ));
                      })()}
                    </div>
                  )}
                </div>
                <button
                  onClick={() => setDiaSeleccionado(null)}
                  style={{
                    fontSize: 14,
                    padding: "7px 14px",
                    borderRadius: 8,
                    border: "1.5px solid #E2B659",
                    background: "transparent",
                    color: "#AFCBE3",
                    cursor: "pointer",
                  }}
                >
                  ✕ Cerrar
                </button>
              </div>
            </div>
            <div style={{ overflowY: "auto", padding: "16px 26px 26px", flex: 1 }}>
              {eventosDelDia === null ? (
                <div style={{ textAlign: "center", padding: "40px 0", color: "#AFCBE3" }}>
                  Cargando eventos del día...
                </div>
              ) : eventosDelDia.length === 0 ? (
                <div style={{ textAlign: "center", padding: "40px 0", color: "#AFCBE3" }}>
                  No se encontraron eventos detallados para esta fecha.
                </div>
              ) : (
                <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                  {eventosDelDia.map((ev) => {
                    const color = EVENTO_COLOR[ev.tipo] || "#2E9FD8";
                    const icono = EVENTO_ICONO[ev.tipo] || "•";
                    const fechaEv = ev.timestamp?.toDate ? ev.timestamp.toDate() : null;
                    const horaStr = fechaEv ? fechaEv.toLocaleTimeString("es-PE", { hour: "2-digit", minute: "2-digit", second: "2-digit" }) : "";
                    return (
                      <div
                        key={ev.id}
                        style={{
                          background: "#06182B",
                          border: "1.5px solid #2E9FD866",
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
                          <div style={{ fontSize: 14, color: "#F4F8FC" }}>
                            <strong>{ev.usuario || "Usuario"}</strong> <span style={{ color }}>{EVENTO_LABEL[ev.tipo] || ev.tipo}</span> <strong style={{ color: "#2E9FD8" }}>{ev.item}</strong>
                          </div>
                          {ev.ruta && (
                            <div style={{ fontSize: 12, color: "#AFCBE3", marginTop: 3, wordBreak: "break-all" }}>
                              📁 {ev.ruta}
                            </div>
                          )}
                          <div style={{ fontSize: 11, color: "#2E9FD8", marginTop: 6, display: "flex", justifyContent: "space-between" }}>
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
