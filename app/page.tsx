"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import {
  directionsFor,
  discoverGtfsFeeds,
  GtfsDataset,
  GtfsFeedOption,
  importLine,
  LineImport,
  parseGtfsFeed,
  servicesFor,
} from "./lib/gtfs";

type Stop = { code: string; name: string; times: string[] };
type Layout = "map-table" | "table";
type LineKind = "AMB" | "Exprés" | "TMB";
type BadgeInk = "white" | "black";
type PeriodMode = "single" | "all";
type AppView = "home" | "projects" | "editor" | "contact";
type RouteShape = "auto" | "normal" | "circular";
type ScheduleMode = "auto" | "exact" | "frequency";

const initialStops: Stop[] = [
  { code: "1042", name: "Plaça de la Vila", times: ["06:20", "06:50", "07:20", "07:50", "08:20", "08:50"] },
  { code: "1047", name: "Mercat Central", times: ["06:24", "06:54", "07:24", "07:54", "08:24", "08:54"] },
  { code: "1051", name: "Av. Catalunya", times: ["06:29", "06:59", "07:29", "07:59", "08:29", "08:59"] },
  { code: "1058", name: "Hospital Comarcal", times: ["06:35", "07:05", "07:35", "08:05", "08:35", "09:05"] },
  { code: "1064", name: "Estació d’autobusos", times: ["06:42", "07:12", "07:42", "08:12", "08:42", "09:12"] },
  { code: "1070", name: "Parc Tecnològic", times: ["06:48", "07:18", "07:48", "08:18", "08:48", "09:18"] },
];

const layoutNames: Record<Layout, string> = {
  "map-table": "A4 · recorregut + horaris",
  table: "A4 · horaris complets",
};

const kindColors: Record<LineKind, { main: string; pastel: string }> = {
  AMB: { main: "#FFD800", pastel: "#FFF3A6" },
  Exprés: { main: "#93D500", pastel: "#E4F5B8" },
  TMB: { main: "#E30613", pastel: "#F9C9CC" },
};

function contrastText(hex: string) {
  const value = hex.replace("#", "");
  const rgb = [0, 2, 4].map((index) => parseInt(value.slice(index, index + 2), 16) / 255);
  const luminance = rgb.map((channel) => channel <= 0.03928 ? channel / 12.92 : Math.pow((channel + 0.055) / 1.055, 2.4));
  return 0.2126 * luminance[0] + 0.7152 * luminance[1] + 0.0722 * luminance[2] > 0.45 ? "#111111" : "#ffffff";
}

function parseCSV(text: string): Stop[] {
  const lines = text.trim().split(/\r?\n/).filter(Boolean);
  if (lines.length < 2) return [];
  const separator = lines[0].includes(";") ? ";" : ",";
  const rows = lines.map((line) => line.split(separator).map((c) => c.trim().replace(/^"|"$/g, "")));
  const headers = rows[0].map((h) => h.toLowerCase());
  const codeI = headers.findIndex((h) => h.includes("codi") || h === "code");
  const nameI = headers.findIndex((h) => h.includes("parada") || h.includes("stop") || h.includes("nom"));
  return rows.slice(1).map((row, i) => ({
    code: row[codeI] || String(1000 + i),
    name: row[nameI] || `Parada ${i + 1}`,
    times: row.filter((v, idx) => idx !== codeI && idx !== nameI && /^\d{1,2}[:.]\d{2}$/.test(v)).map((v) => v.replace(".", ":")),
  })).filter((s) => s.name);
}

type PagePlan = {
  stops: Stop[];
  stopOffset: number;
  tripOffset: number;
  current: number;
  page: number;
  total: number;
  section: string;
};

function paginate(stops: Stop[], current: number, layout: Layout, frequency: boolean): PagePlan[] {
  const rowLimit = layout === "table" ? 54 : 46;
  const columnLimit = layout === "table" ? 16 : 10;
  const tripCount = Math.max(...stops.map((stop) => stop.times.length), 0);
  const stopChunks = Array.from({ length: Math.max(1, Math.ceil(stops.length / rowLimit)) }, (_, index) => ({
    offset: index * rowLimit,
    stops: stops.slice(index * rowLimit, (index + 1) * rowLimit),
  }));
  const tripOffsets = frequency
    ? [0]
    : Array.from({ length: Math.max(1, Math.ceil(tripCount / columnLimit)) }, (_, index) => index * columnLimit);
  const raw = stopChunks.flatMap((stopChunk) => tripOffsets.map((tripOffset) => {
    const pageStops = stopChunk.stops.map((stop) => ({ ...stop, times: frequency ? stop.times : stop.times.slice(tripOffset, tripOffset + columnLimit) }));
    const localCurrent = current >= stopChunk.offset && current < stopChunk.offset + stopChunk.stops.length ? current - stopChunk.offset : -1;
    const from = stopChunk.offset + 1;
    const to = stopChunk.offset + stopChunk.stops.length;
    const timeLabel = frequency || tripOffsets.length === 1 ? "" : ` · expedicions ${tripOffset + 1}–${Math.min(tripOffset + columnLimit, tripCount)}`;
    return { stops: pageStops, stopOffset: stopChunk.offset, tripOffset, current: localCurrent, page: 0, total: 0, section: `Parades ${from}–${to}${timeLabel}` };
  }));
  return raw.map((page, index) => ({ ...page, page: index + 1, total: raw.length }));
}

function timeMinutes(value: string) {
  const clean = value.replace("⁺", "");
  const [hours, minutes] = clean.split(":").map(Number);
  if (!Number.isFinite(hours) || !Number.isFinite(minutes)) return -1;
  return hours * 60 + minutes + (value.includes("⁺") ? 24 * 60 : 0);
}

function frequencyBands(times: string[]) {
  const values = times.filter((value) => value !== "—").map(timeMinutes).filter((value) => value >= 0).sort((a, b) => a - b);
  if (!values.length) return [];
  const windows = [
    { label: "Primeres hores", from: 0, to: 7 * 60 },
    { label: "Matí", from: 7 * 60, to: 12 * 60 },
    { label: "Migdia", from: 12 * 60, to: 16 * 60 },
    { label: "Tarda", from: 16 * 60, to: 20 * 60 },
    { label: "Vespre i nit", from: 20 * 60, to: 48 * 60 },
  ];
  return windows.flatMap((window) => {
    const departures = values.filter((value) => value >= window.from && value < window.to);
    if (!departures.length) return [];
    const gaps = departures.slice(1).map((value, index) => value - departures[index]).filter((gap) => gap > 0 && gap < 180);
    const sorted = gaps.sort((a, b) => a - b);
    const median = sorted.length ? sorted[Math.floor(sorted.length / 2)] : 0;
    const format = (minutes: number) => `${String(Math.floor((minutes % (24 * 60)) / 60)).padStart(2, "0")}:${String(minutes % 60).padStart(2, "0")}`;
    return [{ label: window.label, range: `${format(departures[0])}–${format(departures[departures.length - 1])}`, headway: median ? `Cada ${median} min` : "Una sortida", count: departures.length }];
  });
}

export default function Home() {
  const [appView, setAppView] = useState<AppView>("home");
  const [stops, setStops] = useState(initialStops);
  const [layout, setLayout] = useState<Layout>("map-table");
  const [routeShape, setRouteShape] = useState<RouteShape>("auto");
  const [detectedCircular, setDetectedCircular] = useState(false);
  const [scheduleMode, setScheduleMode] = useState<ScheduleMode>("auto");
  const [kind, setKind] = useState<LineKind>("Exprés");
  const [night, setNight] = useState(false);
  const [gradient, setGradient] = useState(true);
  const [current, setCurrent] = useState(2);
  const [lineCode, setLineCode] = useState("E12.2");
  const [lineColor, setLineColor] = useState("#93D500");
  const [hexDraft, setHexDraft] = useState("#93D500");
  const [customColorOpen, setCustomColorOpen] = useState(false);
  const [badgeInk, setBadgeInk] = useState<BadgeInk>("white");
  const [origin, setOrigin] = useState("Vila Nova");
  const [destination, setDestination] = useState("Barcelona");
  const [period, setPeriod] = useState("Dilluns a divendres feiners");
  const [periodMode, setPeriodMode] = useState<PeriodMode>("single");
  const [operator, setOperator] = useState("Operador de transport");
  const [validity, setValidity] = useState("Darrera actualització: 15.09.2026");
  const [contact, setContact] = useState("012");
  const [infoUrl, setInfoUrl] = useState("mobilitat.gencat.cat");
  const [accessible, setAccessible] = useState(true);
  const [activeTab, setActiveTab] = useState<"design" | "data">("design");
  const fileRef = useRef<HTMLInputElement>(null);
  const legacyFileRef = useRef<HTMLInputElement>(null);
  const [feedOptions, setFeedOptions] = useState<GtfsFeedOption[]>([]);
  const [dataset, setDataset] = useState<GtfsDataset | null>(null);
  const [selectedRoute, setSelectedRoute] = useState("");
  const [selectedDirection, setSelectedDirection] = useState("0");
  const [selectedService, setSelectedService] = useState("");
  const [gtfsStatus, setGtfsStatus] = useState<"idle" | "reading" | "ready" | "error">("idle");
  const [gtfsMessage, setGtfsMessage] = useState("");
  const [importWarnings, setImportWarnings] = useState<string[]>([]);
  const [importNotes, setImportNotes] = useState<string[]>([]);
  const [importedTrips, setImportedTrips] = useState(0);
  const [batchPrint, setBatchPrint] = useState(false);
  const [allPeriodImports, setAllPeriodImports] = useState<LineImport[]>([]);
  useEffect(() => {
    const saved = localStorage.getItem("emma-project");
    if (!saved) return;
    const restore = window.setTimeout(() => {
      try {
        const data = JSON.parse(saved);
        if (data.stops?.length) setStops(data.stops);
        if (data.lineCode) setLineCode(data.lineCode);
        if (data.lineColor) setLineColor(data.lineColor);
        if (data.badgeInk === "white" || data.badgeInk === "black") setBadgeInk(data.badgeInk);
        if (data.origin) setOrigin(data.origin);
        if (data.destination) setDestination(data.destination);
        if (data.operator) setOperator(data.operator);
        if (data.validity) setValidity(data.validity.replace(/^Vigent des del\s*/i, "Darrera actualització: "));
        if (data.contact) setContact(data.contact);
        if (data.infoUrl) setInfoUrl(data.infoUrl);
      } catch { /* Ignore an invalid local draft. */ }
    }, 0);
    return () => window.clearTimeout(restore);
  }, []);

  useEffect(() => {
    localStorage.setItem("emma-project", JSON.stringify({ stops, lineCode, lineColor, badgeInk, origin, destination, operator, validity, contact, infoUrl }));
  }, [stops, lineCode, lineColor, badgeInk, origin, destination, operator, validity, contact, infoUrl]);

  useEffect(() => {
    const syncDraft = window.setTimeout(() => setHexDraft(lineColor.toUpperCase()), 0);
    return () => window.clearTimeout(syncDraft);
  }, [lineColor]);

  useEffect(() => {
    const finishPrint = () => setBatchPrint(false);
    window.addEventListener("afterprint", finishPrint);
    return () => window.removeEventListener("afterprint", finishPrint);
  }, []);

  const allTrips = useMemo(() => {
    const max = Math.max(...stops.map((s) => s.times.length), 0);
    return Array.from({ length: max }, (_, trip) => stops.map((s) => s.times[trip] || "—"));
  }, [stops]);
  const effectiveCircular = routeShape === "circular" || (routeShape === "auto" && detectedCircular);
  const useFrequency = scheduleMode === "frequency" || (scheduleMode === "auto" && allTrips.length > 18);
  const effectiveLayout = layout;
  const periodVariants = useMemo(() => periodMode === "all" && allPeriodImports.length
    ? allPeriodImports.map((item) => ({ label: item.period, stops: item.stops, notes: item.notes }))
    : [{ label: period, stops, notes: importNotes }], [periodMode, allPeriodImports, period, stops, importNotes]);
  const pages = useMemo(() => periodVariants.flatMap((variant) => {
    const selected = stops[current];
    const variantCurrent = Math.max(0, variant.stops.findIndex((stop) => (selected?.code && stop.code === selected.code) || stop.name === selected?.name));
    return paginate(variant.stops, variantCurrent, effectiveLayout, useFrequency);
  }), [periodVariants, stops, current, effectiveLayout, useFrequency]);
  const renderedVersions = useMemo(() => batchPrint
    ? periodVariants.flatMap((variant) => variant.stops.flatMap((stop, stopIndex) => paginate(variant.stops, stopIndex, effectiveLayout, useFrequency).map((page) => ({ page, stop, stopIndex, periodLabel: variant.label, periodNotes: variant.notes }))))
    : periodVariants.flatMap((variant) => {
      const selected = stops[current];
      const stopIndex = Math.max(0, variant.stops.findIndex((stop) => (selected?.code && stop.code === selected.code) || stop.name === selected?.name));
      return paginate(variant.stops, stopIndex, effectiveLayout, useFrequency).map((page) => ({ page, stop: variant.stops[stopIndex], stopIndex, periodLabel: variant.label, periodNotes: variant.notes }));
    }), [batchPrint, periodVariants, stops, current, effectiveLayout, useFrequency]);
  const directionOptions = dataset ? directionsFor(dataset, selectedRoute) : [];
  const serviceOptions = dataset ? servicesFor(dataset, selectedRoute, selectedDirection) : [];

  const updateStop = (index: number, key: "name" | "code", value: string) => {
    setStops((old) => old.map((s, i) => i === index ? { ...s, [key]: value } : s));
  };

  const removeStop = (index: number) => {
    if (stops.length <= 1) return;
    setStops((old) => old.filter((_, i) => i !== index));
    setCurrent((selected) => {
      if (selected > index) return selected - 1;
      if (selected === index) return Math.min(index, stops.length - 2);
      return selected;
    });
  };

  const applyGtfsLine = (source: GtfsDataset, routeId: string, direction: string, serviceId: string) => {
    setGtfsStatus("reading");
    setGtfsMessage("Construint el patró principal i els horaris…");
    setAllPeriodImports([]);
    setPeriodMode("single");
    window.setTimeout(() => {
      try {
        const imported = importLine(source, routeId, direction, serviceId);
        setStops(imported.stops);
        setCurrent(0);
        setLineCode(imported.lineCode);
        setLineColor(imported.lineColor);
        setBadgeInk(imported.lineTextColor);
        setOrigin(imported.origin);
        setDestination(imported.destination);
        setOperator(imported.operator);
        setPeriod(imported.period);
        setValidity(imported.validity || "Darrera actualització: no informada");
        setDetectedCircular(imported.circular);
        setNight(imported.night);
        setAccessible(imported.accessible);
        setImportWarnings(imported.warnings);
        setImportNotes(imported.notes);
        setImportedTrips(imported.trips);
        setGtfsStatus("ready");
        setGtfsMessage(`${imported.stops.length} parades · ${imported.trips} expedicions · ${imported.circular ? "línia circular" : "línia normal"}`);
      } catch (error) {
        setGtfsStatus("error");
        setGtfsMessage(error instanceof Error ? error.message : "No s’ha pogut generar aquesta línia.");
      }
    }, 20);
  };

  const chooseDataset = (option: GtfsFeedOption) => {
    setGtfsStatus("reading");
    setGtfsMessage("Llegint línies, calendaris i parades…");
    window.setTimeout(() => {
      try {
        const parsed = parseGtfsFeed(option);
        const firstRoute = parsed.routes[0];
        if (!firstRoute) throw new Error("Aquest feed no conté línies d’autobús.");
        const direction = directionsFor(parsed, firstRoute.id)[0] || "0";
        const service = servicesFor(parsed, firstRoute.id, direction)[0]?.id || "";
        setDataset(parsed);
        setSelectedRoute(firstRoute.id);
        setSelectedDirection(direction);
        setSelectedService(service);
        applyGtfsLine(parsed, firstRoute.id, direction, service);
      } catch (error) {
        setGtfsStatus("error");
        setGtfsMessage(error instanceof Error ? error.message : "No s’ha pogut llegir el GTFS.");
      }
    }, 20);
  };

  const handleGtfs = async (file?: File) => {
    if (!file) return;
    setGtfsStatus("reading");
    setGtfsMessage("Validant el paquet GTFS…");
    setDataset(null);
    setFeedOptions([]);
    try {
      const options = discoverGtfsFeeds(await file.arrayBuffer(), file.name);
      setFeedOptions(options);
      if (options.length === 1) chooseDataset(options[0]);
      else {
        setGtfsStatus("idle");
        setGtfsMessage(`Hem trobat ${options.length} feeds. Tria quin vols importar.`);
      }
    } catch (error) {
      setGtfsStatus("error");
      setGtfsMessage(error instanceof Error ? error.message : "El fitxer no és un GTFS vàlid.");
    } finally {
      if (fileRef.current) fileRef.current.value = "";
    }
  };

  const changeRoute = (routeId: string) => {
    if (!dataset) return;
    const direction = directionsFor(dataset, routeId)[0] || "0";
    const service = servicesFor(dataset, routeId, direction)[0]?.id || "";
    setSelectedRoute(routeId);
    setSelectedDirection(direction);
    setSelectedService(service);
    applyGtfsLine(dataset, routeId, direction, service);
  };

  const changeDirection = (direction: string) => {
    if (!dataset) return;
    const service = servicesFor(dataset, selectedRoute, direction)[0]?.id || "";
    setSelectedDirection(direction);
    setSelectedService(service);
    applyGtfsLine(dataset, selectedRoute, direction, service);
  };

  const changeService = (service: string) => {
    if (!dataset) return;
    setSelectedService(service);
    applyGtfsLine(dataset, selectedRoute, selectedDirection, service);
  };

  const loadAllPeriods = () => {
    setPeriodMode("all");
    if (!dataset) return;
    setGtfsStatus("reading");
    setGtfsMessage("Agrupant els períodes de servei en una mateixa publicació…");
    window.setTimeout(() => {
      try {
        const seen = new Set<string>();
        const representativeServices = serviceOptions.filter((service) => {
          if (seen.has(service.label)) return false;
          seen.add(service.label);
          return true;
        });
        const imports = representativeServices.map((service) => importLine(dataset, selectedRoute, selectedDirection, service.id));
        setAllPeriodImports(imports);
        setGtfsStatus("ready");
        setGtfsMessage(`${imports.length} períodes agrupats · ${imports.reduce((total, item) => total + item.trips, 0)} expedicions analitzades`);
      } catch (error) {
        setPeriodMode("single");
        setGtfsStatus("error");
        setGtfsMessage(error instanceof Error ? error.message : "No s’han pogut agrupar els períodes.");
      }
    }, 20);
  };

  const handleFile = async (file?: File) => {
    if (!file) return;
    if (/\.xlsx?$/i.test(file.name)) {
      const workbook = XLSX.read(await file.arrayBuffer(), { type: "array" });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const parsed = parseCSV(XLSX.utils.sheet_to_csv(sheet));
      if (parsed.length) { setStops(parsed); setCurrent(0); }
      else alert("No he trobat files vàlides. Usa: codi, parada, 06:20, 06:50…");
      return;
    }
    const parsed = parseCSV(await file.text());
    if (parsed.length) { setStops(parsed); setCurrent(0); }
    else alert("No he trobat files vàlides. Usa: codi, parada, 06:20, 06:50…");
  };

  const exportVersions = () => {
    setBatchPrint(true);
    window.setTimeout(() => window.print(), 350);
  };

  if (appView === "home") return <><SiteHeader view={appView} onNavigate={setAppView} /><HomePage onNavigate={setAppView} /></>;
  if (appView === "projects") return <><SiteHeader view={appView} onNavigate={setAppView} /><ProjectsPage onNavigate={setAppView} /></>;
  if (appView === "contact") return <><SiteHeader view={appView} onNavigate={setAppView} /><ContactPage /></>;

  return (
    <main className="app-shell">
      <SiteHeader view={appView} onNavigate={setAppView} />
      <header className="topbar editor-subheader">
        <div className="brand">
          <button className="back-projects" onClick={() => setAppView("projects")} aria-label="Torna als projectes">←</button>
          <nav className="workspace-path" aria-label="Ubicació actual"><button onClick={() => setAppView("projects")}>Projectes</button><i>/</i><span>Horaris</span><i>/</i><b>{lineCode}</b></nav>
        </div>
        <div className="top-actions">
          <span className="saved"><i /> Desat en local</span>
          <button className="ghost" onClick={() => localStorage.setItem("emma-project", JSON.stringify({ stops, lineCode, lineColor, badgeInk, origin, destination, operator, validity, contact, infoUrl }))}>Desa</button>
          <button className="ghost" onClick={() => window.print()}>Previsualitza PDF</button>
          <button className="primary" onClick={exportVersions}>Genera {periodVariants.reduce((total, variant) => total + variant.stops.length, 0)} versions <span>↗</span></button>
        </div>
      </header>

      <section className="workspace">
        <aside className="sidebar">
          <div className="project-line"><span>Projecte</span><button>Emma · línia {lineCode}⌄</button></div>
          <div className="tabs">
            <button className={activeTab === "design" ? "active" : ""} onClick={() => setActiveTab("design")}>Disseny</button>
            <button className={activeTab === "data" ? "active" : ""} onClick={() => setActiveTab("data")}>Dades <b>{stops.length}</b></button>
          </div>

          {activeTab === "design" ? <div className="panel-scroll">
            <section className="tool-block">
              <h3>Informació de línia</h3>
              <label className="section-label first-section">Servei</label>
              <div className="segmented">{(["AMB", "Exprés", "TMB"] as LineKind[]).map((v) => <button key={v} className={kind === v ? "on" : ""} onClick={() => { setKind(v); setLineColor(kindColors[v].main); }}>{v}</button>)}</div>
              <details className="color-accordion">
                <summary>Altres colors</summary>
                <div className="color-accordion-body">
                  <div className="color-presets" aria-label="Colors de línia">
                    {["#FFD800", "#93D500", "#E30613", "#008C81"].map((color) => <button key={color} aria-label={`Color ${color}`} className={lineColor.toUpperCase() === color ? "active" : ""} style={{ background: color }} onClick={() => setLineColor(color)} />)}
                    <button className={`add-color ${customColorOpen ? "active" : ""}`} aria-label="Afegeix un color personalitzat" aria-expanded={customColorOpen} onClick={() => setCustomColorOpen(!customColorOpen)}>＋</button>
                  </div>
                  {customColorOpen && <div className="custom-color-row">
                    <input className="custom-color-swatch" aria-label="Selector de color personalitzat" type="color" value={lineColor} onChange={(e) => setLineColor(e.target.value)} />
                    <label><span>Hexadecimal</span><input value={hexDraft} maxLength={7} spellCheck={false} onChange={(e) => { const value = e.target.value; setHexDraft(value); if (/^#[0-9a-fA-F]{6}$/.test(value)) setLineColor(value); }} /></label>
                  </div>}
                  <label className="section-label">Text del caixetí de línia</label>
                  <div className="ink-choice" aria-label="Color del text del caixetí"><button className={badgeInk === "white" ? "on" : ""} onClick={() => setBadgeInk("white")}>Blanc</button><button className={badgeInk === "black" ? "on" : ""} onClick={() => setBadgeInk("black")}>Negre</button></div>
                </div>
              </details>
              <Field label="Identificador de línia" value={lineCode} onChange={setLineCode} />
              <div className="two"><Field label="Origen" value={origin} onChange={setOrigin} /><Field label="Destí" value={destination} onChange={setDestination} /></div>
              <Field label="Operador" value={operator} onChange={setOperator} />
            </section>

            <section className="tool-block">
              <h3>Format</h3>
              <div className="layout-grid">{(Object.keys(layoutNames) as Layout[]).map((item) => <button key={item} onClick={() => setLayout(item)} className={layout === item ? "selected" : ""}><span className={`layout-icon ${item}`} />{layoutNames[item]}</button>)}</div>
              <label className="section-label">Geometria del recorregut</label>
              <div className="segmented shape-choice"><button className={routeShape === "auto" ? "on" : ""} onClick={() => setRouteShape("auto")}>Auto</button><button className={routeShape === "normal" ? "on" : ""} onClick={() => setRouteShape("normal")}>Normal</button><button className={routeShape === "circular" ? "on" : ""} onClick={() => setRouteShape("circular")}>Circular</button></div>
              <label className="section-label">Tractament de l’horari</label>
              <div className="segmented shape-choice"><button className={scheduleMode === "auto" ? "on" : ""} onClick={() => setScheduleMode("auto")}>Auto</button><button className={scheduleMode === "exact" ? "on" : ""} onClick={() => setScheduleMode("exact")}>Exacte</button><button className={scheduleMode === "frequency" ? "on" : ""} onClick={() => setScheduleMode("frequency")}>Freqüència</button></div>
              <div className="toggle-row"><span><b>Mode de servei</b><small>{night ? "Nocturn" : "Diürn"}</small></span><button className={`switch ${night ? "on" : ""}`} onClick={() => setNight(!night)}><i /></button></div>
              <div className="toggle-row"><span><b>Capçalera corporativa</b><small>{gradient ? "Degradat 3285 → 2299" : "Vermell corporatiu"}</small></span><button className={`switch ${gradient ? "on" : ""}`} onClick={() => setGradient(!gradient)}><i /></button></div>
            </section>

            <section className="tool-block">
              <h3>Període</h3>
              <div className="period-choice"><button className={periodMode === "single" ? "on" : ""} onClick={() => setPeriodMode("single")}><b>Un període</b><small>Mostra un únic horari</small></button><button className={periodMode === "all" ? "on" : ""} onClick={loadAllPeriods}><b>Tots a la peça</b><small>Una secció per calendari</small></button></div>
              {periodMode === "single" && <select value={period} onChange={(e) => setPeriod(e.target.value)}><option>Dilluns a divendres feiners</option><option>Dissabtes feiners</option><option>Diumenges i festius</option><option>Servei especial</option></select>}
            </section>

            <section className="tool-block">
              <h3>Footer</h3>
              <Field label="Darrera actualització" value={validity} onChange={setValidity} />
              <div className="two"><Field label="Web" value={infoUrl} onChange={setInfoUrl} /><Field label="Telèfon" value={contact} onChange={setContact} /></div>
              <div className="toggle-row"><span><b>Servei accessible</b><small>Mostra informació PMR</small></span><button className={`switch ${accessible ? "on" : ""}`} onClick={() => setAccessible(!accessible)}><i /></button></div>
            </section>
            <div className="info-card"><b>Sistema modular</b><p>Els blocs buits desapareixen i la composició redistribueix l’espai automàticament.</p></div>
          </div> : <div className="panel-scroll data-panel">
            <div className="gtfs-intro"><span>01 · FONT DE DADES</span><h3>Importa el GTFS</h3><p>Emma detecta línies, sentits, períodes, recorreguts circulars i densitat d’horaris.</p></div>
            <button className={`upload gtfs-upload ${gtfsStatus === "reading" ? "is-reading" : ""}`} onClick={() => fileRef.current?.click()} disabled={gtfsStatus === "reading"}><span>{gtfsStatus === "reading" ? "…" : "↑"}</span><b>{gtfsStatus === "reading" ? "Analitzant el GTFS" : "Puja un fitxer GTFS.zip"}</b><small>El processament es fa localment al navegador</small></button>
            <input ref={fileRef} type="file" accept=".zip,application/zip" hidden onChange={(e) => handleGtfs(e.target.files?.[0])} />
            {gtfsMessage && <div className={`import-message ${gtfsStatus}`}><i />{gtfsMessage}</div>}
            {feedOptions.length > 1 && !dataset && <div className="feed-picker"><b>Feeds disponibles</b>{feedOptions.map((feed) => <button key={feed.name} onClick={() => chooseDataset(feed)}><span>▤</span><span><b>{feed.name.replace(/\.zip$/i, "")}</b><small>Obre i analitza aquest feed</small></span><i>→</i></button>)}</div>}
            {dataset && <>
              <div className="gtfs-summary"><div><strong>{dataset.diagnostics.routes}</strong><span>línies bus</span></div><div><strong>{dataset.diagnostics.stops.toLocaleString("ca")}</strong><span>parades</span></div><div><strong>{dataset.diagnostics.trips.toLocaleString("ca")}</strong><span>viatges</span></div></div>
              <section className="gtfs-selection"><span>02 · SELECCIÓ</span><label><b>Línia</b><select value={selectedRoute} onChange={(event) => changeRoute(event.target.value)}>{dataset.routes.map((route) => <option value={route.id} key={route.id}>{route.shortName || route.id} · {route.longName}</option>)}</select></label><div className="two"><label><b>Sentit</b><select value={selectedDirection} onChange={(event) => changeDirection(event.target.value)}>{directionOptions.map((direction) => <option value={direction} key={direction}>Sentit {Number(direction) + 1}</option>)}</select></label><label><b>Període</b><select value={selectedService} onChange={(event) => changeService(event.target.value)}>{serviceOptions.map((service) => <option value={service.id} key={service.id}>{service.label} · {service.count}</option>)}</select></label></div></section>
              <section className="diagnosis"><header><span>03 · DIAGNOSI</span><b>Peça preparada</b></header><dl><div><dt>Recorregut</dt><dd>{detectedCircular ? "Circular" : "Normal"}</dd></div><div><dt>Expedicions</dt><dd>{importedTrips}</dd></div><div><dt>Horari</dt><dd>{useFrequency ? "Freqüències" : "Exacte"}</dd></div><div><dt>Sortida</dt><dd>{pages.length} {pages.length === 1 ? "pàgina" : "pàgines"}</dd></div></dl>{importWarnings.length > 0 && <ul>{importWarnings.map((warning) => <li key={warning}>{warning}</li>)}</ul>}</section>
            </>}
            <details className="legacy-import"><summary>Importació manual de suport</summary><button className="upload" onClick={() => legacyFileRef.current?.click()}><span>↑</span><b>Puja CSV o Excel</b><small>Només per a dades sense GTFS</small></button><input ref={legacyFileRef} type="file" accept=".csv,.xlsx,.xls" hidden onChange={(e) => handleFile(e.target.files?.[0])} /></details>
            <div className="stop-editor-head"><b>Parades</b><span>{stops.length} parades</span></div>
            {stops.map((stop, i) => <div className={`stop-editor ${i === current ? "current" : ""}`} key={i}>
              <button className="select-stop" onClick={() => setCurrent(i)} aria-label={`Marca ${stop.name} com a parada actual`}>{i === current ? "●" : "○"}</button>
              <input value={stop.code} onChange={(e) => updateStop(i, "code", e.target.value)} />
              <input value={stop.name} onChange={(e) => updateStop(i, "name", e.target.value)} />
              <button className="remove-stop" onClick={() => removeStop(i)} disabled={stops.length <= 1} aria-label={`Elimina la parada ${stop.name}`} title="Elimina la parada">×</button>
            </div>)}
            <button className="add-stop" onClick={() => setStops([...stops, { code: "", name: "Nova parada", times: [] }])}>＋ Afegeix parada</button>
          </div>}
        </aside>

        <section className="canvas-area">
          <div className="canvas-toolbar">
            <div><button>−</button><span>78%</span><button>＋</button></div>
            <span>{layoutNames[layout]} · {effectiveCircular ? "Circular" : "Normal"} · {useFrequency ? "Freqüències" : "Exacte"}</span>
            <div><button>⌗</button><button>↗</button></div>
          </div>
          <div className={`paper-stack layout-${effectiveLayout}`}>
            {renderedVersions.map(({ page, stop, stopIndex, periodLabel, periodNotes }, versionIndex) => <TimetablePage
              key={`${versionIndex}-${stopIndex}-${page.stopOffset}-${page.tripOffset}`}
              page={page}
              layout={effectiveLayout}
              circular={effectiveCircular}
              frequency={useFrequency}
              currentStop={stop}
              lineCode={lineCode}
              lineColor={lineColor}
              badgeInk={badgeInk}
              origin={origin}
              destination={destination}
              operator={operator}
              validity={validity}
              period={periodLabel}
              night={night}
              gradient={gradient}
              notes={periodNotes}
              infoUrl={infoUrl}
              contact={contact}
              accessible={accessible}
            />)}
          </div>
        </section>
      </section>
    </main>
  );
}

function TimetablePage({ page, layout, circular, frequency, currentStop, lineCode, lineColor, badgeInk, origin, destination, operator, validity, period, night, gradient, notes, infoUrl, contact, accessible }: {
  page: PagePlan;
  layout: Layout;
  circular: boolean;
  frequency: boolean;
  currentStop?: Stop;
  lineCode: string;
  lineColor: string;
  badgeInk: BadgeInk;
  origin: string;
  destination: string;
  operator: string;
  validity: string;
  period: string;
  night: boolean;
  gradient: boolean;
  notes: string[];
  infoUrl: string;
  contact: string;
  accessible: boolean;
}) {
  const trips = useMemo(() => {
    const max = Math.max(...page.stops.map((stop) => stop.times.length), 0);
    return Array.from({ length: max }, (_, trip) => page.stops.map((stop) => stop.times[trip] || "—"));
  }, [page.stops]);
  const currentTimes = currentStop?.times || [];
  return <div className={`paper-wrap layout-${layout}`}>
    <article className={`timetable ${night ? "night" : ""} ${gradient ? "gradient" : "red"}`} style={{ "--line": lineColor, "--line-ink": contrastText(lineColor), "--pastel": `color-mix(in srgb, ${lineColor} 24%, white)` } as React.CSSProperties}>
      <header className="sheet-header">
        <div className="route-title"><div className="line-badge" style={{ background: lineColor, color: badgeInk === "white" ? "#ffffff" : "#111111" }}>{lineCode}</div><div><span className="operator-line"><small>OPERAT PER</small><b>{operator.toUpperCase()}</b></span><h1>{origin} <i>→</i> {destination}</h1><p>{validity}</p></div></div>
        <div className="stop-code"><small>CODI DE PARADA</small><b>{currentStop?.code || "—"}</b></div>
      </header>
      {page.total > 1 && <div className="continuation"><span>{page.section}</span><b>{page.page} / {page.total}</b></div>}
      <div className="sheet-body">
        {layout !== "table" && <RouteDiagram stops={page.stops} current={page.current} circle={circular} color={lineColor} />}
        <div className="schedule-area">
          <div className="you-are"><i className="stop-marker-icon" aria-hidden="true" /><b>{currentStop?.name}</b><HereTag /></div>
          <div className="schedule-title"><div><small>{frequency ? "FREQÜÈNCIES ORIENTATIVES" : "HORARIS DE PAS APROXIMATS"}</small><h2>{period}</h2></div><span>{night ? "☾" : "☀"} {night ? "Servei nocturn" : "Servei diürn"}</span></div>
          {frequency ? <FrequencyTable times={currentTimes} /> : <ScheduleTable stops={page.stops} trips={trips} current={page.current} showHereLabel />}
          <div className="notes"><b>Informació important</b><p>{notes[0] || "Els horaris estan subjectes a les incidències i contratemps de la via pública. Consulta les alteracions del servei abans de viatjar."}</p></div>
        </div>
      </div>
      <footer className="sheet-footer"><div className="logo-slot">LOGO ESQUERRA</div><p><b>{infoUrl}</b><br/>Informació: {contact}{accessible ? " · ♿ Servei accessible" : ""}</p><div className="logo-slot right">LOGO DRETA</div></footer>
    </article>
  </div>;
}

function SiteHeader({ view, onNavigate }: { view: AppView; onNavigate: (view: AppView) => void }) {
  return <header className="site-header">
    <button className="site-logo" onClick={() => onNavigate("home")}><i className="emma-pulse" aria-hidden="true" /><strong>Emma</strong><span>Creador d’horaris</span></button>
    <nav aria-label="Navegació principal">
      <button className={view === "home" ? "active" : ""} onClick={() => onNavigate("home")}>Inici</button>
      <button className={view === "projects" || view === "editor" ? "active" : ""} onClick={() => onNavigate("projects")}>Projectes</button>
      <button className={view === "contact" ? "active" : ""} onClick={() => onNavigate("contact")}>Contacte</button>
    </nav>
    <button className="account-chip"><span>MG</span><b>Compte demo</b><i>⌄</i></button>
  </header>;
}

function HomePage({ onNavigate }: { onNavigate: (view: AppView) => void }) {
  return <main className="marketing-page">
    <section className="hero">
      <div className="hero-copy"><span className="eyebrow">HORARIS DE TRANSPORT, BEN FETS</span><h1>Del GTFS al cartell,<br/><em>en un sol lloc.</em></h1><p>Crea, adapta i genera totes les versions d’un horari de bus amb un sistema modular pensat per a qualsevol xarxa de Catalunya.</p><div className="hero-actions"><button className="cta" onClick={() => onNavigate("projects")}>Importa un GTFS <span>→</span></button><button onClick={() => onNavigate("contact")}>Parlem-ne</button></div><div className="hero-proof"><span>✓ GTFS estàndard</span><span>✓ Plantilles adaptatives</span><span>✓ PDFs per a cada parada</span></div></div>
      <div className="hero-product" aria-label="Vista prèvia del producte"><div className="mini-window"><div className="mini-bar"><i/><i/><i/><span>Emma / E12.2</span></div><div className="mini-body"><aside><b>Informació de línia</b><span/><span/><b>Format</b><span/><span/></aside><div className="mini-canvas"><article><header><strong>E12.2</strong><div><small>OPERAT PER BUS</small><b>Vila Nova → Barcelona</b></div></header><section><div className="mini-line"><i/><i/><i className="on"/><i/><i/></div><div className="mini-table"><b>HORARIS DE PAS</b>{[1,2,3,4,5].map(i=><span key={i}/>)}</div></section></article></div></div></div></div>
    </section>
    <section className="home-features"><div><span>01</span><h2>Puja el GTFS</h2><p>Emma valida el feed i detecta línies, sentits, calendaris, variants i recorreguts circulars.</p></div><div><span>02</span><h2>Revisa la proposta</h2><p>El sistema escull format, densitat, freqüències i paginació sense deixar blocs buits.</p></div><div><span>03</span><h2>Genera les versions</h2><p>Crea una peça per cada parada i exporta-les juntes, preparades per imprimir.</p></div></section>
  </main>;
}

function ProjectsPage({ onNavigate }: { onNavigate: (view: AppView) => void }) {
  const [activeMenu, setActiveMenu] = useState<string | null>(null);
  const [shareTarget, setShareTarget] = useState<string | null>(null);
  const [folders, setFolders] = useState([
    { name: "AMB", count: 8, tone: "green" },
    { name: "TMB", count: 4, tone: "red" },
    { name: "Exprés.cat", count: 12, tone: "teal" },
  ]);
  const createFolder = () => {
    const name = window.prompt("Nom de la carpeta nova");
    if (name?.trim()) setFolders((currentFolders) => [...currentFolders, { name: name.trim(), count: 0, tone: "blue" }]);
  };
  return <main className="projects-page">
    <aside className="library-nav"><button className="new-project" onClick={() => onNavigate("editor")}>＋ Nou projecte</button><nav><button className="active">▦ Tots els projectes</button><button>◷ Recents</button><button>☆ Favorits</button><button>⇄ Compartit amb mi</button></nav><div className="folder-title"><b>Carpetes</b><button>＋</button></div><nav><button>▱ AMB</button><button>▱ Exprés.cat</button><button>▱ TMB</button><button>▱ Proves</button></nav><nav className="library-bottom"><button>♲ Paperera</button></nav></aside>
    <section className="library-content"><header><div><span>ESPAI DE TREBALL</span><h1>Projectes</h1><p>Organitza, comparteix i genera els teus horaris.</p></div><div className="library-tools"><label>⌕ <input placeholder="Cerca projectes" /></label><button className="create-folder" onClick={createFolder}><span>＋</span> Nova carpeta</button><button className="create-file" onClick={() => onNavigate("editor")}><span>＋</span> Nou projecte</button><button>Ordena ⌄</button><button aria-label="Vista en graella">▦</button></div></header>
      <section className="quick-folders">{folders.map((folder) => <div className="quick-folder-card menu-host" key={folder.name}><button className="folder-open"><i className={`folder ${folder.tone}`}/><span><b>{folder.name}</b><small>{folder.count} projectes</small></span></button><button className="more-button" aria-label={`Més opcions per a ${folder.name}`} onClick={() => setActiveMenu(activeMenu === `folder-${folder.name}` ? null : `folder-${folder.name}`)}>···</button>{activeMenu === `folder-${folder.name}` && <ItemMenu kind="folder" onClose={() => setActiveMenu(null)} onShare={() => { setShareTarget(folder.name); setActiveMenu(null); }} />}</div>)}</section>
      <div className="section-heading"><h2>Projectes recents</h2><button>Veure’ls tots</button></div>
      <section className="project-grid"><ProjectCard color="#93D500" code="E12.2" title="Vila Nova — Barcelona" meta="Editat ara mateix" onOpen={() => onNavigate("editor")} menuOpen={activeMenu === "project-e12"} onMenu={() => setActiveMenu(activeMenu === "project-e12" ? null : "project-e12")} onShare={() => { setShareTarget("Vila Nova — Barcelona"); setActiveMenu(null); }} /><ProjectCard color="#FFD800" code="B8" title="Sant Boi — Barcelona" meta="Editat ahir" onOpen={() => onNavigate("editor")} menuOpen={activeMenu === "project-b8"} onMenu={() => setActiveMenu(activeMenu === "project-b8" ? null : "project-b8")} onShare={() => { setShareTarget("Sant Boi — Barcelona"); setActiveMenu(null); }} /><ProjectCard color="#E30613" code="H12" title="Gornal — Besòs/Verneda" meta="Editat fa 3 dies" onOpen={() => onNavigate("editor")} menuOpen={activeMenu === "project-h12"} onMenu={() => setActiveMenu(activeMenu === "project-h12" ? null : "project-h12")} onShare={() => { setShareTarget("Gornal — Besòs/Verneda"); setActiveMenu(null); }} /><button className="empty-project" onClick={() => onNavigate("editor")}><span>＋</span><b>Crea un projecte</b><small>GTFS, Excel o manual</small></button></section>
    </section>
    {shareTarget && <ShareDialog target={shareTarget} onClose={() => setShareTarget(null)} />}
  </main>;
}

function ProjectCard({ color, code, title, meta, onOpen, menuOpen, onMenu, onShare }: { color: string; code: string; title: string; meta: string; onOpen: () => void; menuOpen: boolean; onMenu: () => void; onShare: () => void }) {
  return <article className="project-card menu-host"><button className="project-open" onClick={onOpen}><div className="project-preview"><div style={{ background: color }}><strong>{code}</strong><span>{title}</span></div><i/><i/><i/><i/></div><div className="project-info"><span className="file-icon">▤</span><span><b>{title}</b><small>{meta} · 6 versions</small></span></div></button><button className="more-button card-more" aria-label={`Més opcions per a ${title}`} onClick={onMenu}>···</button>{menuOpen && <ItemMenu kind="project" onClose={onMenu} onShare={onShare} />}</article>;
}

function ItemMenu({ kind, onClose, onShare }: { kind: "folder" | "project"; onClose: () => void; onShare: () => void }) {
  const action = (message: string) => { alert(`${message}. Aquesta acció es connectarà al backend a la fase següent.`); onClose(); };
  return <div className="item-menu" role="menu"><button role="menuitem" onClick={onShare}><span>↗</span> Comparteix</button><button role="menuitem" onClick={() => action("Enllaç copiat")}><span>⌁</span> Copia l’enllaç</button><hr/><button role="menuitem" onClick={() => action("Element canviat de nom")}><span>✎</span> Canvia el nom</button><button role="menuitem" onClick={() => action("Element mogut")}><span>↳</span> Mou a una carpeta</button>{kind === "project" && <button role="menuitem" onClick={() => action("Projecte duplicat")}><span>⧉</span> Duplica</button>}<hr/>{kind === "project" && <button role="menuitem" onClick={() => action("Exportació preparada")}><span>↓</span> Baixa</button>}<button className="danger" role="menuitem" onClick={() => action("Element mogut a la paperera")}><span>⌫</span> Mou a la paperera</button></div>;
}

function ShareDialog({ target, onClose }: { target: string; onClose: () => void }) {
  const [permission, setPermission] = useState("Pot editar");
  return <div className="modal-backdrop" role="presentation" onMouseDown={onClose}><section className="share-dialog" role="dialog" aria-modal="true" aria-labelledby="share-title" onMouseDown={(event) => event.stopPropagation()}><header><div><small>COMPARTEIX</small><h2 id="share-title">{target}</h2></div><button onClick={onClose} aria-label="Tanca">×</button></header><div className="share-field"><label>Persones o grups</label><div><input autoFocus type="email" placeholder="nom@organitzacio.cat"/><select value={permission} onChange={(event) => setPermission(event.target.value)}><option>Pot editar</option><option>Pot visualitzar</option></select></div><textarea rows={3} placeholder="Afegeix un missatge (opcional)"/></div><button className="send-share" onClick={() => { alert(`Invitació preparada amb el permís: ${permission}.`); onClose(); }}>Envia la invitació</button><div className="share-divider"><span>o comparteix un enllaç</span></div><div className="link-share"><div><b>Només persones convidades</b><small>Caldrà iniciar sessió per accedir-hi</small></div><button onClick={() => alert("Enllaç copiat")}>⌁ Copia l’enllaç</button></div><footer><span className="avatar-stack"><i>MG</i><i>+</i></span><button>Gestiona l’accés</button></footer></section></div>;
}

function ContactPage() {
  return <main className="contact-page"><section><span className="eyebrow">CONTACTE</span><h1>Parlem del teu<br/>sistema d’horaris.</h1><p>Explica’ns com treballeu ara, quins formats utilitzeu i quantes línies gestioneu. T’ajudarem a plantejar la millor automatització.</p><div className="contact-details"><div><small>CORREU</small><b>hola@emma.cat</b></div><div><small>ÀMBIT</small><b>Transport públic · Catalunya</b></div></div></section><form onSubmit={(e)=>{e.preventDefault(); alert("Missatge desat localment. Connectarem l’enviament quan activem el backend.");}}><div className="two"><label><span>Nom</span><input required placeholder="El teu nom" /></label><label><span>Organització</span><input placeholder="Empresa o administració" /></label></div><label><span>Correu electrònic</span><input required type="email" placeholder="nom@organitzacio.cat" /></label><label><span>En què et podem ajudar?</span><select><option>Vull conèixer Emma</option><option>Necessito automatitzar horaris</option><option>Vull aportar dades GTFS</option><option>Altres</option></select></label><label><span>Missatge</span><textarea required rows={6} placeholder="Explica’ns breument el projecte…" /></label><label className="privacy"><input type="checkbox" required /> He llegit i accepto la política de privacitat.</label><button className="cta" type="submit">Envia el missatge <span>→</span></button></form></main>;
}

function Field({ label, value, onChange }: { label: string; value: string; onChange: (s: string) => void }) {
  return <label className="field"><span>{label}</span><input value={value} onChange={(e) => onChange(e.target.value)} /></label>;
}

function RouteDiagram({ stops, current, circle, color }: { stops: Stop[]; current: number; circle: boolean; color: string }) {
  if (circle) return <CircularRouteDiagram stops={stops} current={current} color={color} />;
  const stopHeight = Math.max(18, Math.min(66, 570 / Math.max(stops.length, 1)));
  const density = stops.length > 24 ? "ultra-dense" : stops.length > 12 ? "dense" : "";
  return <section className={`route-diagram ${density}`}><div className="diagram-heading"><small>RECORREGUT</small><b>{stops.length} parades · 34 min</b></div><div className="route-line" style={{ "--route": color, "--stop-height": `${stopHeight}px` } as React.CSSProperties}>{stops.map((s, i) => <div className={`route-stop ${i === current ? "here" : ""}`} key={i}><i /><div><b>{s.name}</b>{i === current && <HereTag />}</div></div>)}</div></section>;
}

function CircularRouteDiagram({ stops, current, color }: { stops: Stop[]; current: number; color: string }) {
  const midpoint = Math.ceil(stops.length / 2);
  const left = stops.slice(0, midpoint).map((stop, index) => ({ stop, index }));
  const right = stops.slice(midpoint).map((stop, index) => ({ stop, index: midpoint + index })).reverse();
  const rows = Math.max(left.length, right.length);
  const rowHeight = Math.max(22, Math.min(48, 520 / Math.max(rows, 1)));
  const currentSide = current < midpoint ? "points-left" : "points-right";
  const currentRow = current < midpoint ? current : Math.max(0, stops.length - 1 - current);
  const currentTop = `${((currentRow + 0.5) / Math.max(rows, 1)) * 100}%`;
  return <section className={`route-diagram circular-diagram ${stops.length > 24 ? "ultra-dense" : stops.length > 14 ? "dense" : ""}`} style={{ "--route": color, "--loop-row": `${rowHeight}px` } as React.CSSProperties}>
    <div className="diagram-heading"><small>RECORREGUT CIRCULAR</small><b>{stops.length} parades · sentit indicat</b></div>
    <div className="circular-route">
      <div className="circular-branch left-branch">{left.map(({ stop, index }) => <div className={`circular-stop ${index === current ? "here" : ""}`} key={index}><div><b>{stop.name}</b></div><i /></div>)}</div>
      <div className="loop-track"><span className="arrow down" aria-hidden="true">↓</span><span className="arrow up" aria-hidden="true">↑</span>{stops[current] && <HereTag className="loop-current-tag" direction={currentSide === "points-right" ? "right" : "left"} style={{ top: currentTop }} />}</div>
      <div className="circular-branch right-branch">{right.map(({ stop, index }) => <div className={`circular-stop ${index === current ? "here" : ""}`} key={index}><i /><div><b>{stop.name}</b></div></div>)}</div>
    </div>
  </section>;
}

function FrequencyTable({ times, compact = false }: { times: string[]; compact?: boolean }) {
  const bands = frequencyBands(times);
  return <div className={`frequency-table ${compact ? "compact" : ""}`}>
    {bands.length ? bands.map((band) => <div className="frequency-row" key={band.label}><span><small>{band.label}</small><b>{band.range}</b></span><strong>{band.headway}</strong><em>{band.count} sortides</em></div>) : <div className="frequency-empty">No hi ha prou expedicions per calcular una freqüència.</div>}
    <p>Consulta les hores exactes al web o al codi QR de l’operador.</p>
  </div>;
}

function ScheduleTable({ stops, trips, current, showHereLabel }: { stops: Stop[]; trips: string[][]; current: number; showHereLabel: boolean }) {
  const density = trips.length > 16 ? "hours-ultra" : trips.length > 10 ? "hours-dense" : "";
  return <div className={`table-scroll ${density}`}><table className="transposed-table"><thead><tr><th className="stop-column"><small>RECORREGUT</small><span>Parada</span></th>{trips.map((trip, i) => <th key={i}><span>{trip.find((value) => value !== "—") || "—"}</span></th>)}</tr></thead><tbody>{stops.map((stop, stopIndex) => <tr className={stopIndex === current ? "active-row" : ""} key={stopIndex}><th scope="row" className="stop-name">{showHereLabel && stopIndex === current && <HereTag compact />}<span>{stop.name}</span></th>{trips.map((trip, tripIndex) => <td key={tripIndex}>{trip[stopIndex] || "—"}</td>)}</tr>)}</tbody></table></div>;
}

function HereTag({ compact = false, direction = "left", className = "", style }: { compact?: boolean; direction?: "left" | "right"; className?: string; style?: React.CSSProperties }) {
  return <span className={`here-tag points-${direction}${compact ? " is-compact" : ""}${className ? ` ${className}` : ""}`} style={style}>
    <svg viewBox="0 0 74.8 18.4" preserveAspectRatio="none" aria-hidden="true"><path d="M72.5,18.4H9.4c-.8,0-1.6-.4-2.2-1L1.5,11.3c-1.1-1.1-1.1-2.9,0-4.1L7.2,1c.6-.6,1.4-1,2.2-1h63.1c1.7,0,3,1.3,3,3v12.4c0,1.7-1.3,3-3,3Z" /></svg>
    <span>ETS AQUÍ</span>
  </span>;
}
