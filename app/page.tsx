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

type Stop = { code: string; name: string; times: string[]; sourceKey?: string; sourceStopId?: string };
type Layout = "map-table" | "table";
type LineKind = "AMB" | "Exprés" | "TMB";
type BadgeInk = "white" | "black";
type PeriodMode = "single" | "all";
type RouteShape = "auto" | "normal" | "circular";
type ScheduleMode = "auto" | "exact" | "frequency";

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

function mixWithWhite(hex: string, amount: number) {
  const value = hex.replace("#", "");
  if (!/^[0-9a-f]{6}$/i.test(value)) return "#e4f5b8";
  const channels = [0, 2, 4].map((index) => parseInt(value.slice(index, index + 2), 16));
  return `#${channels.map((channel) => Math.round(channel + (255 - channel) * amount).toString(16).padStart(2, "0")).join("")}`;
}

function currentDateLabel() {
  const parts = new Intl.DateTimeFormat("ca-ES", { day: "2-digit", month: "2-digit", year: "numeric" }).formatToParts(new Date());
  const value = Object.fromEntries(parts.map((part) => [part.type, part.value]));
  return `Generat: ${value.day}.${value.month}.${value.year}`;
}

function canvasWithUniformBleed(source: HTMLCanvasElement) {
  const pixelsPerMm = 6;
  const bleed = 3 * pixelsPerMm;
  const trimWidth = 210 * pixelsPerMm;
  const trimHeight = 297 * pixelsPerMm;
  const trim = document.createElement("canvas");
  trim.width = trimWidth;
  trim.height = trimHeight;
  const trimContext = trim.getContext("2d");
  if (!trimContext) throw new Error("No s’ha pogut preparar el full A4.");
  trimContext.imageSmoothingEnabled = true;
  trimContext.imageSmoothingQuality = "high";
  trimContext.drawImage(source, 0, 0, trimWidth, trimHeight);

  const output = document.createElement("canvas");
  output.width = trimWidth + bleed * 2;
  output.height = trimHeight + bleed * 2;
  const context = output.getContext("2d");
  if (!context) throw new Error("No s’han pogut preparar les sangs del PDF.");
  context.fillStyle = "#ffffff";
  context.fillRect(0, 0, output.width, output.height);
  context.drawImage(trim, bleed, bleed);

  context.drawImage(trim, 0, 0, 1, trimHeight, 0, bleed, bleed, trimHeight);
  context.drawImage(trim, trimWidth - 1, 0, 1, trimHeight, bleed + trimWidth, bleed, bleed, trimHeight);
  context.drawImage(trim, 0, 0, trimWidth, 1, bleed, 0, trimWidth, bleed);
  context.drawImage(trim, 0, trimHeight - 1, trimWidth, 1, bleed, bleed + trimHeight, trimWidth, bleed);
  context.drawImage(trim, 0, 0, 1, 1, 0, 0, bleed, bleed);
  context.drawImage(trim, trimWidth - 1, 0, 1, 1, bleed + trimWidth, 0, bleed, bleed);
  context.drawImage(trim, 0, trimHeight - 1, 1, 1, 0, bleed + trimHeight, bleed, bleed);
  context.drawImage(trim, trimWidth - 1, trimHeight - 1, 1, 1, bleed + trimWidth, bleed + trimHeight, bleed, bleed);
  return output;
}

function setPrintPageBoxes(pdf: import("jspdf").jsPDF) {
  const pointsPerMm = 72 / 25.4;
  const box = (left: number, bottom: number, right: number, top: number) => ({
    bottomLeftX: left * pointsPerMm,
    bottomLeftY: bottom * pointsPerMm,
    topRightX: right * pointsPerMm,
    topRightY: top * pointsPerMm,
  });
  const page = pdf.getCurrentPageInfo().pageContext;
  page.cropBox = box(0, 0, 216, 303);
  page.bleedBox = box(0, 0, 216, 303);
  page.trimBox = box(3, 3, 213, 300);
}

function sameStopOccurrence(a?: Stop, b?: Stop) {
  if (!a || !b) return false;
  if (a.sourceKey && b.sourceKey) return a.sourceKey === b.sourceKey;
  if (a.code && b.code) return a.code === b.code && a.name === b.name;
  return a.name === b.name;
}

function samePhysicalStop(a?: Stop, b?: Stop) {
  if (!a || !b) return false;
  if (a.sourceStopId && b.sourceStopId) return a.sourceStopId === b.sourceStopId;
  if (a.code && b.code) return a.code === b.code;
  return a.name === b.name;
}

function parseCSV(text: string): Stop[] {
  const lines = text.trim().split(/\r?\n/).filter(Boolean);
  if (lines.length < 2) return [];
  const separator = lines[0].includes(";") ? ";" : ",";
  const rows = lines.map((line) => line.split(separator).map((c) => c.trim().replace(/^"|"$/g, "")));
  const headers = rows[0].map((h) => h.toLowerCase());
  if (headers.includes("trip_id") || headers.includes("stop_sequence")) {
    throw new Error("Aquest fitxer és una taula interna del GTFS. Puja el paquet GTFS.zip complet o utilitza la plantilla Emma.");
  }
  const codeI = headers.findIndex((h) => h.includes("codi") || h === "code");
  const nameI = headers.findIndex((h) => h.includes("parada") || h.includes("stop") || h.includes("nom"));
  if (nameI < 0) throw new Error("La plantilla necessita una columna «parada» o «nom».");
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
  selectedStop?: Stop;
};

type RenderedVersion = {
  page: PagePlan;
  stop?: Stop;
  stopIndex: number;
  overview: boolean;
  periodLabel: string;
  periodNotes: string[];
};

function paginate(stops: Stop[], current: number, layout: Layout, frequency: boolean): PagePlan[] {
  const rowLimit = layout === "table" ? 44 : 36;
  const tripCount = Math.max(...stops.map((stop) => stop.times.length), 0);
  const columnLimit = layout === "table" ? (tripCount <= 20 ? 20 : 18) : (tripCount <= 12 ? 12 : 10);
  const charsPerRow = layout === "table" ? 54 : 38;
  const stopChunks: { offset: number; stops: Stop[] }[] = [];
  let chunk: Stop[] = [];
  let chunkUnits = 0;
  let chunkStart = 0;
  stops.forEach((stop, index) => {
    const units = Math.max(1, Math.ceil(stop.name.length / charsPerRow));
    if (chunk.length && chunkUnits + units > rowLimit) {
      stopChunks.push({ offset: chunkStart, stops: chunk });
      chunk = [];
      chunkUnits = 0;
      chunkStart = index;
    }
    chunk.push(stop);
    chunkUnits += units;
  });
  if (chunk.length || !stops.length) stopChunks.push({ offset: chunkStart, stops: chunk });
  const tripOffsets = frequency
    ? [0]
    : Array.from({ length: Math.max(1, Math.ceil(tripCount / columnLimit)) }, (_, index) => index * columnLimit);
  const raw = stopChunks.flatMap((stopChunk) => tripOffsets.map((tripOffset) => {
    const pageStops = stopChunk.stops.map((stop) => ({ ...stop, times: frequency ? stop.times : stop.times.slice(tripOffset, tripOffset + columnLimit) }));
    const localCurrent = current >= stopChunk.offset && current < stopChunk.offset + stopChunk.stops.length ? current - stopChunk.offset : -1;
    const from = stopChunk.offset + 1;
    const to = stopChunk.offset + stopChunk.stops.length;
    const timeLabel = frequency || tripOffsets.length === 1 ? "" : ` · expedicions ${tripOffset + 1}–${Math.min(tripOffset + columnLimit, tripCount)}`;
    return { stops: pageStops, stopOffset: stopChunk.offset, tripOffset, current: localCurrent, page: 0, total: 0, section: `Parades ${from}–${to}${timeLabel}`, selectedStop: current >= 0 ? stops[current] : undefined };
  }));
  return raw.map((page, index) => ({ ...page, page: index + 1, total: raw.length }));
}

function singlePageOverview(stops: Stop[], layout: Layout, frequency: boolean): PagePlan {
  const firstStopPage = paginate(stops, 0, layout, frequency)[0];
  const visibleTripCount = Math.max(...(firstStopPage?.stops || []).map((stop) => stop.times.length), 0);
  return {
    stops: stops.map((stop) => ({
      ...stop,
      times: frequency ? stop.times : stop.times.slice(0, visibleTripCount),
    })),
    stopOffset: 0,
    tripOffset: 0,
    current: -1,
    page: 1,
    total: 1,
    section: "Visió general",
    selectedStop: undefined,
  };
}

function afterNextPaint() {
  return new Promise<void>((resolve) => {
    window.requestAnimationFrame(() => window.requestAnimationFrame(() => resolve()));
  });
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
  const [stops, setStops] = useState<Stop[]>([]);
  const [layout, setLayout] = useState<Layout>("map-table");
  const [routeShape, setRouteShape] = useState<RouteShape>("auto");
  const [detectedCircular, setDetectedCircular] = useState(false);
  const [scheduleMode, setScheduleMode] = useState<ScheduleMode>("auto");
  const [kind, setKind] = useState<LineKind>("Exprés");
  const [night, setNight] = useState(false);
  const [gradient, setGradient] = useState(true);
  const [current, setCurrent] = useState(0);
  const [lineCode, setLineCode] = useState("");
  const [lineTitle, setLineTitle] = useState("");
  const [lineColor, setLineColor] = useState("#93D500");
  const [hexDraft, setHexDraft] = useState("#93D500");
  const [customColorOpen, setCustomColorOpen] = useState(false);
  const [badgeInk, setBadgeInk] = useState<BadgeInk>("white");
  const [origin, setOrigin] = useState("");
  const [destination, setDestination] = useState("");
  const [period, setPeriod] = useState("Dilluns a divendres feiners");
  const [periodMode, setPeriodMode] = useState<PeriodMode>("single");
  const [operator, setOperator] = useState("");
  const [validity, setValidity] = useState("");
  const [generatedOn, setGeneratedOn] = useState(currentDateLabel);
  const [contact, setContact] = useState("");
  const [infoUrl, setInfoUrl] = useState("");
  const [accessible, setAccessible] = useState(true);
  const [generalView, setGeneralView] = useState(false);
  const [leftLogo, setLeftLogo] = useState("");
  const [rightLogo, setRightLogo] = useState("");
  const [activeTab, setActiveTab] = useState<"design" | "data">("data");
  const fileRef = useRef<HTMLInputElement>(null);
  const legacyFileRef = useRef<HTMLInputElement>(null);
  const leftLogoRef = useRef<HTMLInputElement>(null);
  const rightLogoRef = useRef<HTMLInputElement>(null);
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
  const [gtfsUsesFrequencies, setGtfsUsesFrequencies] = useState(false);
  const [exportPage, setExportPage] = useState<RenderedVersion | null>(null);
  const [pdfProgress, setPdfProgress] = useState("");
  const [pdfPreview, setPdfPreview] = useState<{ url: string; filename: string; pages: string[] } | null>(null);
  const [zoom, setZoom] = useState(78);
  const [allPeriodImports, setAllPeriodImports] = useState<LineImport[]>([]);
  const [newProjectDialog, setNewProjectDialog] = useState(false);
  const [savedSnapshot, setSavedSnapshot] = useState("");
  const [editingTimes, setEditingTimes] = useState<number | null>(null);
  const [timesDraft, setTimesDraft] = useState("");
  const [draggedStop, setDraggedStop] = useState<number | null>(null);
  const [dragOverStop, setDragOverStop] = useState<number | null>(null);
  const [pendingRemoval, setPendingRemoval] = useState<{ stop: Stop; index: number } | null>(null);
  const [lastRemoved, setLastRemoved] = useState<{ stop: Stop; index: number; previousCurrent: number } | null>(null);
  const paperStackRef = useRef<HTMLDivElement>(null);
  const dragTargetRef = useRef<number | null>(null);

  const projectData = useMemo(() => ({
    stops, layout, routeShape, scheduleMode, kind, night, gradient, current, lineCode, lineTitle,
    lineColor, badgeInk, origin, destination, period, periodMode, operator, validity, generatedOn,
    contact, infoUrl, accessible, generalView, leftLogo, rightLogo,
  }), [stops, layout, routeShape, scheduleMode, kind, night, gradient, current, lineCode, lineTitle, lineColor, badgeInk, origin, destination, period, periodMode, operator, validity, generatedOn, contact, infoUrl, accessible, generalView, leftLogo, rightLogo]);
  const projectSnapshot = useMemo(() => JSON.stringify(projectData), [projectData]);
  const hasProjectData = stops.length > 0 || Boolean(lineCode.trim() || origin.trim() || destination.trim() || operator.trim());
  const isDirty = hasProjectData && projectSnapshot !== savedSnapshot;

  useEffect(() => {
    const syncDraft = window.setTimeout(() => setHexDraft(lineColor.toUpperCase()), 0);
    return () => window.clearTimeout(syncDraft);
  }, [lineColor]);

  useEffect(() => {
    if (!isDirty) return;
    const warnBeforeLeaving = (event: BeforeUnloadEvent) => event.preventDefault();
    window.addEventListener("beforeunload", warnBeforeLeaving);
    return () => window.removeEventListener("beforeunload", warnBeforeLeaving);
  }, [isDirty]);

  useEffect(() => {
    if (!lastRemoved) return;
    const timer = window.setTimeout(() => setLastRemoved(null), 7000);
    return () => window.clearTimeout(timer);
  }, [lastRemoved]);

  const allTrips = useMemo(() => {
    const max = Math.max(...stops.map((s) => s.times.length), 0);
    return Array.from({ length: max }, (_, trip) => stops.map((s) => s.times[trip] || "—"));
  }, [stops]);
  const effectiveCircular = routeShape === "circular" || (routeShape === "auto" && detectedCircular);
  const useFrequency = scheduleMode === "frequency" || (scheduleMode === "auto" && gtfsUsesFrequencies);
  const effectiveLayout = layout;
  const physicalStops = useMemo(() => stops.filter((stop, index) => stops.findIndex((candidate) => samePhysicalStop(candidate, stop)) === index), [stops]);
  const periodVariants = useMemo(() => periodMode === "all" && allPeriodImports.length
    ? allPeriodImports.map((item) => ({ label: item.period, stops: item.stops, notes: item.notes }))
    : [{ label: period, stops, notes: importNotes }], [periodMode, allPeriodImports, period, stops, importNotes]);
  const overviewVariant = useMemo(() => periodVariants.find((variant) => variant.label === period) || periodVariants[0], [periodVariants, period]);
  const pages = useMemo(() => generalView && overviewVariant
    ? [singlePageOverview(overviewVariant.stops, effectiveLayout, useFrequency)]
    : periodVariants.flatMap((variant) => {
      const selected = stops[current];
      const variantCurrent = Math.max(0, variant.stops.findIndex((stop) => sameStopOccurrence(stop, selected)));
      return paginate(variant.stops, variantCurrent, effectiveLayout, useFrequency);
    }), [periodVariants, overviewVariant, stops, current, generalView, effectiveLayout, useFrequency]);
  const currentRenderedVersions = useMemo<RenderedVersion[]>(() => {
    if (generalView && overviewVariant) {
      return [{ page: singlePageOverview(overviewVariant.stops, effectiveLayout, useFrequency), stop: undefined, stopIndex: -1, overview: true, periodLabel: overviewVariant.label, periodNotes: overviewVariant.notes }];
    }
    return periodVariants.flatMap((variant) => {
      const selected = stops[current];
      const stopIndex = Math.max(0, variant.stops.findIndex((stop) => sameStopOccurrence(stop, selected)));
      return paginate(variant.stops, stopIndex, effectiveLayout, useFrequency).map((page) => ({ page, stop: variant.stops[stopIndex], stopIndex, overview: false, periodLabel: variant.label, periodNotes: variant.notes }));
    });
  }, [periodVariants, overviewVariant, stops, current, generalView, effectiveLayout, useFrequency]);
  const allRenderedVersions = useMemo<RenderedVersion[]>(() => periodVariants.flatMap((variant) => [
      { page: singlePageOverview(variant.stops, effectiveLayout, useFrequency), stop: undefined, stopIndex: -1, overview: true, periodLabel: variant.label, periodNotes: variant.notes },
      ...physicalStops.flatMap((physicalStop) => {
        const stopIndex = Math.max(0, variant.stops.findIndex((stop) => samePhysicalStop(stop, physicalStop)));
        const stop = variant.stops[stopIndex];
        return paginate(variant.stops, stopIndex, effectiveLayout, useFrequency).map((page) => ({ page, stop, stopIndex, overview: false, periodLabel: variant.label, periodNotes: variant.notes }));
      }),
    ]), [periodVariants, physicalStops, effectiveLayout, useFrequency]);
  const renderedVersions = exportPage ? [exportPage] : currentRenderedVersions;
  const directionOptions = dataset ? directionsFor(dataset, selectedRoute) : [];
  const serviceOptions = dataset ? servicesFor(dataset, selectedRoute, selectedDirection) : [];
  const versionCount = stops.length
    ? (generalView ? 1 : periodVariants.reduce((total) => total + physicalStops.length + 1, 0))
    : 0;

  const updateStop = (index: number, key: "name" | "code", value: string) => {
    setStops((old) => old.map((s, i) => i === index ? { ...s, [key]: value } : s));
  };

  const updateStopTimes = (index: number, value: string) => {
    const times = value.split(/[;,\s]+/).map((time) => time.trim().replace(".", ":")).filter(Boolean);
    setStops((old) => old.map((stop, stopIndex) => stopIndex === index ? { ...stop, times } : stop));
  };

  const moveStop = (index: number, direction: -1 | 1) => {
    const target = index + direction;
    if (target < 0 || target >= stops.length) return;
    setStops((old) => {
      const next = [...old];
      [next[index], next[target]] = [next[target], next[index]];
      return next;
    });
    setCurrent((selected) => selected === index ? target : selected === target ? index : selected);
    setEditingTimes((selected) => selected === index ? target : selected === target ? index : selected);
  };

  const reorderStop = (from: number, to: number) => {
    if (from === to || from < 0 || to < 0 || from >= stops.length || to >= stops.length) return;
    const remapIndex = (selected: number) => {
      if (selected === from) return to;
      if (from < to && selected > from && selected <= to) return selected - 1;
      if (from > to && selected >= to && selected < from) return selected + 1;
      return selected;
    };
    setStops((old) => {
      const next = [...old];
      const [moved] = next.splice(from, 1);
      next.splice(to, 0, moved);
      return next;
    });
    setCurrent(remapIndex);
    setEditingTimes((selected) => selected === null ? null : remapIndex(selected));
  };

  const finishStopDrag = () => {
    dragTargetRef.current = null;
    setDraggedStop(null);
    setDragOverStop(null);
  };

  const trackStopDrag = (event: React.PointerEvent<HTMLButtonElement>) => {
    const group = document.elementFromPoint(event.clientX, event.clientY)?.closest<HTMLElement>(".stop-editor-group");
    const target = Number(group?.dataset.stopIndex);
    if (!Number.isInteger(target) || target === dragTargetRef.current) return;
    dragTargetRef.current = target;
    setDragOverStop(target);
  };

  const removeStop = (index: number) => {
    if (stops.length <= 1) return;
    setLastRemoved({ stop: stops[index], index, previousCurrent: current });
    setStops((old) => old.filter((_, i) => i !== index));
    setCurrent((selected) => {
      if (selected > index) return selected - 1;
      if (selected === index) return Math.min(index, stops.length - 2);
      return selected;
    });
    setEditingTimes(null);
    setTimesDraft("");
    setPendingRemoval(null);
  };

  const requestRemoveStop = (index: number) => {
    if (stops.length <= 1) return;
    setPendingRemoval({ stop: stops[index], index });
  };

  const undoRemoveStop = () => {
    if (!lastRemoved) return;
    setStops((old) => [...old.slice(0, lastRemoved.index), lastRemoved.stop, ...old.slice(lastRemoved.index)]);
    setCurrent(lastRemoved.previousCurrent);
    setLastRemoved(null);
  };

  const addStopAfterCurrent = () => {
    const index = stops.length ? Math.min(current + 1, stops.length) : 0;
    const stop: Stop = { code: "", name: "Nova parada", times: [] };
    setStops((old) => [...old.slice(0, index), stop, ...old.slice(index)]);
    setCurrent(index);
    setEditingTimes(index);
    setTimesDraft("");
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
        setLineTitle(imported.routeLongName);
        setLineColor(imported.lineColor);
        setBadgeInk(imported.lineTextColor);
        setOrigin(imported.origin);
        setDestination(imported.destination);
        setOperator(imported.operator);
        setPeriod(imported.period);
        setValidity(imported.validity);
        setGeneratedOn(currentDateLabel());
        setDetectedCircular(imported.circular);
        setNight(imported.night);
        setAccessible(imported.accessible);
        setImportWarnings(imported.warnings);
        setImportNotes(imported.notes);
        setImportedTrips(imported.trips);
        setGtfsUsesFrequencies(imported.frequencyBased);
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
    try {
      if (/\.xlsx?$/i.test(file.name)) {
        const workbook = XLSX.read(await file.arrayBuffer(), { type: "array" });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const parsed = parseCSV(XLSX.utils.sheet_to_csv(sheet));
        if (parsed.length) { setStops(parsed); setCurrent(0); }
        else throw new Error("La plantilla no conté cap parada.");
        return;
      }
      const parsed = parseCSV(await file.text());
      if (parsed.length) { setStops(parsed); setCurrent(0); }
      else throw new Error("La plantilla no conté cap parada.");
    } catch (error) {
      alert(error instanceof Error ? error.message : "No s’ha pogut importar la plantilla Emma.");
    }
  };

  const downloadEmmaTemplate = () => {
    const csv = "codi,parada,06:20,06:50\n1001,Parada d’exemple,06:20,06:50\n";
    const link = document.createElement("a");
    link.href = URL.createObjectURL(new Blob([csv], { type: "text/csv;charset=utf-8" }));
    link.download = "plantilla-emma.csv";
    link.click();
    window.setTimeout(() => URL.revokeObjectURL(link.href), 1000);
  };

  const handleLogo = (side: "left" | "right", file?: File) => {
    if (!file) return;
    if (!file.type.startsWith("image/") && !/\.svg$/i.test(file.name)) {
      alert("Puja un logo en format SVG, PNG, JPG o WebP.");
      return;
    }
    if (file.size > 1024 * 1024) {
      alert("El logo no pot superar 1 MB. Per a millor qualitat i menys pes, prioritza SVG.");
      return;
    }
    const reader = new FileReader();
    reader.onload = () => {
      const value = typeof reader.result === "string" ? reader.result : "";
      if (side === "left") setLeftLogo(value);
      else setRightLogo(value);
    };
    reader.readAsDataURL(file);
  };

  const exportPdf = async (includeAllVersions: boolean, preview: boolean) => {
    if (!stops.length || pdfProgress) return;
    const includeStopVersions = includeAllVersions && !generalView;
    const pagesToExport = includeStopVersions ? allRenderedVersions : currentRenderedVersions;
    setPdfProgress("Preparant el PDF…");
    try {
      if (!pagesToExport.length) throw new Error("No hi ha cap pàgina per exportar.");
      const [{ default: html2canvas }, { jsPDF }] = await Promise.all([import("html2canvas"), import("jspdf")]);
      await document.fonts?.ready;
      document.body.classList.add("pdf-exporting");
      const pdf = new jsPDF({ orientation: "portrait", unit: "mm", format: [216, 303], compress: true, putOnlyUsedFonts: true });
      const previewPages: string[] = [];
      pdf.setProperties({ title: `${lineCode || "Emma"} · ${lineTitle || "Horaris"}`, subject: "Horaris de transport públic", creator: "Emma · Creador d’horaris" });
      for (let index = 0; index < pagesToExport.length; index += 1) {
        setPdfProgress(`Generant pàgina ${index + 1} de ${pagesToExport.length}…`);
        setExportPage(pagesToExport[index]);
        await afterNextPaint();
        const element = paperStackRef.current?.querySelector<HTMLElement>(".paper-wrap .timetable");
        if (!element) throw new Error(`No s’ha pogut preparar la pàgina ${index + 1}.`);
        if (index > 0) pdf.addPage([216, 303], "portrait");
        const canvas = await html2canvas(element, {
          scale: 1.5,
          backgroundColor: "#ffffff",
          useCORS: true,
          logging: false,
          imageTimeout: 15000,
        });
        const printCanvas = canvasWithUniformBleed(canvas);
        const image = printCanvas.toDataURL("image/jpeg", 0.94);
        if (preview) previewPages.push(image);
        const alias = `emma-page-${index}`;
        setPrintPageBoxes(pdf);
        pdf.addImage(image, "JPEG", 0, 0, 216, 303, alias, "FAST");
        pdf.setDrawColor(0, 0, 0);
        pdf.setLineWidth(0.15);
        [[0.5, 3, 2.5, 3], [3, 0.5, 3, 2.5], [213.5, 3, 215.5, 3], [213, 0.5, 213, 2.5], [0.5, 300, 2.5, 300], [3, 300.5, 3, 302.5], [213.5, 300, 215.5, 300], [213, 300.5, 213, 302.5]].forEach(([x1, y1, x2, y2]) => pdf.line(x1, y1, x2, y2));
      }
      const filename = `emma-${(lineCode || "horaris").replace(/[^a-z0-9_-]+/gi, "-").toLowerCase()}.pdf`;
      if (preview) {
        const url = URL.createObjectURL(pdf.output("blob"));
        setPdfPreview((previous) => {
          if (previous) URL.revokeObjectURL(previous.url);
          return { url, filename, pages: previewPages };
        });
      } else {
        pdf.save(filename);
      }
    } catch (error) {
      alert(error instanceof Error ? error.message : "No s’ha pogut generar el PDF.");
    } finally {
      document.body.classList.remove("pdf-exporting");
      setExportPage(null);
      setPdfProgress("");
    }
  };

  const saveProject = () => {
    if (!hasProjectData) return;
    try {
      localStorage.setItem("emma-project", projectSnapshot);
      setSavedSnapshot(projectSnapshot);
    } catch {
      alert("No s’ha pogut desar el projecte. Redueix el pes dels logos i torna-ho a provar.");
    }
  };

  const closePdfPreview = () => {
    setPdfPreview((previous) => {
      if (previous) URL.revokeObjectURL(previous.url);
      return null;
    });
  };

  const resetProject = () => {
    setStops([]);
    setLayout("map-table");
    setRouteShape("auto");
    setDetectedCircular(false);
    setScheduleMode("auto");
    setKind("Exprés");
    setNight(false);
    setGradient(true);
    setCurrent(0);
    setLineCode("");
    setLineTitle("");
    setLineColor("#93D500");
    setBadgeInk("white");
    setOrigin("");
    setDestination("");
    setPeriod("Dilluns a divendres feiners");
    setPeriodMode("single");
    setOperator("");
    setValidity("");
    setGeneratedOn(currentDateLabel());
    setContact("");
    setInfoUrl("");
    setAccessible(true);
    setGeneralView(false);
    setLeftLogo("");
    setRightLogo("");
    setActiveTab("data");
    setFeedOptions([]);
    setDataset(null);
    setSelectedRoute("");
    setSelectedDirection("0");
    setSelectedService("");
    setGtfsStatus("idle");
    setGtfsMessage("");
    setImportWarnings([]);
    setImportNotes([]);
    setImportedTrips(0);
    setGtfsUsesFrequencies(false);
    setAllPeriodImports([]);
    setZoom(78);
    setSavedSnapshot("");
    setEditingTimes(null);
    setTimesDraft("");
    setPendingRemoval(null);
    setLastRemoved(null);
    setNewProjectDialog(false);
  };

  const requestNewProject = () => {
    if (isDirty) setNewProjectDialog(true);
    else resetProject();
  };

  return (
    <main className="app-shell">
      <SiteHeader />
      <input ref={fileRef} type="file" accept=".zip,application/zip" hidden onChange={(event) => handleGtfs(event.target.files?.[0])} />
      <input ref={leftLogoRef} type="file" accept="image/svg+xml,image/png,image/jpeg,image/webp" hidden onChange={(event) => { handleLogo("left", event.target.files?.[0]); event.currentTarget.value = ""; }} />
      <input ref={rightLogoRef} type="file" accept="image/svg+xml,image/png,image/jpeg,image/webp" hidden onChange={(event) => { handleLogo("right", event.target.files?.[0]); event.currentTarget.value = ""; }} />
      <header className="topbar editor-subheader">
        <div className="brand">
          <nav className="workspace-path" aria-label="Document actual"><span>Eina d’horaris</span><i>/</i><b>{lineCode || "Nou projecte"}</b></nav>
        </div>
        <div className="top-actions">
          <button className="new-project-action" onClick={requestNewProject}><span>＋</span> Nou projecte</button>
          <span className={`saved ${isDirty ? "dirty" : ""}`}><i /> {hasProjectData ? (isDirty ? "Canvis sense desar" : "Desat en local") : "Projecte buit"}</span>
          <button className="ghost" onClick={saveProject} disabled={!hasProjectData}>Desa</button>
          <button className="ghost" onClick={() => exportPdf(false, true)} disabled={!stops.length || Boolean(pdfProgress)}>{pdfProgress || "Previsualitza PDF"}</button>
          <button className="primary" onClick={() => exportPdf(true, false)} disabled={!stops.length || Boolean(pdfProgress)}>Genera {versionCount} {versionCount === 1 ? "versió" : "versions"} <span>↓</span></button>
        </div>
      </header>

      <section className="workspace">
        <aside className="sidebar">
          <div className="project-line"><span>Projecte</span><button>{lineCode ? `Emma · línia ${lineCode}` : "Nou projecte"}</button></div>
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
              <Field label="Nom oficial de la línia (route_long_name)" value={lineTitle} onChange={setLineTitle} />
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
              <div className="toggle-row"><span><b>Visió general</b><small>Sense parada actual ni «ETS AQUÍ»</small></span><button className={`switch ${generalView ? "on" : ""}`} aria-pressed={generalView} onClick={() => setGeneralView(!generalView)}><i /></button></div>
              <div className="toggle-row"><span><b>Mode de servei</b><small>{night ? "Nocturn" : "Diürn"}</small></span><button className={`switch ${night ? "on" : ""}`} onClick={() => setNight(!night)}><i /></button></div>
              <div className="toggle-row"><span><b>Capçalera corporativa</b><small>{gradient ? "Degradat 3285 → 2299" : "Vermell corporatiu"}</small></span><button className={`switch ${gradient ? "on" : ""}`} onClick={() => setGradient(!gradient)}><i /></button></div>
            </section>

            <section className="tool-block">
              <h3>Període</h3>
              <div className="period-choice"><button className={periodMode === "single" ? "on" : ""} onClick={() => setPeriodMode("single")}><b>Un període</b><small>Mostra un únic horari</small></button><button className={periodMode === "all" ? "on" : ""} onClick={loadAllPeriods}><b>Tots a la peça</b><small>Una secció per calendari</small></button></div>
              {periodMode === "single" && <select value={period} onChange={(e) => setPeriod(e.target.value)}><option>Dilluns a divendres feiners</option><option>Dissabtes feiners</option><option>Diumenges i festius</option><option>Servei especial</option></select>}
            </section>

            <div className="info-card"><b>Sistema modular</b><p>Els blocs buits desapareixen i la composició redistribueix l’espai automàticament.</p></div>
            <section className="tool-block footer-tool">
              <details className="footer-accordion">
                <summary>Footer</summary>
                <div className="footer-accordion-body">
                  <label className="section-label first-section">Logos</label>
                  <div className="footer-logo-controls">
                    <button onClick={() => leftLogoRef.current?.click()}>{leftLogo ? <img src={leftLogo} alt="Logo esquerre" /> : <><span>＋</span><small>Logo esquerre</small></>}</button>
                    <button onClick={() => rightLogoRef.current?.click()}>{rightLogo ? <img src={rightLogo} alt="Logo dret" /> : <><span>＋</span><small>Logo dret</small></>}</button>
                  </div>
                  <Field label="Vigència del servei" value={validity} onChange={setValidity} />
                  <Field label="Data de generació" value={generatedOn} onChange={setGeneratedOn} />
                  <div className="two"><Field label="Web" value={infoUrl} onChange={setInfoUrl} /><Field label="Telèfon" value={contact} onChange={setContact} /></div>
                  <div className="toggle-row"><span><b>Servei accessible</b><small>Mostra informació PMR</small></span><button className={`switch ${accessible ? "on" : ""}`} onClick={() => setAccessible(!accessible)}><i /></button></div>
                </div>
              </details>
            </section>
          </div> : <div className="panel-scroll data-panel">
            <div className="gtfs-intro"><span>01 · FONT DE DADES</span><h3>Importa el GTFS</h3><p>Emma detecta línies, sentits, períodes, recorreguts circulars i densitat d’horaris.</p></div>
            <button className={`upload gtfs-upload ${gtfsStatus === "reading" ? "is-reading" : ""}`} onClick={() => fileRef.current?.click()} disabled={gtfsStatus === "reading"}><span>{gtfsStatus === "reading" ? "…" : "↑"}</span><b>{gtfsStatus === "reading" ? "Analitzant el GTFS" : "Puja un fitxer GTFS.zip"}</b><small>El processament es fa localment al navegador</small></button>
            {gtfsMessage && <div className={`import-message ${gtfsStatus}`}><i />{gtfsMessage}</div>}
            {feedOptions.length > 1 && !dataset && <div className="feed-picker"><b>Feeds disponibles</b>{feedOptions.map((feed) => <button key={feed.name} onClick={() => chooseDataset(feed)}><span>▤</span><span><b>{feed.name.replace(/\.zip$/i, "")}</b><small>Obre i analitza aquest feed</small></span><i>→</i></button>)}</div>}
            {dataset && <>
              <div className="gtfs-summary"><div><strong>{dataset.diagnostics.routes}</strong><span>línies bus</span></div><div><strong>{dataset.diagnostics.stops.toLocaleString("ca")}</strong><span>parades</span></div><div><strong>{dataset.diagnostics.trips.toLocaleString("ca")}</strong><span>viatges</span></div></div>
              <section className="gtfs-selection"><span>02 · SELECCIÓ</span><label><b>Línia</b><select value={selectedRoute} onChange={(event) => changeRoute(event.target.value)}>{dataset.routes.map((route) => <option value={route.id} key={route.id}>{route.shortName || route.id} · {route.longName}</option>)}</select></label><div className="two"><label><b>Sentit</b><select value={selectedDirection} onChange={(event) => changeDirection(event.target.value)}>{directionOptions.map((direction) => <option value={direction} key={direction}>Sentit {Number(direction) + 1}</option>)}</select></label><label><b>Període</b><select value={selectedService} onChange={(event) => changeService(event.target.value)}>{serviceOptions.map((service) => <option value={service.id} key={service.id}>{service.label} · {service.count}</option>)}</select></label></div></section>
              <section className="diagnosis"><header><span>03 · DIAGNOSI</span><b>Peça preparada</b></header><dl><div><dt>Recorregut</dt><dd>{detectedCircular ? "Circular" : "Normal"}</dd></div><div><dt>Expedicions</dt><dd>{importedTrips}</dd></div><div><dt>Horari</dt><dd>{useFrequency ? "Freqüències" : "Exacte"}</dd></div><div><dt>Sortida</dt><dd>{pages.length} {pages.length === 1 ? "pàgina" : "pàgines"}</dd></div></dl>{importWarnings.length > 0 && <ul>{importWarnings.map((warning) => <li key={warning}>{warning}</li>)}</ul>}</section>
            </>}
            <details className="legacy-import"><summary>Importació manual de suport</summary><button className="upload" onClick={() => legacyFileRef.current?.click()}><span>↑</span><b>Importa una plantilla Emma</b><small>CSV o Excel amb una fila per parada</small></button><button className="download-template" onClick={downloadEmmaTemplate}>↓ Descarrega la plantilla CSV</button><input ref={legacyFileRef} type="file" accept=".csv,.xlsx,.xls" hidden onChange={(e) => handleFile(e.target.files?.[0])} /></details>
            <div className="stop-editor-head"><b>Parades</b><span>{stops.length} parades</span></div>
            {stops.map((stop, i) => <div
              className={`stop-editor-group${draggedStop === i ? " is-dragging" : ""}${dragOverStop === i && draggedStop !== i ? " is-drag-over" : ""}`}
              key={stop.sourceKey || `${stop.code}-${i}`}
              data-stop-index={i}
            >
              <div className={`stop-editor ${i === current ? "current" : ""}`}>
                <button
                  className="stop-drag-handle"
                  onPointerDown={(event) => { event.currentTarget.setPointerCapture(event.pointerId); dragTargetRef.current = i; setDraggedStop(i); setDragOverStop(i); }}
                  onPointerMove={trackStopDrag}
                  onPointerUp={(event) => { const target = dragTargetRef.current; if (target !== null) reorderStop(i, target); if (event.currentTarget.hasPointerCapture(event.pointerId)) event.currentTarget.releasePointerCapture(event.pointerId); finishStopDrag(); }}
                  onPointerCancel={finishStopDrag}
                  aria-label={`Arrossega per reordenar ${stop.name}`}
                  title="Arrossega per reordenar"
                ><span aria-hidden="true">{Array.from({ length: 6 }, (_, dot) => <i key={dot} />)}</span></button>
                <button className="select-stop" onClick={() => { setCurrent(i); setGeneralView(false); }} aria-label={`Marca ${stop.name} com a parada actual`}>{i === current && !generalView ? "●" : "○"}</button>
                <input value={stop.code} onChange={(e) => updateStop(i, "code", e.target.value)} aria-label={`Codi de ${stop.name}`} />
                <input value={stop.name} onChange={(e) => updateStop(i, "name", e.target.value)} aria-label={`Nom de la parada ${i + 1}`} />
                <div className="stop-actions">
                  <button onClick={() => moveStop(i, -1)} disabled={i === 0} aria-label={`Mou ${stop.name} amunt`} title="Mou amunt">↑</button>
                  <button onClick={() => moveStop(i, 1)} disabled={i === stops.length - 1} aria-label={`Mou ${stop.name} avall`} title="Mou avall">↓</button>
                  <button className={editingTimes === i ? "active" : ""} onClick={() => { if (editingTimes === i) { setEditingTimes(null); setTimesDraft(""); } else { setEditingTimes(i); setTimesDraft(stop.times.join(", ")); } }} aria-label={`Edita els horaris de ${stop.name}`} title="Edita els horaris">◷</button>
                  <button className="remove-stop" onClick={() => requestRemoveStop(i)} disabled={stops.length <= 1} aria-label={`Elimina la parada ${stop.name}`} title="Elimina la parada">×</button>
                </div>
              </div>
              {editingTimes === i && <label className="stop-times-editor"><span>Hores de pas, separades per comes</span><textarea value={timesDraft} onChange={(event) => { setTimesDraft(event.target.value); updateStopTimes(i, event.target.value); }} placeholder="06:20, 06:50, 07:20" /></label>}
            </div>)}
            <button className="add-stop" onClick={addStopAfterCurrent}>＋ Afegeix parada després de l’actual</button>
          </div>}
        </aside>

        <section className="canvas-area">
          <div className="canvas-toolbar">
            <div className="zoom-controls">
              <button onClick={() => setZoom((value) => Math.max(40, value - 10))} disabled={zoom <= 40} aria-label="Redueix el zoom" title="Redueix el zoom">−</button>
              <span aria-live="polite">{zoom}%</span>
              <button onClick={() => setZoom((value) => Math.min(148, value + 10))} disabled={zoom >= 148} aria-label="Augmenta el zoom" title="Augmenta el zoom">＋</button>
            </div>
            <span>{layoutNames[layout]} · {effectiveCircular ? "Circular" : "Normal"} · {useFrequency ? "Freqüències" : "Exacte"}</span>
            <div><button onClick={() => setZoom(78)} aria-label="Ajusta el full a la vista" title="Ajusta el full a la vista">⌗</button><button>↗</button></div>
          </div>
          <div ref={paperStackRef} className={`paper-stack layout-${effectiveLayout}${stops.length ? "" : " empty-project-canvas"}`} style={{ "--preview-zoom": zoom / 100, "--preview-width": `${8 * zoom}px`, "--preview-height": `${11.312 * zoom}px` } as React.CSSProperties}>
            {!stops.length ? <div className="paper-wrap empty-paper-wrap"><article className="timetable empty-timetable">
              <div className="empty-sheet-upload">
                <span className="empty-sheet-icon" aria-hidden="true">↑</span>
                <small>PROJECTE NOU</small>
                <h2>Comença amb un fitxer GTFS</h2>
                <p>Puja el paquet <b>.zip</b> de l’operador. Emma prepararà les línies, els sentits, els períodes i totes les versions per parada.</p>
                <button onClick={() => fileRef.current?.click()}>Puja un GTFS.zip <span>→</span></button>
                <em>El fitxer es processa localment al navegador</em>
              </div>
            </article></div> : renderedVersions.map(({ page, stop, stopIndex, overview, periodLabel, periodNotes }, versionIndex) => <TimetablePage
              key={`${versionIndex}-${stopIndex}-${page.stopOffset}-${page.tripOffset}`}
              page={page}
              layout={effectiveLayout}
              circular={effectiveCircular}
              frequency={useFrequency}
              currentStop={stop}
              overview={overview}
              lineCode={lineCode}
              lineTitle={lineTitle}
              lineColor={lineColor}
              badgeInk={badgeInk}
              origin={origin}
              destination={destination}
              operator={operator}
              validity={validity}
              generatedOn={generatedOn}
              period={periodLabel}
              night={night}
              gradient={gradient}
              notes={periodNotes}
              infoUrl={infoUrl}
              contact={contact}
              accessible={accessible}
              leftLogo={leftLogo}
              rightLogo={rightLogo}
              onLeftLogoClick={() => leftLogoRef.current?.click()}
              onRightLogoClick={() => rightLogoRef.current?.click()}
            />)}
          </div>
        </section>
      </section>
      {newProjectDialog && <div className="modal-backdrop" role="presentation" onMouseDown={() => setNewProjectDialog(false)}>
        <section className="new-project-dialog" role="dialog" aria-modal="true" aria-labelledby="new-project-title" onMouseDown={(event) => event.stopPropagation()}>
          <div className="dialog-icon" aria-hidden="true">＋</div>
          <small>NOU PROJECTE</small>
          <h2 id="new-project-title">Vols desar els canvis?</h2>
          <p>El projecte actual té canvis sense desar. Si continues sense desar-los, es perdran.</p>
          <div className="dialog-actions">
            <button className="quiet" onClick={() => setNewProjectDialog(false)}>Cancel·la</button>
            <button className="danger" onClick={resetProject}>Descarta i crea’n un de nou</button>
            <button className="confirm" onClick={() => { saveProject(); resetProject(); }}>Desa i crea’n un de nou</button>
          </div>
        </section>
      </div>}
      {pendingRemoval && <div className="modal-backdrop editor-confirm-backdrop" role="presentation" onMouseDown={() => setPendingRemoval(null)}>
        <section className="new-project-dialog remove-stop-dialog" role="dialog" aria-modal="true" aria-labelledby="remove-stop-title" onMouseDown={(event) => event.stopPropagation()}>
          <div className="dialog-icon" aria-hidden="true">!</div>
          <small>ACCIÓ DELICADA</small>
          <h2 id="remove-stop-title">Eliminar aquesta parada?</h2>
          <p>Eliminaràs <b>{pendingRemoval.stop.name}</b> del recorregut i també totes les seves hores de pas.</p>
          <div className="dialog-actions">
            <button className="quiet" autoFocus onClick={() => setPendingRemoval(null)}>Cancel·la</button>
            <button className="danger" onClick={() => removeStop(pendingRemoval.index)}>Elimina la parada</button>
          </div>
        </section>
      </div>}
      {pdfPreview && <div className="modal-backdrop pdf-preview-backdrop" role="presentation" onMouseDown={closePdfPreview}>
        <section className="pdf-preview-dialog" role="dialog" aria-modal="true" aria-labelledby="pdf-preview-title" onMouseDown={(event) => event.stopPropagation()}>
          <header><div><small>PDF A4 D’IMPREMTA</small><h2 id="pdf-preview-title">Previsualització</h2></div><div><a href={pdfPreview.url} download={pdfPreview.filename}>Descarrega PDF</a><button onClick={closePdfPreview} aria-label="Tanca la previsualització">×</button></div></header>
          <div className="pdf-preview-pages" role="document" aria-label="Previsualització del PDF">{pdfPreview.pages.map((pageImage, index) => <figure key={index}><img src={pageImage} alt={`Pàgina ${index + 1} del PDF`} /><figcaption>Pàgina {index + 1} de {pdfPreview.pages.length}</figcaption></figure>)}</div>
        </section>
      </div>}
      {lastRemoved && <div className="undo-toast" role="status"><span><b>Parada eliminada</b>{lastRemoved.stop.name}</span><button onClick={undoRemoveStop}>Desfés</button><button className="dismiss" onClick={() => setLastRemoved(null)} aria-label="Tanca">×</button></div>}
    </main>
  );
}

function TimetablePage({ page, layout, circular, frequency, currentStop, overview, lineCode, lineTitle, lineColor, badgeInk, origin, destination, operator, validity, generatedOn, period, night, gradient, notes, infoUrl, contact, accessible, leftLogo, rightLogo, onLeftLogoClick, onRightLogoClick }: {
  page: PagePlan;
  layout: Layout;
  circular: boolean;
  frequency: boolean;
  currentStop?: Stop;
  overview: boolean;
  lineCode: string;
  lineTitle: string;
  lineColor: string;
  badgeInk: BadgeInk;
  origin: string;
  destination: string;
  operator: string;
  validity: string;
  generatedOn: string;
  period: string;
  night: boolean;
  gradient: boolean;
  notes: string[];
  infoUrl: string;
  contact: string;
  accessible: boolean;
  leftLogo: string;
  rightLogo: string;
  onLeftLogoClick: () => void;
  onRightLogoClick: () => void;
}) {
  const trips = useMemo(() => {
    const max = Math.max(...page.stops.map((stop) => stop.times.length), 0);
    return Array.from({ length: max }, (_, trip) => page.stops.map((stop) => stop.times[trip] || "—"));
  }, [page.stops]);
  const currentTimes = currentStop?.times || page.stops[0]?.times || [];
  const overviewDensity = overview
    ? (page.stops.length > 32 || trips.length > 28 ? " overview-ultra" : page.stops.length > 22 || trips.length > 16 ? " overview-dense" : "")
    : "";
  return <div className={`paper-wrap layout-${layout}`}>
    <article className={`timetable ${night ? "night" : ""} ${gradient ? "gradient" : "red"}${overview ? ` overview-version${overviewDensity}` : ""}`} style={{ "--line": lineColor, "--line-ink": contrastText(lineColor), "--pastel": mixWithWhite(lineColor, .76), "--active-row": mixWithWhite(lineColor, .86) } as React.CSSProperties}>
      <header className="sheet-header">
        <div className="route-title"><div className="line-badge" style={{ background: lineColor, color: badgeInk === "white" ? "#ffffff" : "#111111" }}>{lineCode}</div><div><span className="operator-line"><small>OPERAT PER</small><b>{operator.toUpperCase()}</b></span><h1>{lineTitle || [origin, destination].filter(Boolean).join(" — ")}</h1><p>{[validity, generatedOn].filter(Boolean).join(" · ")}</p></div></div>
        {!overview && <div className="stop-code"><small>CODI DE PARADA</small><b>{currentStop?.code || "—"}</b></div>}
      </header>
      {page.total > 1 && <div className="continuation"><span><strong>{page.page === 1 ? "INICI" : "CONTINUACIÓ"}</strong>{page.section}</span><b>Pàgina {page.page} / {page.total}</b></div>}
      <div className="sheet-body">
        {layout !== "table" && <RouteDiagram stops={page.stops} currentStop={page.selectedStop} circle={circular} color={lineColor} />}
        <div className="schedule-area">
          {!overview && currentStop && <div className="you-are"><i className="stop-marker-icon" aria-hidden="true" /><b>{currentStop.name}</b><HereTag /></div>}
          <div className="schedule-title"><div><small>{frequency ? "FREQÜÈNCIES ORIENTATIVES" : "HORARIS DE PAS APROXIMATS"}</small><h2>{period}</h2></div><span>{night ? "☾" : "☀"} {night ? "Servei nocturn" : "Servei diürn"}</span></div>
          {frequency ? <FrequencyTable times={currentTimes} /> : <ScheduleTable stops={page.stops} trips={trips} currentStop={page.selectedStop} showHereLabel={!overview} />}
          <div className="notes"><b>Informació important</b><p>{notes[0] || "Els horaris estan subjectes a les incidències i contratemps de la via pública. Consulta les alteracions del servei abans de viatjar."}</p></div>
        </div>
      </div>
      <footer className="sheet-footer"><button type="button" className={`logo-slot ${leftLogo ? "has-logo" : "empty"}`} onClick={onLeftLogoClick} aria-label="Puja o canvia el logo esquerre">{leftLogo ? <img src={leftLogo} alt="Logo esquerre" /> : <span>＋ AFEGEIX LOGO</span>}</button><p><b>{infoUrl}</b><br/>Informació: {contact}{accessible ? " · ♿ Servei accessible" : ""}</p><button type="button" className={`logo-slot right ${rightLogo ? "has-logo" : "empty"}`} onClick={onRightLogoClick} aria-label="Puja o canvia el logo dret">{rightLogo ? <img src={rightLogo} alt="Logo dret" /> : <span>＋ AFEGEIX LOGO</span>}</button></footer>
    </article>
  </div>;
}

function SiteHeader() {
  return <header className="site-header">
    <div className="site-logo" aria-label="Emma, creador d’horaris"><i className="emma-pulse" aria-hidden="true" /><strong>Emma</strong><span>Creador d’horaris</span></div>
    <button className="account-chip"><span>MG</span><b>Compte demo</b><i>⌄</i></button>
  </header>;
}

function Field({ label, value, onChange }: { label: string; value: string; onChange: (s: string) => void }) {
  return <label className="field"><span>{label}</span><input value={value} onChange={(e) => onChange(e.target.value)} /></label>;
}

function RouteDiagram({ stops, currentStop, circle, color }: { stops: Stop[]; currentStop?: Stop; circle: boolean; color: string }) {
  if (circle) return <CircularRouteDiagram stops={stops} currentStop={currentStop} color={color} />;
  const stopHeight = Math.max(18, Math.min(66, 570 / Math.max(stops.length, 1)));
  const density = stops.length > 24 ? "ultra-dense" : stops.length > 12 ? "dense" : "";
  return <section className={`route-diagram ${density}`}><div className="diagram-heading"><small>RECORREGUT</small><b>{stops.length} parades</b></div><div className="route-line" style={{ "--route": color, "--stop-height": `${stopHeight}px` } as React.CSSProperties}>{stops.map((s, i) => { const isHere = samePhysicalStop(s, currentStop); const isTerminal = i === 0 || i === stops.length - 1; return <div className={`route-stop ${isHere ? "here" : ""} ${isTerminal ? "is-terminal" : ""}`} key={s.sourceKey || i}><i /><div><b>{s.name}</b>{isHere && <HereTag />}</div></div>; })}</div></section>;
}

function CircularRouteDiagram({ stops, currentStop, color }: { stops: Stop[]; currentStop?: Stop; color: string }) {
  const midpoint = Math.ceil(stops.length / 2);
  const left = stops.slice(0, midpoint).map((stop, index) => ({ stop, index }));
  const right = stops.slice(midpoint).map((stop, index) => ({ stop, index: midpoint + index })).reverse();
  const rows = Math.max(left.length, right.length);
  const rowHeight = Math.max(22, Math.min(48, 520 / Math.max(rows, 1)));
  return <section className={`route-diagram circular-diagram ${stops.length > 24 ? "ultra-dense" : stops.length > 14 ? "dense" : ""}`} style={{ "--route": color, "--loop-row": `${rowHeight}px` } as React.CSSProperties}>
    <div className="diagram-heading"><small>RECORREGUT CIRCULAR</small><b>{stops.length} parades · sentit indicat</b></div>
    <div className="circular-route">
      <div className="circular-branch left-branch">{left.map(({ stop, index }) => { const isHere = samePhysicalStop(stop, currentStop); return <div className={`circular-stop ${isHere ? "here" : ""}`} key={stop.sourceKey || index}><div><b>{stop.name}</b></div><i /></div>; })}</div>
      <div className="loop-track"><span className="arrow down" aria-hidden="true">↓</span><span className="arrow up" aria-hidden="true">↑</span></div>
      <div className="circular-branch right-branch">{right.map(({ stop, index }) => { const isHere = samePhysicalStop(stop, currentStop); return <div className={`circular-stop ${isHere ? "here" : ""}`} key={stop.sourceKey || index}><i /><div><b>{stop.name}</b></div></div>; })}</div>
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

function ScheduleTable({ stops, trips, currentStop, showHereLabel }: { stops: Stop[]; trips: string[][]; currentStop?: Stop; showHereLabel: boolean }) {
  const density = trips.length > 16 ? "hours-ultra" : trips.length > 10 ? "hours-dense" : "";
  return <div className={`table-scroll ${density}`}><table className="transposed-table"><thead><tr><th className="stop-column"><small>RECORREGUT</small><span>Parada</span></th>{trips.map((trip, i) => <th key={i}><span>{trip.find((value) => value !== "—") || "—"}</span></th>)}</tr></thead><tbody>{stops.map((stop, stopIndex) => { const isHere = samePhysicalStop(stop, currentStop); return <tr className={isHere ? "active-row" : ""} key={stop.sourceKey || stopIndex}><th scope="row" className="stop-name">{showHereLabel && isHere && <HereTag compact />}<span>{stop.name}</span></th>{trips.map((trip, tripIndex) => <td key={tripIndex}>{trip[stopIndex] || "—"}</td>)}</tr>; })}</tbody></table></div>;
}

function HereTag({ compact = false, direction = "left", className = "", style }: { compact?: boolean; direction?: "left" | "right"; className?: string; style?: React.CSSProperties }) {
  return <span className={`here-tag points-${direction}${compact ? " is-compact" : ""}${className ? ` ${className}` : ""}`} style={style}>
    <svg viewBox="0 0 74.8 18.4" preserveAspectRatio="none" aria-hidden="true"><path d="M72.5,18.4H9.4c-.8,0-1.6-.4-2.2-1L1.5,11.3c-1.1-1.1-1.1-2.9,0-4.1L7.2,1c.6-.6,1.4-1,2.2-1h63.1c1.7,0,3,1.3,3,3v12.4c0,1.7-1.3,3-3,3Z" /></svg>
    <span>ETS AQUÍ</span>
  </span>;
}
