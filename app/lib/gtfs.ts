import { unzipSync } from "fflate";

export type GtfsRoute = {
  id: string;
  agencyId: string;
  shortName: string;
  longName: string;
  type: string;
  color: string;
  textColor: string;
};

export type GtfsTrip = {
  id: string;
  routeId: string;
  serviceId: string;
  direction: string;
  shapeId: string;
  headsign: string;
  wheelchair: string;
};

export type GtfsStop = {
  id: string;
  code: string;
  name: string;
  wheelchair: string;
};

export type GtfsDataset = {
  name: string;
  agencyName: string;
  publisher: string;
  updated: string;
  routes: GtfsRoute[];
  trips: GtfsTrip[];
  stops: Map<string, GtfsStop>;
  serviceLabels: Map<string, string>;
  serviceMeta: Map<string, GtfsServiceMeta>;
  agencies: Map<string, string>;
  frequencyTripIds: Set<string>;
  stopTimesText: string;
  notesText: string;
  warnings: string[];
  diagnostics: {
    routes: number;
    trips: number;
    stops: number;
    hasHeadsigns: boolean;
    hasColors: boolean;
    hasAccessibility: boolean;
    hasNamedPeriods: boolean;
    hasFrequencies: boolean;
  };
};

export type GtfsServiceMeta = {
  startDate: string;
  endDate: string;
  addedDates: number;
  removedDates: number;
};

export type GtfsFeedOption = {
  name: string;
  bytes: Uint8Array;
};

export type LineStop = {
  code: string;
  name: string;
  times: string[];
  sourceKey?: string;
  sourceStopId?: string;
};

export type LineImport = {
  stops: LineStop[];
  lineCode: string;
  routeLongName: string;
  lineColor: string;
  lineTextColor: "white" | "black";
  origin: string;
  destination: string;
  operator: string;
  period: string;
  validity: string;
  circular: boolean;
  night: boolean;
  accessible: boolean;
  trips: number;
  patternCount: number;
  notes: string[];
  warnings: string[];
  frequencyBased: boolean;
};

const decoder = new TextDecoder("utf-8");
const required = ["agency.txt", "routes.txt", "trips.txt", "stops.txt", "stop_times.txt"];
const wanted = new Set([
  ...required,
  "calendar.txt",
  "calendar_dates.txt",
  "feed_info.txt",
  "frequencies.txt",
  "nota.txt",
]);

function basename(path: string) {
  return path.split("/").pop()?.toLowerCase() || path.toLowerCase();
}

function unzipGtfs(bytes: Uint8Array) {
  const unpacked = unzipSync(bytes, {
    filter: (file) => wanted.has(basename(file.name)),
  });
  const normalized: Record<string, Uint8Array> = {};
  Object.entries(unpacked).forEach(([path, contents]) => {
    normalized[basename(path)] = contents;
  });
  return normalized;
}

export function discoverGtfsFeeds(buffer: ArrayBuffer, filename: string): GtfsFeedOption[] {
  const bytes = new Uint8Array(buffer);
  const direct = unzipGtfs(bytes);
  if (required.every((file) => direct[file])) return [{ name: filename, bytes }];

  const outer = unzipSync(bytes, { filter: (file) => file.name.toLowerCase().endsWith(".zip") });
  const feeds = Object.entries(outer)
    .filter(([, nested]) => {
      try {
        const contents = unzipGtfs(nested);
        return required.every((file) => contents[file]);
      } catch {
        return false;
      }
    })
    .map(([path, nested]) => ({ name: path.split("/").pop() || path, bytes: nested }));
  if (!feeds.length) throw new Error("El ZIP no conté un GTFS estàndard ni cap GTFS comprimit a dins.");
  return feeds;
}

function decode(bytes?: Uint8Array) {
  if (!bytes) return "";
  return decoder.decode(bytes).replace(/^\uFEFF/, "");
}

function parseLine(line: string) {
  const values: string[] = [];
  let value = "";
  let quoted = false;
  for (let index = 0; index < line.length; index += 1) {
    const char = line[index];
    if (char === '"') {
      if (quoted && line[index + 1] === '"') {
        value += '"';
        index += 1;
      } else {
        quoted = !quoted;
      }
    } else if (char === "," && !quoted) {
      values.push(value);
      value = "";
    } else {
      value += char;
    }
  }
  values.push(value.replace(/\r$/, ""));
  return values;
}

function eachCsvRow(text: string, callback: (row: Record<string, string>) => void) {
  if (!text) return;
  let start = 0;
  let end = text.indexOf("\n", start);
  if (end < 0) return;
  const headers = parseLine(text.slice(start, end)).map((header) => header.trim());
  start = end + 1;
  while (start < text.length) {
    end = text.indexOf("\n", start);
    if (end < 0) end = text.length;
    const line = text.slice(start, end);
    start = end + 1;
    if (!line.trim()) continue;
    const values = parseLine(line);
    const row: Record<string, string> = {};
    headers.forEach((header, column) => { row[header] = values[column] ?? ""; });
    callback(row);
  }
}

function csvRows(text: string) {
  const rows: Record<string, string>[] = [];
  eachCsvRow(text, (row) => rows.push(row));
  return rows;
}

function csvHeaders(text: string) {
  if (!text) return [];
  const end = text.indexOf("\n");
  return parseLine(text.slice(0, end < 0 ? text.length : end)).map((header) => header.trim());
}

function validateGtfsEntries(entries: Record<string, Uint8Array>) {
  const requirements: Record<string, string[]> = {
    "agency.txt": ["agency_name", "agency_url", "agency_timezone"],
    "routes.txt": ["route_id", "route_type"],
    "trips.txt": ["route_id", "service_id", "trip_id"],
    "stops.txt": ["stop_id", "stop_name"],
    "stop_times.txt": ["trip_id", "arrival_time", "departure_time", "stop_id", "stop_sequence"],
  };
  const errors: string[] = [];
  Object.entries(requirements).forEach(([file, columns]) => {
    if (!entries[file]) {
      errors.push(`Falta ${file}`);
      return;
    }
    const headers = new Set(csvHeaders(decode(entries[file])));
    const missing = columns.filter((column) => !headers.has(column));
    if (missing.length) errors.push(`${file}: falten ${missing.join(", ")}`);
  });

  const hasCalendar = Boolean(entries["calendar.txt"]);
  const hasCalendarDates = Boolean(entries["calendar_dates.txt"]);
  if (!hasCalendar && !hasCalendarDates) errors.push("Falta calendar.txt o calendar_dates.txt");
  if (hasCalendar) {
    const headers = new Set(csvHeaders(decode(entries["calendar.txt"])));
    const requiredCalendar = ["service_id", "monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday", "start_date", "end_date"];
    const missing = requiredCalendar.filter((column) => !headers.has(column));
    if (missing.length) errors.push(`calendar.txt: falten ${missing.join(", ")}`);
  }
  if (hasCalendarDates) {
    const headers = new Set(csvHeaders(decode(entries["calendar_dates.txt"])));
    const missing = ["service_id", "date", "exception_type"].filter((column) => !headers.has(column));
    if (missing.length) errors.push(`calendar_dates.txt: falten ${missing.join(", ")}`);
  }
  if (errors.length) throw new Error(`GTFS incomplet. ${errors.join(" · ")}.`);
}

function formatDate(value: string) {
  if (!/^\d{8}$/.test(value)) return value;
  return `${value.slice(6, 8)}.${value.slice(4, 6)}.${value.slice(0, 4)}`;
}

function calendarLabel(row: Record<string, string>) {
  if (row.name?.trim()) return row.name.trim();
  const active = ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"]
    .map((day) => row[day] === "1");
  let label = "Servei especial";
  if (active.every(Boolean)) label = "Cada dia";
  else if (active.slice(0, 5).every(Boolean) && !active[5] && !active[6]) label = "Dilluns a divendres feiners";
  else if (!active.slice(0, 5).some(Boolean) && active[5] && !active[6]) label = "Dissabtes";
  else if (!active.slice(0, 6).some(Boolean) && active[6]) label = "Diumenges i festius";
  else {
    const names = ["Dl", "Dt", "Dc", "Dj", "Dv", "Ds", "Dg"];
    const days = names.filter((_, index) => active[index]);
    if (days.length) label = days.join(", ");
  }
  return label;
}

export function parseGtfsFeed(option: GtfsFeedOption): GtfsDataset {
  const entries = unzipGtfs(option.bytes);
  validateGtfsEntries(entries);
  const validationWarnings: string[] = [];
  const agencies = new Map<string, string>();
  let defaultAgency = "Operador de transport";
  csvRows(decode(entries["agency.txt"])).forEach((row, index) => {
    const id = row.agency_id || String(index);
    const name = row.agency_name || defaultAgency;
    agencies.set(id, name);
    if (index === 0) defaultAgency = name;
  });

  const allRouteRows = csvRows(decode(entries["routes.txt"]));
  const allRoutes = allRouteRows.filter((row) => row.route_id).map((row): GtfsRoute => ({
    id: row.route_id,
    agencyId: row.agency_id || "0",
    shortName: row.route_short_name || "",
    longName: row.route_long_name || "",
    type: row.route_type || "3",
    color: row.route_color || "",
    textColor: row.route_text_color || "",
  }));
  const buses = allRoutes.filter((route) => route.type === "3" || route.type === "700");
  const routes = buses.length ? buses : allRoutes;
  const routeIds = new Set(routes.map((route) => route.id));

  const allTripRows = csvRows(decode(entries["trips.txt"]));
  const trips = allTripRows
    .filter((row) => row.trip_id && row.service_id && routeIds.has(row.route_id))
    .map((row): GtfsTrip => ({
      id: row.trip_id,
      routeId: row.route_id,
      serviceId: row.service_id,
      direction: row.direction_id || "0",
      shapeId: row.shape_id || "",
      headsign: row.trip_headsign || "",
      wheelchair: row.wheelchair_accessible || "",
    }));

  const stops = new Map<string, GtfsStop>();
  csvRows(decode(entries["stops.txt"])).forEach((row) => {
    if (!row.stop_id || !row.stop_name) return;
    stops.set(row.stop_id, {
      id: row.stop_id,
      code: row.stop_code && row.stop_code.toLowerCase() !== "null" ? row.stop_code : "",
      name: row.stop_name || "Parada sense nom",
      wheelchair: row.wheelchair_boarding || "",
    });
  });

  const serviceLabels = new Map<string, string>();
  const serviceMeta = new Map<string, GtfsServiceMeta>();
  csvRows(decode(entries["calendar.txt"])).forEach((row) => {
    if (!row.service_id) return;
    serviceLabels.set(row.service_id, calendarLabel(row));
    serviceMeta.set(row.service_id, {
      startDate: row.start_date || "",
      endDate: row.end_date || "",
      addedDates: 0,
      removedDates: 0,
    });
  });
  csvRows(decode(entries["calendar_dates.txt"])).forEach((row) => {
    if (!row.service_id) return;
    if (!serviceLabels.has(row.service_id)) serviceLabels.set(row.service_id, "Calendari especial");
    const meta = serviceMeta.get(row.service_id) || { startDate: row.date || "", endDate: row.date || "", addedDates: 0, removedDates: 0 };
    if (row.date && (!meta.startDate || row.date < meta.startDate)) meta.startDate = row.date;
    if (row.date && (!meta.endDate || row.date > meta.endDate)) meta.endDate = row.date;
    if (row.exception_type === "1") meta.addedDates += 1;
    if (row.exception_type === "2") meta.removedDates += 1;
    serviceMeta.set(row.service_id, meta);
  });

  const knownTripIds = new Set(allTripRows.map((row) => row.trip_id).filter(Boolean));
  const knownStopIds = new Set(stops.keys());
  let orphanStopTimes = 0;
  let invalidSequences = 0;
  eachCsvRow(decode(entries["stop_times.txt"]), (row) => {
    if (!knownTripIds.has(row.trip_id) || !knownStopIds.has(row.stop_id)) orphanStopTimes += 1;
    if (!Number.isFinite(Number(row.stop_sequence))) invalidSequences += 1;
  });
  if (orphanStopTimes) validationWarnings.push(`${orphanStopTimes} registres de stop_times.txt fan referència a viatges o parades inexistents`);
  if (invalidSequences) validationWarnings.push(`${invalidSequences} registres tenen un stop_sequence invàlid`);
  if (!routes.length) throw new Error("El GTFS no conté cap línia utilitzable.");
  if (!trips.length) throw new Error("El GTFS no conté cap expedició utilitzable per a les línies detectades.");
  if (!stops.size) throw new Error("El GTFS no conté cap parada utilitzable.");

  const frequencyTripIds = new Set(csvRows(decode(entries["frequencies.txt"])).map((row) => row.trip_id).filter(Boolean));

  const feedInfo = csvRows(decode(entries["feed_info.txt"]))[0] || {};
  const updated = formatDate(feedInfo.feed_version || feedInfo.feed_end_date || "");
  return {
    name: option.name,
    agencyName: defaultAgency,
    publisher: feedInfo.feed_publisher_name || defaultAgency,
    updated,
    routes: routes.sort((a, b) => (a.shortName || a.longName).localeCompare(b.shortName || b.longName, "ca", { numeric: true })),
    trips,
    stops,
    serviceLabels,
    serviceMeta,
    agencies,
    frequencyTripIds,
    stopTimesText: decode(entries["stop_times.txt"]),
    notesText: decode(entries["nota.txt"]),
    warnings: validationWarnings,
    diagnostics: {
      routes: routes.length,
      trips: trips.length,
      stops: stops.size,
      hasHeadsigns: trips.some((trip) => Boolean(trip.headsign)),
      hasColors: routes.some((route) => Boolean(route.color)),
      hasAccessibility: trips.some((trip) => trip.wheelchair === "1"),
      hasNamedPeriods: csvRows(decode(entries["calendar.txt"])).some((row) => Boolean(row.name)),
      hasFrequencies: frequencyTripIds.size > 0,
    },
  };
}

export function directionsFor(dataset: GtfsDataset, routeId: string) {
  const directions = new Set(dataset.trips.filter((trip) => trip.routeId === routeId).map((trip) => trip.direction));
  return Array.from(directions).sort();
}

export function servicesFor(dataset: GtfsDataset, routeId: string, direction: string) {
  const counts = new Map<string, number>();
  dataset.trips.filter((trip) => trip.routeId === routeId && trip.direction === direction).forEach((trip) => {
    counts.set(trip.serviceId, (counts.get(trip.serviceId) || 0) + 1);
  });
  return Array.from(counts, ([id, count]) => ({ id, count, label: dataset.serviceLabels.get(id) || "Servei especial" }))
    .sort((a, b) => b.count - a.count || a.label.localeCompare(b.label, "ca"));
}

function timeValue(value: string) {
  const [hours = "0", minutes = "0", seconds = "0"] = value.split(":");
  return Number(hours) * 3600 + Number(minutes) * 60 + Number(seconds);
}

function displayTime(value: string) {
  const [hours = "0", minutes = "00"] = value.split(":");
  const hour = Number(hours);
  return `${String(hour % 24).padStart(2, "0")}:${minutes}${hour >= 24 ? "⁺" : ""}`;
}

function publicCode(route: GtfsRoute) {
  const embedded = route.longName.match(/^\s*\(([^)]+)\)/)?.[1];
  if (/^L\d{3,}$/i.test(route.shortName) && embedded) return embedded;
  return route.shortName || embedded || route.id;
}

function contrastInk(hex: string): "white" | "black" {
  if (!/^[0-9a-f]{6}$/i.test(hex)) return "white";
  const channels = [0, 2, 4].map((index) => parseInt(hex.slice(index, index + 2), 16));
  return (channels[0] * 299 + channels[1] * 587 + channels[2] * 114) / 1000 > 150 ? "black" : "white";
}

function suppliedInk(hex: string): "white" | "black" {
  if (!/^[0-9a-f]{6}$/i.test(hex)) return "white";
  const channels = [0, 2, 4].map((index) => parseInt(hex.slice(index, index + 2), 16));
  return (channels[0] * 299 + channels[1] * 587 + channels[2] * 114) / 1000 > 150 ? "white" : "black";
}

export function importLine(dataset: GtfsDataset, routeId: string, direction: string, serviceId: string): LineImport {
  const route = dataset.routes.find((item) => item.id === routeId);
  if (!route) throw new Error("No s’ha trobat la línia seleccionada.");
  let candidates = dataset.trips.filter((trip) => trip.routeId === routeId && trip.direction === direction && trip.serviceId === serviceId);
  if (!candidates.length) candidates = dataset.trips.filter((trip) => trip.routeId === routeId && trip.direction === direction);

  const candidateIds = new Set(candidates.map((trip) => trip.id));

  type Call = { stopId: string; sequence: number; time: string; key: string };
  const calls = new Map<string, Call[]>();
  eachCsvRow(dataset.stopTimesText, (row) => {
    if (!candidateIds.has(row.trip_id)) return;
    const list = calls.get(row.trip_id) || [];
    list.push({
      stopId: row.stop_id,
      sequence: Number(row.stop_sequence || list.length),
      time: row.departure_time || row.arrival_time || "",
      key: "",
    });
    calls.set(row.trip_id, list);
  });
  calls.forEach((list) => {
    list.sort((a, b) => a.sequence - b.sequence);
    const occurrences = new Map<string, number>();
    list.forEach((call) => {
      const occurrence = (occurrences.get(call.stopId) || 0) + 1;
      occurrences.set(call.stopId, occurrence);
      call.key = `${call.stopId}\u001e${occurrence}`;
    });
  });

  const patterns = new Map<string, { ids: string[]; trips: string[] }>();
  calls.forEach((list, tripId) => {
    const ids = list.map((call) => call.stopId);
    const key = ids.join("\u001f");
    const pattern = patterns.get(key) || { ids, trips: [] };
    pattern.trips.push(tripId);
    patterns.set(key, pattern);
  });
  const dominant = Array.from(patterns.values()).sort((a, b) => b.ids.length - a.ids.length || b.trips.length - a.trips.length)[0];
  if (!dominant) throw new Error("Aquesta selecció no conté hores de pas.");
  const tripIds = candidates.map((trip) => trip.id).filter((tripId) => calls.has(tripId))
    .sort((a, b) => timeValue(calls.get(a)?.[0]?.time || "") - timeValue(calls.get(b)?.[0]?.time || ""));
  const closesAtOrigin = dominant.ids.length > 2 && dominant.ids[0] === dominant.ids[dominant.ids.length - 1];
  const circular = closesAtOrigin || /circular/i.test(route.longName) || candidates.some((trip) => /circular/i.test(trip.headsign));
  const referenceCalls = calls.get(dominant.trips[0]) || [];
  const visibleCalls = closesAtOrigin ? referenceCalls.slice(0, -1) : referenceCalls;
  const importedStops = visibleCalls.map((referenceCall): LineStop => {
    const stopId = referenceCall.stopId;
    const stop = dataset.stops.get(stopId);
    return {
      code: stop?.code || "",
      name: stop?.name || stopId,
      sourceKey: referenceCall.key,
      sourceStopId: stopId,
      times: tripIds.map((tripId) => {
        const call = calls.get(tripId)?.find((item) => item.key === referenceCall.key);
        return call?.time ? displayTime(call.time) : "—";
      }),
    };
  });

  const tripMeta = new Map(candidates.map((trip) => [trip.id, trip]));
  const selectedTrip = tripMeta.get(tripIds[0]);
  const first = importedStops[0]?.name || route.longName;
  const last = importedStops[importedStops.length - 1]?.name || route.longName;
  const headsign = selectedTrip?.headsign || "";
  const rawColor = route.color.replace(/^#/, "");
  const color = /^[0-9a-f]{6}$/i.test(rawColor) ? `#${rawColor}` : "#93D500";
  const rawText = route.textColor.replace(/^#/, "");
  const textColor = /^[0-9a-f]{6}$/i.test(rawText) ? suppliedInk(rawText) : contrastInk(rawColor);
  const warnings: string[] = [];
  if (!selectedTrip?.headsign) warnings.push("Destinació inferida amb les parades terminals");
  if (!route.color) warnings.push("Color corporatiu no informat; s’aplica el verd d’Emma");
  warnings.push(...dataset.warnings);
  if (patterns.size > 1) warnings.push(`${patterns.size} variants de recorregut integrades; les parades no compartides es mostren amb un guió`);

  const notes: string[] = [];
  if (dataset.notesText) {
    const selectedIds = new Set(tripIds);
    eachCsvRow(dataset.notesText, (row) => {
      if ((row.route_id === routeId || selectedIds.has(row.trip_id)) && row.nota && row.nota !== "." && !notes.includes(row.nota)) notes.push(row.nota);
    });
  }
  const serviceMeta = dataset.serviceMeta.get(serviceId);
  const exceptionCount = (serviceMeta?.addedDates || 0) + (serviceMeta?.removedDates || 0);
  if (serviceMeta && exceptionCount > 0 && exceptionCount <= 20) {
    const parts = [
      serviceMeta.addedDates ? `${serviceMeta.addedDates} dies afegits` : "",
      serviceMeta.removedDates ? `${serviceMeta.removedDates} dies sense servei` : "",
    ].filter(Boolean);
    notes.unshift(`Excepcions del calendari: ${parts.join(" · ")}.`);
  }
  const firstTimes = tripIds.map((id) => calls.get(id)?.[0]?.time || "").filter(Boolean);
  const nightTrips = firstTimes.filter((time) => {
    const hour = Number(time.split(":")[0]) % 24;
    return hour >= 20 || hour < 5;
  }).length;
  const routeAgency = dataset.agencies.get(route.agencyId) || dataset.agencyName;
  return {
    stops: importedStops,
    lineCode: publicCode(route),
    routeLongName: route.longName || [first, headsign || last].filter(Boolean).join(" — "),
    lineColor: color,
    lineTextColor: textColor,
    origin: circular ? (route.longName.replace(/^\s*\([^)]+\)\s*/, "") || first) : first,
    destination: circular ? "Circular" : (headsign || last),
    operator: routeAgency,
    period: dataset.serviceLabels.get(serviceId) || "Servei seleccionat",
    validity: serviceMeta?.startDate || serviceMeta?.endDate
      ? `Vigència: ${formatDate(serviceMeta.startDate)} — ${formatDate(serviceMeta.endDate)}`
      : (dataset.updated ? `Versió GTFS: ${dataset.updated}` : "Vigència no informada"),
    circular,
    night: firstTimes.length > 0 && nightTrips / firstTimes.length > 0.5,
    accessible: candidates.some((trip) => trip.wheelchair === "1"),
    trips: tripIds.length,
    patternCount: patterns.size,
    notes: notes.slice(0, 3),
    warnings,
    frequencyBased: tripIds.some((tripId) => dataset.frequencyTripIds.has(tripId)),
  };
}
