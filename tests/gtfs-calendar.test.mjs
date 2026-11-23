import assert from "node:assert/strict";
import test from "node:test";
import { strToU8, zipSync } from "fflate";
import { importLine, parseGtfsFeed, servicesFor } from "../app/lib/gtfs.ts";

function gtfsFeed(serviceIds, options = {}) {
  const activeDays = options.activeDays || "1,1,1,1,1,1,1";
  const calendarDates = options.calendarDates || ["S1,20261225,2", "S1,20260101,2"];
  const trips = serviceIds.map((serviceId, index) => `R1,${serviceId},T${index + 1},Destí,0,1`).join("\n");
  const stopTimes = serviceIds.flatMap((_, index) => [
    `T${index + 1},08:0${index}:00,08:0${index}:00,S1,1`,
    `T${index + 1},08:1${index}:00,08:1${index}:00,S2,2`,
  ]).join("\n");
  const calendars = serviceIds.map((serviceId) => `${serviceId},${activeDays},20250101,20261231`).join("\n");
  const files = {
    "agency.txt": "agency_id,agency_name,agency_url,agency_timezone\nA1,Operador de prova,https://example.com,Europe/Madrid\n",
    "routes.txt": "route_id,agency_id,route_short_name,route_long_name,route_type,route_color\nR1,A1,1,Origen - Destí,3,93D500\n",
    "trips.txt": `route_id,service_id,trip_id,trip_headsign,direction_id,wheelchair_accessible\n${trips}\n`,
    "stops.txt": "stop_id,stop_code,stop_name,wheelchair_boarding\nS1,001,Origen,1\nS2,002,Destí,1\n",
    "stop_times.txt": `trip_id,arrival_time,departure_time,stop_id,stop_sequence\n${stopTimes}\n`,
    "calendar.txt": `service_id,monday,tuesday,wednesday,thursday,friday,saturday,sunday,start_date,end_date\n${calendars}\n`,
    "calendar_dates.txt": `service_id,date,exception_type\n${calendarDates.join("\n")}\n`,
  };
  return {
    name: "calendari.zip",
    bytes: zipSync(Object.fromEntries(Object.entries(files).map(([name, contents]) => [name, strToU8(contents)]))),
  };
}

test("lists exact no-service dates when an operator has a single calendar", () => {
  const dataset = parseGtfsFeed(gtfsFeed(["S1"]));
  const imported = importLine(dataset, "R1", "0", "S1");
  assert.equal(
    imported.notes[0],
    "Circula de dilluns a diumenges i festius, excepte l'1 de gener i el 25 de desembre.",
  );
});

test("groups repeated exception days by month and year", () => {
  const dataset = parseGtfsFeed(gtfsFeed(["S1"], {
    activeDays: "0,0,1,0,0,0,0",
    calendarDates: [
      "S1,20251225,2",
      "S1,20251226,2",
      "S1,20260101,2",
      "S1,20261225,2",
      "S1,20261226,2",
    ],
  }));
  const imported = importLine(dataset, "R1", "0", "S1");
  assert.equal(
    imported.notes[0],
    "Circula els dimecres, excepte el 25 i 26 de desembre de 2025, l'1 de gener, el 25 i 26 de desembre de 2026.",
  );
});

test("omits calendar exceptions when an operator has multiple calendars", () => {
  const dataset = parseGtfsFeed(gtfsFeed(["S1", "S2"]));
  const imported = importLine(dataset, "R1", "0", "S1");
  assert.equal(imported.notes.length, 0);
});

test("lists equivalent service calendars only once", () => {
  const dataset = parseGtfsFeed(gtfsFeed(["S1", "S2", "S3", "S4"]));
  assert.deepEqual(servicesFor(dataset, "R1", "0"), [{ id: "S1", count: 1, label: "Cada dia" }]);
});
