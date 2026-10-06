import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

async function render() {
  const workerUrl = new URL("../dist/server/index.js", import.meta.url);
  workerUrl.searchParams.set("test", `${process.pid}-${Date.now()}`);
  const { default: worker } = await import(workerUrl.href);
  return worker.fetch(
    new Request("http://localhost/", { headers: { accept: "text/html" } }),
    { ASSETS: { fetch: async () => new Response("Not found", { status: 404 }) } },
    { waitUntil() {}, passThroughOnException() {} },
  );
}

test("server-renders the Emma GTFS product shell", async () => {
  const response = await render();
  assert.equal(response.status, 200);
  assert.match(response.headers.get("content-type") ?? "", /^text\/html\b/i);
  const html = await response.text();
  assert.match(html, /<html lang="ca">/i);
  assert.match(html, /<title>Emma — Generador d’horaris de bus<\/title>/i);
  assert.match(html, /Creador d’horaris/i);
  assert.match(html, /Eina d’horaris/i);
  assert.match(html, /Comença amb un fitxer GTFS/i);
  assert.match(html, /Puja un GTFS\.zip/i);
  assert.match(html, /Nou projecte/i);
  assert.doesNotMatch(html, />Inici<|>Projectes<|>Contacte</i);
});

test("keeps GTFS import, adaptive templates and pagination connected", async () => {
  const [page, gtfs, css, packageJson] = await Promise.all([
    readFile(new URL("../app/page.tsx", import.meta.url), "utf8"),
    readFile(new URL("../app/lib/gtfs.ts", import.meta.url), "utf8"),
    readFile(new URL("../app/globals.css", import.meta.url), "utf8"),
    readFile(new URL("../package.json", import.meta.url), "utf8"),
  ]);
  assert.match(packageJson, /"fflate": "0\.7\.5"/);
  assert.match(gtfs, /discoverGtfsFeeds/);
  assert.match(gtfs, /importLine/);
  assert.match(gtfs, /route_type/);
  assert.match(page, /A4 · recorregut \+ horaris/);
  assert.match(page, /A4 · horaris complets/);
  assert.doesNotMatch(page, /Auto · recomanat|Mig A4 · línia/);
  assert.match(page, /function paginate/);
  assert.match(page, /charsPerRow/);
  assert.match(page, /CONTINUACIÓ/);
  assert.match(page, /function FrequencyTable/);
  assert.match(page, /detectedCircular/);
  assert.match(page, /circular-current-tag/);
  assert.match(page, /setZoom/);
  assert.match(css, /paper-stack/);
  assert.match(css, /panel-scroll[^}]*overflow-y:auto/);
  assert.match(css, /--preview-zoom/);
  assert.match(css, /transform:scale\(var\(--preview-zoom,1\)\)/);
  assert.match(css, /break-after:\s*page/);
  assert.match(css, /@page\{size:216mm 303mm;margin:0\}/);
  assert.match(css, /\.site-header,.topbar,.sidebar,.canvas-toolbar,.modal-backdrop\{display:none!important\}/);
  assert.match(css, /paper-stack>\.paper-wrap:before/);
});
