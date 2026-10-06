# Emma · Creador d’horaris

Eina web per importar dades GTFS i generar cartells d’horaris de transport
públic modulars, accessibles i preparats per imprimir.

## Funcions actuals

- Importació de fitxers GTFS ZIP al navegador.
- Selecció de línia, sentit i període de servei.
- Plantilla A4 de recorregut i horaris.
- Plantilla A4 d’horaris complets.
- Adaptació per a línies normals i circulars.
- Paginació automàtica dels horaris extensos.
- Personalització visual i modes de servei diürn i nocturn.

Les dades GTFS no surten del dispositiu mentre no s’hi connecti un backend.

## Desenvolupament local

Requereix Node.js 22 o superior.

```bash
npm ci
npm run dev
```

El projecte queda disponible a `http://localhost:3000`.

## Verificació

```bash
npm run build
npm test
```

La compilació genera un web estàtic dins de `dist/client`.

## Publicació a GitHub Pages

El workflow `.github/workflows/deploy-pages.yml` compila i publica Emma quan
s’envia un canvi a la branca `main`.

Per evitar haver de configurar rutes internes, es recomana crear un repositori
de pàgina d’usuari o d’organització amb el nom exacte:

```text
<compte>.github.io
```

Després cal activar **Settings → Pages → Source → GitHub Actions** al repositori.

```bash
git remote add origin https://github.com/<compte>/<compte>.github.io.git
git push -u origin main
```

La publicació resultant estarà disponible a `https://<compte>.github.io`.
