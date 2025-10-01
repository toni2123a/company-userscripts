# Company Userscripts

Dieses Repository bündelt mehrere firmenspezifische Userscripts samt einer kleinen Landingpage, über die Kolleg:innen die Skripte bequem installieren können. Ein optionales Bootstrap-Skript lädt alle Tools dynamisch anhand der `manifest.json`.

## Inhalte

- **Landingpage (`index.html`)** – Listet alle verfügbaren Tools, beschreibt sie und verlinkt direkt zur Installation. Zusätzlich enthält die Seite eine kurze Schritt-für-Schritt-Anleitung zur Einrichtung von Tampermonkey sowie ein Formular zum Hinterlegen der Standortnummer für das ASEA-Tool.
- **`manifest.json`** – Verzeichnet alle Userscripts samt Metadaten (Name, Version, Beschreibung, @match/@grant Informationen).
- **`bootstrap.user.js`** – Tampermonkey-Skript, das die `manifest.json` lädt und anschließend nacheinander alle Tools einbindet.
- **`tools/`** – Enthält die eigentlichen Userscripts (z. B. `tool-SSW-letzter08.user.js`, `tool-OpenPricer.user.js`, `tool-Zerberus.user.js`).
- **`build-manifest.js`** – Node.js-Skript, das die Meta-Header der Userscripts einliest und daraus automatisiert die `manifest.json` erstellt.

## Landingpage lokal testen

```bash
# Statisches Hosting, z. B. via http-server
npm install --global http-server
cd company-userscripts
http-server -p 8080
# Danach http://localhost:8080 im Browser öffnen.
```

Alternativ lässt sich die Datei auch direkt in Visual Studio Code per Live Server oder einem anderen statischen Webserver öffnen.

## manifest.json generieren

Falls neue Tools hinzukommen oder Meta-Angaben geändert werden, sollte die `manifest.json` neu erstellt werden:

```bash
cd company-userscripts
node build-manifest.js
```

Das Skript liest alle `*.user.js` Dateien im Ordner `tools/` aus, extrahiert die UserScript-Metadaten und schreibt eine aktualisierte `manifest.json`.

## Bootstrap-Skript nutzen

1. Tampermonkey in Edge (oder einem anderen Browser) installieren.
2. `bootstrap.user.js` in Tampermonkey importieren.
3. In der Datei die Platzhalter `https://github.com/<deinUser>/company-userscripts` bzw. `<deinUser>` anpassen oder die URLs auf dieses Repository zeigen lassen.
4. Tampermonkey lädt daraufhin automatisch alle Tools, die in der `manifest.json` aufgelistet sind.

## Lizenz

Falls nicht anders angegeben, gelten die internen Unternehmensrichtlinien. Eine explizite Open-Source-Lizenz ist nicht hinterlegt.
