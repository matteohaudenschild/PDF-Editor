# PDF Desktop Editor

Lokale Windows-Desktop-App fuer 1:1-PDF-Ansicht mit editierbarem Textlayer.

## Stack

- Electron
- FastAPI
- PyMuPDF

## Was die App kann

- PDF lokal laden
- Originalseite als textfreien Hintergrund rendern
- Vorhandene Textbloecke positionsgenau als editierbare Overlays anzeigen
- Entwurf in `<datei>.pdfedit.json` speichern
- Neue PDF als `<datei>-bearbeitet.pdf` exportieren
- Texte ablehnen, die auch mit `6pt` nicht in den Originalblock passen

## Was V1 bewusst ablehnt

- Scan-PDFs ohne echten Text
- verschluesselte PDFs
- XFA-/Formular-PDFs
- rotierte oder nicht-horizontal gesetzte Textobjekte
- Fonts, die nicht sauber rekonstruiert werden koennen

## Start

1. `npm.cmd install`
2. `npm.cmd start`

Die App startet den lokalen Python-Service automatisch mit der vorhandenen virtuellen Umgebung unter [backend/.venv](C:/Users/forc3/CascadeProjects/codex-workspace/pdf-desktop-editor/backend/.venv).

## Windows-Installer fuer andere PCs

Fuer Zielrechner ohne vorinstalliertes Python gibt es jetzt einen Setup-Installer:

- Build: `npm.cmd run build:installer`
- Ausgabe: [dist/PDF Desktop Editor Setup.exe](C:/Users/forc3/CascadeProjects/codex-workspace/pdf-desktop-editor/dist/PDF%20Desktop%20Editor%20Setup.exe)

Der Installer:

- installiert die Desktop-App unter `%LOCALAPPDATA%\\Programs\\PDF Desktop Editor`
- entpackt die portable App intern
- laedt waehrend der Installation die offizielle eingebettete Python-3.12.10-Laufzeit von `python.org`
- richtet diese direkt im Backend-Ordner ein, sodass auf dem Ziel-PC keine separate Python-Installation noetig ist

Der bisherige Ordner [dist/portable-win](C:/Users/forc3/CascadeProjects/codex-workspace/pdf-desktop-editor/dist/portable-win) ist weiter fuer manuelle portable Nutzung da, aber fuer andere PCs sollte bevorzugt die `Setup.exe` verschickt werden.

## Wichtige Pfade

- Electron Einstieg: [app/src/main/main.js](C:/Users/forc3/CascadeProjects/codex-workspace/pdf-desktop-editor/app/src/main/main.js)
- Renderer: [app/src/renderer/renderer.js](C:/Users/forc3/CascadeProjects/codex-workspace/pdf-desktop-editor/app/src/renderer/renderer.js)
- Backend API: [backend/pdf_editor_service/app.py](C:/Users/forc3/CascadeProjects/codex-workspace/pdf-desktop-editor/backend/pdf_editor_service/app.py)
- PDF-Analyse/Export: [backend/pdf_editor_service/pdf_engine.py](C:/Users/forc3/CascadeProjects/codex-workspace/pdf-desktop-editor/backend/pdf_editor_service/pdf_engine.py)
