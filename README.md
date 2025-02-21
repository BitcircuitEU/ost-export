# OST to PST Converter

Ein TypeScript-basiertes Tool zur Konvertierung von Outlook OST-Dateien in das PST-Format.

## Schnellstart (Kompilierte Version)

1. Laden Sie die `ost-to-pst.exe` aus dem Release-Bereich herunter
2. Kopieren Sie die .exe in das Verzeichnis, in dem sich Ihre .ost-Dateien befinden
3. Führen Sie die .exe aus
4. Das Tool wird automatisch:
   - Alle .ost-Dateien im Verzeichnis finden
   - Einen `outlook_export` Ordner erstellen
   - Für jede .ost-Datei einen Unterordner mit dem gleichen Namen anlegen
   - Die Konvertierung durchführen
   - Die Ergebnisse in folgendem Format speichern:
     ```
     outlook_export/
     ├── beispiel1.ost/
     │   ├── Emails/
     │   │   ├── 2024-01-20_Betreff1.eml
     │   │   └── Attachments/
     │   ├── Contacts/
     │   │   └── Kontakt1.vcf
     │   └── Calendar/
     │       └── 2024-01-20_Meeting1.ics
     └── beispiel2.ost/
         └── ...
     ```

## Entwicklung und Build

Wenn Sie das Tool selbst kompilieren möchten, folgen Sie diesen Schritten:

### Voraussetzungen

- Node.js (v14 oder höher)
- npm
- Windows (für COM-Automation)

### Installation

1. Klone dieses Repository
2. Installiere die Abhängigkeiten:
```bash
npm install
```

### Verwendung während der Entwicklung

Platziere deine .ost Datei im Projektverzeichnis und führe aus:

```bash
npm start
```

### Entwicklungsmodus

Zum Ausführen im Entwicklungsmodus mit Auto-Reload:

```bash
npm run dev
```

### Build

Zum Bauen des TypeScript-Codes:

```bash
npm run build
```

Zum Erstellen der eigenständigen .exe:

```bash
npm run package
```

Die fertige .exe finden Sie dann im `dist`-Ordner.

### Nur Lesen der OST-Dateistruktur

Zum Lesen und Anzeigen der OST-Dateistruktur (ohne Konvertierung):

```bash
npm run read
```

## Wie es funktioniert

Das Tool verwendet die `pst-extractor` Bibliothek um:
- Die OST-Datei zu lesen
- Emails als .eml-Dateien zu exportieren
- Kontakte als .vcf-Dateien zu speichern
- Kalendereinträge als .ics-Dateien zu speichern
- Anhänge in separaten Ordnern zu organisieren

## Hinweise

- Die exportierten Dateien können in verschiedene Email-Clients importiert werden
- Große OST-Dateien können einige Zeit zur Konvertierung benötigen
- Alle Warnungen und Fehler werden in der Konsole angezeigt
- Das Tool verarbeitet auch mehrere OST-Dateien nacheinander
- Bei Problemen können Sie das Tool einfach erneut ausführen

## Note

This is a basic implementation. The current version uses pst-extractor which is primarily designed for reading PST files. For full OST to PST conversion functionality, additional libraries or implementations may be needed. 