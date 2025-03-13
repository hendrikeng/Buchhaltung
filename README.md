# Buchhaltungs-App für Google Sheets

Eine umfassende Buchhaltungslösung für kleine Unternehmen, die auf Google Sheets basiert. Unterstützt sowohl operative GmbHs als auch Holding-Strukturen mit SKR04-konformen Kontenrahmen.

## Features

- **Dokumenten-Management**: Automatischer Import von Rechnungen aus Google Drive
- **Steuerliche Auswertungen**: Berechnung von UStVA nach deutschem Steuerrecht
- **Betriebswirtschaftliche Auswertung (BWA)**: Monatliche und quartalsweise Auswertungen mit relevanten Kennzahlen
- **Bilanzerstellung**: Automatische Generierung einer SKR04-konformen Bilanz
- **Validierung**: Umfassende Validierung aller Eingaben mit detaillierten Fehlermeldungen
- **Fehlertoleranz**: Robuste Fehlerbehandlung und Logging
- **Flexibilität**: Unterstützung für verschiedene Steuerszenarien (Holding/Operativ)
- **Performance-Optimierung**: Caching-Mechanismen und Batch-Operationen für schnellere Berechnungen

## Module

Die App besteht aus folgenden Modulen:

- **importModule.js**: Import von Dateien aus Google Drive mit automatischer Kategorisierung
- **refreshModule.js**: Aktualisierung der Sheets, Formeln und Berechnungen
- **uStVACalculator.js**: Berechnung der Umsatzsteuervoranmeldung nach aktuellem Steuerrecht
- **bWACalculator.js**: Erstellung der Betriebswirtschaftlichen Auswertung mit wichtigen Kennzahlen (EBIT, EBITDA)
- **bilanzCalculator.js**: Generierung einer SKR04-konformen Bilanz mit Aktiva und Passiva
- **validator.js**: Validierung der Eingaben und Berechnungen
- **helpers.js**: Hilfsfunktionen für Datum, Währung, MwSt-Berechnungen und mehr
- **config.js**: Zentrale Konfigurationsdatei mit SKR04-Mapping
- **code.js**: Hauptmodul mit UI-Integration und Fehlerbehandlung

## Installation

1. Klone dieses Repository:
   ```
   git clone https://github.com/hendrikeng/Buchhaltung.git
   ```

2. Installiere die Abhängigkeiten:
   ```
   npm install
   ```

3. Baue das Projekt:
   ```
   npm run build
   ```

4. Kopiere den generierten Code aus `dist/code.js` in dein Google Apps Script Projekt.

## Einrichtung in Google Sheets

1. Erstelle ein neues Google Sheets Dokument
2. Gehe zu Erweiterungen > Apps Script
3. Füge den generierten Code ein
4. Speichere und führe die Funktion `onOpen` aus
5. Erlaube die notwendigen Berechtigungen
6. Das Buchhaltungsmenü sollte nun in deinem Sheet erscheinen

## Ordnerstruktur

Für die automatische Dokumentenverarbeitung wird folgende Ordnerstruktur in Google Drive benötigt:

```
└── [Dein Spreadsheet]
    ├── Einnahmen/
    │   └── [Eingangsrechnungen als PDF]
    └── Ausgaben/
        └── [Ausgangsrechnungen als PDF]
```

Die App erkennt automatisch Dateinamen mit Datumsangaben in verschiedenen Formaten.

## Sheet-Struktur

Die folgenden Sheets werden von der App verwendet oder erstellt:

- **Einnahmen**: Eingabe der Einnahmen
- **Ausgaben**: Eingabe der Ausgaben
- **Eigenbelege**: Eingabe von Eigenbelegen
- **Bankbewegungen**: Erfassung aller Bankbewegungen
- **Gesellschafterkonto**: Eingabe von Gesellschafterkonten
- **Holding Transfers**: Eingabe von Holding-Transfers
- **Änderungshistorie**: Protokollierung aller Änderungen und Importe
- **UStVA**: Generierte Umsatzsteuervoranmeldung (monatlich und quartalsweise)
- **BWA**: Generierte Betriebswirtschaftliche Auswertung mit Kennzahlen
- **Bilanz**: Generierte SKR04-konforme Bilanz

## Konfiguration

Die grundlegende Konfiguration erfolgt in der `config.js` Datei:

- **Geschäftsjahr**: Das aktuelle Geschäftsjahr
- **Stammkapital**: Höhe des Stammkapitals
- **Holding/Operativ**: Art des Unternehmens
- **Steuersätze**: Konfiguration der Steuersätze (MwSt, Gewerbesteuer, etc.)
- **Kategorien**: Einnahmen- und Ausgabenkategorien mit Steuertypen
- **Kontenmapping**: SKR04-konforme Konten für jede Kategorie
- **BWA-Mapping**: Zuordnung der Kategorien zu BWA-Positionen

## Funktionen

### Import von Dateien
- Automatische Erkennung von Rechnungsdaten aus Dateinamen
- Protokollierung aller Importe in der Änderungshistorie
- Automatische Erstellung fehlender Ordner
- Batch-Import für bessere Performance

### Berechnungen
- **UStVA**: Automatische Berechnung von Umsatzsteuer und Vorsteuer nach deutschem Steuerrecht
- **BWA**: Berechnung wichtiger Kennzahlen wie Rohertrag, EBIT und EBITDA
- **Bilanz**: Erstellung einer Bilanz mit Anlagevermögen, Umlaufvermögen, Eigenkapital und Verbindlichkeiten
- **Caching**: Intelligentes Caching von Berechnungsergebnissen für bessere Performance

### Validierung
- Validierung aller Eingaben auf Vollständigkeit und Korrektheit
- Prüfung auf erlaubte MwSt-Sätze
- Validierung von Datums- und Zahlenwerten
- Verbesserte Fehlerberichterstattung

### Performance-Optimierung
- Effiziente Batch-Operationen für API-Calls
- Caching-Mechanismen für rechenintensive Operationen
- Optimierte Suche und Matching-Algorithmen

## Lizenz

MIT

## Mitwirkende

- Hendrik Werner ([@hendrikeng](https://github.com/hendrikeng))

## Hinweise

- Diese App ersetzt keine professionelle Buchhaltungssoftware oder Steuerberatung
- Für die steuerliche Verwendung konsultiere bitte einen Steuerberater
- Die App wird ständig weiterentwickelt und verbessert