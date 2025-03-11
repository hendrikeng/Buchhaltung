# Buchhaltungs-App für Google Sheets

Eine umfassende Buchhaltungslösung für kleine Unternehmen, die auf Google Sheets basiert. Unterstützt sowohl operative GmbHs als auch Holding-Strukturen mit SKR04-konformen Kontenrahmen.

## Features

- **Dokumenten-Management**: Automatischer Import von Rechnungen aus Google Drive
- **Steuerliche Auswertungen**: Berechnung von UStVA mit ELSTER-Unterstützung
- **Betriebswirtschaftliche Auswertung (BWA)**: Monatliche und quartalsweise Auswertungen
- **Bilanzerstellung**: Automatische Generierung einer einfachen Bilanz
- **Kontenjournal**: Erstellung von Buchungsjournalen, Kontoblättern und Monatsjournalen
- **Validierung**: Umfassende Validierung aller Eingaben
- **Flexibilität**: Unterstützung für verschiedene Steuerszenarien (Holding/Operativ)

## Module

Die App besteht aus folgenden Modulen:

- **importModule.js**: Import von Dateien aus Google Drive
- **refreshModule.js**: Aktualisierung der Sheets und Formeln
- **uStVACalculator.js**: Berechnung der Umsatzsteuervoranmeldung
- **bWACalculator.js**: Erstellung der Betriebswirtschaftlichen Auswertung
- **calculateBilanz.js**: Generierung einer einfachen Bilanz
- **journalModule.js**: Erstellung von Kontenjournalen und Buchungsberichten
- **settings.js**: Verwaltung der Anwendungseinstellungen
- **validator.js**: Validierung der Eingaben
- **helpers.js**: Hilfsfunktionen für Datum, Währung etc.
- **config.js**: Zentrale Konfigurationsdatei
- **code.js**: Hauptmodul mit UI-Integration

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

## Sheet-Struktur

Die folgenden Sheets werden von der App verwendet oder erstellt:

- **Einnahmen**: Eingabe der Einnahmen
- **Ausgaben**: Eingabe der Ausgaben
- **Eigenbelege**: Eingabe von Eigenbelegen
- **Bankbewegungen**: Erfassung aller Bankbewegungen
- **UStVA**: Generierte Umsatzsteuervoranmeldung
- **ELSTER-UStVA**: ELSTER-kompatible Umsatzsteuervoranmeldung
- **BWA**: Generierte Betriebswirtschaftliche Auswertung
- **Bilanz**: Generierte Bilanz
- **Kontenjournal**: Generiertes Kontenjournal
- **Konto [Kontonummer]**: Generierte Kontoblätter
- **Journal [Monat]**: Generierte Monatsjournale
- **Einstellungen**: Konfiguration der App

## Konfiguration

Die grundlegende Konfiguration erfolgt in der `config.js` Datei:

- **Geschäftsjahr**: Das aktuelle Geschäftsjahr
- **Stammkapital**: Höhe des Stammkapitals
- **Holding/Operativ**: Art des Unternehmens
- **Steuersätze**: Konfiguration der Steuersätze
- **Kategorien**: Einnahmen- und Ausgabenkategorien
- **Kontenmapping**: SKR04-konforme Konten

Weitere Einstellungen können über das Menü "Einstellungen" im Sheet vorgenommen werden.

## Lizenz

MIT

## Mitwirkende

- Hendrik Werner ([@hendrikeng](https://github.com/hendrikeng))

## Hinweise

- Diese App ersetzt keine professionelle Buchhaltungssoftware oder Steuerberatung
- Für die steuerliche Verwendung konsultiere bitte einen Steuerberater
- Die App wird ständig weiterentwickelt und verbessert