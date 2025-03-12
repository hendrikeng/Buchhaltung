// file: src/code.js
// imports
import ImportModule from "./importModule.js";
import RefreshModule from "./refreshModule.js";
import UStVACalculator from "./uStVACalculator.js";
import BWACalculator from "./bWACalculator.js";
import BilanzCalculator from "./bilanzCalculator.js";
import config from "./config.js";

// =================== Globale Funktionen ===================
/**
 * Erstellt das Menü in der Google Sheets UI beim Öffnen der Tabelle
 */
const onOpen = () => {
    SpreadsheetApp.getUi()
        .createMenu("📂 Buchhaltung")
        .addItem("📥 Dateien importieren", "importDriveFiles")
        .addItem("🔄 Aktuelles Blatt aktualisieren", "refreshSheet")
        .addItem("📊 UStVA berechnen", "calculateUStVA")
        .addItem("📈 BWA berechnen", "calculateBWA")
        .addItem("📝 Bilanz erstellen", "calculateBilanz")
        .addToUi();
};

/**
 * Wird bei jeder Bearbeitung des Spreadsheets ausgelöst
 * Fügt Zeitstempel hinzu, wenn bestimmte Blätter bearbeitet werden
 *
 * @param {Object} e - Event-Objekt von Google Sheets
 */
const onEdit = e => {
    const {range} = e;
    const sheet = range.getSheet();
    const name = sheet.getName();
    const sheetKey = name.toLowerCase();

    // Prüfen, ob wir für dieses Sheet eine Konfiguration haben
    if (!config.sheets[sheetKey] || !config.sheets[sheetKey].columns.zeitstempel) return;

    // Header-Zeile ignorieren
    if (range.getRow() === 1) return;

    // Spalte für Zeitstempel aus der Konfiguration
    const timestampCol = config.sheets[sheetKey].columns.zeitstempel;

    // Prüfen, ob die bearbeitete Spalte bereits die Zeitstempel-Spalte ist
    if (range.getColumn() === timestampCol ||
        (range.getNumColumns() > 1 && range.getColumn() <= timestampCol &&
            range.getColumn() + range.getNumColumns() > timestampCol)) {
        return; // Vermeide Endlosschleife
    }

    // Multi-Zellen-Editierung unterstützen
    const numRows = range.getNumRows();
    const timestampValues = [];
    const now = new Date();

    // Für jede bearbeitete Zeile
    for (let i = 0; i < numRows; i++) {
        const rowIndex = range.getRow() + i;
        const headerLen = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].length;

        // Prüfen, ob die Zeile leer ist
        const rowValues = sheet.getRange(rowIndex, 1, 1, headerLen).getValues()[0];
        if (rowValues.every(cell => cell === "")) continue;

        // Zeitstempel für diese Zeile hinzufügen
        timestampValues.push([now]);
    }

    // Wenn es Zeitstempel zu setzen gibt
    if (timestampValues.length > 0) {
        sheet.getRange(range.getRow(), timestampCol, timestampValues.length, 1)
            .setValues(timestampValues);
    }
};

/**
 * Richtet die notwendigen Trigger für das Spreadsheet ein
 */
const setupTrigger = () => {
    const triggers = ScriptApp.getProjectTriggers();
    // Prüfe, ob der onOpen Trigger bereits existiert
    if (!triggers.some(t => t.getHandlerFunction() === "onOpen")) {
        ScriptApp.newTrigger("onOpen")
            .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
            .onOpen()
            .create();
    }

    // Prüfe, ob der onEdit Trigger bereits existiert
    if (!triggers.some(t => t.getHandlerFunction() === "onEdit")) {
        ScriptApp.newTrigger("onEdit")
            .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
            .onEdit()
            .create();
    }
};

/**
 * Gemeinsame Fehlerbehandlungsfunktion für alle Berechnungsfunktionen
 * @param {Function} fn - Die auszuführende Funktion
 * @param {string} errorMessage - Die Fehlermeldung bei einem Fehler
 */
const executeWithErrorHandling = (fn, errorMessage) => {
    try {
        RefreshModule.refreshAllSheets();
        fn();
    } catch (error) {
        SpreadsheetApp.getUi().alert(`${errorMessage}: ${error.message}`);
        console.error(`${errorMessage}:`, error);
    }
};

/**
 * Aktualisiert das aktive Tabellenblatt
 */
const refreshSheet = () => RefreshModule.refreshActiveSheet();

/**
 * Berechnet die Umsatzsteuervoranmeldung
 */
const calculateUStVA = () => {
    executeWithErrorHandling(
        UStVACalculator.calculateUStVA,
        "Fehler bei der UStVA-Berechnung"
    );
};

/**
 * Berechnet die BWA (Betriebswirtschaftliche Auswertung)
 */
const calculateBWA = () => {
    executeWithErrorHandling(
        BWACalculator.calculateBWA,
        "Fehler bei der BWA-Berechnung"
    );
};

/**
 * Erstellt die Bilanz
 */
const calculateBilanz = () => {
    executeWithErrorHandling(
        BilanzCalculator.calculateBilanz,
        "Fehler bei der Bilanzerstellung"
    );
};

/**
 * Importiert Dateien aus Google Drive und aktualisiert alle Tabellenblätter
 */
const importDriveFiles = () => {
    try {
        ImportModule.importDriveFiles();
        RefreshModule.refreshAllSheets();
    } catch (error) {
        SpreadsheetApp.getUi().alert("Fehler beim Dateiimport: " + error.message);
        console.error("Import-Fehler:", error);
    }
};