// file: src/code.js
// imports
import ImportModule from "./importModule.js";
import RefreshModule from "./refreshModule.js";
import UStVACalculator from "./uStVACalculator.js";
import BWACalculator from "./bWACalculator.js";
import BilanzCalculator from "./bilanzCalculator.js"; // Aktualisierter Import

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
        SpreadsheetApp.newTrigger("onOpen")
            .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
            .onOpen()
            .create();
    }

    // Prüfe, ob der onEdit Trigger bereits existiert
    if (!triggers.some(t => t.getHandlerFunction() === "onEdit")) {
        SpreadsheetApp.newTrigger("onEdit")
            .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
            .onEdit()
            .create();
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
    try {
        RefreshModule.refreshAllSheets();
        UStVACalculator.calculateUStVA();
    } catch (error) {
        SpreadsheetApp.getUi().alert("Fehler bei der UStVA-Berechnung: " + error.message);
        console.error("UStVA-Fehler:", error);
    }
};

/**
 * Berechnet die BWA (Betriebswirtschaftliche Auswertung)
 */
const calculateBWA = () => {
    try {
        RefreshModule.refreshAllSheets();
        BWACalculator.calculateBWA();
    } catch (error) {
        SpreadsheetApp.getUi().alert("Fehler bei der BWA-Berechnung: " + error.message);
        console.error("BWA-Fehler:", error);
    }
};

/**
 * Erstellt die Bilanz
 */
const calculateBilanz = () => {
    try {
        RefreshModule.refreshAllSheets();
        BilanzCalculator.calculateBilanz(); // Aktualisierter Aufruf
    } catch (error) {
        SpreadsheetApp.getUi().alert("Fehler bei der Bilanzerstellung: " + error.message);
        console.error("Bilanz-Fehler:", error);
    }
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
