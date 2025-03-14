// file: src/code.js
// imports
import config from './config/index.js';
import ImportModule from './modules/importModule/index.js';
import RefreshModule from './modules/refreshModule/index.js';
import UStVAModule from './modules/ustvaModule/index.js';
import BWAModule from './modules/bwaModule/index.js';
import BilanzModule from './modules/bilanzModule/index.js';
import ValidatorModule from './modules/validatorModule/index.js';

// =================== Globale Funktionen ===================
/**
 * Erstellt das Menü in der Google Sheets UI beim Öffnen der Tabelle
 */
function onOpen() {
    SpreadsheetApp.getUi()
        .createMenu("📂 Buchhaltung")
        .addItem("📥 Dateien importieren", "importDriveFiles")
        .addItem("🔄 Aktuelles Blatt aktualisieren", "refreshSheet")
        .addItem("📊 UStVA berechnen", "calculateUStVA")
        .addItem("📈 BWA berechnen", "calculateBWA")
        .addItem("📝 Bilanz erstellen", "calculateBilanz")
        .addToUi();
}

/**
 * Wird bei jeder Bearbeitung des Spreadsheets ausgelöst
 * Fügt Zeitstempel hinzu, wenn bestimmte Blätter bearbeitet werden
 *
 * @param {Object} e - Event-Objekt von Google Sheets
 */
function onEdit(e) {
    const {range} = e;
    const sheet = range.getSheet();
    const name = sheet.getName();

    // Header-Zeile überspringen (Zeile 1)
    if (range.getRow() <= 1) return;

    // Konvertieren in kleinbuchstaben und leerzeichen entfernen
    let sheetKey = name.toLowerCase().replace(/\s+/g, '');

    // Nach entsprechendem Schlüssel in config suchen (case-insensitive)
    let configKey = null;
    Object.keys(config).forEach(key => {
        if (key.toLowerCase() === sheetKey) {
            configKey = key;
        }
    });

    // Wenn kein passender Schlüssel gefunden wurde, abbrechen
    if (!configKey || !config[configKey].columns.zeitstempel) return;

    // Spalte für Zeitstempel aus der Konfiguration
    const timestampCol = config[configKey].columns.zeitstempel;

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

        // Header-Zeile überspringen
        if (rowIndex <= 1) continue;

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
}

/**
 * Richtet die notwendigen Trigger für das Spreadsheet ein
 */
function setupTrigger() {
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
}

/**
 * Validiert alle relevanten Sheets
 * @returns {boolean} - True wenn alle Sheets valide sind, False sonst
 */
function validateSheets() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const revenueSheet = ss.getSheetByName("Einnahmen");
    const expenseSheet = ss.getSheetByName("Ausgaben");
    const bankSheet = ss.getSheetByName("Bankbewegungen");
    const eigenSheet = ss.getSheetByName("Eigenbelege");

    return ValidatorModule.validateAllSheets(revenueSheet, expenseSheet, bankSheet, eigenSheet, config);
}

/**
 * Gemeinsame Fehlerbehandlungsfunktion für alle Berechnungsfunktionen
 * @param {Function} fn - Die auszuführende Funktion
 * @param {string} errorMessage - Die Fehlermeldung bei einem Fehler
 */
function executeWithErrorHandling(fn, errorMessage) {
    try {
        // Zuerst alle Sheets aktualisieren
        RefreshModule.refreshAllSheets(config);

        // Dann Validierung durchführen
        if (!validateSheets()) {
            // Validierung fehlgeschlagen - Berechnung abbrechen
            console.error(`${errorMessage}: Validierung fehlgeschlagen`);
            return;
        }

        // Wenn Validierung erfolgreich, Berechnung ausführen
        fn();
    } catch (error) {
        SpreadsheetApp.getUi().alert(`${errorMessage}: ${error.message}`);
        console.error(`${errorMessage}:`, error);
    }
}

/**
 * Aktualisiert das aktive Tabellenblatt
 */
function refreshSheet() {
    RefreshModule.refreshActiveSheet(config);
}

/**
 * Berechnet die Umsatzsteuervoranmeldung
 */
function calculateUStVA() {
    executeWithErrorHandling(
        () => UStVAModule.calculateUStVA(config),
        "Fehler bei der UStVA-Berechnung"
    );
}

/**
 * Berechnet die BWA (Betriebswirtschaftliche Auswertung)
 */
function calculateBWA() {
    executeWithErrorHandling(
        () => BWAModule.calculateBWA(config),
        "Fehler bei der BWA-Berechnung"
    );
}

/**
 * Erstellt die Bilanz
 */
function calculateBilanz() {
    executeWithErrorHandling(
        () => BilanzModule.calculateBilanz(config),
        "Fehler bei der Bilanzerstellung"
    );
}

/**
 * Importiert Dateien aus Google Drive und aktualisiert alle Tabellenblätter
 */
function importDriveFiles() {
    try {
        ImportModule.importDriveFiles(config);
        RefreshModule.refreshAllSheets(config);

        // Optional: Nach dem Import auch eine Validierung durchführen
        // validateSheets();
    } catch (error) {
        SpreadsheetApp.getUi().alert("Fehler beim Dateiimport: " + error.message);
        console.error("Import-Fehler:", error);
    }
}

// Exportiere alle relevanten Funktionen in den globalen Namensraum,
// damit sie von Google Apps Script als Trigger und Menüpunkte aufgerufen werden können.
global.onOpen = onOpen;
global.onEdit = onEdit;
global.setupTrigger = setupTrigger;
global.refreshSheet = refreshSheet;
global.calculateUStVA = calculateUStVA;
global.calculateBWA = calculateBWA;
global.calculateBilanz = calculateBilanz;
global.importDriveFiles = importDriveFiles;