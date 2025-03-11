/**
 * Haupteinstiegspunkt für das Google Apps Script
 * Importiert alle Module und stellt globale Funktionen bereit
 */

// Importiere Konfiguration und Module
// Die Module werden inline expandiert, nachdem die Code-Datei gebündelt wurde

// Konfiguration importieren
// @include "config.js"

// Hilfs- und Dienstprogramme
// @include "helperModule.js"
// @include "accountingModule.js"
// @include "validationModule.js"

// Hauptfunktionalitätsmodule
// @include "bankModule.js"
// @include "einnahmenModule.js"
// @include "ausgabenModule.js"
// @include "eigenbelagModule.js"
// @include "ustvaModule.js"
// @include "bwaModule.js"
// @include "importModule.js"
// @include "bilanzModule.js"

// =================== Globale Funktionen ===================

/**
 * Wird beim Öffnen des Spreadsheets automatisch ausgeführt
 * Fügt das Hauptmenü hinzu
 */
function onOpen() {
    const ui = SpreadsheetApp.getUi();

    ui.createMenu('📂 Buchhaltung')
        .addSubMenu(ui.createMenu('Daten')
            .addItem('📥 Dateien importieren', 'importDriveFiles')
            .addItem('🔄 Aktuelles Sheet aktualisieren', 'refreshSheet')
            .addItem('🔍 Bankbewegungen automatisch zuordnen', 'matchBankMovements')
            .addSeparator()
            .addItem('📊 Alle Daten aktualisieren', 'refreshAllData'))
        .addSubMenu(ui.createMenu('Berichte')
            .addItem('📊 UStVA berechnen', 'calculateUStVA')
            .addItem('📈 BWA berechnen', 'calculateBWA')
            .addItem('📝 Bilanz erstellen', 'generateBilanz')
            .addSeparator()
            .addItem('📤 ELSTER-Export erstellen', 'createElsterExport'))
        .addSubMenu(ui.createMenu('Einstellungen')
            .addItem('⚙️ Konfiguration', 'showConfigDialog')
            .addItem('👤 Benutzereinstellungen', 'showUserSettings')
            .addItem('🔧 System-Diagnose', 'runSystemDiagnostics'))
        .addSubMenu(ui.createMenu('Hilfe')
            .addItem('📖 Anleitung', 'showHelp')
            .addItem('📅 Steuertermine', 'showTaxCalendar')
            .addItem('ℹ️ Über', 'showAbout'))
        .addToUi();

    Logger.log('Buchhaltungsmenü erfolgreich erstellt');
}

/**
 * Wird bei jeder Änderung im Spreadsheet automatisch ausgeführt
 * Aktualisiert die "Letzte Aktualisierung"-Spalte
 */
function onEdit(e) {
    const {range} = e;
    const sheet = range.getSheet();
    const name = sheet.getName();

    // Spaltenmapping für "Letzte Aktualisierung"
    const mapping = {
        "Einnahmen": CONFIG.SYSTEM.EINNAHMEN_COLS.LETZTE_AKTUALISIERUNG,
        "Ausgaben": CONFIG.SYSTEM.AUSGABEN_COLS.LETZTE_AKTUALISIERUNG,
        "Eigenbelege": CONFIG.SYSTEM.EIGENBELEGE_COLS.LETZTE_AKTUALISIERUNG,
        "Bankbewegungen": CONFIG.SYSTEM.BANKBEWEGUNGEN_COLS.LETZTE_AKTUALISIERUNG,
        "Gesellschafterkonto": CONFIG.SYSTEM.GESELLSCHAFTERKONTO_COLS.LETZTE_AKTUALISIERUNG,
        "Holding Transfers": CONFIG.SYSTEM.HOLDING_TRANSFERS_COLS.LETZTE_AKTUALISIERUNG
    };

    // Prüfen, ob das Sheet im Mapping enthalten ist
    if (!(name in mapping)) return;

    // Header-Zeile ignorieren
    if (range.getRow() === 1) return;

    // Spalten außerhalb des Headers ignorieren
    const headerLen = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].length;
    if (range.getColumn() > headerLen) return;

    // Änderungen in der Spalte "Letzte Aktualisierung" ignorieren
    if (range.getColumn() === mapping[name] + 1) return;

    // Leere Zeilen ignorieren
    const rowValues = sheet.getRange(range.getRow(), 1, 1, headerLen).getValues()[0];
    if (rowValues.every(cell => cell === "")) return;

    // Timestamp aktualisieren
    const ts = new Date();
    sheet.getRange(range.getRow(), mapping[name] + 1).setValue(ts);
}

/**
 * Richtet die notwendigen Trigger für das Projekt ein
 */
function setupTriggers() {
    const triggers = ScriptApp.getProjectTriggers();

    // onOpen-Trigger
    if (!triggers.some(t => t.getHandlerFunction() === "onOpen")) {
        ScriptApp.newTrigger("onOpen")
            .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
            .onOpen()
            .create();
    }

    // onEdit-Trigger
    if (!triggers.some(t => t.getHandlerFunction() === "onEdit")) {
        ScriptApp.newTrigger("onEdit")
            .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
            .onEdit()
            .create();
    }

    return triggers.length;
}

// =================== Wrapper-Funktionen für die UI ===================

/**
 * Aktualisiert das aktive Sheet
 */
function refreshSheet() {
    return refreshActiveSheet();
}

/**
 * Importiert Dateien aus Google Drive
 */
function importDriveFiles() {
    ImportModule.importDriveFiles();
    refreshAllData();
}

/**
 * Startet die automatische Zuordnung von Bankbewegungen
 */
function matchBankMovements() {
    return BankModule.matchBankbewegungToRechnung();
}

/**
 * Berechnet die UStVA
 */
function calculateUStVA() {
    refreshAllData();
    return UStVAModule.calculateUStVA();
}

/**
 * Berechnet die BWA
 */
function calculateBWA() {
    refreshAllData();
    return BWAModule.calculateBWA();
}

/**
 * Erstellt die Bilanz
 */
function generateBilanz() {
    refreshAllData();
    return BilanzModule.generateBilanz();
}

/**
 * Erstellt einen ELSTER-Export
 */
function createElsterExport() {
    return UStVAModule.createElsterExport();
}

/**
 * Zeigt den Dialog für die Konfiguration an
 */
function showConfigDialog() {
    // Platzhalter bis SettingsModule implementiert ist
    const ui = SpreadsheetApp.getUi();
    ui.alert(
        'Konfiguration',
        'Diese Funktion wird in einer zukünftigen Version implementiert.',
        ui.ButtonSet.OK
    );
}

/**
 * Zeigt den Dialog für die Benutzereinstellungen an
 */
function showUserSettings() {
    // Platzhalter bis SettingsModule implementiert ist
    const ui = SpreadsheetApp.getUi();
    ui.alert(
        'Benutzereinstellungen',
        'Diese Funktion wird in einer zukünftigen Version implementiert.',
        ui.ButtonSet.OK
    );
}

/**
 * Führt eine Systemdiagnose durch
 */
function runSystemDiagnostics() {
    // Platzhalter bis DiagnosticsModule implementiert ist
    const ui = SpreadsheetApp.getUi();
    ui.alert(
        'System-Diagnose',
        'Diese Funktion wird in einer zukünftigen Version implementiert.',
        ui.ButtonSet.OK
    );
}

/**
 * Zeigt die Hilfe an
 */
function showHelp() {
    // Platzhalter bis HelpModule implementiert ist
    const ui = SpreadsheetApp.getUi();
    ui.alert(
        'Hilfe',
        'Diese Funktion wird in einer zukünftigen Version implementiert.',
        ui.ButtonSet.OK
    );
}

/**
 * Zeigt den Steuerkalender an
 */
function showTaxCalendar() {
    // Platzhalter bis CalendarModule implementiert ist
    const ui = SpreadsheetApp.getUi();
    ui.alert(
        'Steuertermine',
        'Diese Funktion wird in einer zukünftigen Version implementiert.',
        ui.ButtonSet.OK
    );
}

/**
 * Zeigt Informationen über die Anwendung an
 */
function showAbout() {
    const ui = SpreadsheetApp.getUi();
    ui.alert(
        'Über Buchhaltungssystem',
        'Buchhaltungssystem für GmbHs\n' +
        'Version: 1.0.0\n' +
        'Entwickelt mit Google Apps Script\n\n' +
        'Copyright © ' + new Date().getFullYear(),
        ui.ButtonSet.OK
    );
}