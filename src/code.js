// file: src/code.js
// imports
import config from './config/index.js';
import ImportModule from './modules/importModule/index.js';
import RefreshModule from './modules/refreshModule/index.js';
import UStVAModule from './modules/ustvaModule/index.js';
import BWAModule from './modules/bwaModule/index.js';
import BilanzModule from './modules/bilanzModule/index.js';
import ValidatorModule from './modules/validatorModule/index.js';
import SetupModule from './modules/setupModule/index.js';
import BankReconciliationModule from './modules/bankReconciliationModule/index.js';

// =================== Globale Funktionen ===================
/**
 * Erstellt das Menü in der Google Sheets UI beim Öffnen der Tabelle
 * Optimierte Version mit strukturiertem Menüaufbau
 */
function onOpen() {
    try {
        SpreadsheetApp.getUi()
            .createMenu('📂 Buchhaltung')
            .addItem('🛠️ Buchhaltung einrichten', 'setupSpreadsheet')
            .addSeparator()
            .addItem('📥 Dateien importieren', 'importDriveFiles')
            .addItem('🔄 Aktuelles Blatt aktualisieren', 'refreshSheet')
            .addItem('🏦 Bankabgleich durchführen', 'bankReconciliation')
            .addSeparator()
            .addItem('📊 UStVA berechnen', 'calculateUStVA')
            .addItem('📈 BWA berechnen', 'calculateBWA')
            .addItem('📝 Bilanz erstellen', 'calculateBilanz')
            .addToUi();
    } catch (e) {
        console.error('Fehler beim Erstellen des Menüs:', e);
        // Silent failure - don't interrupt user experience
    }
}

/**
 * Wird bei jeder Bearbeitung des Spreadsheets ausgelöst
 * Optimierte Version mit besserer Performance und Fehlerbehandlung
 *
 * @param {Object} e - Event-Objekt von Google Sheets
 */
function onEdit(e) {
    try {
        const {range} = e;
        const sheet = range.getSheet();
        const name = sheet.getName();

        // Header-Zeile überspringen
        if (range.getRow() <= 1) return;

        // Cache für Sheet-Konfigurationen
        const sheetConfigCache = {};

        // Effizientere Sheet-Typ-Bestimmung
        const getSheetConfig = (sheetName) => {
            if (sheetConfigCache[sheetName]) return sheetConfigCache[sheetName];

            // Konvertieren in kleinbuchstaben und leerzeichen entfernen
            const normalized = sheetName.toLowerCase().replace(/\s+/g, '');

            // Direkte Zuordnung für schnelleren Lookup
            const directMap = {
                'einnahmen': 'einnahmen',
                'ausgaben': 'ausgaben',
                'eigenbelege': 'eigenbelege',
                'gesellschafterkonto': 'gesellschafterkonto',
                'holdingtransfers': 'holdingTransfers',
                'bankbewegungen': 'bankbewegungen',
            };

            if (directMap[normalized]) {
                sheetConfigCache[sheetName] = directMap[normalized];
                return directMap[normalized];
            }

            // Fallback auf vollständige Suche
            for (const key in config) {
                if (key.toLowerCase() === normalized) {
                    sheetConfigCache[sheetName] = key;
                    return key;
                }
            }

            return null;
        };

        // Bestimme Sheet-Typ
        const configKey = getSheetConfig(name);
        if (!configKey || !config[configKey].columns || !config[configKey].columns.zeitstempel) {
            return;
        }

        // Spalte für Zeitstempel
        const timestampCol = config[configKey].columns.zeitstempel;

        // Prüfen, ob die bearbeitete Spalte der Zeitstempel ist (verhindert Endlosschleife)
        if (range.getColumn() === timestampCol ||
            (range.getNumColumns() > 1 && range.getColumn() <= timestampCol &&
                range.getColumn() + range.getNumColumns() > timestampCol)) {
            return;
        }

        // Optimierte Verarbeitung von Multi-Zellen-Editierungen
        const numRows = range.getNumRows();
        const rowsToUpdate = [];
        const now = new Date();

        // Identifiziere nur nicht-leere Zeilen für Updates
        for (let i = 0; i < numRows; i++) {
            const rowIndex = range.getRow() + i;
            if (rowIndex <= 1) continue; // Header überspringen

            // Prüfe, ob die Zeile nicht leer ist (nur erste 5 Spalten zur Performance-Optimierung)
            const checkRange = Math.min(5, sheet.getLastColumn());
            const rowValues = sheet.getRange(rowIndex, 1, 1, checkRange).getValues()[0];

            if (rowValues.some(cell => cell !== '' && cell !== null && cell !== undefined)) {
                rowsToUpdate.push(rowIndex);
            }
        }

        // Batch-Update für Zeitstempel, wenn Zeilen gefunden wurden
        if (rowsToUpdate.length > 0) {
            // Gruppiere zusammenhängende Zeilen für effizienteres Update
            const ranges = [];
            let currentRange = [rowsToUpdate[0], rowsToUpdate[0]];

            for (let i = 1; i < rowsToUpdate.length; i++) {
                if (rowsToUpdate[i] === currentRange[1] + 1) {
                    currentRange[1] = rowsToUpdate[i];
                } else {
                    ranges.push([...currentRange]);
                    currentRange = [rowsToUpdate[i], rowsToUpdate[i]];
                }
            }
            ranges.push([...currentRange]);

            // Update jede Range mit Zeitstempeln
            ranges.forEach(([startRow, endRow]) => {
                const numRowsInRange = endRow - startRow + 1;
                const timestampValues = Array(numRowsInRange).fill([now]);
                sheet.getRange(startRow, timestampCol, numRowsInRange, 1)
                    .setValues(timestampValues);
            });
        }
    } catch (e) {
        console.error('Fehler beim onEdit-Handler:', e);
        // Silent failure to avoid interrupting user
    }
}

/**
 * Richtet die notwendigen Trigger für das Spreadsheet ein
 * Optimierte Version mit Überprüfung auf bereits vorhandene Trigger
 */
function setupTrigger() {
    try {
        const triggers = ScriptApp.getProjectTriggers();
        const ss = SpreadsheetApp.getActiveSpreadsheet();

        // Definiere erforderliche Trigger
        const requiredTriggers = [
            { func: 'onOpen', event: ScriptApp.EventType.ON_OPEN },
            { func: 'onEdit', event: ScriptApp.EventType.ON_EDIT },
        ];

        // Prüfe und erstelle jeden erforderlichen Trigger
        requiredTriggers.forEach(({func, event}) => {
            // Prüfe, ob der Trigger bereits existiert
            const exists = triggers.some(t =>
                t.getHandlerFunction() === func &&
                t.getEventType() === event &&
                t.getTriggerSourceId() === ss.getId(),
            );

            // Erstelle den Trigger nur, wenn er nicht existiert
            if (!exists) {
                ScriptApp.newTrigger(func)
                    .forSpreadsheet(ss)
                    .on(event)
                    .create();
                console.log(`Trigger für ${func} erstellt`);
            }
        });

        return true;
    } catch (e) {
        console.error('Fehler beim Einrichten der Trigger:', e);
        SpreadsheetApp.getUi().alert('Fehler beim Einrichten der automatischen Trigger: ' + e.toString());
        return false;
    }
}

/**
 * Richtet die grundlegende Struktur für die Buchhaltung ein
 * Optimierte Version mit besserer Fehlerbehandlung
 */
function setupSpreadsheet() {
    return SetupModule.setupSpreadsheet(config);
}

/**
 * Validiert alle relevanten Sheets
 * @returns {boolean} - True wenn alle Sheets valide sind, False sonst
 */
function validateSheets() {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();

        // Alle benötigten Sheets in einem Batch laden
        const sheets = {
            revenueSheet: ss.getSheetByName('Einnahmen'),
            expenseSheet: ss.getSheetByName('Ausgaben'),
            bankSheet: ss.getSheetByName('Bankbewegungen'),
            eigenSheet: ss.getSheetByName('Eigenbelege'),
        };

        // Fehler-Check für kritische Sheets
        if (!sheets.revenueSheet || !sheets.expenseSheet) {
            SpreadsheetApp.getUi().alert('Kritische Sheets (Einnahmen oder Ausgaben) fehlen!');
            return false;
        }

        return ValidatorModule.validateAllSheets(
            sheets.revenueSheet,
            sheets.expenseSheet,
            sheets.bankSheet,
            sheets.eigenSheet,
            config,
        );
    } catch (e) {
        console.error('Fehler bei der Sheet-Validierung:', e);
        SpreadsheetApp.getUi().alert('Fehler bei der Validierung: ' + e.toString());
        return false;
    }
}

/**
 * Gemeinsame Fehlerbehandlungsfunktion für alle Berechnungsfunktionen
 * Optimierte Version mit besserer Fehlerbehandlung und UI-Feedback
 * @param {Function} fn - Die auszuführende Funktion
 * @param {string} errorMessage - Die Fehlermeldung bei einem Fehler
 * @param {string} successMessage - Die Erfolgsmeldung
 * @returns {boolean} - Erfolg der Operation
 */
function executeWithErrorHandling(fn, errorMessage, successMessage) {
    const ui = SpreadsheetApp.getUi();
    try {
        ui.alert('Verarbeitung läuft...', 'Bitte warten Sie, während die Daten aktualisiert werden.', ui.ButtonSet.OK);

        // Zuerst alle Sheets aktualisieren
        RefreshModule.refreshAllSheets(config);

        // Dann Validierung durchführen
        if (!validateSheets()) {
            // Validierung fehlgeschlagen - Berechnung abbrechen
            console.error(`${errorMessage}: Validierung fehlgeschlagen`);
            ui.alert('Validierungsfehler', 'Die Berechnung wurde abgebrochen, da Fehler in den Daten gefunden wurden. Bitte korrigieren Sie die angezeigten Fehler.', ui.ButtonSet.OK);
            return false;
        }

        // Wenn Validierung erfolgreich, Berechnung ausführen
        const result = fn();

        if (result) {
            if (successMessage) {
                ui.alert('Erfolgreich', successMessage, ui.ButtonSet.OK);
            }
            return true;
        } else {
            ui.alert('Fehler', 'Die Berechnung konnte nicht abgeschlossen werden.', ui.ButtonSet.OK);
            return false;
        }
    } catch (error) {
        console.error(`${errorMessage}:`, error);
        ui.alert('Fehler', `${errorMessage}: ${error.message}`, ui.ButtonSet.OK);
        return false;
    }
}

/**
 * Aktualisiert das aktive Tabellenblatt
 * Optimierte Version mit Statusanzeige
 */
function refreshSheet() {
    try {
        const success = RefreshModule.refreshActiveSheet(config);
        return success;
    } catch (error) {
        SpreadsheetApp.getUi().alert('Fehler beim Aktualisieren: ' + error.message);
        console.error('Refresh-Fehler:', error);
        return false;
    }
}

/**
 * Berechnet die Umsatzsteuervoranmeldung
 * @returns {boolean} - Erfolg der Berechnung
 */
function calculateUStVA() {
    return executeWithErrorHandling(
        () => UStVAModule.calculateUStVA(config),
        'Fehler bei der UStVA-Berechnung',
        'Die UStVA wurde erfolgreich berechnet!',
    );
}

/**
 * Berechnet die BWA (Betriebswirtschaftliche Auswertung)
 * @returns {boolean} - Erfolg der Berechnung
 */
function calculateBWA() {
    return executeWithErrorHandling(
        () => BWAModule.calculateBWA(config),
        'Fehler bei der BWA-Berechnung',
        'Die BWA wurde erfolgreich berechnet!',
    );
}

/**
 * Erstellt die Bilanz
 * @returns {boolean} - Erfolg der Berechnung
 */
function calculateBilanz() {
    return executeWithErrorHandling(
        () => BilanzModule.calculateBilanz(config),
        'Fehler bei der Bilanzerstellung',
        'Die Bilanz wurde erfolgreich erstellt!',
    );
}

/**
 * Importiert Dateien aus Google Drive und aktualisiert alle Tabellenblätter
 * Optimierte Version mit Fortschrittsanzeige und besserer Fehlerbehandlung
 * @returns {boolean} - Erfolg des Imports
 */
function importDriveFiles() {
    const ui = SpreadsheetApp.getUi();
    try {
        ui.alert('Import wird gestartet', 'Bitte warten Sie, während die Dateien importiert werden.', ui.ButtonSet.OK);

        // Import durchführen
        const importCount = ImportModule.importDriveFiles(config);

        if (importCount > 0) {
            // Wenn Dateien importiert wurden, Sheets aktualisieren
            RefreshModule.refreshAllSheets(config);
            ui.alert('Import abgeschlossen', `Es wurden ${importCount} Dateien erfolgreich importiert und alle Sheets aktualisiert.`, ui.ButtonSet.OK);
            return true;
        } else if (importCount === 0) {
            // Keine Dateien gefunden oder importiert
            ui.alert('Kein Import erfolgt', 'Es wurden keine neuen Dateien gefunden.', ui.ButtonSet.OK);
            return true;
        } else {
            // Fehler beim Import
            ui.alert('Import fehlgeschlagen', 'Es ist ein Fehler beim Importieren der Dateien aufgetreten.', ui.ButtonSet.OK);
            return false;
        }
    } catch (error) {
        console.error('Import-Fehler:', error);
        ui.alert('Fehler beim Dateiimport', error.message, ui.ButtonSet.OK);
        return false;
    }
}

/**
 * Führt den Bankabgleich durch
 * Optimierte Version mit Statusmeldungen
 * @returns {boolean} - Erfolg des Bankabgleichs
 */
function bankReconciliation() {
    const ui = SpreadsheetApp.getUi();
    try {
        ui.alert('Bankabgleich wird gestartet', 'Bitte warten Sie, während der Bankabgleich durchgeführt wird.', ui.ButtonSet.OK);

        // Zuerst Sheets aktualisieren
        RefreshModule.refreshAllSheets(config);

        // Dann Bankabgleich durchführen
        const success = BankReconciliationModule.reconcileBank(config);

        if (success) {
            ui.alert('Bankabgleich abgeschlossen', 'Der Bankabgleich wurde erfolgreich durchgeführt.', ui.ButtonSet.OK);
        }

        return success;
    } catch (error) {
        console.error('Bankabgleich-Fehler:', error);
        ui.alert('Fehler beim Bankabgleich', error.message, ui.ButtonSet.OK);
        return false;
    }
}


// Exportiere alle relevanten Funktionen in den globalen Namensraum,
// damit sie von Google Apps Script als Trigger und Menüpunkte aufgerufen werden können.
global.onOpen = onOpen;
global.onEdit = onEdit;
global.setupTrigger = setupTrigger;
global.setupSpreadsheet = setupSpreadsheet;
global.refreshSheet = refreshSheet;
global.calculateUStVA = calculateUStVA;
global.calculateBWA = calculateBWA;
global.calculateBilanz = calculateBilanz;
global.importDriveFiles = importDriveFiles;
global.bankReconciliation = bankReconciliation;