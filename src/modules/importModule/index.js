// modules/importModule/index.js
import fileManager from './fileManager.js';
import dataProcessor from './dataProcessor.js';
import historyTracker from './historyTracker.js';

/**
 * Modul für den Import von Dateien aus Google Drive in die Buchhaltungstabelle
 */
const ImportModule = {
    /**
     * Importiert Dateien aus Google Drive und aktualisiert alle Tabellenblätter
     * Optimierte Version mit Batch-Verarbeitung und Fortschrittsanzeige
     * @param {Object} config - Die Konfiguration
     * @returns {number} - Anzahl der importierten Dateien
     */
    importDriveFiles(config) {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const ui = SpreadsheetApp.getUi();
        let totalImported = 0;

        try {
            console.log('Starting file import...');

            // Optimierung: Cache für die Dateihistorie leeren, um aktuelle Daten zu erhalten
            historyTracker.clearFileHistoryCache();

            // Hauptsheets für Einnahmen, Ausgaben und Eigenbelege abrufen
            const sheets = {
                revenue: ss.getSheetByName('Einnahmen'),
                expense: ss.getSheetByName('Ausgaben'),
                receipts: ss.getSheetByName('Eigenbelege'),
            };

            // Optimierung: Fehlerhafte Sheets in einem Block prüfen
            if (!sheets.revenue || !sheets.expense || !sheets.receipts) {
                ui.alert("Fehler: Die Sheets 'Einnahmen', 'Ausgaben' oder 'Eigenbelege' existieren nicht!");
                return 0;
            }

            // Änderungshistorie abrufen oder erstellen
            const history = historyTracker.getOrCreateHistorySheet(ss, config);

            // Auf übergeordneten Ordner zugreifen
            const parentFolder = fileManager.getParentFolder();
            if (!parentFolder) {
                ui.alert('Fehler: Kein übergeordneter Ordner gefunden.');
                return 0;
            }

            // Unterordner für Einnahmen, Ausgaben und Eigenbelege finden oder erstellen
            const folders = {
                revenue: fileManager.findOrCreateFolder(parentFolder, 'Einnahmen', ui),
                expense: fileManager.findOrCreateFolder(parentFolder, 'Ausgaben', ui),
                receipts: fileManager.findOrCreateFolder(parentFolder, 'Eigenbelege', ui),
            };

            // Optimierung: Wenn kein Ordner gefunden wurde, Import abbrechen
            if (!folders.revenue && !folders.expense && !folders.receipts) {
                ui.alert('Fehler: Keine Importordner gefunden oder erstellt.');
                return 0;
            }

            // Bereits importierte Dateien aus der Historie abrufen
            const existingFiles = historyTracker.collectExistingFiles(history);
            console.log(`Found ${existingFiles.size} already imported files`);

            // Statusdialog anzeigen
            ui.alert(
                'Import gestartet',
                'Der Import der Dateien wurde gestartet. Dies kann je nach Datenmenge einige Zeit dauern.',
                ui.ButtonSet.OK,
            );

            // Import mit Fehlerbehandlung pro Ordner durchführen
            // Hier die Variablen als Parameter übergeben, um ESLint-Fehler zu vermeiden
            const imports = {
                revenue: folders.revenue ? importFilesFromFolder(folders.revenue, sheets.revenue, 'Einnahme', history, existingFiles, config, ui) : 0,
                expense: folders.expense ? importFilesFromFolder(folders.expense, sheets.expense, 'Ausgabe', history, existingFiles, config, ui) : 0,
                receipts: folders.receipts ? importFilesFromFolder(folders.receipts, sheets.receipts, 'Eigenbeleg', history, existingFiles, config, ui) : 0,
            };

            // Gesamtzahl der Importe berechnen
            totalImported = imports.revenue + imports.expense + imports.receipts;

            // Abschluss-Meldung anzeigen
            if (totalImported === 0) {
                ui.alert('Es wurden keine neuen Dateien gefunden.');
            } else {
                ui.alert(
                    'Import abgeschlossen.\n\n' +
                    `${imports.revenue} Einnahmen importiert.\n` +
                    `${imports.expense} Ausgaben importiert.\n` +
                    `${imports.receipts} Eigenbelege importiert.`,
                );
            }

            return totalImported;
        } catch (e) {
            console.error('Unerwarteter Fehler beim Import:', e);
            ui.alert('Ein unerwarteter Fehler ist aufgetreten: ' + e.toString());
            return 0;
        }
    },
};

/**
 * Hilfsfunktion zum Importieren von Dateien aus einem Ordner
 * @param {Folder} folder - Der Ordner
 * @param {Sheet} sheet - Das Ziel-Sheet
 * @param {string} type - Der Dateityp
 * @param {Sheet} history - Das Historie-Sheet
 * @param {Set} existingFiles - Set mit bereits importierten Dateien
 * @param {Object} config - Die Konfiguration
 * @param {UI} ui - Die UI für Fehlermeldungen
 * @returns {number} - Anzahl der importierten Dateien
 */
function importFilesFromFolder(folder, sheet, type, history, existingFiles, config, ui) {
    try {
        return dataProcessor.importFilesFromFolder(
            folder,
            sheet,
            type,
            history,
            existingFiles,
            config,
        );
    } catch (e) {
        console.error(`Fehler beim Import der ${type}:`, e);
        ui.alert(`Fehler beim Import der ${type}: ${e.toString()}`);
        return 0;
    }
}

export default ImportModule;