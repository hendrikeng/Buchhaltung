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
     * @param {Object} config - Die Konfiguration
     * @returns {number} - Anzahl der importierten Dateien
     */
    importDriveFiles(config) {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const ui = SpreadsheetApp.getUi();
        let totalImported = 0;

        try {
            // Hauptsheets für Einnahmen, Ausgaben und Eigenbelege abrufen
            const revenueMain = ss.getSheetByName("Einnahmen");
            const expenseMain = ss.getSheetByName("Ausgaben");
            const receiptsMain = ss.getSheetByName("Eigenbelege");

            if (!revenueMain || !expenseMain || !receiptsMain) {
                ui.alert("Fehler: Die Sheets 'Einnahmen', 'Ausgaben' oder 'Eigenbelege' existieren nicht!");
                return 0;
            }

            // Änderungshistorie abrufen oder erstellen
            const history = historyTracker.getOrCreateHistorySheet(ss, config);

            // Auf übergeordneten Ordner zugreifen
            const parentFolder = fileManager.getParentFolder();
            if (!parentFolder) {
                ui.alert("Fehler: Kein übergeordneter Ordner gefunden.");
                return 0;
            }

            // Unterordner für Einnahmen, Ausgaben und Eigenbelege finden oder erstellen
            const revenueFolder = fileManager.findOrCreateFolder(parentFolder, "Einnahmen", ui);
            const expenseFolder = fileManager.findOrCreateFolder(parentFolder, "Ausgaben", ui);
            const receiptsFolder = fileManager.findOrCreateFolder(parentFolder, "Eigenbelege", ui);

            // Bereits importierte Dateien sammeln
            const existingFiles = historyTracker.collectExistingFiles(history);

            // Import durchführen wenn Ordner existieren
            let importedRevenue = 0, importedExpense = 0, importedReceipts = 0;

            if (revenueFolder) {
                try {
                    importedRevenue = dataProcessor.importFilesFromFolder(
                        revenueFolder,
                        revenueMain,
                        "Einnahme",
                        history,
                        existingFiles,
                        config
                    );
                    totalImported += importedRevenue;
                } catch (e) {
                    console.error("Fehler beim Import der Einnahmen:", e);
                    ui.alert("Fehler beim Import der Einnahmen: " + e.toString());
                }
            }

            if (expenseFolder) {
                try {
                    importedExpense = dataProcessor.importFilesFromFolder(
                        expenseFolder,
                        expenseMain,
                        "Ausgabe",
                        history,
                        existingFiles,
                        config
                    );
                    totalImported += importedExpense;
                } catch (e) {
                    console.error("Fehler beim Import der Ausgaben:", e);
                    ui.alert("Fehler beim Import der Ausgaben: " + e.toString());
                }
            }

            if (receiptsFolder) {
                try {
                    importedReceipts = dataProcessor.importFilesFromFolder(
                        receiptsFolder,
                        receiptsMain,
                        "Eigenbeleg",
                        history,
                        existingFiles,
                        config
                    );
                    totalImported += importedReceipts;
                } catch (e) {
                    console.error("Fehler beim Import der Eigenbelege:", e);
                    ui.alert("Fehler beim Import der Eigenbelege: " + e.toString());
                }
            }

            // Abschluss-Meldung anzeigen
            if (totalImported === 0) {
                ui.alert("Es wurden keine neuen Dateien gefunden.");
            } else {
                ui.alert(
                    `Import abgeschlossen.\n\n` +
                    `${importedRevenue} Einnahmen importiert.\n` +
                    `${importedExpense} Ausgaben importiert.\n` +
                    `${importedReceipts} Eigenbelege importiert.`
                );
            }

            return totalImported;
        } catch (e) {
            console.error("Unerwarteter Fehler beim Import:", e);
            ui.alert("Ein unerwarteter Fehler ist aufgetreten: " + e.toString());
            return 0;
        }
    }
};

export default ImportModule;