// file: src/importModule.js
/**
 * importModule.js - Modul für den Import von Dateien
 *
 * Dieses Modul verwaltet den Import von Rechnungen und anderen Dokumenten
 * aus Google Drive in die Buchhaltungstabelle. Es ermöglicht sowohl den automatischen
 * als auch den manuellen Import und bietet eine übersichtliche Benutzeroberfläche.
 */

import HelperModule from "./helperModule.js";
import CONFIG from "./config.js";

const ImportModule = (function() {
    // Private Variablen und Funktionen

    /**
     * Importiert Dateien aus einem Ordner in ein entsprechendes Sheet
     * @param {Object} folder - Google Drive Ordner
     * @param {Object} importSheet - Sheet für importierte Dateien
     * @param {Object} mainSheet - Hauptsheet für die Daten (Einnahmen/Ausgaben)
     * @param {string} type - Typ der Dateien ('Einnahme' oder 'Ausgabe')
     * @param {Object} historySheet - Sheet für die Änderungshistorie
     * @returns {Object} Statistik über importierte Dateien
     */
    function importFilesFromFolder(folder, importSheet, mainSheet, type, historySheet) {
        // Dateien im Ordner abrufen
        const files = folder.getFiles();

        // Bestehende Dateien in den Sheets ermitteln, um Duplikate zu vermeiden
        const getExistingFiles = (sheet, colIndex) =>
            new Set(sheet.getDataRange().getValues().slice(1).map(row => row[colIndex]));

        const existingMain = getExistingFiles(mainSheet,
            type === 'Einnahme' ? CONFIG.SYSTEM.EINNAHMEN_COLS.DATEINAME : CONFIG.SYSTEM.AUSGABEN_COLS.DATEINAME);
        const existingImport = getExistingFiles(importSheet, 0);

        // Listen für neue Einträge
        const newMainRows = [];
        const newImportRows = [];
        const newHistoryRows = [];

        // Aktueller Zeitstempel
        const timestamp = new Date();

        // Alle Dateien durchgehen
        while (files.hasNext()) {
            const file = files.next();
            const fileName = file.getName();

            // Dateinamen bereinigen (wenn nötig)
            const baseName = fileName.replace(/\.[^/.]+$/, ""); // Entferne Dateierweiterung

            // Rechnungsname (ohne Präfix, falls vorhanden)
            const invoiceName = baseName.replace(/^[^ ]* /, "");

            // Rechnungsdatum aus Dateinamen extrahieren (falls möglich)
            const invoiceDate = HelperModule.extractDateFromFilename(fileName);

            // URL der Datei
            const fileUrl = file.getUrl();

            // Flag, ob die Datei importiert wurde
            let wasImported = false;

            // In Hauptsheet (Einnahmen/Ausgaben) eintragen, falls noch nicht vorhanden
            if (!existingMain.has(fileName)) {
                // Neue Zeile für das Hauptsheet erstellen
                const newRow = Array(mainSheet.getLastColumn()).fill("");

                // Spaltenindizes für das entsprechende Sheet
                const cols = type === 'Einnahme' ?
                    CONFIG.SYSTEM.EINNAHMEN_COLS :
                    CONFIG.SYSTEM.AUSGABEN_COLS;

                // Daten setzen
                newRow[cols.DATUM] = invoiceDate ? new Date(invoiceDate) : "";
                newRow[cols.RECHNUNGSNUMMER] = invoiceName;
                // Kategorie, Kunde/Lieferant, Betrag etc. müssen manuell ausgefüllt werden

                // Metadaten
                newRow[cols.LETZTE_AKTUALISIERUNG] = timestamp;
                newRow[cols.DATEINAME] = fileName;
                newRow[cols.RECHNUNG_LINK] = fileUrl;

                // Zur Liste der neuen Einträge hinzufügen
                newMainRows.push(newRow);

                // Als vorhanden markieren
                existingMain.add(fileName);

                // Als importiert markieren
                wasImported = true;
            }

            // In das entsprechende Rechnungs-Importsheet eintragen, falls noch nicht vorhanden
            if (!existingImport.has(fileName)) {
                // Neue Zeile für das Importsheet
                const importRow = [fileName, fileUrl, invoiceName];

                // Zur Liste der neuen Einträge hinzufügen
                newImportRows.push(importRow);

                // Als vorhanden markieren
                existingImport.add(fileName);

                // Als importiert markieren
                wasImported = true;
            }

            // Wenn die Datei importiert wurde, Eintrag in der Änderungshistorie erstellen
            if (wasImported) {
                newHistoryRows.push([timestamp, type, fileName, fileUrl]);
            }
        }

        // Neue Einträge in die Sheets schreiben
        if (newImportRows.length > 0) {
            importSheet.getRange(
                importSheet.getLastRow() + 1,
                1,
                newImportRows.length,
                newImportRows[0].length
            ).setValues(newImportRows);
        }

        if (newMainRows.length > 0) {
            mainSheet.getRange(
                mainSheet.getLastRow() + 1,
                1,
                newMainRows.length,
                newMainRows[0].length
            ).setValues(newMainRows);
        }

        if (newHistoryRows.length > 0) {
            historySheet.getRange(
                historySheet.getLastRow() + 1,
                1,
                newHistoryRows.length,
                newHistoryRows[0].length
            ).setValues(newHistoryRows);
        }

        return {
            importCount: newImportRows.length,
            mainCount: newMainRows.length,
            historyCount: newHistoryRows.length
        };
    }

    /**
     * Erstellt ein Sheet, falls es noch nicht existiert
     * @param {Object} ss - Spreadsheet-Objekt
     * @param {string} sheetName - Name des Sheets
     * @param {Array} headerRow - Header-Zeile (optional)
     * @returns {Object} Sheet-Objekt
     */
    function getOrCreateSheet(ss, sheetName, headerRow = null) {
        let sheet = ss.getSheetByName(sheetName);

        if (!sheet) {
            // Neues Sheet erstellen
            sheet = ss.insertSheet(sheetName);

            // Header-Zeile hinzufügen, falls vorhanden
            if (headerRow) {
                sheet.appendRow(headerRow);
            }
        }

        return sheet;
    }

    // Öffentliche API
    return {
        /**
         * Führt den ursprünglichen Importvorgang aus der alten Version durch
         * Kompatibilitätsfunktion für die vorherige API
         */
        importDriveFiles: function() {
            try {
                const ss = SpreadsheetApp.getActiveSpreadsheet();

                // Haupt-Sheets
                const revenueMain = ss.getSheetByName(CONFIG.SYSTEM.SHEET_NAMES.EINNAHMEN);
                const expenseMain = ss.getSheetByName(CONFIG.SYSTEM.SHEET_NAMES.AUSGABEN);

                // Import-Sheets erstellen oder abrufen
                const revenue = getOrCreateSheet(ss, CONFIG.SYSTEM.SHEET_NAMES.RECHNUNGEN_EINNAHMEN,
                    ["Dateiname", "Link zur Datei", "Rechnungsnummer"]);

                const expense = getOrCreateSheet(ss, CONFIG.SYSTEM.SHEET_NAMES.RECHNUNGEN_AUSGABEN,
                    ["Dateiname", "Link zur Datei", "Rechnungsnummer"]);

                const history = getOrCreateSheet(ss, CONFIG.SYSTEM.SHEET_NAMES.AENDERUNGSHISTORIE,
                    ["Datum", "Rechnungstyp", "Dateiname", "Link zur Datei"]);

                // Übergeordneten Ordner der Tabelle finden
                const file = DriveApp.getFileById(ss.getId());
                const parents = file.getParents();
                const parentFolder = parents.hasNext() ? parents.next() : null;

                if (!parentFolder) {
                    SpreadsheetApp.getUi().alert("Kein übergeordneter Ordner gefunden.");
                    return { error: "Kein übergeordneter Ordner gefunden" };
                }

                // Statistik für den Import
                const stats = {
                    einnahmen: { imported: 0, files: 0 },
                    ausgaben: { imported: 0, files: 0 }
                };

                // Einnahmen-Ordner suchen und Dateien importieren
                const revenueFolder = HelperModule.getFolderByName(parentFolder, "Einnahmen");
                if (revenueFolder) {
                    const result = importFilesFromFolder(revenueFolder, revenue, revenueMain, "Einnahme", history);
                    stats.einnahmen.imported = result.mainCount;
                    stats.einnahmen.files = result.importCount;
                } else {
                    SpreadsheetApp.getUi().alert("Fehler: 'Einnahmen'-Ordner nicht gefunden.");
                    stats.einnahmen.error = "Ordner nicht gefunden";
                }

                // Ausgaben-Ordner suchen und Dateien importieren
                const expenseFolder = HelperModule.getFolderByName(parentFolder, "Ausgaben");
                if (expenseFolder) {
                    const result = importFilesFromFolder(expenseFolder, expense, expenseMain, "Ausgabe", history);
                    stats.ausgaben.imported = result.mainCount;
                    stats.ausgaben.files = result.importCount;
                } else {
                    SpreadsheetApp.getUi().alert("Fehler: 'Ausgaben'-Ordner nicht gefunden.");
                    stats.ausgaben.error = "Ordner nicht gefunden";
                }

                // Erfolgsmeldung
                const totalImported = stats.einnahmen.imported + stats.ausgaben.imported;
                const totalFiles = stats.einnahmen.files + stats.ausgaben.files;

                if (totalImported > 0 || totalFiles > 0) {
                    HelperModule.showToast(
                        `Import abgeschlossen: ${totalImported} Dateien in Buchhaltung übernommen`,
                        'Erfolg'
                    );
                } else {
                    HelperModule.showToast(
                        'Keine neuen Dateien zum Importieren gefunden',
                        'Info'
                    );
                }

                return stats;
            } catch (e) {
                Logger.log(`Fehler beim Importieren von Dateien: ${e.message}`);
                SpreadsheetApp.getUi().alert(`Fehler beim Importieren: ${e.message}`);
                return { error: e.message };
            }
        },

        /**
         * Sucht nach neuen Dateien in den Einnahmen und Ausgaben Ordnern
         * @returns {Object} Liste der gefundenen Dateien
         */
        findNewFiles: function() {
            try {
                const ss = SpreadsheetApp.getActiveSpreadsheet();

                // Hauptsheets abrufen
                const revenueMain = ss.getSheetByName(CONFIG.SYSTEM.SHEET_NAMES.EINNAHMEN);
                const expenseMain = ss.getSheetByName(CONFIG.SYSTEM.SHEET_NAMES.AUSGABEN);

                if (!revenueMain || !expenseMain) {
                    throw new Error("Einnahmen- oder Ausgabensheet fehlt");
                }

                // Bestehende Dateien ermitteln
                const existingRevenueFiles = new Set(
                    revenueMain.getDataRange().getValues()
                        .slice(1)
                        .map(row => row[CONFIG.SYSTEM.EINNAHMEN_COLS.DATEINAME])
                        .filter(name => name)
                );

                const existingExpenseFiles = new Set(
                    expenseMain.getDataRange().getValues()
                        .slice(1)
                        .map(row => row[CONFIG.SYSTEM.AUSGABEN_COLS.DATEINAME])
                        .filter(name => name)
                );

                // Übergeordneten Ordner der Tabelle finden
                const file = DriveApp.getFileById(ss.getId());
                const parents = file.getParents();
                const parentFolder = parents.hasNext() ? parents.next() : null;

                if (!parentFolder) {
                    throw new Error("Kein übergeordneter Ordner gefunden");
                }

                // Einnahmen und Ausgaben Ordner finden
                const revenueFolder = HelperModule.getFolderByName(parentFolder, "Einnahmen");
                const expenseFolder = HelperModule.getFolderByName(parentFolder, "Ausgaben");

                if (!revenueFolder && !expenseFolder) {
                    throw new Error("Weder Einnahmen- noch Ausgabenordner gefunden");
                }

                // Listen für neue Dateien
                const newRevenueFiles = [];
                const newExpenseFiles = [];

                // Dateien im Einnahmenordner prüfen
                if (revenueFolder) {
                    const files = revenueFolder.getFiles();
                    while (files.hasNext()) {
                        const file = files.next();
                        const fileName = file.getName();

                        if (!existingRevenueFiles.has(fileName)) {
                            newRevenueFiles.push({
                                name: fileName,
                                id: file.getId(),
                                url: file.getUrl(),
                                date: file.getDateCreated()
                            });
                        }
                    }
                }

                // Dateien im Ausgabenordner prüfen
                if (expenseFolder) {
                    const files = expenseFolder.getFiles();
                    while (files.hasNext()) {
                        const file = files.next();
                        const fileName = file.getName();

                        if (!existingExpenseFiles.has(fileName)) {
                            newExpenseFiles.push({
                                name: fileName,
                                id: file.getId(),
                                url: file.getUrl(),
                                date: file.getDateCreated()
                            });
                        }
                    }
                }

                return {
                    success: true,
                    einnahmen: newRevenueFiles,
                    ausgaben: newExpenseFiles,
                    total: newRevenueFiles.length + newExpenseFiles.length
                };
            } catch (e) {
                Logger.log(`Fehler beim Suchen neuer Dateien: ${e.message}`);
                return {
                    success: false,
                    error: e.message
                };
            }
        },

        /**
         * Zeigt einen Dialog mit allen neuen Dateien an
         */
        showNewFilesDialog: function() {
            try {
                const result = this.findNewFiles();

                if (!result.success) {
                    SpreadsheetApp.getUi().alert(`Fehler: ${result.error}`);
                    return;
                }

                if (result.total === 0) {
                    SpreadsheetApp.getUi().alert("Keine neuen Dateien gefunden.");
                    return;
                }

                // HTML für den Dialog erstellen
                let html = '<html><head><style>';
                html += 'body {font-family: Arial, sans-serif; margin: 10px;}';
                html += '.file-list {margin-bottom: 20px;}';
                html += '.file-item {margin: 5px 0; padding: 5px; border: 1px solid #eee; border-radius: 5px;}';
                html += '.file-item:hover {background-color: #f5f5f5;}';
                html += '.section {margin-bottom: 20px; padding: 10px; background-color: #f9f9f9; border-radius: 5px;}';
                html += 'button {background-color: #4285F4; color: white; border: none; padding: 8px 12px; border-radius: 4px; cursor: pointer; margin-right: 10px;}';
                html += 'button:hover {background-color: #3367D6;}';
                html += '</style></head><body>';

                html += `<h2>Neue Dateien gefunden (${result.total})</h2>`;

                // Einnahmen-Dateien anzeigen
                if (result.einnahmen.length > 0) {
                    html += `<div class="section"><h3>Einnahmen (${result.einnahmen.length})</h3>`;
                    html += '<div class="file-list">';

                    result.einnahmen.forEach((file, index) => {
                        html += `<div class="file-item">`;
                        html += `<input type="checkbox" id="einnahme_${index}" checked>`;
                        html += `<label for="einnahme_${index}">${file.name}</label>`;
                        html += `<div><small>${new Date(file.date).toLocaleDateString()} | <a href="${file.url}" target="_blank">Ansehen</a></small></div>`;
                        html += `<input type="hidden" id="file_id_einnahme_${index}" value="${file.id}">`;
                        html += `</div>`;
                    });

                    html += '</div></div>';
                }

                // Ausgaben-Dateien anzeigen
                if (result.ausgaben.length > 0) {
                    html += `<div class="section"><h3>Ausgaben (${result.ausgaben.length})</h3>`;
                    html += '<div class="file-list">';

                    result.ausgaben.forEach((file, index) => {
                        html += `<div class="file-item">`;
                        html += `<input type="checkbox" id="ausgabe_${index}" checked>`;
                        html += `<label for="ausgabe_${index}">${file.name}</label>`;
                        html += `<div><small>${new Date(file.date).toLocaleDateString()} | <a href="${file.url}" target="_blank">Ansehen</a></small></div>`;
                        html += `<input type="hidden" id="file_id_ausgabe_${index}" value="${file.id}">`;
                        html += `</div>`;
                    });

                    html += '</div></div>';
                }

                // Buttons für Import-Aktionen
                html += '<div style="text-align: center; margin-top: 20px;">';
                html += '<button onclick="importSelected()">Ausgewählte importieren</button>';
                html += '<button onclick="importAll()">Alle importieren</button>';
                html += '<button onclick="cancel()">Abbrechen</button>';
                html += '</div>';

                // JavaScript für Import-Funktionen
                html += '<script>';

                // Ausgewählte Dateien importieren
                html += 'function importSelected() {';
                html += '  const selectedFiles = {einnahmen: [], ausgaben: []};';

                // Ausgewählte Einnahmen sammeln
                html += `  for (let i = 0; i < ${result.einnahmen.length}; i++) {`;
                html += '    if (document.getElementById(`einnahme_${i}`).checked) {';
                html += '      selectedFiles.einnahmen.push(document.getElementById(`file_id_einnahme_${i}`).value);';
                html += '    }';
                html += '  }';

                // Ausgewählte Ausgaben sammeln
                html += `  for (let i = 0; i < ${result.ausgaben.length}; i++) {`;
                html += '    if (document.getElementById(`ausgabe_${i}`).checked) {';
                html += '      selectedFiles.ausgaben.push(document.getElementById(`file_id_ausgabe_${i}`).value);';
                html += '    }';
                html += '  }';

                // Import durchführen
                html += '  google.script.run';
                html += '    .withSuccessHandler(closeDialogWithMessage)';
                html += '    .withFailureHandler(showError)';
                html += '    .importSelectedFiles(selectedFiles);';
                html += '}';

                // Alle Dateien importieren
                html += 'function importAll() {';
                html += '  google.script.run';
                html += '    .withSuccessHandler(closeDialogWithMessage)';
                html += '    .withFailureHandler(showError)';
                html += '    .importAllNewFiles();';
                html += '}';

                // Dialog schließen
                html += 'function cancel() {';
                html += '  google.script.host.close();';
                html += '}';

                // Erfolgsmeldung anzeigen und Dialog schließen
                html += 'function closeDialogWithMessage(result) {';
                html += '  if (result && result.success) {';
                html += '    google.script.host.close();';
                html += '  } else {';
                html += '    showError(result.error || "Unbekannter Fehler");';
                html += '  }';
                html += '}';

                // Fehlermeldung anzeigen
                html += 'function showError(error) {';
                html += '  alert(`Fehler: ${error}`);';
                html += '}';

                html += '</script></body></html>';

                // Dialog anzeigen
                const htmlOutput = HtmlService
                    .createHtmlOutput(html)
                    .setWidth(500)
                    .setHeight(450)
                    .setTitle('Neue Dateien importieren');

                SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Neue Dateien importieren');
            } catch (e) {
                Logger.log(`Fehler beim Anzeigen des Dialogs: ${e.message}`);
                SpreadsheetApp.getUi().alert(`Fehler: ${e.message}`);
            }
        },

        /**
         * Importiert ausgewählte Dateien
         * @param {Object} selectedFiles - Objekt mit den IDs der ausgewählten Dateien
         * @returns {Object} Ergebnis des Imports
         */
        importSelectedFiles: function(selectedFiles) {
            try {
                let importCount = 0;

                // Einnahmen importieren
                for (const fileId of selectedFiles.einnahmen) {
                    const result = this.importFileById(fileId, 'Einnahme');
                    if (result.success) importCount++;
                }

                // Ausgaben importieren
                for (const fileId of selectedFiles.ausgaben) {
                    const result = this.importFileById(fileId, 'Ausgabe');
                    if (result.success) importCount++;
                }

                // Erfolgsmeldung
                if (importCount > 0) {
                    HelperModule.showToast(`${importCount} Dateien erfolgreich importiert`, 'Erfolg');
                }

                return {
                    success: true,
                    count: importCount
                };
            } catch (e) {
                Logger.log(`Fehler beim Importieren ausgewählter Dateien: ${e.message}`);
                return {
                    success: false,
                    error: e.message
                };
            }
        },

        /**
         * Importiert alle neuen Dateien
         * @returns {Object} Ergebnis des Imports
         */
        importAllNewFiles: function() {
            try {
                // Alle neuen Dateien finden
                const result = this.findNewFiles();

                if (!result.success) {
                    return {
                        success: false,
                        error: result.error
                    };
                }

                if (result.total === 0) {
                    return {
                        success: true,
                        count: 0,
                        message: "Keine neuen Dateien gefunden"
                    };
                }

                let importCount = 0;

                // Alle Einnahmen importieren
                for (const file of result.einnahmen) {
                    const importResult = this.importFileById(file.id, 'Einnahme');
                    if (importResult.success) importCount++;
                }

                // Alle Ausgaben importieren
                for (const file of result.ausgaben) {
                    const importResult = this.importFileById(file.id, 'Ausgabe');
                    if (importResult.success) importCount++;
                }

                // Erfolgsmeldung
                if (importCount > 0) {
                    HelperModule.showToast(`${importCount} Dateien erfolgreich importiert`, 'Erfolg');
                }

                return {
                    success: true,
                    count: importCount
                };
            } catch (e) {
                Logger.log(`Fehler beim Importieren aller Dateien: ${e.message}`);
                return {
                    success: false,
                    error: e.message
                };
            }
        },

        /**
         * Zeigt einen Dialog zum manuellen Import einer Datei
         */
        showImportDialog: function() {
            const ui = SpreadsheetApp.getUi();

            const typeResponse = ui.alert(
                'Datei-Import',
                'Welchen Typ von Datei möchten Sie importieren?',
                ui.ButtonSet.YES_NO_CANCEL
            );

            let type;
            if (typeResponse === ui.Button.YES) {
                type = 'Einnahme';
            } else if (typeResponse === ui.Button.NO) {
                type = 'Ausgabe';
            } else {
                return; // Abgebrochen
            }

            const fileResponse = ui.prompt(
                'Datei-Import',
                'Bitte geben Sie die Google Drive Datei-ID oder URL ein:',
                ui.ButtonSet.OK_CANCEL
            );

            if (fileResponse.getSelectedButton() !== ui.Button.OK) {
                return;
            }

            const input = fileResponse.getResponseText().trim();
            let fileId;

            if (input.startsWith('https://')) {
                // URL: ID extrahieren
                const match = input.match(/\/d\/([a-zA-Z0-9_-]+)/);
                if (match && match[1]) {
                    fileId = match[1];
                } else {
                    ui.alert('Fehler', 'Die eingegebene URL konnte nicht verarbeitet werden.', ui.ButtonSet.OK);
                    return;
                }
            } else {
                // Direkt die ID
                fileId = input;
            }

            this.importFileById(fileId, type);
        },

        /**
         * Importiert eine einzelne Datei anhand ihrer ID
         * @param {string} fileId - Google Drive Datei-ID
         * @param {string} type - Typ der Datei ('Einnahme' oder 'Ausgabe')
         * @returns {Object} Ergebnis des Imports
         */
        importFileById: function(fileId, type) {
            try {
                const ss = SpreadsheetApp.getActiveSpreadsheet();

                // Datei abrufen
                let file;
                try {
                    file = DriveApp.getFileById(fileId);
                } catch (e) {
                    throw new Error(`Die Datei mit der ID ${fileId} wurde nicht gefunden.`);
                }

                // Ziel-Sheets basierend auf dem Dateityp bestimmen
                let mainSheet, importSheet;
                if (type === 'Einnahme') {
                    mainSheet = ss.getSheetByName(CONFIG.SYSTEM.SHEET_NAMES.EINNAHMEN);
                    importSheet = getOrCreateSheet(ss, CONFIG.SYSTEM.SHEET_NAMES.RECHNUNGEN_EINNAHMEN,
                        ["Dateiname", "Link zur Datei", "Rechnungsnummer"]);
                } else {
                    mainSheet = ss.getSheetByName(CONFIG.SYSTEM.SHEET_NAMES.AUSGABEN);
                    importSheet = getOrCreateSheet(ss, CONFIG.SYSTEM.SHEET_NAMES.RECHNUNGEN_AUSGABEN,
                        ["Dateiname", "Link zur Datei", "Rechnungsnummer"]);
                }

                const history = getOrCreateSheet(ss, CONFIG.SYSTEM.SHEET_NAMES.AENDERUNGSHISTORIE,
                    ["Datum", "Rechnungstyp", "Dateiname", "Link zur Datei"]);

                // Dateiinformationen
                const fileName = file.getName();
                const fileUrl = file.getUrl();
                const baseName = fileName.replace(/\.[^/.]+$/, "");
                const invoiceName = baseName.replace(/^[^ ]* /, "");
                const invoiceDate = HelperModule.extractDateFromFilename(fileName);

                // Prüfen, ob die Datei bereits importiert wurde
                const mainData = mainSheet.getDataRange().getValues();
                for (let i = 1; i < mainData.length; i++) {
                    const col = type === 'Einnahme' ?
                        CONFIG.SYSTEM.EINNAHMEN_COLS.DATEINAME :
                        CONFIG.SYSTEM.AUSGABEN_COLS.DATEINAME;

                    if (mainData[i][col] === fileName) {
                        SpreadsheetApp.getUi().alert(`Die Datei "${fileName}" wurde bereits importiert.`);
                        return { success: false, message: 'Datei bereits importiert' };
                    }
                }

                // Neue Zeile für das Hauptsheet erstellen
                const timestamp = new Date();
                const cols = type === 'Einnahme' ?
                    CONFIG.SYSTEM.EINNAHMEN_COLS :
                    CONFIG.SYSTEM.AUSGABEN_COLS;

                const newMainRow = Array(mainSheet.getLastColumn()).fill("");
                newMainRow[cols.DATUM] = invoiceDate ? new Date(invoiceDate) : "";
                newMainRow[cols.RECHNUNGSNUMMER] = invoiceName;
                newMainRow[cols.LETZTE_AKTUALISIERUNG] = timestamp;
                newMainRow[cols.DATEINAME] = fileName;
                newMainRow[cols.RECHNUNG_LINK] = fileUrl;

                // In Hauptsheet einfügen
                const mainLastRow = mainSheet.getLastRow();
                mainSheet.getRange(mainLastRow + 1, 1, 1, newMainRow.length).setValues([newMainRow]);

                // In Import-Sheet einfügen
                const importRow = [fileName, fileUrl, invoiceName];
                const importLastRow = importSheet.getLastRow();
                importSheet.getRange(importLastRow + 1, 1, 1, importRow.length).setValues([importRow]);

                // In Änderungshistorie eintragen
                const historyRow = [timestamp, type, fileName, fileUrl];
                const historyLastRow = history.getLastRow();
                history.getRange(historyLastRow + 1, 1, 1, historyRow.length).setValues([historyRow]);

                // Erfolgsmeldung
                HelperModule.showToast(`Datei "${fileName}" erfolgreich importiert`, 'Erfolg');

                // Zur neuen Zeile springen
                mainSheet.setActiveRange(mainSheet.getRange(mainLastRow + 1, 1));

                return { success: true, message: 'Datei erfolgreich importiert' };
            } catch (e) {
                Logger.log(`Fehler beim Importieren der Datei: ${e.message}`);
                SpreadsheetApp.getUi().alert(`Fehler beim Importieren: ${e.message}`);
                return { success: false, error: e.message };
            }
        }
    };
})();
