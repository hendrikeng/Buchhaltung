// file: src/importModule.js
import Helpers from "./helpers.js";
import config from "./config.js";

/**
 * Modul für den Import von Dateien aus Google Drive in die Buchhaltungstabelle
 */
const ImportModule = (() => {
    /**
     * Importiert Dateien aus einem Ordner in die entsprechenden Sheets
     *
     * @param {Folder} folder - Google Drive Ordner mit den zu importierenden Dateien
     * @param {Sheet} mainSheet - Hauptsheet (Einnahmen oder Ausgaben)
     * @param {string} type - Typ der Dateien ("Einnahme" oder "Ausgabe")
     * @param {Sheet} historySheet - Sheet für die Änderungshistorie
     * @param {Set} existingFiles - Set mit bereits importierten Dateinamen
     * @returns {number} - Anzahl der importierten Dateien
     */
    const importFilesFromFolder = (folder, mainSheet, type, historySheet, existingFiles) => {
        const files = folder.getFiles();
        const newMainRows = [];
        const newHistoryRows = [];
        const timestamp = new Date();
        let importedCount = 0;

        // Konfiguration für das richtige Sheet auswählen
        const sheetConfig = type === "Einnahme"
            ? config.sheets.einnahmen.columns
            : config.sheets.ausgaben.columns;

        // Konfiguration für das Änderungshistorie-Sheet
        const historyConfig = config.sheets.aenderungshistorie.columns;

        while (files.hasNext()) {
            const file = files.next();
            const fileName = file.getName().replace(/\.[^/.]+$/, ""); // Entfernt Dateiendung
            const invoiceName = fileName.replace(/^[^ ]* /, ""); // Entfernt Präfix vor erstem Leerzeichen
            const invoiceDate = Helpers.extractDateFromFilename(fileName);
            const fileUrl = file.getUrl();

            // Prüfe, ob die Datei bereits importiert wurde
            if (!existingFiles.has(fileName)) {
                // Neue Zeile für das Hauptsheet erstellen
                const row = Array(mainSheet.getLastColumn()).fill("");

                // Daten in die richtigen Spalten setzen (0-basiert)
                row[sheetConfig.datum - 1] = invoiceDate;                  // Datum
                row[sheetConfig.rechnungsnummer - 1] = invoiceName;        // Rechnungsnummer
                row[sheetConfig.zeitstempel - 1] = timestamp;              // Zeitstempel
                row[sheetConfig.dateiname - 1] = fileName;                 // Dateiname
                row[sheetConfig.dateilink - 1] = fileUrl;                  // Link zur Originaldatei

                newMainRows.push(row);

                // Protokolliere den Import in der Änderungshistorie
                const historyRow = Array(historySheet.getLastColumn()).fill("");

                // Daten in die richtigen Historie-Spalten setzen (0-basiert)
                historyRow[historyConfig.datum - 1] = timestamp;           // Zeitstempel
                historyRow[historyConfig.typ - 1] = type;                  // Typ (Einnahme/Ausgabe)
                historyRow[historyConfig.dateiname - 1] = fileName;        // Dateiname
                historyRow[historyConfig.dateilink - 1] = fileUrl;         // Link zur Datei

                newHistoryRows.push(historyRow);

                existingFiles.add(fileName); // Zur Liste der importierten Dateien hinzufügen
                importedCount++;
            }
        }

        // Neue Zeilen in die entsprechenden Sheets schreiben
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

        return importedCount;
    };

    /**
     * Hauptfunktion zum Importieren von Dateien aus den Einnahmen- und Ausgabenordnern
     * @returns {number} Anzahl der importierten Dateien
     */
    const importDriveFiles = () => {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const ui = SpreadsheetApp.getUi();
        let totalImported = 0;

        try {
            // Hauptsheets für Einnahmen und Ausgaben abrufen
            const revenueMain = ss.getSheetByName("Einnahmen");
            const expenseMain = ss.getSheetByName("Ausgaben");

            if (!revenueMain || !expenseMain) {
                ui.alert("Fehler: Die Sheets 'Einnahmen' oder 'Ausgaben' existieren nicht!");
                return 0;
            }

            // Änderungshistorie abrufen oder erstellen
            const history = ss.getSheetByName("Änderungshistorie") || ss.insertSheet("Änderungshistorie");

            // Header-Zeile für Änderungshistorie initialisieren, falls nötig
            if (history.getLastRow() === 0) {
                const historyConfig = config.sheets.aenderungshistorie.columns;
                const headerRow = ["", "", "", ""];
                headerRow[historyConfig.datum - 1] = "Datum";
                headerRow[historyConfig.typ - 1] = "Rechnungstyp";
                headerRow[historyConfig.dateiname - 1] = "Dateiname";
                headerRow[historyConfig.dateilink - 1] = "Link zur Datei";

                history.appendRow(headerRow);
                history.getRange(1, 1, 1, 4).setFontWeight("bold");
            }

            // Bereits importierte Dateien aus der Änderungshistorie erfassen
            const historyData = history.getDataRange().getValues();
            const existingFiles = new Set();
            const historyConfig = config.sheets.aenderungshistorie.columns;

            // Überschriftenzeile überspringen und alle Dateinamen sammeln
            for (let i = 1; i < historyData.length; i++) {
                existingFiles.add(historyData[i][historyConfig.dateiname - 1]); // Dateiname aus der entsprechenden Spalte
            }

            // Auf übergeordneten Ordner zugreifen
            let parentFolder;
            try {
                const file = DriveApp.getFileById(ss.getId());
                const parents = file.getParents();
                parentFolder = parents.hasNext() ? parents.next() : null;

                if (!parentFolder) {
                    ui.alert("Fehler: Kein übergeordneter Ordner gefunden.");
                    return 0;
                }
            } catch (e) {
                ui.alert("Fehler beim Zugriff auf Google Drive: " + e.toString());
                return 0;
            }

            // Unterordner für Einnahmen und Ausgaben finden oder erstellen
            let revenueFolder, expenseFolder;
            let importedRevenue = 0, importedExpense = 0;

            try {
                revenueFolder = Helpers.getFolderByName(parentFolder, "Einnahmen");
                if (!revenueFolder) {
                    const createFolder = ui.alert(
                        "Der Ordner 'Einnahmen' existiert nicht. Soll er erstellt werden?",
                        ui.ButtonSet.YES_NO
                    );
                    if (createFolder === ui.Button.YES) {
                        revenueFolder = parentFolder.createFolder("Einnahmen");
                    }
                }
            } catch (e) {
                ui.alert("Fehler beim Zugriff auf den Einnahmen-Ordner: " + e.toString());
            }

            try {
                expenseFolder = Helpers.getFolderByName(parentFolder, "Ausgaben");
                if (!expenseFolder) {
                    const createFolder = ui.alert(
                        "Der Ordner 'Ausgaben' existiert nicht. Soll er erstellt werden?",
                        ui.ButtonSet.YES_NO
                    );
                    if (createFolder === ui.Button.YES) {
                        expenseFolder = parentFolder.createFolder("Ausgaben");
                    }
                }
            } catch (e) {
                ui.alert("Fehler beim Zugriff auf den Ausgaben-Ordner: " + e.toString());
            }

            // Import durchführen wenn Ordner existieren
            if (revenueFolder) {
                try {
                    importedRevenue = importFilesFromFolder(
                        revenueFolder,
                        revenueMain,
                        "Einnahme",
                        history,
                        existingFiles
                    );
                    totalImported += importedRevenue;
                } catch (e) {
                    console.error("Fehler beim Import der Einnahmen:", e);
                    ui.alert("Fehler beim Import der Einnahmen: " + e.toString());
                }
            }

            if (expenseFolder) {
                try {
                    importedExpense = importFilesFromFolder(
                        expenseFolder,
                        expenseMain,
                        "Ausgabe",
                        history,
                        existingFiles  // Das gleiche Set wird für beide Importe verwendet
                    );
                    totalImported += importedExpense;
                } catch (e) {
                    console.error("Fehler beim Import der Ausgaben:", e);
                    ui.alert("Fehler beim Import der Ausgaben: " + e.toString());
                }
            }

            // Abschluss-Meldung anzeigen
            if (totalImported === 0) {
                ui.alert("Es wurden keine neuen Dateien gefunden.");
            } else {
                ui.alert(
                    `Import abgeschlossen.\n\n` +
                    `${importedRevenue} Einnahmen importiert.\n` +
                    `${importedExpense} Ausgaben importiert.`
                );
            }

            return totalImported;
        } catch (e) {
            console.error("Unerwarteter Fehler beim Import:", e);
            ui.alert("Ein unerwarteter Fehler ist aufgetreten: " + e.toString());
            return 0;
        }
    };

    // Öffentliche API des Moduls
    return {
        importDriveFiles
    };
})();

export default ImportModule;