// file: src/importModule.js
import Helpers from "./helpers.js";
import config from "./config.js";

/**
 * Modul für den Import von Dateien aus Google Drive in die Buchhaltungstabelle
 */
const ImportModule = (() => {
    /**
     * Initialisiert die Änderungshistorie, falls sie nicht existiert
     * @param {Sheet} history - Das Änderungshistorie-Sheet
     * @returns {boolean} - true bei erfolgreicher Initialisierung
     */
    const initializeHistorySheet = (history) => {
        try {
            if (history.getLastRow() === 0) {
                const historyConfig = config.aenderungshistorie.columns;
                const headerRow = Array(history.getLastColumn()).fill("");

                headerRow[historyConfig.datum - 1] = "Datum";
                headerRow[historyConfig.typ - 1] = "Rechnungstyp";
                headerRow[historyConfig.dateiname - 1] = "Dateiname";
                headerRow[historyConfig.dateilink - 1] = "Link zur Datei";

                history.appendRow(headerRow);
                history.getRange(1, 1, 1, 4).setFontWeight("bold");
            }
            return true;
        } catch (e) {
            console.error("Fehler bei der Initialisierung des History-Sheets:", e);
            return false;
        }
    };

    /**
     * Sammelt bereits importierte Dateien aus der Änderungshistorie
     * @param {Sheet} history - Das Änderungshistorie-Sheet
     * @returns {Set} - Set mit bereits importierten Dateinamen
     */
    const collectExistingFiles = (history) => {
        const existingFiles = new Set();
        try {
            const historyData = history.getDataRange().getValues();
            const historyConfig = config.aenderungshistorie.columns;

            // Überschriftenzeile überspringen und alle Dateinamen sammeln
            for (let i = 1; i < historyData.length; i++) {
                const fileName = historyData[i][historyConfig.dateiname - 1];
                if (fileName) existingFiles.add(fileName);
            }
        } catch (e) {
            console.error("Fehler beim Sammeln bereits importierter Dateien:", e);
        }
        return existingFiles;
    };

    /**
     * Ruft den übergeordneten Ordner des aktuellen Spreadsheets ab
     * @returns {Folder|null} - Der übergeordnete Ordner oder null bei Fehler
     */
    const getParentFolder = () => {
        try {
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            const file = DriveApp.getFileById(ss.getId());
            const parents = file.getParents();
            return parents.hasNext() ? parents.next() : null;
        } catch (e) {
            console.error("Fehler beim Abrufen des übergeordneten Ordners:", e);
            return null;
        }
    };

    /**
     * Findet oder erstellt einen Ordner mit dem angegebenen Namen
     * @param {Folder} parentFolder - Der übergeordnete Ordner
     * @param {string} folderName - Der zu findende oder erstellende Ordnername
     * @returns {Folder|null} - Der gefundene oder erstellte Ordner oder null bei Fehler
     */
    const findOrCreateFolder = (parentFolder, folderName) => {
        if (!parentFolder) return null;

        try {
            let folder = Helpers.getFolderByName(parentFolder, folderName);

            if (!folder) {
                const ui = SpreadsheetApp.getUi();
                const createFolder = ui.alert(
                    `Der Ordner '${folderName}' existiert nicht. Soll er erstellt werden?`,
                    ui.ButtonSet.YES_NO
                );

                if (createFolder === ui.Button.YES) {
                    folder = parentFolder.createFolder(folderName);
                }
            }

            return folder;
        } catch (e) {
            console.error(`Fehler beim Finden/Erstellen des Ordners ${folderName}:`, e);
            return null;
        }
    };

    /**
     * Importiert Dateien aus einem Ordner in die entsprechenden Sheets
     *
     * @param {Folder} folder - Google Drive Ordner mit den zu importierenden Dateien
     * @param {Sheet} mainSheet - Hauptsheet (Einnahmen, Ausgaben oder Eigenbelege)
     * @param {string} type - Typ der Dateien ("Einnahme", "Ausgabe" oder "Eigenbeleg")
     * @param {Sheet} historySheet - Sheet für die Änderungshistorie
     * @param {Set} existingFiles - Set mit bereits importierten Dateinamen
     * @returns {number} - Anzahl der importierten Dateien
     */
    const importFilesFromFolder = (folder, mainSheet, type, historySheet, existingFiles) => {
        if (!folder || !mainSheet || !historySheet) return 0;

        const files = folder.getFiles();
        const newMainRows = [];
        const newHistoryRows = [];
        const timestamp = new Date();
        let importedCount = 0;

        // Konfiguration für das richtige Sheet auswählen
        const sheetTypeMap = {
            "Einnahme": config.einnahmen.columns,
            "Ausgabe": config.ausgaben.columns,
            "Eigenbeleg": config.eigenbelege.columns
        };

        const sheetConfig = sheetTypeMap[type];
        if (!sheetConfig) {
            console.error("Unbekannter Dateityp:", type);
            return 0;
        }

        // Konfiguration für das Änderungshistorie-Sheet
        const historyConfig = config.aenderungshistorie.columns;

        // Batch-Verarbeitung der Dateien
        const batchSize = 20;
        let fileCount = 0;
        let currentBatch = [];

        while (files.hasNext()) {
            const file = files.next();
            const fileName = file.getName().replace(/\.[^/.]+$/, ""); // Entfernt Dateiendung

            // Prüfe, ob die Datei bereits importiert wurde
            if (!existingFiles.has(fileName)) {
                // Extraktion der Rechnungsinformationen
                const invoiceName = fileName.replace(/^[^ ]* /, ""); // Entfernt Präfix vor erstem Leerzeichen
                const invoiceDate = Helpers.extractDateFromFilename(fileName);
                const fileUrl = file.getUrl();

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
                historyRow[historyConfig.typ - 1] = type;                  // Typ (Einnahme/Ausgabe/Eigenbeleg)
                historyRow[historyConfig.dateiname - 1] = fileName;        // Dateiname
                historyRow[historyConfig.dateilink - 1] = fileUrl;         // Link zur Datei

                newHistoryRows.push(historyRow);
                existingFiles.add(fileName); // Zur Liste der importierten Dateien hinzufügen
                importedCount++;

                // Batch-Verarbeitung zum Verbessern der Performance
                fileCount++;
                currentBatch.push(file);

                // Verarbeitungslimit erreicht oder letzte Datei
                if (fileCount % batchSize === 0 || !files.hasNext()) {
                    // Hier könnte zusätzliche Batch-Verarbeitung erfolgen
                    // z.B. Metadaten extrahieren, etc.

                    // Batch zurücksetzen
                    currentBatch = [];

                    // Kurze Pause einfügen um API-Limits zu vermeiden
                    Utilities.sleep(50);
                }
            }
        }

        // Optimierte Schreibvorgänge mit Helpers
        if (newMainRows.length > 0) {
            Helpers.batchWriteToSheet(
                mainSheet,
                newMainRows,
                mainSheet.getLastRow() + 1,
                1
            );
        }

        if (newHistoryRows.length > 0) {
            Helpers.batchWriteToSheet(
                historySheet,
                newHistoryRows,
                historySheet.getLastRow() + 1,
                1
            );
        }

        return importedCount;
    };

    /**
     * Hauptfunktion zum Importieren von Dateien aus den Einnahmen- Ausgaben- und Eigenbelegeordnern
     * @returns {number} Anzahl der importierten Dateien
     */
    const importDriveFiles = () => {
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
            const history = ss.getSheetByName("Änderungshistorie") || ss.insertSheet("Änderungshistorie");

            // Änderungshistorie initialisieren
            if (!initializeHistorySheet(history)) {
                ui.alert("Fehler: Die Änderungshistorie konnte nicht initialisiert werden!");
                return 0;
            }

            // Bereits importierte Dateien sammeln
            const existingFiles = collectExistingFiles(history);

            // Auf übergeordneten Ordner zugreifen
            const parentFolder = getParentFolder();
            if (!parentFolder) {
                ui.alert("Fehler: Kein übergeordneter Ordner gefunden.");
                return 0;
            }

            // Unterordner für Einnahmen, Ausgaben und Eigenbelege finden oder erstellen
            const revenueFolder = findOrCreateFolder(parentFolder, "Einnahmen");
            const expenseFolder = findOrCreateFolder(parentFolder, "Ausgaben");
            const receiptsFolder = findOrCreateFolder(parentFolder, "Eigenbelege");

            // Import durchführen wenn Ordner existieren
            let importedRevenue = 0, importedExpense = 0, importedReceipts = 0;

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

            if (receiptsFolder) {
                try {
                    importedReceipts = importFilesFromFolder(
                        receiptsFolder,
                        receiptsMain,
                        "Eigenbeleg",
                        history,
                        existingFiles  // Das gleiche Set wird für beide Importe verwendet
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
    };

    // Öffentliche API des Moduls
    return {
        importDriveFiles
    };
})();

export default ImportModule;