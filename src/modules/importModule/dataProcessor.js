// modules/importModule/dataProcessor.js
import dateUtils from '../../utils/dateUtils.js';
import sheetUtils from '../../utils/sheetUtils.js';

/**
 * Importiert Dateien aus einem Ordner in die entsprechenden Sheets
 *
 * @param {Folder} folder - Google Drive Ordner mit den zu importierenden Dateien
 * @param {Sheet} mainSheet - Hauptsheet (Einnahmen, Ausgaben oder Eigenbelege)
 * @param {string} type - Typ der Dateien ("Einnahme", "Ausgabe" oder "Eigenbeleg")
 * @param {Sheet} historySheet - Sheet für die Änderungshistorie
 * @param {Set} existingFiles - Set mit bereits importierten Dateinamen
 * @param {Object} config - Die Konfiguration
 * @returns {number} - Anzahl der importierten Dateien
 */
function importFilesFromFolder(folder, mainSheet, type, historySheet, existingFiles, config) {
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
            const invoiceDate = dateUtils.extractDateFromFilename(fileName);
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
        sheetUtils.batchWriteToSheet(
            mainSheet,
            newMainRows,
            mainSheet.getLastRow() + 1,
            1
        );
    }

    if (newHistoryRows.length > 0) {
        sheetUtils.batchWriteToSheet(
            historySheet,
            newHistoryRows,
            historySheet.getLastRow() + 1,
            1
        );
    }

    return importedCount;
}

export default {
    importFilesFromFolder
};