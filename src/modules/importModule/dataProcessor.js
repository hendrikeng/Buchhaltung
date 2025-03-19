// modules/importModule/dataProcessor.js
import dateUtils from '../../utils/dateUtils.js';
import sheetUtils from '../../utils/sheetUtils.js';

/**
 * Importiert Dateien aus einem Ordner in die entsprechenden Sheets
 * Optimierte Version mit Batch-Operationen und effizientem Caching
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

    // Optimierung: Dateien-Batchverarbeitung
    const files = folder.getFiles();
    const newMainRows = [];
    const newHistoryRows = [];
    const timestamp = new Date();
    let importedCount = 0;

    // Konfiguration für das richtige Sheet auswählen
    const sheetTypeMap = {
        'Einnahme': config.einnahmen.columns,
        'Ausgabe': config.ausgaben.columns,
        'Eigenbeleg': config.eigenbelege.columns,
    };

    const sheetConfig = sheetTypeMap[type];
    if (!sheetConfig) {
        console.error('Unbekannter Dateityp:', type);
        return 0;
    }

    // Konfiguration für das Änderungshistorie-Sheet
    const historyConfig = config.aenderungshistorie.columns;

    // Batch-Verarbeitung mit optimierten Limits
    const batchSize = 50; // Größere Batches für bessere Performance
    let fileCount = 0;
    const fileCache = new Map(); // Cache für bereits verarbeitet Dateien

    // Optimierung: Collect all files first to avoid API streaming issues
    const allFiles = [];
    while (files.hasNext()) {
        allFiles.push(files.next());
    }

    console.log(`Found ${allFiles.length} files in folder for type ${type}`);

    // Verarbeite alle Dateien in optimierten Batches
    for (const file of allFiles) {
        const fileName = file.getName().replace(/\.[^/.]+$/, ''); // Entfernt Dateiendung

        // Prüfe, ob die Datei bereits importiert wurde
        if (existingFiles.has(fileName)) {
            continue;
        }

        try {
            // Extraktion der Rechnungsinformationen
            const invoiceName = fileName.replace(/^[^ ]* /, ''); // Entfernt Präfix vor erstem Leerzeichen
            const invoiceDate = dateUtils.extractDateFromFilename(fileName);
            const fileUrl = file.getUrl();

            // Neue Zeile für das Hauptsheet erstellen (optimiert)
            const row = Array(mainSheet.getLastColumn()).fill('');

            // Daten in die richtigen Spalten setzen (0-basiert)
            row[sheetConfig.datum - 1] = invoiceDate;                  // Datum
            row[sheetConfig.rechnungsnummer - 1] = invoiceName;        // Rechnungsnummer
            row[sheetConfig.zeitstempel - 1] = timestamp;              // Zeitstempel
            row[sheetConfig.dateiname - 1] = fileName;                 // Dateiname
            row[sheetConfig.dateilink - 1] = fileUrl;                  // Link zur Originaldatei

            newMainRows.push(row);

            // Protokolliere den Import in der Änderungshistorie
            const historyRow = Array(historySheet.getLastColumn()).fill('');

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

            // Verarbeitungslimit alle 50 Dateien - kurze Pause
            if (fileCount % batchSize === 0) {
                // Pause einfügen um API-Limits zu vermeiden
                Utilities.sleep(100);
            }
        } catch (e) {
            console.error(`Error processing file ${fileName}:`, e);
            // Continue with other files even if one fails
        }
    }

    // Optimierte Schreibvorgänge mit Helpers in einem Batch
    if (newMainRows.length > 0) {
        const startRow = mainSheet.getLastRow() + 1;
        console.log(`Writing ${newMainRows.length} rows to main sheet starting at row ${startRow}`);

        // Batch-Schreiben mit optimierter Fehlerbehandlung
        try {
            sheetUtils.batchWriteToSheet(
                mainSheet,
                newMainRows,
                startRow,
                1,
            );
        } catch (e) {
            console.error(`Error batch writing to main sheet ${type}:`, e);

            // Fallback: Einzelne Zeilen hinzufügen
            try {
                for (let i = 0; i < newMainRows.length; i++) {
                    mainSheet.appendRow(newMainRows[i]);

                    // Kurze Pause nach jedem Batch einfügen
                    if (i % batchSize === 0 && i > 0) {
                        Utilities.sleep(100);
                    }
                }
            } catch (fallbackError) {
                console.error('Fallback import error:', fallbackError);
                // Continue with history at least
            }
        }
    }

    // Historie in einem Batch schreiben
    if (newHistoryRows.length > 0) {
        const startRow = historySheet.getLastRow() + 1;
        console.log(`Writing ${newHistoryRows.length} rows to history sheet starting at row ${startRow}`);

        try {
            sheetUtils.batchWriteToSheet(
                historySheet,
                newHistoryRows,
                startRow,
                1,
            );
        } catch (e) {
            console.error('Error batch writing to history sheet:', e);

            // Fallback: Einzelne Zeilen hinzufügen
            try {
                for (let i = 0; i < newHistoryRows.length; i++) {
                    historySheet.appendRow(newHistoryRows[i]);

                    // Kurze Pause nach jedem Batch einfügen
                    if (i % batchSize === 0 && i > 0) {
                        Utilities.sleep(100);
                    }
                }
            } catch (fallbackError) {
                console.error('Fallback history error:', fallbackError);
            }
        }
    }

    return importedCount;
}

export default {
    importFilesFromFolder,
};