// modules/importModule/historyTracker.js
import sheetUtils from '../../utils/sheetUtils.js';

/**
 * Initialisiert die Änderungshistorie, falls sie nicht existiert
 * @param {Spreadsheet} ss - Das Spreadsheet
 * @param {Object} config - Die Konfiguration
 * @returns {Sheet} - Das initialisierte Historie-Sheet
 */
function getOrCreateHistorySheet(ss, config) {
    const history = sheetUtils.getOrCreateSheet(ss, "Änderungshistorie");

    try {
        if (history.getLastRow() === 0) {
            const historyConfig = config.aenderungshistorie.columns;
            const headerRow = Array(Math.max(...Object.values(historyConfig))).fill("");

            headerRow[historyConfig.datum - 1] = "Datum";
            headerRow[historyConfig.typ - 1] = "Rechnungstyp";
            headerRow[historyConfig.dateiname - 1] = "Dateiname";
            headerRow[historyConfig.dateilink - 1] = "Link zur Datei";

            history.appendRow(headerRow);
            history.getRange(1, 1, 1, 4).setFontWeight("bold");
        }

        return history;
    } catch (e) {
        console.error("Fehler bei der Initialisierung des History-Sheets:", e);
        throw e;
    }
}

/**
 * Sammelt bereits importierte Dateien aus der Änderungshistorie
 * @param {Sheet} history - Das Änderungshistorie-Sheet
 * @returns {Set} - Set mit bereits importierten Dateinamen
 */
function collectExistingFiles(history) {
    const existingFiles = new Set();

    try {
        const historyData = history.getDataRange().getValues();

        // Spaltenindex für Dateinamen ermitteln (sollte Spalte C sein)
        const fileNameIndex = 2;

        // Überschriftenzeile überspringen und alle Dateinamen sammeln
        for (let i = 1; i < historyData.length; i++) {
            const fileName = historyData[i][fileNameIndex];
            if (fileName) existingFiles.add(fileName);
        }
    } catch (e) {
        console.error("Fehler beim Sammeln bereits importierter Dateien:", e);
    }

    return existingFiles;
}

export default {
    getOrCreateHistorySheet,
    collectExistingFiles
};