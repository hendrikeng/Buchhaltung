// modules/importModule/historyTracker.js
import sheetUtils from '../../utils/sheetUtils.js';

// Cache für bereits geladene Dateien
const fileHistoryCache = new Set();

/**
 * Initialisiert die Änderungshistorie, falls sie nicht existiert
 * Optimierte Version mit Batch-Operationen
 * @param {Spreadsheet} ss - Das Spreadsheet
 * @param {Object} config - Die Konfiguration
 * @returns {Sheet} - Das initialisierte Historie-Sheet
 */
function getOrCreateHistorySheet(ss, config) {
    const history = sheetUtils.getOrCreateSheet(ss, 'Änderungshistorie');

    try {
        if (history.getLastRow() === 0) {
            // Optimierung: Header-Zeile in einem Batch setzen
            const historyConfig = config.aenderungshistorie.columns;

            // Vorbereitung der Header-Zeile (effizient)
            const maxColumn = Math.max(...Object.values(historyConfig));
            const headerRow = Array(maxColumn).fill('');

            // Header basierend auf Konfiguration setzen
            headerRow[historyConfig.datum - 1] = 'Datum';
            headerRow[historyConfig.typ - 1] = 'Rechnungstyp';
            headerRow[historyConfig.dateiname - 1] = 'Dateiname';
            headerRow[historyConfig.dateilink - 1] = 'Link zur Datei';

            // Header in einem API-Call setzen und formatieren
            history.appendRow(headerRow);
            history.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#f3f3f3');
        }

        return history;
    } catch (e) {
        console.error('Fehler bei der Initialisierung des History-Sheets:', e);
        throw e;
    }
}

/**
 * Sammelt bereits importierte Dateien aus der Änderungshistorie
 * Optimierte Version mit Caching und effizienterer Datenverarbeitung
 * @param {Sheet} history - Das Änderungshistorie-Sheet
 * @returns {Set} - Set mit bereits importierten Dateinamen
 */
function collectExistingFiles(history) {
    // Wenn Cache bereits gefüllt ist, diesen verwenden
    if (fileHistoryCache.size > 0) {
        return fileHistoryCache;
    }

    try {
        const historyData = history.getDataRange().getValues();

        // Optimierung: Spaltenkonfiguration dynamisch bestimmen
        // Suche nach dem Header "Dateiname" in der ersten Zeile
        const fileNameIndex = historyData[0].findIndex(header =>
            header.toString().toLowerCase() === 'dateiname') || 2;

        // Überschriftenzeile überspringen und alle Dateinamen effizient sammeln
        for (let i = 1; i < historyData.length; i++) {
            const fileName = historyData[i][fileNameIndex];
            if (fileName) fileHistoryCache.add(fileName);
        }
    } catch (e) {
        console.error('Fehler beim Sammeln bereits importierter Dateien:', e);
    }

    return fileHistoryCache;
}

/**
 * Leert den Cache für die Dateihistorie
 */
function clearFileHistoryCache() {
    fileHistoryCache.clear();
}

export default {
    getOrCreateHistorySheet,
    collectExistingFiles,
    clearFileHistoryCache,
};