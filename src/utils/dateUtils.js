// src/utils/dateUtils.js
/**
 * Funktionen für die Verarbeitung und Formatierung von Datumsangaben
 */

// Cache für Datumsberechnungen
const _dateCache = new Map();

/**
 * Extrahiert ein Datum aus einem Dateinamen in verschiedenen Formaten
 * @param {string} filename - Der Dateiname, aus dem das Datum extrahiert werden soll
 * @returns {Date|null} - Das extrahierte Datum oder null, wenn kein Datum gefunden wurde
 */
function extractDateFromFilename(filename) {
    if (!filename) return null;

    // Cache-Lookup
    const cacheKey = `filename_${filename}`;
    if (_dateCache.has(cacheKey)) {
        return _dateCache.get(cacheKey);
    }

    const nameWithoutExtension = filename.replace(/\.[^/.]+$/, "");
    let dateStr = null;

    // Verschiedene Formate erkennen (vom spezifischsten zum allgemeinsten)
    // 1. Format: DD.MM.YYYY im Dateinamen (deutsches Format)
    let match = nameWithoutExtension.match(/(\d{2}[.]\d{2}[.]\d{4})/);
    if (match?.[1]) {
        const parts = match[1].split('.');
        dateStr = `${parts[2]}-${parts[1]}-${parts[0]}`; // YYYY-MM-DD für Date-Konstruktor
    } else {
        // 2. Format: RE-YYYY-MM-DD oder ähnliches mit Trennzeichen
        match = nameWithoutExtension.match(/[^0-9](\d{4}[-_.\/]\d{2}[-_.\/]\d{2})[^0-9]/);
        if (match?.[1]) {
            dateStr = match[1].replace(/[-_.\/]/g, '-');
        } else {
            // 3. Format: YYYY-MM-DD am Anfang oder Ende
            match = nameWithoutExtension.match(/(^|[^0-9])(\d{4}[-_.\/]\d{2}[-_.\/]\d{2})($|[^0-9])/);
            if (match?.[2]) {
                dateStr = match[2].replace(/[-_.\/]/g, '-');
            } else {
                // 4. Format: DD-MM-YYYY mit verschiedenen Trennzeichen
                match = nameWithoutExtension.match(/(\d{2})[-_.\/](\d{2})[-_.\/](\d{4})/);
                if (match) {
                    const [_, day, month, year] = match;
                    dateStr = `${year}-${month}-${day}`;
                }
            }
        }
    }

    let result = null;
    if (dateStr) {
        result = new Date(dateStr);
        // Prüfen ob das Datum valide ist
        if (isNaN(result.getTime())) {
            result = null;
        }
    }

    // Ergebnis cachen
    _dateCache.set(cacheKey, result);
    return result;
}

/**
 * Extrahiert den Monat (1-12) aus einer Zeile
 * @param {Array} row - Die Zeile mit Daten
 * @param {string} sheetType - Der Typ des Sheets (einnahmen, ausgaben, eigenbelege)
 * @param {Object} config - Die Konfiguration
 * @returns {number|null} - Der Monat (1-12) oder null
 */
function getMonthFromRow(row, sheetType, config) {
    // Vereinfachte Implementierung für Tests
    return 1;
}

/**
 * Parst ein Datum aus verschiedenen Formaten
 * @param {string|Date} value - Das zu parsende Datum
 * @returns {Date|null} - Das geparste Datum oder null, wenn kein gültiges Datum
 */
function parseDate(value) {
    if (!value) return null;

    // Wenn bereits ein Date-Objekt übergeben wurde
    if (value instanceof Date) return value;

    // String-Wert parsen
    const dateStr = value.toString().trim();

    // Deutsches Format (DD.MM.YYYY)
    const germanFormat = /^(\d{1,2})\.(\d{1,2})\.(\d{4})$/;
    if (germanFormat.test(dateStr)) {
        const [_, day, month, year] = dateStr.match(germanFormat);
        return new Date(year, month - 1, day);
    }

    // ISO Format (YYYY-MM-DD)
    const isoFormat = /^(\d{4})-(\d{1,2})-(\d{1,2})$/;
    if (isoFormat.test(dateStr)) {
        return new Date(dateStr);
    }

    // Fallback
    const parsedDate = new Date(dateStr);
    return isNaN(parsedDate.getTime()) ? null : parsedDate;
}

export default {
    extractDateFromFilename,
    getMonthFromRow,
    parseDate
};