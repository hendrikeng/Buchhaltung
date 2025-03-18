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

    const nameWithoutExtension = filename.replace(/\.[^/.]+$/, '');
    let dateStr = null;

    // Verschiedene Formate erkennen (vom spezifischsten zum allgemeinsten)
    // 1. Format: DD.MM.YYYY im Dateinamen (deutsches Format)
    let match = nameWithoutExtension.match(/(\d{2}[.]\d{2}[.]\d{4})/);
    if (match?.[1]) {
        const parts = match[1].split('.');
        dateStr = `${parts[2]}-${parts[1]}-${parts[0]}`; // YYYY-MM-DD für Date-Konstruktor
    } else {
        // 2. Format: RE-YYYY-MM-DD oder ähnliches mit Trennzeichen
        match = nameWithoutExtension.match(/[^0-9](\d{4}[-_./]\d{2}[-_./]\d{2})[^0-9]/);
        if (match?.[1]) {
            dateStr = match[1].replace(/[-_.]/g, '-');
        } else {
            // 3. Format: YYYY-MM-DD am Anfang oder Ende
            match = nameWithoutExtension.match(/(^|[^0-9])(\d{4}[-_./]\d{2}[-_./]\d{2})($|[^0-9])/);
            if (match?.[2]) {
                dateStr = match[2].replace(/[-_.]/g, '-');
            } else {
                // 4. Format: DD-MM-YYYY mit verschiedenen Trennzeichen
                match = nameWithoutExtension.match(/(\d{2})[-_./](\d{2})[-_./](\d{4})/);
                if (match) {
                    const [, day, month, year] = match;
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
    if (!row || !sheetType || !config) return null;

    // Get the correct column configuration for this sheet type
    const columns = config[sheetType]?.columns;
    if (!columns) return null;

    // First try to use the payment date, which is most relevant for tax calculations
    let dateToUse = null;

    if (columns.zahlungsdatum && row[columns.zahlungsdatum - 1]) {
        dateToUse = parseDate(row[columns.zahlungsdatum - 1]);
    }

    // If no payment date, fall back to invoice/document date
    if (!dateToUse && columns.datum && row[columns.datum - 1]) {
        dateToUse = parseDate(row[columns.datum - 1]);
    }

    // If we have a valid date, extract month (1-12)
    if (dateToUse && !isNaN(dateToUse.getTime())) {
        // Check if the year matches the configured year
        const targetYear = config?.tax?.year || new Date().getFullYear();

        if (dateToUse.getFullYear() === targetYear) {
            return dateToUse.getMonth() + 1; // JavaScript months are 0-based
        }
    }

    // Could not determine a valid month
    return null;
}

// src/utils/dateUtils.js
// We only need standard date parsing/formatting without timezone complexity

/**
 * Parst ein Datum aus verschiedenen Formaten
 * @param {string|Date} value - Das zu parsende Datum
 * @returns {Date|null} - Das geparste Datum oder null, wenn kein gültiges Datum
 */
function parseDate(value) {
    if (!value) return null;

    // If already a Date object, return it
    if (value instanceof Date) {
        return value;
    }

    // String-Wert parsen
    const dateStr = value.toString().trim();

    // Deutsches Format (DD.MM.YYYY)
    const germanFormat = /^(\d{1,2})\.(\d{1,2})\.(\d{4})$/;
    if (germanFormat.test(dateStr)) {
        const [, day, month, year] = dateStr.match(germanFormat);
        return new Date(parseInt(year), parseInt(month) - 1, parseInt(day));
    }

    // ISO Format (YYYY-MM-DD)
    const isoFormat = /^(\d{4})-(\d{1,2})-(\d{1,2})$/;
    if (isoFormat.test(dateStr)) {
        const [, year, month, day] = dateStr.match(isoFormat);
        return new Date(parseInt(year), parseInt(month) - 1, parseInt(day));
    }

    // Fallback
    const parsedDate = new Date(dateStr);
    return isNaN(parsedDate.getTime()) ? null : parsedDate;
}

/**
 * Formats a date in DD.MM.YYYY format
 * @param {Date|string} date - The date to format
 * @returns {string} - Formatted date string
 */
function formatDate(date) {
    if (!date) return '';

    let dateObj;
    if (typeof date === 'string') {
        // If already in DD.MM.YYYY format, return it
        if (/^\d{1,2}\.\d{1,2}\.\d{4}$/.test(date)) {
            return date;
        }
        dateObj = parseDate(date);
    } else if (date instanceof Date) {
        dateObj = date;
    } else {
        return String(date);
    }

    if (!dateObj || isNaN(dateObj.getTime())) {
        return String(date);
    }

    const day = String(dateObj.getDate()).padStart(2, '0');
    const month = String(dateObj.getMonth() + 1).padStart(2, '0');
    const year = dateObj.getFullYear();

    return `${day}.${month}.${year}`;
}

export default {
    extractDateFromFilename,
    getMonthFromRow,
    parseDate,
    formatDate,
};