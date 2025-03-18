// utils/dateUtils.js
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

    // Simplified regex approach - test in order of probability
    const datePatterns = [
        // DD.MM.YYYY (German format)
        {
            regex: /(\d{2})[.](\d{2})[.](\d{4})/,
            process: (m) => new Date(parseInt(m[3]), parseInt(m[2]) - 1, parseInt(m[1])),
        },
        // YYYY-MM-DD format
        {
            regex: /(\d{4})[-_./](\d{2})[-_./](\d{2})/,
            process: (m) => new Date(parseInt(m[1]), parseInt(m[2]) - 1, parseInt(m[3])),
        },
        // DD-MM-YYYY format with various separators
        {
            regex: /(\d{2})[-_./](\d{2})[-_./](\d{4})/,
            process: (m) => new Date(parseInt(m[3]), parseInt(m[2]) - 1, parseInt(m[1])),
        },
    ];

    let result = null;
    for (const pattern of datePatterns) {
        const match = nameWithoutExtension.match(pattern.regex);
        if (match) {
            result = pattern.process(match);
            if (!isNaN(result.getTime())) {
                break;
            }
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

    // Cache key for frequent lookups
    const cacheKey = `month_${sheetType}_${JSON.stringify(row)}`;
    if (_dateCache.has(cacheKey)) {
        return _dateCache.get(cacheKey);
    }

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
    let result = null;
    if (dateToUse && !isNaN(dateToUse.getTime())) {
        // Check if the year matches the configured year
        const targetYear = config?.tax?.year || new Date().getFullYear();

        if (dateToUse.getFullYear() === targetYear) {
            result = dateToUse.getMonth() + 1; // JavaScript months are 0-based
        }
    }

    // Cache the result
    _dateCache.set(cacheKey, result);
    return result;
}

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

    // Cache for parsed dates
    const cacheKey = `parse_${value}`;
    if (_dateCache.has(cacheKey)) {
        return _dateCache.get(cacheKey);
    }

    // String-Wert parsen
    const dateStr = value.toString().trim();
    let result = null;

    // Unified regex approach to date parsing
    const datePatterns = [
        // DD.MM.YYYY (German format)
        {
            regex: /^(\d{1,2})\.(\d{1,2})\.(\d{4})$/,
            process: (m) => new Date(parseInt(m[3]), parseInt(m[2]) - 1, parseInt(m[1])),
        },
        // YYYY-MM-DD (ISO format)
        {
            regex: /^(\d{4})-(\d{1,2})-(\d{1,2})$/,
            process: (m) => new Date(parseInt(m[1]), parseInt(m[2]) - 1, parseInt(m[3])),
        },
    ];

    for (const pattern of datePatterns) {
        const match = dateStr.match(pattern.regex);
        if (match) {
            result = pattern.process(match);
            if (!isNaN(result.getTime())) {
                break;
            }
            result = null;
        }
    }

    // Fallback to standard Date constructor if patterns didn't match
    if (!result) {
        const fallbackDate = new Date(dateStr);
        if (!isNaN(fallbackDate.getTime())) {
            result = fallbackDate;
        }
    }

    // Cache the result
    _dateCache.set(cacheKey, result);
    return result;
}

/**
 * Formats a date in DD.MM.YYYY format
 * @param {Date|string} date - The date to format
 * @returns {string} - Formatted date string
 */
function formatDate(date) {
    if (!date) return '';

    // Cache for formatted dates
    const cacheKey = `format_${date}`;
    if (_dateCache.has(cacheKey)) {
        return _dateCache.get(cacheKey);
    }

    let result = '';

    if (typeof date === 'string') {
        // If already in DD.MM.YYYY format, return it
        if (/^\d{1,2}\.\d{1,2}\.\d{4}$/.test(date)) {
            result = date;
        } else {
            const dateObj = parseDate(date);
            if (dateObj && !isNaN(dateObj.getTime())) {
                const day = String(dateObj.getDate()).padStart(2, '0');
                const month = String(dateObj.getMonth() + 1).padStart(2, '0');
                const year = dateObj.getFullYear();
                result = `${day}.${month}.${year}`;
            } else {
                result = String(date);
            }
        }
    } else if (date instanceof Date) {
        if (!isNaN(date.getTime())) {
            const day = String(date.getDate()).padStart(2, '0');
            const month = String(date.getMonth() + 1).padStart(2, '0');
            const year = date.getFullYear();
            result = `${day}.${month}.${year}`;
        } else {
            result = '';
        }
    } else {
        result = String(date);
    }

    // Cache the result
    _dateCache.set(cacheKey, result);
    return result;
}

export default {
    extractDateFromFilename,
    getMonthFromRow,
    parseDate,
    formatDate,
};