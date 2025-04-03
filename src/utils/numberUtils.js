// utils/numberUtils.js
/**
 * Funktionen für die Verarbeitung und Formatierung von Zahlen und Währungen
 */

// Cache for number operations
const _numberCache = new Map();

/**
 * Konvertiert einen String oder eine Zahl in einen numerischen Währungswert
 * @param {number|string} value - Der zu parsende Wert
 * @returns {number} - Der geparste Währungswert oder 0 bei ungültigem Format
 */
function parseCurrency(value) {
    if (value === null || value === undefined || value === '') return 0;
    if (typeof value === 'number') return value;

    // Use cache for string values
    const cacheKey = `currency_${value}`;
    if (_numberCache.has(cacheKey)) {
        return _numberCache.get(cacheKey);
    }

    // Handle German formatted number (e.g., "29.335,49 €")
    const str = value.toString().trim().replace(/[€$£¥]/g, '');
    let result;

    // Check if this is likely a German-formatted number (with period as thousand separator)
    if (str.includes(',') && (str.includes('.') || /\d{4,}/.test(str))) {
        // German format: remove all periods, replace comma with period
        const normalized = str.replace(/\./g, '').replace(',', '.');
        result = parseFloat(normalized);
    } else if (str.includes(',') && !str.includes('.')) {
        // Format with comma as decimal separator only
        const normalized = str.replace(',', '.');
        result = parseFloat(normalized);
    } else {
        // Standard or US format
        const cleaned = str.replace(/[^\d,.-]/g, '')
            .replace(/,/g, '.'); // Replace all commas with dots
        result = parseFloat(cleaned);
    }

    result = isNaN(result) ? 0 : result;

    // Cache the result
    _numberCache.set(cacheKey, result);
    return result;
}

/**
 * Parst einen MwSt-Satz und normalisiert ihn
 * @param {number|string} value - Der zu parsende MwSt-Satz
 * @param {number} defaultMwst - Standard-MwSt-Satz, falls value ungültig ist
 * @returns {number} - Der normalisierte MwSt-Satz (0-100)
 */
function parseMwstRate(value, defaultMwst = 19) {
    if (value === null || value === undefined || value === '') {
        return defaultMwst;
    }

    // Use cache for common values
    const cacheKey = `mwst_${value}_${defaultMwst}`;
    if (_numberCache.has(cacheKey)) {
        return _numberCache.get(cacheKey);
    }

    let result;

    if (typeof value === 'number') {
        // Numbers < 1 are assumed to be decimal (e.g. 0.19)
        result = value < 1 ? value * 100 : value;
    } else {
        // Parse and clean the string value
        const rateStr = value.toString()
            .replace(/%/g, '')
            .replace(/,/g, '.')
            .trim();

        const rate = parseFloat(rateStr);

        if (isNaN(rate)) {
            result = defaultMwst;
        } else {
            // Normalize: values < 1 are interpreted as decimal (e.g. 0.19 -> 19)
            result = rate < 1 ? rate * 100 : rate;
        }
    }

    // Cache the result
    _numberCache.set(cacheKey, result);
    return result;
}

/**
 * Formatiert einen Währungsbetrag im deutschen Format
 * @param {number|string} amount - Der zu formatierende Betrag
 * @param {string} currency - Das Währungssymbol (Standard: "€")
 * @returns {string} - Der formatierte Betrag
 */
function formatCurrency(amount, currency = '€') {
    const value = parseCurrency(amount);

    // Use cache for formatted currencies
    const cacheKey = `format_${value}_${currency}`;
    if (_numberCache.has(cacheKey)) {
        return _numberCache.get(cacheKey);
    }

    const formatted = value.toLocaleString('de-DE', {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2,
    }) + ' ' + currency;

    // Cache the result
    _numberCache.set(cacheKey, formatted);
    return formatted;
}

/**
 * Prüft, ob zwei Zahlenwerte im Rahmen einer bestimmten Toleranz gleich sind
 * @param {number} a - Erster Wert
 * @param {number} b - Zweiter Wert
 * @param {number} tolerance - Toleranzwert (Standard: 0.01)
 * @returns {boolean} - true wenn Werte innerhalb der Toleranz gleich sind
 */
function isApproximatelyEqual(a, b, tolerance = 0.01) {
    // Pre-check for exact equality to avoid unnecessary calculations
    if (a === b) return true;
    return Math.abs(a - b) <= tolerance;
}

/**
 * Sicheres Runden eines Werts auf n Dezimalstellen
 * @param {number} value - Der zu rundende Wert
 * @param {number} decimals - Anzahl der Dezimalstellen (Standard: 2)
 * @returns {number} - Gerundeter Wert
 */
function round(value, decimals = 2) {
    const factor = 10 ** decimals;
    return Math.round((value + Number.EPSILON) * factor) / factor;
}

/**
 * Prüft, ob ein Wert keine gültige Zahl ist
 * @param {*} v - Der zu prüfende Wert
 * @returns {boolean} - True, wenn der Wert keine gültige Zahl ist
 */
function isInvalidNumber(v) {
    if (v === null || v === undefined || v === '') return true;
    return isNaN(parseFloat(v.toString().trim()));
}

export default {
    parseCurrency,
    parseMwstRate,
    formatCurrency,
    isApproximatelyEqual,
    round,
    isInvalidNumber,
};