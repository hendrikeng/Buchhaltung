// utils/stringUtils.js
/**
 * Funktionen für die Verarbeitung von Strings
 */

// Cache für häufig verwendete Operationen
const _stringCache = new Map();

/**
 * Prüft, ob ein Wert leer oder undefiniert ist
 * @param {*} value - Der zu prüfende Wert
 * @returns {boolean} - true wenn der Wert leer ist
 */
function isEmpty(value) {
    return value === null || value === undefined || value.toString().trim() === '';
}

/**
 * Bereinigt einen Text von Sonderzeichen und macht ihn vergleichbar
 * @param {string} text - Der zu bereinigende Text
 * @returns {string} - Der bereinigte Text
 */
function normalizeText(text) {
    if (!text) return '';

    // Cache-Lookup für häufige Normalisierungen
    const cacheKey = `normalize_${text}`;
    if (_stringCache.has(cacheKey)) {
        return _stringCache.get(cacheKey);
    }

    // Optimierte Zeichenersetzung mit Map-Lookup statt mehreren replace-Aufrufen
    const replacements = {
        'ä': 'ae', 'ö': 'oe', 'ü': 'ue', 'ß': 'ss',
        'Ä': 'Ae', 'Ö': 'Oe', 'Ü': 'Ue',
    };

    const result = text.toString()
        .toLowerCase()
        .replace(/[äöüßÄÖÜ]/g, match => replacements[match] || match)
        .replace(/[^a-z0-9]/g, '');

    // Ergebnis cachen
    _stringCache.set(cacheKey, result);
    return result;
}

/**
 * Konvertiert einen Spaltenindex (1-basiert) in einen Spaltenbuchstaben (A, B, C, ...)
 * @param {number} columnIndex - 1-basierter Spaltenindex
 * @returns {string} - Spaltenbuchstabe(n)
 */
function getColumnLetter(columnIndex) {
    // Cache für häufig verwendete Spaltenindizes
    const cacheKey = `column_${columnIndex}`;
    if (_stringCache.has(cacheKey)) {
        return _stringCache.get(cacheKey);
    }

    // Quick return für häufige Spalten
    if (columnIndex >= 1 && columnIndex <= 26) {
        const letter = String.fromCharCode(64 + columnIndex);
        _stringCache.set(cacheKey, letter);
        return letter;
    }

    let letter = '';
    let colIndex = columnIndex;

    while (colIndex > 0) {
        const modulo = (colIndex - 1) % 26;
        letter = String.fromCharCode(65 + modulo) + letter;
        colIndex = Math.floor((colIndex - modulo) / 26);
    }

    // Ergebnis cachen
    _stringCache.set(cacheKey, letter);
    return letter;
}

/**
 * Generiert eine eindeutige ID (für Referenzzwecke)
 * @param {string} prefix - Optional ein Präfix für die ID
 * @returns {string} - Eine eindeutige ID
 */
function generateUniqueId(prefix = '') {
    const timestamp = new Date().getTime();
    const random = Math.floor(Math.random() * 10000);
    return `${prefix}${timestamp}${random}`;
}

export default {
    isEmpty,
    normalizeText,
    getColumnLetter,
    generateUniqueId,
};