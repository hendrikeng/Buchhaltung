// utils/numberUtils.js
/**
 * Funktionen für die Verarbeitung und Formatierung von Zahlen und Währungen
 */

/**
 * Konvertiert einen String oder eine Zahl in einen numerischen Währungswert
 * @param {number|string} value - Der zu parsende Wert
 * @returns {number} - Der geparste Währungswert oder 0 bei ungültigem Format
 */
function parseCurrency(value) {
    if (value === null || value === undefined || value === '') return 0;
    if (typeof value === 'number') return value;

    // Entferne alle Zeichen außer Ziffern, Komma, Punkt und Minus
    const str = value.toString()
        .replace(/[^\d,.-]/g, '')
        .replace(/,/g, '.'); // Alle Kommas durch Punkte ersetzen

    // Bei mehreren Punkten nur den letzten als Dezimaltrenner behandeln
    const parts = str.split('.');
    let result;

    if (parts.length > 2) {
        const last = parts.pop();
        result = parseFloat(parts.join('') + '.' + last);
    } else {
        result = parseFloat(str);
    }

    return isNaN(result) ? 0 : result;
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

    let result;

    if (typeof value === 'number') {
        // Wenn der Wert < 1 ist, nehmen wir an, dass es sich um einen Dezimalwert handelt (z.B. 0.19)
        result = value < 1 ? value * 100 : value;
    } else {
        // String-Wert parsen und bereinigen
        const rateStr = value.toString()
            .replace(/%/g, '')
            .replace(/,/g, '.')
            .trim();

        const rate = parseFloat(rateStr);

        // Wenn der geparste Wert ungültig ist, Standardwert zurückgeben
        if (isNaN(rate)) {
            result = defaultMwst;
        } else {
            // Normalisieren: Werte < 1 werden als Dezimalwerte interpretiert (z.B. 0.19 -> 19)
            result = rate < 1 ? rate * 100 : rate;
        }
    }

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
    return value.toLocaleString('de-DE', {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2,
    }) + ' ' + currency;
}

/**
 * Prüft, ob zwei Zahlenwerte im Rahmen einer bestimmten Toleranz gleich sind
 * @param {number} a - Erster Wert
 * @param {number} b - Zweiter Wert
 * @param {number} tolerance - Toleranzwert (Standard: 0.01)
 * @returns {boolean} - true wenn Werte innerhalb der Toleranz gleich sind
 */
function isApproximatelyEqual(a, b, tolerance = 0.01) {
    return Math.abs(a - b) <= tolerance;
}

/**
 * Sicheres Runden eines Werts auf n Dezimalstellen
 * @param {number} value - Der zu rundende Wert
 * @param {number} decimals - Anzahl der Dezimalstellen (Standard: 2)
 * @returns {number} - Gerundeter Wert
 */
function round(value, decimals = 2) {
    const factor = Math.pow(10, decimals);
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