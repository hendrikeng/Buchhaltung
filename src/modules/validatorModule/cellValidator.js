// modules/validatorModule/cellValidator.js
import stringUtils from '../../utils/stringUtils.js';
import numberUtils from '../../utils/numberUtils.js';
import dateUtils from '../../utils/dateUtils.js';

// Cache für häufig verwendete Validierungsergebnisse
const validationCache = new Map();

/**
 * Erstellt eine Dropdown-Validierung für einen Bereich mit optimierter Batch-Verarbeitung
 * @param {Sheet} sheet - Das Sheet, in dem validiert werden soll
 * @param {number} row - Die Start-Zeile
 * @param {number} col - Die Start-Spalte
 * @param {number} numRows - Die Anzahl der Zeilen
 * @param {number} numCols - Die Anzahl der Spalten
 * @param {Array<string>} list - Die Liste der gültigen Werte
 * @returns {Range} - Der validierte Bereich
 */
function validateDropdown(sheet, row, col, numRows, numCols, list) {
    if (!sheet || !list || !list.length) return null;

    try {
        // Optimierung: Einheitliche Validierung für größere Bereiche
        console.log(`Validating dropdown for range [${row}:${col}] with ${list.length} items`);

        // Optimierung: Unique Values und Sortierung für bessere Dropdown-Performance
        const uniqueList = [...new Set(list)].sort();

        return sheet.getRange(row, col, numRows, numCols).setDataValidation(
            SpreadsheetApp.newDataValidation()
                .requireValueInList(uniqueList, true)
                .setAllowInvalid(false)
                .build(),
        );
    } catch (e) {
        console.error('Fehler beim Erstellen der Dropdown-Validierung:', e);

        // Fallback: Versuche in kleineren Blöcken zu validieren
        try {
            const BATCH_SIZE = 100;
            for (let i = 0; i < numRows; i += BATCH_SIZE) {
                const batchRows = Math.min(BATCH_SIZE, numRows - i);
                sheet.getRange(row + i, col, batchRows, numCols).setDataValidation(
                    SpreadsheetApp.newDataValidation()
                        .requireValueInList(list, true)
                        .setAllowInvalid(false)
                        .build(),
                );

                // Kurze Pause zur Vermeidung von Quota-Limits
                if (i + BATCH_SIZE < numRows) {
                    Utilities.sleep(50);
                }
            }
            return sheet.getRange(row, col, numRows, numCols);
        } catch (fallbackError) {
            console.error('Fallback-Fehler bei Dropdown-Validierung:', fallbackError);
            return null;
        }
    }
}

/**
 * Validiert eine einzelne Zelle anhand eines festgelegten Formats mit Caching
 * @param {*} value - Der zu validierende Wert
 * @param {string} type - Der Validierungstyp (date, number, currency, mwst, text)
 * @param {Object} config - Die Konfiguration
 * @returns {Object} - Validierungsergebnis {isValid, message}
 */
function validateCellValue(value, type, config) {
    // Optimierung: Cache-Schlüssel erstellen
    const cacheKey = `${type}_${value}_${JSON.stringify(config?.tax?.allowedMwst || [])}`;

    // Cache-Lookup für häufige Validierungen
    if (validationCache.has(cacheKey)) {
        return validationCache.get(cacheKey);
    }

    let result;

    switch (type.toLowerCase()) {
    case 'date':
        const date = dateUtils.parseDate(value);
        result = {
            isValid: !!date,
            message: date ? '' : 'Ungültiges Datumsformat. Bitte verwenden Sie DD.MM.YYYY.',
        };
        break;

    case 'number':
        const num = parseFloat(value);
        result = {
            isValid: !isNaN(num),
            message: !isNaN(num) ? '' : 'Ungültige Zahl.',
        };
        break;

    case 'currency':
        const amount = numberUtils.parseCurrency(value);
        result = {
            isValid: !isNaN(amount),
            message: !isNaN(amount) ? '' : 'Ungültiger Geldbetrag.',
        };
        break;

    case 'mwst':
        const mwst = numberUtils.parseMwstRate(value, config.tax.defaultMwst);
        // Optimierung: Verwende Set für O(1) Lookup statt Array.includes
        const allowedRates = new Set(config?.tax?.allowedMwst || [0, 7, 19]);
        result = {
            isValid: allowedRates.has(Math.round(mwst)),
            message: allowedRates.has(Math.round(mwst))
                ? ''
                : `Ungültiger MwSt-Satz. Erlaubt sind: ${Array.from(allowedRates).join('%, ')}%.`,
        };
        break;

    case 'text':
        result = {
            isValid: !stringUtils.isEmpty(value),
            message: !stringUtils.isEmpty(value) ? '' : 'Text darf nicht leer sein.',
        };
        break;

    default:
        result = {
            isValid: true,
            message: '',
        };
    }

    // Speichere Ergebnis im Cache für zukünftige Validierungen
    validationCache.set(cacheKey, result);

    return result;
}

/**
 * Validiert ein komplettes Sheet und liefert einen detaillierten Fehlerbericht
 * Optimierte Version mit Batch-Verarbeitung
 * @param {Sheet} sheet - Das zu validierende Sheet
 * @param {Object} validationRules - Regeln für jede Spalte {spaltenIndex: {type, required}}
 * @param {Object} config - Die Konfiguration
 * @returns {Object} - Validierungsergebnis mit Fehlern pro Zeile/Spalte
 */
function validateSheetWithRules(sheet, validationRules, config) {
    const results = {
        isValid: true,
        errors: [],
        errorsByRow: {},
        errorsByColumn: {},
    };

    if (!sheet) {
        results.isValid = false;
        results.errors.push('Sheet nicht gefunden');
        return results;
    }

    // Optimierung: Lade alle Daten in einem API-Call
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return results; // Nur Header oder leer

    // Optimierung: Spaltenindizes vorab für schnelleren Zugriff vorbereiten
    const columnIndices = Object.keys(validationRules).map(idx => parseInt(idx, 10));
    const maxCol = Math.max(...columnIndices);

    // Optimierung: Nur relevante Spalten prüfen
    for (let rowIdx = 1; rowIdx < data.length; rowIdx++) {
        const row = data[rowIdx];

        // Prüfe jede Spalte mit definierten Regeln
        for (const colIndex of columnIndices) {
            if (colIndex >= row.length) continue;

            const rules = validationRules[colIndex];
            const cellValue = row[colIndex];

            // Prüfe Pflichtfeld
            if (rules.required && stringUtils.isEmpty(cellValue)) {
                addError(results, {
                    row: rowIdx + 1,
                    column: colIndex + 1,
                    message: 'Pflichtfeld nicht ausgefüllt',
                });
                continue;
            }

            // Prüfe Typ wenn nicht leer
            if (!stringUtils.isEmpty(cellValue) && rules.type) {
                const validation = validateCellValue(cellValue, rules.type, config);
                if (!validation.isValid) {
                    addError(results, {
                        row: rowIdx + 1,
                        column: colIndex + 1,
                        message: validation.message,
                    });
                }
            }
        }
    }

    results.isValid = results.errors.length === 0;
    return results;
}

/**
 * Fügt einen Fehler zum Validierungsergebnis hinzu
 * @param {Object} results - Das Validierungsergebnis
 * @param {Object} error - Der Fehler {row, column, message}
 */
function addError(results, error) {
    results.errors.push(error);

    // Nach Zeile gruppieren
    if (!results.errorsByRow[error.row]) {
        results.errorsByRow[error.row] = [];
    }
    results.errorsByRow[error.row].push(error);

    // Nach Spalte gruppieren
    if (!results.errorsByColumn[error.column]) {
        results.errorsByColumn[error.column] = [];
    }
    results.errorsByColumn[error.column].push(error);
}

export default {
    validateDropdown,
    validateCellValue,
    validateSheetWithRules,
};