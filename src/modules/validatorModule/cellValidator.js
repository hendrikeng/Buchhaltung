// modules/validatorModule/cellValidator.js
import stringUtils from '../../utils/stringUtils.js';
import numberUtils from '../../utils/numberUtils.js';
import dateUtils from '../../utils/dateUtils.js';

/**
 * Erstellt eine Dropdown-Validierung für einen Bereich
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
        return sheet.getRange(row, col, numRows, numCols).setDataValidation(
            SpreadsheetApp.newDataValidation()
                .requireValueInList(list, true)
                .setAllowInvalid(false)
                .build()
        );
    } catch (e) {
        console.error("Fehler beim Erstellen der Dropdown-Validierung:", e);
        return null;
    }
}

/**
 * Validiert eine einzelne Zelle anhand eines festgelegten Formats
 * @param {*} value - Der zu validierende Wert
 * @param {string} type - Der Validierungstyp (date, number, currency, mwst, text)
 * @param {Object} config - Die Konfiguration
 * @returns {Object} - Validierungsergebnis {isValid, message}
 */
function validateCellValue(value, type, config) {
    switch (type.toLowerCase()) {
        case 'date':
            const date = dateUtils.parseDate(value);
            return {
                isValid: !!date,
                message: date ? "" : "Ungültiges Datumsformat. Bitte verwenden Sie DD.MM.YYYY."
            };

        case 'number':
            const num = parseFloat(value);
            return {
                isValid: !isNaN(num),
                message: !isNaN(num) ? "" : "Ungültige Zahl."
            };

        case 'currency':
            const amount = numberUtils.parseCurrency(value);
            return {
                isValid: !isNaN(amount),
                message: !isNaN(amount) ? "" : "Ungültiger Geldbetrag."
            };

        case 'mwst':
            const mwst = numberUtils.parseMwstRate(value, config.tax.defaultMwst);
            const allowedRates = config?.tax?.allowedMwst || [0, 7, 19];
            return {
                isValid: allowedRates.includes(Math.round(mwst)),
                message: allowedRates.includes(Math.round(mwst))
                    ? ""
                    : `Ungültiger MwSt-Satz. Erlaubt sind: ${allowedRates.join('%, ')}%.`
            };

        case 'text':
            return {
                isValid: !stringUtils.isEmpty(value),
                message: !stringUtils.isEmpty(value) ? "" : "Text darf nicht leer sein."
            };

        default:
            return {
                isValid: true,
                message: ""
            };
    }
}

/**
 * Validiert ein komplettes Sheet und liefert einen detaillierten Fehlerbericht
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
        errorsByColumn: {}
    };

    if (!sheet) {
        results.isValid = false;
        results.errors.push("Sheet nicht gefunden");
        return results;
    }

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return results; // Nur Header oder leer

    // Header-Zeile überspringen
    for (let rowIdx = 1; rowIdx < data.length; rowIdx++) {
        const row = data[rowIdx];

        // Jede Spalte mit Validierungsregeln prüfen
        Object.entries(validationRules).forEach(([colIndex, rules]) => {
            const colIdx = parseInt(colIndex, 10);
            if (isNaN(colIdx) || colIdx >= row.length) return;

            const cellValue = row[colIdx];
            const { type, required } = rules;

            // Pflichtfeld-Prüfung
            if (required && stringUtils.isEmpty(cellValue)) {
                const error = {
                    row: rowIdx + 1,
                    column: colIdx + 1,
                    message: "Pflichtfeld nicht ausgefüllt"
                };
                addError(results, error);
                return;
            }

            // Wenn nicht leer, auf Typ prüfen
            if (!stringUtils.isEmpty(cellValue) && type) {
                const validation = validateCellValue(cellValue, type, config);
                if (!validation.isValid) {
                    const error = {
                        row: rowIdx + 1,
                        column: colIdx + 1,
                        message: validation.message
                    };
                    addError(results, error);
                }
            }
        });
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
    validateSheetWithRules
};