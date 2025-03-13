// file: src/validator.js
import Helpers from "./helpers.js";
import config from "./config.js";

/**
 * Modul zur Validierung der Eingaben in den verschiedenen Tabellen
 */
const Validator = (() => {
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
    const validateDropdown = (sheet, row, col, numRows, numCols, list) => {
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
    };

    /**
     * Validiert eine Einnahmen- oder Ausgaben-Zeile
     * @param {Array} row - Die zu validierende Zeile
     * @param {number} rowIndex - Der Index der Zeile (für Fehlermeldungen)
     * @param {string} sheetType - Der Typ des Sheets ("einnahmen" oder "ausgaben")
     * @returns {Array<string>} - Array mit Warnungen
     */
    const validateRevenueAndExpenses = (row, rowIndex, sheetType = "einnahmen") => {
        const warnings = [];
        const columns = config.sheets[sheetType].columns;

        /**
         * Validiert eine Zeile anhand einer Liste von Validierungsregeln
         * @param {Array} row - Die zu validierende Zeile
         * @param {number} idx - Der Index der Zeile (für Fehlermeldungen)
         * @param {Array<Object>} rules - Array mit Regeln ({check, message})
         */
        const validateRow = (row, idx, rules) => {
            rules.forEach(({check, message}) => {
                if (check(row)) warnings.push(`Zeile ${idx}: ${message}`);
            });
        };

        // Grundlegende Validierungsregeln
        const baseRules = [
            {check: r => Helpers.isEmpty(r[columns.datum - 1]), message: "Rechnungsdatum fehlt."},
            {check: r => Helpers.isEmpty(r[columns.rechnungsnummer - 1]), message: "Rechnungsnummer fehlt."},
            {check: r => Helpers.isEmpty(r[columns.kategorie - 1]), message: "Kategorie fehlt."},
            {check: r => Helpers.isEmpty(r[columns.kunde - 1]), message: "Kunde/Lieferant fehlt."},
            {check: r => isInvalidNumber(r[columns.nettobetrag - 1]), message: "Nettobetrag fehlt oder ungültig."},
            {
                check: r => {
                    const mwstStr = r[columns.mwstSatz - 1] == null ? "" : r[columns.mwstSatz - 1].toString().trim();
                    if (Helpers.isEmpty(mwstStr)) return false; // Wird schon durch andere Regel geprüft

                    // MwSt-Satz extrahieren und normalisieren
                    const mwst = Helpers.parseMwstRate(mwstStr);
                    if (isNaN(mwst)) return true;

                    // Prüfe auf erlaubte MwSt-Sätze aus der Konfiguration
                    const allowedRates = config?.tax?.allowedMwst || [0, 7, 19];
                    return !allowedRates.includes(Math.round(mwst));
                },
                message: `Ungültiger MwSt-Satz. Erlaubt sind: ${config?.tax?.allowedMwst?.join('%, ')}% oder leer.`
            }
        ];

        // Status-abhängige Regeln
        const status = row[columns.zahlungsstatus - 1] ? row[columns.zahlungsstatus - 1].toString().trim().toLowerCase() : "";

        // Regeln für offene Zahlungen
        const openPaymentRules = [
            {
                check: r => !Helpers.isEmpty(r[columns.zahlungsart - 1]),
                message: 'Zahlungsart darf bei offener Zahlung nicht gesetzt sein.'
            },
            {
                check: r => !Helpers.isEmpty(r[columns.zahlungsdatum - 1]),
                message: 'Zahlungsdatum darf bei offener Zahlung nicht gesetzt sein.'
            }
        ];

        // Regeln für bezahlte Zahlungen
        const paidPaymentRules = [
            {
                check: r => Helpers.isEmpty(r[columns.zahlungsart - 1]),
                message: 'Zahlungsart muss bei bezahlter/teilbezahlter Zahlung gesetzt sein.'
            },
            {
                check: r => Helpers.isEmpty(r[columns.zahlungsdatum - 1]),
                message: 'Zahlungsdatum muss bei bezahlter/teilbezahlter Zahlung gesetzt sein.'
            },
            {
                check: r => {
                    if (Helpers.isEmpty(r[columns.zahlungsdatum - 1])) return false; // Wird schon durch andere Regel geprüft

                    const paymentDate = Helpers.parseDate(r[columns.zahlungsdatum - 1]);
                    return paymentDate ? paymentDate > new Date() : false;
                },
                message: "Zahlungsdatum darf nicht in der Zukunft liegen."
            },
            {
                check: r => {
                    if (Helpers.isEmpty(r[columns.zahlungsdatum - 1]) || Helpers.isEmpty(r[columns.datum - 1])) return false;

                    const paymentDate = Helpers.parseDate(r[columns.zahlungsdatum - 1]);
                    const invoiceDate = Helpers.parseDate(r[columns.datum - 1]);
                    return paymentDate && invoiceDate ? paymentDate < invoiceDate : false;
                },
                message: "Zahlungsdatum darf nicht vor dem Rechnungsdatum liegen."
            }
        ];

        // Regeln basierend auf Zahlungsstatus zusammenstellen
        const paymentRules = status === "offen" ? openPaymentRules : paidPaymentRules;

        // Alle Regeln kombinieren und anwenden
        const rules = [...baseRules, ...paymentRules];
        validateRow(row, rowIndex, rules);

        return warnings;
    };

    /**
     * Prüft, ob ein Wert keine gültige Zahl ist
     * @param {*} v - Der zu prüfende Wert
     * @returns {boolean} - True, wenn der Wert keine gültige Zahl ist
     */
    const isInvalidNumber = v => Helpers.isEmpty(v) || isNaN(parseFloat(v.toString().trim()));

    /**
     * Validiert das Bankbewegungen-Sheet
     * @param {Sheet} bankSheet - Das zu validierende Sheet
     * @returns {Array<string>} - Array mit Warnungen
     */
    const validateBanking = bankSheet => {
        if (!bankSheet) return ["Bankbewegungen-Sheet nicht gefunden"];

        const data = bankSheet.getDataRange().getValues();
        const warnings = [];
        const columns = config.sheets.bankbewegungen.columns;

        // Regeln für Header- und Footer-Zeilen
        const headerFooterRules = [
            {check: r => Helpers.isEmpty(r[columns.datum - 1]), message: "Buchungsdatum fehlt."},
            {check: r => Helpers.isEmpty(r[columns.buchungstext - 1]), message: "Buchungstext fehlt."},
            {
                check: r => !Helpers.isEmpty(r[columns.betrag - 1]) && !isNaN(parseFloat(r[columns.betrag - 1].toString().trim())),
                message: "Betrag darf nicht gesetzt sein."
            },
            {check: r => Helpers.isEmpty(r[columns.saldo - 1]) || isInvalidNumber(r[columns.saldo - 1]), message: "Saldo fehlt oder ungültig."},
            {check: r => !Helpers.isEmpty(r[columns.transaktionstyp - 1]), message: "Typ darf nicht gesetzt sein."},
            {check: r => !Helpers.isEmpty(r[columns.kategorie - 1]), message: "Kategorie darf nicht gesetzt sein."},
            {check: r => !Helpers.isEmpty(r[columns.kontoSoll - 1]), message: "Konto (Soll) darf nicht gesetzt sein."},
            {check: r => !Helpers.isEmpty(r[columns.kontoHaben - 1]), message: "Gegenkonto (Haben) darf nicht gesetzt sein."}
        ];

        // Regeln für Datenzeilen
        const dataRowRules = [
            {check: r => Helpers.isEmpty(r[columns.datum - 1]), message: "Buchungsdatum fehlt."},
            {check: r => Helpers.isEmpty(r[columns.buchungstext - 1]), message: "Buchungstext fehlt."},
            {check: r => Helpers.isEmpty(r[columns.betrag - 1]) || isInvalidNumber(r[columns.betrag - 1]), message: "Betrag fehlt oder ungültig."},
            {check: r => Helpers.isEmpty(r[columns.saldo - 1]) || isInvalidNumber(r[columns.saldo - 1]), message: "Saldo fehlt oder ungültig."},
            {check: r => Helpers.isEmpty(r[columns.transaktionstyp - 1]), message: "Typ fehlt."},
            {check: r => Helpers.isEmpty(r[columns.kategorie - 1]), message: "Kategorie fehlt."},
            {check: r => Helpers.isEmpty(r[columns.kontoSoll - 1]), message: "Konto (Soll) fehlt."},
            {check: r => Helpers.isEmpty(r[columns.kontoHaben - 1]), message: "Gegenkonto (Haben) fehlt."}
        ];

        /**
         * Validiert eine Zeile anhand von Regeln
         * @param {Array} row - Die zu validierende Zeile
         * @param {number} idx - Der Index der Zeile (für Fehlermeldungen)
         * @param {Array<Object>} rules - Array mit Regeln ({check, message})
         */
        const validateRow = (row, idx, rules) => {
            rules.forEach(({check, message}) => {
                if (check(row)) warnings.push(`Zeile ${idx}: ${message}`);
            });
        };

        // Zeilen validieren
        data.forEach((row, i) => {
            const idx = i + 1;

            // Header oder Footer
            if (i === 0 || i === data.length - 1) {
                validateRow(row, idx, headerFooterRules);
            }
            // Datenzeilen
            else if (i > 0 && i < data.length - 1) {
                validateRow(row, idx, dataRowRules);
            }
        });

        return warnings;
    };

    /**
     * Validiert alle Sheets auf Fehler
     * @param {Sheet} revenueSheet - Das Einnahmen-Sheet
     * @param {Sheet} expenseSheet - Das Ausgaben-Sheet
     * @param {Sheet|null} bankSheet - Das Bankbewegungen-Sheet (optional)
     * @returns {boolean} - True, wenn keine Fehler gefunden wurden
     */
    const validateAllSheets = (revenueSheet, expenseSheet, bankSheet = null) => {
        if (!revenueSheet || !expenseSheet) {
            SpreadsheetApp.getUi().alert("Fehler: Benötigte Sheets nicht gefunden!");
            return false;
        }

        try {
            // Warnungen für alle Sheets sammeln
            const allWarnings = [];

            // Einnahmen validieren (wenn Daten vorhanden)
            if (revenueSheet.getLastRow() > 1) {
                const revenueData = revenueSheet.getDataRange().getValues().slice(1); // Header überspringen
                const revenueWarnings = validateSheet(revenueData, "einnahmen");
                if (revenueWarnings.length) {
                    allWarnings.push("Fehler in 'Einnahmen':\n" + revenueWarnings.join("\n"));
                }
            }

            // Ausgaben validieren (wenn Daten vorhanden)
            if (expenseSheet.getLastRow() > 1) {
                const expenseData = expenseSheet.getDataRange().getValues().slice(1); // Header überspringen
                const expenseWarnings = validateSheet(expenseData, "ausgaben");
                if (expenseWarnings.length) {
                    allWarnings.push("Fehler in 'Ausgaben':\n" + expenseWarnings.join("\n"));
                }
            }

            // Bankbewegungen validieren (wenn vorhanden)
            if (bankSheet) {
                const bankWarnings = validateBanking(bankSheet);
                if (bankWarnings.length) {
                    allWarnings.push("Fehler in 'Bankbewegungen':\n" + bankWarnings.join("\n"));
                }
            }

            // Fehlermeldungen anzeigen, falls vorhanden
            if (allWarnings.length) {
                const ui = SpreadsheetApp.getUi();
                // Bei vielen Fehlern ggf. einschränken, um UI-Limits zu vermeiden
                const maxMsgLength = 1500; // Google Sheets Alert-Dialog hat Beschränkungen
                let alertMsg = allWarnings.join("\n\n");

                if (alertMsg.length > maxMsgLength) {
                    alertMsg = alertMsg.substring(0, maxMsgLength) +
                        "\n\n... und weitere Fehler. Bitte beheben Sie die angezeigten Fehler zuerst.";
                }

                ui.alert("Validierungsfehler gefunden", alertMsg, ui.ButtonSet.OK);
                return false;
            }

            return true;
        } catch (e) {
            console.error("Fehler bei der Validierung:", e);
            SpreadsheetApp.getUi().alert("Ein Fehler ist bei der Validierung aufgetreten: " + e.toString());
            return false;
        }
    };

    /**
     * Validiert alle Zeilen in einem Sheet
     * @param {Array} data - Zeilen-Daten (ohne Header)
     * @param {string} sheetType - Typ des Sheets ('einnahmen' oder 'ausgaben')
     * @returns {Array<string>} - Array mit Warnungen
     */
    const validateSheet = (data, sheetType) => {
        return data.reduce((warnings, row, index) => {
            // Nur nicht-leere Zeilen prüfen
            if (row.some(cell => cell !== "")) {
                const rowWarnings = validateRevenueAndExpenses(row, index + 2, sheetType);
                warnings.push(...rowWarnings);
            }
            return warnings;
        }, []);
    };

    /**
     * Validiert eine einzelne Zelle anhand eines festgelegten Formats
     * @param {*} value - Der zu validierende Wert
     * @param {string} type - Der Validierungstyp (date, number, currency, mwst, text)
     * @returns {Object} - Validierungsergebnis {isValid, message}
     */
    const validateCellValue = (value, type) => {
        switch (type.toLowerCase()) {
            case 'date':
                const date = Helpers.parseDate(value);
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
                const amount = Helpers.parseCurrency(value);
                return {
                    isValid: !isNaN(amount),
                    message: !isNaN(amount) ? "" : "Ungültiger Geldbetrag."
                };

            case 'mwst':
                const mwst = Helpers.parseMwstRate(value);
                const allowedRates = config?.tax?.allowedMwst || [0, 7, 19];
                return {
                    isValid: allowedRates.includes(Math.round(mwst)),
                    message: allowedRates.includes(Math.round(mwst))
                        ? ""
                        : `Ungültiger MwSt-Satz. Erlaubt sind: ${allowedRates.join('%, ')}%.`
                };

            case 'text':
                return {
                    isValid: !Helpers.isEmpty(value),
                    message: !Helpers.isEmpty(value) ? "" : "Text darf nicht leer sein."
                };

            default:
                return {
                    isValid: true,
                    message: ""
                };
        }
    };

    /**
     * Validiert ein komplettes Sheet und liefert einen detaillierten Fehlerbericht
     * @param {Sheet} sheet - Das zu validierende Sheet
     * @param {Object} validationRules - Regeln für jede Spalte {spaltenIndex: {type, required}}
     * @returns {Object} - Validierungsergebnis mit Fehlern pro Zeile/Spalte
     */
    const validateSheetWithRules = (sheet, validationRules) => {
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
                if (required && Helpers.isEmpty(cellValue)) {
                    const error = {
                        row: rowIdx + 1,
                        column: colIdx + 1,
                        message: "Pflichtfeld nicht ausgefüllt"
                    };
                    addError(results, error);
                    return;
                }

                // Wenn nicht leer, auf Typ prüfen
                if (!Helpers.isEmpty(cellValue) && type) {
                    const validation = validateCellValue(cellValue, type);
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
    };

    /**
     * Fügt einen Fehler zum Validierungsergebnis hinzu
     * @param {Object} results - Das Validierungsergebnis
     * @param {Object} error - Der Fehler {row, column, message}
     */
    const addError = (results, error) => {
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
    };

    // Öffentliche API des Moduls
    return {
        validateDropdown,
        validateRevenueAndExpenses,
        validateBanking,
        validateAllSheets,
        validateCellValue,
        validateSheetWithRules,
        isEmpty: Helpers.isEmpty,
        isInvalidNumber
    };
})();

export default Validator;