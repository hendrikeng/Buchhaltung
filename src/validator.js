// file: src/validator.js
import Helpers from "./helpers.js";
import config from "./config.js";

/**
 * Modul zur Validierung der Eingaben in den verschiedenen Tabellen
 */
const Validator = (() => {
    /**
     * Prüft, ob ein Wert leer ist
     * @param {*} v - Der zu prüfende Wert
     * @returns {boolean} - True, wenn der Wert leer ist
     */
    const isEmpty = v => v == null || v.toString().trim() === "";

    /**
     * Prüft, ob ein Wert keine gültige Zahl ist
     * @param {*} v - Der zu prüfende Wert
     * @returns {boolean} - True, wenn der Wert keine gültige Zahl ist
     */
    const isInvalidNumber = v => isEmpty(v) || isNaN(parseFloat(v.toString().trim()));

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

        // Grundlegende Validierungsregeln
        const baseRules = [
            {check: r => isEmpty(r[columns.datum - 1]), message: "Rechnungsdatum fehlt."},
            {check: r => isEmpty(r[columns.rechnungsnummer - 1]), message: "Rechnungsnummer fehlt."},
            {check: r => isEmpty(r[columns.kategorie - 1]), message: "Kategorie fehlt."},
            {check: r => isEmpty(r[columns.kunde - 1]), message: "Kunde/Lieferant fehlt."},
            {check: r => isInvalidNumber(r[columns.nettobetrag - 1]), message: "Nettobetrag fehlt oder ungültig."},
            {
                check: r => {
                    const mwstStr = r[columns.mwstSatz - 1] == null ? "" : r[columns.mwstSatz - 1].toString().trim();
                    if (isEmpty(mwstStr)) return false; // Wird schon durch andere Regel geprüft

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
                check: r => !isEmpty(r[columns.zahlungsart - 1]),
                message: 'Zahlungsart darf bei offener Zahlung nicht gesetzt sein.'
            },
            {
                check: r => !isEmpty(r[columns.zahlungsdatum - 1]),
                message: 'Zahlungsdatum darf bei offener Zahlung nicht gesetzt sein.'
            }
        ];

        // Regeln für bezahlte Zahlungen
        const paidPaymentRules = [
            {
                check: r => isEmpty(r[columns.zahlungsart - 1]),
                message: 'Zahlungsart muss bei bezahlter/teilbezahlter Zahlung gesetzt sein.'
            },
            {
                check: r => isEmpty(r[columns.zahlungsdatum - 1]),
                message: 'Zahlungsdatum muss bei bezahlter/teilbezahlter Zahlung gesetzt sein.'
            },
            {
                check: r => {
                    if (isEmpty(r[columns.zahlungsdatum - 1])) return false; // Wird schon durch andere Regel geprüft

                    const paymentDate = Helpers.parseDate(r[columns.zahlungsdatum - 1]);
                    return paymentDate ? paymentDate > new Date() : false;
                },
                message: "Zahlungsdatum darf nicht in der Zukunft liegen."
            },
            {
                check: r => {
                    if (isEmpty(r[columns.zahlungsdatum - 1]) || isEmpty(r[columns.datum - 1])) return false;

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
     * Validiert das Bankbewegungen-Sheet
     * @param {Sheet} bankSheet - Das zu validierende Sheet
     * @returns {Array<string>} - Array mit Warnungen
     */
    const validateBanking = bankSheet => {
        if (!bankSheet) return ["Bankbewegungen-Sheet nicht gefunden"];

        const data = bankSheet.getDataRange().getValues();
        const warnings = [];
        const columns = config.sheets.bankbewegungen.columns;

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

        // Regeln für Header- und Footer-Zeilen
        const headerFooterRules = [
            {check: r => isEmpty(r[columns.datum - 1]), message: "Buchungsdatum fehlt."},
            {check: r => isEmpty(r[columns.buchungstext - 1]), message: "Buchungstext fehlt."},
            {
                check: r => !isEmpty(r[columns.betrag - 1]) && !isNaN(parseFloat(r[columns.betrag - 1].toString().trim())),
                message: "Betrag darf nicht gesetzt sein."
            },
            {check: r => isEmpty(r[columns.saldo - 1]) || isInvalidNumber(r[columns.saldo - 1]), message: "Saldo fehlt oder ungültig."},
            {check: r => !isEmpty(r[columns.transaktionstyp - 1]), message: "Typ darf nicht gesetzt sein."},
            {check: r => !isEmpty(r[columns.kategorie - 1]), message: "Kategorie darf nicht gesetzt sein."},
            {check: r => !isEmpty(r[columns.kontoSoll - 1]), message: "Konto (Soll) darf nicht gesetzt sein."},
            {check: r => !isEmpty(r[columns.kontoHaben - 1]), message: "Gegenkonto (Haben) darf nicht gesetzt sein."}
        ];

        // Regeln für Datenzeilen
        const dataRowRules = [
            {check: r => isEmpty(r[columns.datum - 1]), message: "Buchungsdatum fehlt."},
            {check: r => isEmpty(r[columns.buchungstext - 1]), message: "Buchungstext fehlt."},
            {check: r => isEmpty(r[columns.betrag - 1]) || isInvalidNumber(r[columns.betrag - 1]), message: "Betrag fehlt oder ungültig."},
            {check: r => isEmpty(r[columns.saldo - 1]) || isInvalidNumber(r[columns.saldo - 1]), message: "Saldo fehlt oder ungültig."},
            {check: r => isEmpty(r[columns.transaktionstyp - 1]), message: "Typ fehlt."},
            {check: r => isEmpty(r[columns.kategorie - 1]), message: "Kategorie fehlt."},
            {check: r => isEmpty(r[columns.kontoSoll - 1]), message: "Konto (Soll) fehlt."},
            {check: r => isEmpty(r[columns.kontoHaben - 1]), message: "Gegenkonto (Haben) fehlt."}
        ];

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
            // Daten aus den Sheets lesen
            const revData = revenueSheet.getDataRange().getValues();
            const expData = expenseSheet.getDataRange().getValues();

            // Einnahmen validieren (Header-Zeile überspringen)
            const revenueWarnings = revData.length > 1
                ? revData.slice(1).reduce((acc, row, i) => {
                    if (row.some(cell => cell !== "")) { // Nur nicht-leere Zeilen prüfen
                        return acc.concat(validateRevenueAndExpenses(row, i + 2, "einnahmen"));
                    }
                    return acc;
                }, [])
                : [];

            // Ausgaben validieren (Header-Zeile überspringen)
            const expenseWarnings = expData.length > 1
                ? expData.slice(1).reduce((acc, row, i) => {
                    if (row.some(cell => cell !== "")) { // Nur nicht-leere Zeilen prüfen
                        return acc.concat(validateRevenueAndExpenses(row, i + 2, "ausgaben"));
                    }
                    return acc;
                }, [])
                : [];

            // Bank validieren (falls verfügbar)
            const bankWarnings = bankSheet ? validateBanking(bankSheet) : [];

            // Fehlermeldungen zusammenstellen
            const msgArr = [];
            if (revenueWarnings.length) {
                msgArr.push("Fehler in 'Einnahmen':\n" + revenueWarnings.join("\n"));
            }

            if (expenseWarnings.length) {
                msgArr.push("Fehler in 'Ausgaben':\n" + expenseWarnings.join("\n"));
            }

            if (bankWarnings.length) {
                msgArr.push("Fehler in 'Bankbewegungen':\n" + bankWarnings.join("\n"));
            }

            // Fehlermeldungen anzeigen, falls vorhanden
            if (msgArr.length) {
                const ui = SpreadsheetApp.getUi();
                // Bei vielen Fehlern ggf. einschränken, um UI-Limits zu vermeiden
                const maxMsgLength = 1500; // Google Sheets Alert-Dialog hat Beschränkungen
                let alertMsg = msgArr.join("\n\n");

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
                    isValid: !isEmpty(value),
                    message: !isEmpty(value) ? "" : "Text darf nicht leer sein."
                };

            default:
                return {
                    isValid: true,
                    message: ""
                };
        }
    };

    // Öffentliche API des Moduls
    return {
        validateDropdown,
        validateRevenueAndExpenses,
        validateBanking,
        validateAllSheets,
        validateCellValue,
        isEmpty,
        isInvalidNumber
    };
})();

export default Validator;