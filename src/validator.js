import Helpers from "./helpers.js";

const Validator = (() => {
    const isEmpty = v => v == null || v.toString().trim() === "";
    const isInvalidNumber = v => isEmpty(v) || isNaN(parseFloat(v.toString().trim()));

    const validateDropdown = (sheet, row, col, numRows, numCols, list) =>
        sheet.getRange(row, col, numRows, numCols).setDataValidation(
            SpreadsheetApp.newDataValidation().requireValueInList(list, true).build()
        );

    const validateRevenueAndExpenses = (row, rowIndex) => {
        const warnings = [];
        const validateRow = (row, idx, rules) => {
            rules.forEach(({check, message}) => {
                if (check(row)) warnings.push(`Zeile ${idx}: ${message}`);
            });
        };
        const baseRules = [
            {check: r => isEmpty(r[0]), message: "Rechnungsdatum fehlt."},
            {check: r => isEmpty(r[1]), message: "Rechnungsnummer fehlt."},
            {check: r => isEmpty(r[2]), message: "Kategorie fehlt."},
            {check: r => isEmpty(r[3]), message: "Kunde fehlt."},
            {check: r => isInvalidNumber(r[4]), message: "Nettobetrag fehlt oder ungültig."},
            {
                check: r => {
                    const mwstStr = r[5] == null ? "" : r[5].toString().trim();
                    return isEmpty(mwstStr) || isNaN(parseFloat(mwstStr.replace("%", "").replace(",", ".")));
                },
                message: "Mehrwertsteuer fehlt oder ungültig."
            }
        ];
        const status = row[11] ? row[11].toString().trim().toLowerCase() : "";
        const paymentRules = status === "offen"
            ? [
                {check: r => !isEmpty(r[12]), message: 'Zahlungsart darf bei offener Zahlung nicht gesetzt sein.'},
                {check: r => !isEmpty(r[13]), message: 'Zahlungsdatum darf bei offener Zahlung nicht gesetzt sein.'}
            ]
            : [
                {
                    check: r => isEmpty(r[12]),
                    message: 'Zahlungsart muss bei bezahlter/teilbezahlter Zahlung gesetzt sein.'
                },
                {
                    check: r => isEmpty(r[13]),
                    message: 'Zahlungsdatum muss bei bezahlter/teilbezahlter Zahlung gesetzt sein.'
                },
                {
                    check: r => {
                        if (isEmpty(r[13]) || isEmpty(r[0])) return false;
                        const paymentDate = Helpers.parseDate(r[13]);
                        return paymentDate ? paymentDate > new Date() : false;
                    },
                    message: "Zahlungsdatum darf nicht in der Zukunft liegen."
                },
                {
                    check: r => {
                        if (isEmpty(r[13]) || isEmpty(r[0])) return false;
                        const paymentDate = Helpers.parseDate(r[13]);
                        const invoiceDate = Helpers.parseDate(r[0]);
                        return paymentDate && invoiceDate ? paymentDate < invoiceDate : false;
                    },
                    message: "Zahlungsdatum darf nicht vor dem Rechnungsdatum liegen."
                }
            ];
        const rules = baseRules.concat(paymentRules);
        validateRow(row, rowIndex, rules);
        return warnings;
    };

    const validateBanking = bankSheet => {
        const data = bankSheet.getDataRange().getValues();
        const warnings = [];
        const validateRow = (row, idx, rules) => {
            rules.forEach(({check, message}) => {
                if (check(row)) warnings.push(`Zeile ${idx}: ${message}`);
            });
        };
        const headerFooterRules = [
            {check: r => isEmpty(r[0]), message: "Buchungsdatum fehlt."},
            {check: r => isEmpty(r[1]), message: "Buchungstext fehlt."},
            {
                check: r => !isEmpty(r[2]) || !isNaN(parseFloat(r[2].toString().trim())),
                message: "Betrag darf nicht gesetzt sein."
            },
            {check: r => isEmpty(r[3]) || isInvalidNumber(r[3]), message: "Saldo fehlt oder ungültig."},
            {check: r => !isEmpty(r[4]), message: "Typ darf nicht gesetzt sein."},
            {check: r => !isEmpty(r[5]), message: "Kategorie darf nicht gesetzt sein."},
            {check: r => !isEmpty(r[6]), message: "Konto (Soll) darf nicht gesetzt sein."},
            {check: r => !isEmpty(r[7]), message: "Gegenkonto (Haben) darf nicht gesetzt sein."}
        ];
        const dataRowRules = [
            {check: r => isEmpty(r[0]), message: "Buchungsdatum fehlt."},
            {check: r => isEmpty(r[1]), message: "Buchungstext fehlt."},
            {check: r => isEmpty(r[2]) || isInvalidNumber(r[2]), message: "Betrag fehlt oder ungültig."},
            {check: r => isEmpty(r[3]) || isInvalidNumber(r[3]), message: "Saldo fehlt oder ungültig."},
            {check: r => isEmpty(r[4]), message: "Typ fehlt."},
            {check: r => isEmpty(r[5]), message: "Kategorie fehlt."},
            {check: r => isEmpty(r[6]), message: "Konto (Soll) fehlt."},
            {check: r => isEmpty(r[7]), message: "Gegenkonto (Haben) fehlt."}
        ];
        data.forEach((row, i) => {
            const idx = i + 1;
            if (i === 1 || i === data.length - 1) {
                validateRow(row, idx, headerFooterRules);
            } else if (i > 1 && i < data.length - 1) {
                validateRow(row, idx, dataRowRules);
            }
        });
        return warnings;
    };

    const validateAllSheets = (revenueSheet, expenseSheet, bankSheet = null) => {
        const revData = revenueSheet.getDataRange().getValues();
        const expData = expenseSheet.getDataRange().getValues();
        const revenueWarnings = revData.length > 1 ? revData.slice(1).reduce((acc, row, i) => acc.concat(validateRevenueAndExpenses(row, i + 2)), []) : [];
        const expenseWarnings = expData.length > 1 ? expData.slice(1).reduce((acc, row, i) => acc.concat(validateRevenueAndExpenses(row, i + 2)), []) : [];
        const bankWarnings = bankSheet ? validateBanking(bankSheet) : [];
        const msgArr = [];
        revenueWarnings.length && msgArr.push("Fehler in 'Einnahmen':\n" + revenueWarnings.join("\n"));
        expenseWarnings.length && msgArr.push("Fehler in 'Ausgaben':\n" + expenseWarnings.join("\n"));
        bankWarnings.length && msgArr.push("Fehler in 'Bankbewegungen':\n" + bankWarnings.join("\n"));
        if (msgArr.length) {
            SpreadsheetApp.getUi().alert(msgArr.join("\n\n"));
            return false;
        }
        return true;
    };

    return {validateDropdown, validateRevenueAndExpenses, validateBanking, validateAllSheets};
})();

export default Validator;
