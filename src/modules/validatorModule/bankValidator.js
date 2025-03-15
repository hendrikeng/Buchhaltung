// modules/validatorModule/bankValidator.js
import stringUtils from '../../utils/stringUtils.js';
import numberUtils from '../../utils/numberUtils.js';

/**
 * Validiert das Bankbewegungen-Sheet
 * @param {Sheet} bankSheet - Das zu validierende Sheet
 * @param {Object} config - Die Konfiguration
 * @returns {Array<string>} - Array mit Warnungen
 */
function validateBanking(bankSheet, config) {
    if (!bankSheet) return ["Bankbewegungen-Sheet nicht gefunden"];

    const data = bankSheet.getDataRange().getValues();
    const warnings = [];
    const columns = config.bankbewegungen.columns;

    // Regeln für Header- und Footer-Zeilen
    const headerFooterRules = [
        {check: r => stringUtils.isEmpty(r[columns.datum - 1]), message: "Buchungsdatum fehlt."},
        {check: r => stringUtils.isEmpty(r[columns.buchungstext - 1]), message: "Buchungstext fehlt."},
        {
            check: r => !stringUtils.isEmpty(r[columns.betrag - 1]) && !isNaN(parseFloat(r[columns.betrag - 1].toString().trim())),
            message: "Betrag darf nicht gesetzt sein."
        },
        {check: r => stringUtils.isEmpty(r[columns.saldo - 1]) || numberUtils.isInvalidNumber(r[columns.saldo - 1]), message: "Saldo fehlt oder ungültig."},
        {check: r => !stringUtils.isEmpty(r[columns.transaktionstyp - 1]), message: "Typ darf nicht gesetzt sein."},
        {check: r => !stringUtils.isEmpty(r[columns.kategorie - 1]), message: "Kategorie darf nicht gesetzt sein."},
        {check: r => !stringUtils.isEmpty(r[columns.kontoSoll - 1]), message: "Konto (Soll) darf nicht gesetzt sein."},
        {check: r => !stringUtils.isEmpty(r[columns.kontoHaben - 1]), message: "Gegenkonto (Haben) darf nicht gesetzt sein."}
    ];

    // Regeln für Datenzeilen
    const dataRowRules = [
        {check: r => stringUtils.isEmpty(r[columns.datum - 1]), message: "Buchungsdatum fehlt."},
        {check: r => stringUtils.isEmpty(r[columns.buchungstext - 1]), message: "Buchungstext fehlt."},
        {check: r => stringUtils.isEmpty(r[columns.betrag - 1]) || numberUtils.isInvalidNumber(r[columns.betrag - 1]), message: "Betrag fehlt oder ungültig."},
        {check: r => stringUtils.isEmpty(r[columns.saldo - 1]) || numberUtils.isInvalidNumber(r[columns.saldo - 1]), message: "Saldo fehlt oder ungültig."},
        {check: r => stringUtils.isEmpty(r[columns.transaktionstyp - 1]), message: "Typ fehlt."},
        {check: r => stringUtils.isEmpty(r[columns.kategorie - 1]), message: "Kategorie fehlt."},
        {check: r => stringUtils.isEmpty(r[columns.kontoSoll - 1]), message: "Konto (Soll) fehlt."},
        {check: r => stringUtils.isEmpty(r[columns.kontoHaben - 1]), message: "Gegenkonto (Haben) fehlt."}
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
}

export default {
    validateBanking
};