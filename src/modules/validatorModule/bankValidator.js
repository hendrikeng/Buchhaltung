// modules/validatorModule/bankValidator.js
import stringUtils from '../../utils/stringUtils.js';
import numberUtils from '../../utils/numberUtils.js';

/**
 * Validiert das Bankbewegungen-Sheet mit optimierter Batch-Verarbeitung
 * @param {Sheet} bankSheet - Das zu validierende Sheet
 * @param {Object} config - Die Konfiguration
 * @returns {Array<string>} - Array mit Warnungen
 */
function validateBanking(bankSheet, config) {
    if (!bankSheet) return ['Bankbewegungen-Sheet nicht gefunden'];

    // Optimierung: Alle Daten in einem Batch laden
    const data = bankSheet.getDataRange().getValues();
    const warnings = [];
    const columns = config.bankbewegungen.columns;

    // Optimierte Regelstruktur mit Maps für schnelleren Zugriff
    const headerFooterRuleMap = new Map([
        ['datum', r => stringUtils.isEmpty(r[columns.datum - 1]), 'Buchungsdatum fehlt.'],
        ['buchungstext', r => stringUtils.isEmpty(r[columns.buchungstext - 1]), 'Buchungstext fehlt.'],
        ['betrag', r => !stringUtils.isEmpty(r[columns.betrag - 1]) &&
            !isNaN(parseFloat(r[columns.betrag - 1].toString().trim())),
        'Betrag darf nicht gesetzt sein.'],
        ['saldo', r => stringUtils.isEmpty(r[columns.saldo - 1]) ||
            numberUtils.isInvalidNumber(r[columns.saldo - 1]), 'Saldo fehlt oder ungültig.'],
        ['transaktionstyp', r => !stringUtils.isEmpty(r[columns.transaktionstyp - 1]), 'Typ darf nicht gesetzt sein.'],
        ['kategorie', r => !stringUtils.isEmpty(r[columns.kategorie - 1]), 'Kategorie darf nicht gesetzt sein.'],
        ['kontoSoll', r => !stringUtils.isEmpty(r[columns.kontoSoll - 1]), 'Konto (Soll) darf nicht gesetzt sein.'],
        ['kontoHaben', r => !stringUtils.isEmpty(r[columns.kontoHaben - 1]), 'Gegenkonto (Haben) darf nicht gesetzt sein.'],
    ]);

    const dataRowRuleMap = new Map([
        ['datum', r => stringUtils.isEmpty(r[columns.datum - 1]), 'Buchungsdatum fehlt.'],
        ['buchungstext', r => stringUtils.isEmpty(r[columns.buchungstext - 1]), 'Buchungstext fehlt.'],
        ['betrag', r => stringUtils.isEmpty(r[columns.betrag - 1]) ||
            numberUtils.isInvalidNumber(r[columns.betrag - 1]), 'Betrag fehlt oder ungültig.'],
        ['saldo', r => stringUtils.isEmpty(r[columns.saldo - 1]) ||
            numberUtils.isInvalidNumber(r[columns.saldo - 1]), 'Saldo fehlt oder ungültig.'],
        ['transaktionstyp', r => stringUtils.isEmpty(r[columns.transaktionstyp - 1]), 'Typ fehlt.'],
        ['kategorie', r => stringUtils.isEmpty(r[columns.kategorie - 1]), 'Kategorie fehlt.'],
        ['kontoSoll', r => stringUtils.isEmpty(r[columns.kontoSoll - 1]), 'Konto (Soll) fehlt.'],
        ['kontoHaben', r => stringUtils.isEmpty(r[columns.kontoHaben - 1]), 'Gegenkonto (Haben) fehlt.'],
    ]);

    /**
     * Validiert eine Zeile anhand von Regeln
     * @param {Array} row - Die zu validierende Zeile
     * @param {number} idx - Der Index der Zeile (für Fehlermeldungen)
     * @param {Map} rules - Map mit Regeln (key, checkFn, message)
     */
    const validateRow = (row, idx, rules) => {
        rules.forEach((message, checkFn, key) => {
            if (checkFn(row)) warnings.push(`Zeile ${idx}: ${message}`);
        });
    };

    // Alle Zeilen in einem Durchgang validieren
    data.forEach((row, i) => {
        const idx = i + 1;

        // Header oder Footer
        if (i === 0 || i === data.length - 1) {
            headerFooterRuleMap.forEach((message, checkFn, key) => {
                if (checkFn(row)) warnings.push(`Zeile ${idx}: ${message}`);
            });
        }
        // Datenzeilen
        else if (i > 0 && i < data.length - 1) {
            dataRowRuleMap.forEach((message, checkFn, key) => {
                if (checkFn(row)) warnings.push(`Zeile ${idx}: ${message}`);
            });
        }
    });

    return warnings;
}

export default {
    validateBanking,
};