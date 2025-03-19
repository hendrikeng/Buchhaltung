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

    // Regeln für normale Datenzeilen
    const dataRowRules = [
        {check: r => stringUtils.isEmpty(r[columns.datum - 1]), message: 'Buchungsdatum fehlt.'},
        {check: r => stringUtils.isEmpty(r[columns.buchungstext - 1]), message: 'Buchungstext fehlt.'},
        {check: r => stringUtils.isEmpty(r[columns.betrag - 1]) ||
                numberUtils.isInvalidNumber(r[columns.betrag - 1]), message: 'Betrag fehlt oder ungültig.'},
        {check: r => stringUtils.isEmpty(r[columns.saldo - 1]) ||
                numberUtils.isInvalidNumber(r[columns.saldo - 1]), message: 'Saldo fehlt oder ungültig.'},
        {check: r => stringUtils.isEmpty(r[columns.transaktionstyp - 1]), message: 'Typ fehlt.'},
        {check: r => stringUtils.isEmpty(r[columns.kategorie - 1]), message: 'Kategorie fehlt.'},
        {check: r => stringUtils.isEmpty(r[columns.kontoSoll - 1]), message: 'Konto (Soll) fehlt.'},
        {check: r => stringUtils.isEmpty(r[columns.kontoHaben - 1]), message: 'Gegenkonto (Haben) fehlt.'},
    ];

    // Regeln für Saldo-Zeilen (Anfangssaldo und Endsaldo)
    const saldoRowRules = [
        {check: r => stringUtils.isEmpty(r[columns.datum - 1]), message: 'Buchungsdatum fehlt.'},
        {check: r => stringUtils.isEmpty(r[columns.buchungstext - 1]), message: 'Buchungstext fehlt.'},
        {check: r => stringUtils.isEmpty(r[columns.saldo - 1]) ||
                numberUtils.isInvalidNumber(r[columns.saldo - 1]), message: 'Saldo fehlt oder ungültig.'},
    ];

    /**
     * Validiert eine Zeile anhand von Regeln
     * @param {Array} row - Die zu validierende Zeile
     * @param {number} idx - Der Index der Zeile (für Fehlermeldungen)
     * @param {Array} rules - Array mit Regeln
     */
    const validateRow = (row, idx, rules) => {
        rules.forEach(rule => {
            if (rule.check(row)) warnings.push(`Zeile ${idx}: ${rule.message}`);
        });
    };

    // Alle Zeilen in einem Durchgang validieren
    data.forEach((row, i) => {
        const idx = i + 1;

        // Zeile 1 ist der Header - überspringen
        if (i === 0) return;

        const buchungstext = (row[columns.buchungstext - 1] || '').toString().trim().toLowerCase();

        // Anfangssaldo oder Endsaldo
        if (buchungstext === 'anfangssaldo' || buchungstext === 'endsaldo') {
            validateRow(row, idx, saldoRowRules);
        }
        // Normale Datenzeilen
        else {
            validateRow(row, idx, dataRowRules);
        }
    });

    return warnings;
}

export default {
    validateBanking,
};