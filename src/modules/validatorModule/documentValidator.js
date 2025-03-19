// modules/validatorModule/documentValidator.js
import stringUtils from '../../utils/stringUtils.js';
import numberUtils from '../../utils/numberUtils.js';
import dateUtils from '../../utils/dateUtils.js';
import bankValidator from './bankValidator.js';

/**
 * Validiert eine Zeile aus einem Dokument (Einnahmen, Ausgaben, Eigenbelege, Gesellschafterkonto oder Holding Transfers)
 * @param {Array} row - Die zu validierende Zeile
 * @param {number} rowIndex - Der Index der Zeile (für Fehlermeldungen)
 * @param {string} sheetType - Der Typ des Sheets ("einnahmen", "ausgaben", "eigenbelege", "gesellschafterkonto" oder "holdingTransfers")
 * @param {Object} config - Die Konfiguration
 * @returns {Array<string>} - Array mit Warnungen
 */
function validateDocumentRow(row, rowIndex, sheetType = 'einnahmen', config) {
    const warnings = [];
    const columns = config[sheetType].columns;

    // Optimierte Validierung mit gemeinsamer Validierungslogik
    function validateRule(check, message) {
        if (check(row)) warnings.push(`Zeile ${rowIndex}: ${message}`);
    }

    // Grundlegende Validierungsregeln für alle Dokumente
    const baseRules = [
        {check: r => stringUtils.isEmpty(r[columns.datum - 1]), message: `${sheetType === 'eigenbelege' ? 'Beleg' : 'Rechnungs'}datum fehlt.`},
        {check: r => stringUtils.isEmpty(r[columns.rechnungsnummer - 1]), message: `${sheetType === 'eigenbelege' ? 'Beleg' : 'Rechnungs'}nummer fehlt.`},
        {check: r => stringUtils.isEmpty(r[columns.kategorie - 1]), message: 'Kategorie fehlt.'},
        {check: r => numberUtils.isInvalidNumber(r[columns.nettobetrag - 1]), message: 'Nettobetrag fehlt oder ungültig.'},
        {
            check: r => {
                const mwstStr = r[columns.mwstSatz - 1] == null ? '' : r[columns.mwstSatz - 1].toString().trim();
                if (stringUtils.isEmpty(mwstStr)) return false;

                // MwSt-Satz extrahieren und normalisieren
                const mwst = numberUtils.parseMwstRate(mwstStr, config.tax.defaultMwst);
                if (isNaN(mwst)) return true;

                // Optimierung: Verwende Set für O(1) Lookup statt Array.includes()
                const allowedRates = new Set(config?.tax?.allowedMwst || [0, 7, 19]);
                return !allowedRates.has(Math.round(mwst));
            },
            message: `Ungültiger MwSt-Satz. Erlaubt sind: ${config?.tax?.allowedMwst?.join('%, ')}% oder leer.`,
        },
    ];

    // Dokument-spezifische Regeln - Optimiert mit Funktion statt Switch
    const sheetSpecificRules = [];
    if (sheetType === 'einnahmen' || sheetType === 'ausgaben') {
        sheetSpecificRules.push({
            check: r => stringUtils.isEmpty(r[columns.kunde - 1]),
            message: `${sheetType === 'einnahmen' ? 'Kunde' : 'Lieferant'} fehlt.`,
        });
    } else if (sheetType === 'eigenbelege') {
        sheetSpecificRules.push(
            {check: r => stringUtils.isEmpty(r[columns.ausgelegtVon - 1]), message: 'Ausgelegt von fehlt.'},
            {check: r => stringUtils.isEmpty(r[columns.beschreibung - 1]), message: 'Beschreibung fehlt.'},
        );
    } else if (sheetType === 'gesellschafterkonto') {
        sheetSpecificRules.push(
            {check: r => stringUtils.isEmpty(r[columns.gesellschafter - 1]), message: 'Gesellschafter fehlt.'},
            {check: r => stringUtils.isEmpty(r[columns.referenz - 1]), message: 'Referenz fehlt.'},
            {check: r => numberUtils.isInvalidNumber(r[columns.betrag - 1]), message: 'Betrag fehlt oder ungültig.'},
        );
    } else if (sheetType === 'holdingTransfers') {
        sheetSpecificRules.push(
            {check: r => stringUtils.isEmpty(r[columns.ziel - 1]), message: 'Zielgesellschaft fehlt.'},
            {check: r => stringUtils.isEmpty(r[columns.referenz - 1]), message: 'Referenz fehlt.'},
            {check: r => numberUtils.isInvalidNumber(r[columns.betrag - 1]), message: 'Betrag fehlt oder ungültig.'},
        );
    }

    // Status-abhängige Regeln
    const zahlungsstatus = row[columns.zahlungsstatus - 1] ?
        row[columns.zahlungsstatus - 1].toString().trim().toLowerCase() : '';
    const isOpen = zahlungsstatus === 'offen';

    // Angepasste Bezeichnungen je nach Dokumenttyp
    const paidStatus = sheetType === 'eigenbelege' ? 'erstattet/teilerstattet' : 'bezahlt/teilbezahlt';
    const paymentType = sheetType === 'eigenbelege' ? 'Erstattungsart' : 'Zahlungsart';
    const paymentDate = sheetType === 'eigenbelege' ? 'Erstattungsdatum' : 'Zahlungsdatum';

    // Regeln für offene/bezahlte Zahlungen
    const paymentRules = isOpen ?
        [
            {check: r => !stringUtils.isEmpty(r[columns.zahlungsart - 1]),
                message: `${paymentType} darf bei offener Zahlung nicht gesetzt sein.`},
            {check: r => !stringUtils.isEmpty(r[columns.zahlungsdatum - 1]),
                message: `${paymentDate} darf bei offener Zahlung nicht gesetzt sein.`},
        ] :
        [
            {check: r => stringUtils.isEmpty(r[columns.zahlungsart - 1]),
                message: `${paymentType} muss bei ${paidStatus} Zahlung gesetzt sein.`},
            {check: r => stringUtils.isEmpty(r[columns.zahlungsdatum - 1]),
                message: `${paymentDate} muss bei ${paidStatus} Zahlung gesetzt sein.`},
            {
                check: r => {
                    if (stringUtils.isEmpty(r[columns.zahlungsdatum - 1])) return false;
                    const paymentDate = dateUtils.parseDate(r[columns.zahlungsdatum - 1]);
                    return paymentDate ? paymentDate > new Date() : false;
                },
                message: `${paymentDate} darf nicht in der Zukunft liegen.`,
            },
            {
                check: r => {
                    if (stringUtils.isEmpty(r[columns.zahlungsdatum - 1]) ||
                        stringUtils.isEmpty(r[columns.datum - 1])) return false;

                    const paymentDate = dateUtils.parseDate(r[columns.zahlungsdatum - 1]);
                    const documentDate = dateUtils.parseDate(r[columns.datum - 1]);
                    return paymentDate && documentDate ? paymentDate < documentDate : false;
                },
                message: `${paymentDate} darf nicht vor dem ${sheetType === 'eigenbelege' ? 'Beleg' : 'Rechnungs'}datum liegen.`,
            },
        ];

    // Alle Regeln in einem durchlauf validieren
    [...baseRules, ...sheetSpecificRules, ...paymentRules].forEach(rule => {
        validateRule(rule.check, rule.message);
    });

    return warnings;
}

/**
 * Validiert alle Zeilen in einem Sheet mit optimierter Batch-Verarbeitung
 * @param {Array} data - Zeilen-Daten (ohne Header)
 * @param {string} sheetType - Typ des Sheets ('einnahmen', 'ausgaben' oder 'eigenbelege')
 * @param {Object} config - Die Konfiguration
 * @returns {Array<string>} - Array mit Warnungen
 */
function validateSheet(data, sheetType, config) {
    // Optimierung: Allocate array with estimated size
    const warnings = [];

    // Optimierung: Filter zuerst leere Zeilen heraus
    const nonEmptyRows = data.filter(row => row.some(cell => cell !== ''));

    // Validiere alle nicht-leeren Zeilen in einem Durchgang
    nonEmptyRows.forEach((row, index) => {
        const rowWarnings = validateDocumentRow(row, index + 2, sheetType, config);
        if (rowWarnings.length > 0) {
            warnings.push(...rowWarnings);
        }
    });

    return warnings;
}

/**
 * Validiert alle Sheets auf Fehler mit optimierter Fehlerbehandlung
 * @param {Sheet} revenueSheet - Das Einnahmen-Sheet
 * @param {Sheet} expenseSheet - Das Ausgaben-Sheet
 * @param {Sheet|null} bankSheet - Das Bankbewegungen-Sheet (optional)
 * @param {Sheet|null} eigenSheet - Das Eigenbelege-Sheet (optional)
 * @param {Object} config - Die Konfiguration
 * @param {Sheet|null} gesellschafterSheet - Das Gesellschafterkonto-Sheet (optional)
 * @param {Sheet|null} holdingSheet - Das Holding Transfers-Sheet (optional)
 * @returns {boolean} - True, wenn keine Fehler gefunden wurden
 */
function validateAllSheets(revenueSheet, expenseSheet, bankSheet = null, eigenSheet = null, config, gesellschafterSheet = null, holdingSheet = null) {
    if (!revenueSheet || !expenseSheet) {
        SpreadsheetApp.getUi().alert('Fehler: Benötigte Sheets nicht gefunden!');
        return false;
    }

    try {
        // Optimierung: Warnungen für alle Sheets in einem strukturierten Objekt sammeln
        const allWarnings = {
            'Einnahmen': [],
            'Ausgaben': [],
            'Eigenbelege': [],
            'Bankbewegungen': [],
            'Gesellschafterkonto': [],
            'Holding Transfers': [],
        };

        // Optimierung: Daten für alle Sheets in einem Durchgang laden
        const sheetsData = {};

        // Einnahmen validieren
        if (revenueSheet.getLastRow() > 1) {
            sheetsData.einnahmen = revenueSheet.getDataRange().getValues().slice(1);
            allWarnings['Einnahmen'] = validateSheet(sheetsData.einnahmen, 'einnahmen', config);
        }

        // Ausgaben validieren
        if (expenseSheet.getLastRow() > 1) {
            sheetsData.ausgaben = expenseSheet.getDataRange().getValues().slice(1);
            allWarnings['Ausgaben'] = validateSheet(sheetsData.ausgaben, 'ausgaben', config);
        }

        // Eigenbelege validieren
        if (eigenSheet && eigenSheet.getLastRow() > 1) {
            sheetsData.eigenbelege = eigenSheet.getDataRange().getValues().slice(1);
            allWarnings['Eigenbelege'] = validateSheet(sheetsData.eigenbelege, 'eigenbelege', config);
        }

        // Bankbewegungen validieren
        if (bankSheet) {
            allWarnings['Bankbewegungen'] = bankValidator.validateBanking(bankSheet, config);
        }

        // Gesellschafterkonto validieren
        if (gesellschafterSheet && gesellschafterSheet.getLastRow() > 1) {
            sheetsData.gesellschafterkonto = gesellschafterSheet.getDataRange().getValues().slice(1);
            allWarnings['Gesellschafterkonto'] = validateSheet(sheetsData.gesellschafterkonto, 'gesellschafterkonto', config);
        }

        // Holding Transfers validieren
        if (holdingSheet && holdingSheet.getLastRow() > 1) {
            sheetsData.holdingTransfers = holdingSheet.getDataRange().getValues().slice(1);
            allWarnings['Holding Transfers'] = validateSheet(sheetsData.holdingTransfers, 'holdingTransfers', config);
        }

        // Fehlerberichtswarnungen zusammenstellen
        const errorMessages = [];
        Object.entries(allWarnings).forEach(([sheetName, warnings]) => {
            if (warnings.length) {
                errorMessages.push(`Fehler in '${sheetName}':\n${warnings.join('\n')}`);
            }
        });

        // Fehlermeldungen anzeigen, falls vorhanden
        if (errorMessages.length) {
            const ui = SpreadsheetApp.getUi();
            const maxMsgLength = 1500; // Google Sheets Alert-Dialog hat Beschränkungen
            let alertMsg = errorMessages.join('\n\n');

            if (alertMsg.length > maxMsgLength) {
                alertMsg = alertMsg.substring(0, maxMsgLength) +
                    '\n\n... und weitere Fehler. Bitte beheben Sie die angezeigten Fehler zuerst.';
            }

            ui.alert('Validierungsfehler gefunden', alertMsg, ui.ButtonSet.OK);
            return false;
        }

        return true;
    } catch (e) {
        console.error('Fehler bei der Validierung:', e);
        SpreadsheetApp.getUi().alert('Ein Fehler ist bei der Validierung aufgetreten: ' + e.toString());
        return false;
    }
}

export default {
    validateDocumentRow,
    validateSheet,
    validateAllSheets,
};