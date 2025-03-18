// modules/validatorModule/documentValidator.js
import stringUtils from '../../utils/stringUtils.js';
import numberUtils from '../../utils/numberUtils.js';
import dateUtils from '../../utils/dateUtils.js';
import bankValidator from './bankValidator.js';

/**
 * Validiert eine Zeile aus einem Dokument (Einnahmen, Ausgaben oder Eigenbelege)
 * @param {Array} row - Die zu validierende Zeile
 * @param {number} rowIndex - Der Index der Zeile (für Fehlermeldungen)
 * @param {string} sheetType - Der Typ des Sheets ("einnahmen", "ausgaben" oder "eigenbelege")
 * @param {Object} config - Die Konfiguration
 * @returns {Array<string>} - Array mit Warnungen
 */
function validateDocumentRow(row, rowIndex, sheetType = 'einnahmen', config) {
    const warnings = [];
    const columns = config[sheetType].columns;

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

    // Grundlegende Validierungsregeln für alle Dokumente
    const baseRules = [
        {check: r => stringUtils.isEmpty(r[columns.datum - 1]), message: `${sheetType === 'eigenbelege' ? 'Beleg' : 'Rechnungs'}datum fehlt.`},
        {check: r => stringUtils.isEmpty(r[columns.rechnungsnummer - 1]), message: `${sheetType === 'eigenbelege' ? 'Beleg' : 'Rechnungs'}nummer fehlt.`},
        {check: r => stringUtils.isEmpty(r[columns.kategorie - 1]), message: 'Kategorie fehlt.'},
        {check: r => numberUtils.isInvalidNumber(r[columns.nettobetrag - 1]), message: 'Nettobetrag fehlt oder ungültig.'},
        {
            check: r => {
                const mwstStr = r[columns.mwstSatz - 1] == null ? '' : r[columns.mwstSatz - 1].toString().trim();
                if (stringUtils.isEmpty(mwstStr)) return false; // Wird schon durch andere Regel geprüft

                // MwSt-Satz extrahieren und normalisieren
                const mwst = numberUtils.parseMwstRate(mwstStr, config.tax.defaultMwst);
                if (isNaN(mwst)) return true;

                // Prüfe auf erlaubte MwSt-Sätze aus der Konfiguration
                const allowedRates = config?.tax?.allowedMwst || [0, 7, 19];
                return !allowedRates.includes(Math.round(mwst));
            },
            message: `Ungültiger MwSt-Satz. Erlaubt sind: ${config?.tax?.allowedMwst?.join('%, ')}% oder leer.`,
        },
    ];

    // Dokument-spezifische Regeln
    if (sheetType === 'einnahmen' || sheetType === 'ausgaben') {
        baseRules.push({
            check: r => stringUtils.isEmpty(r[columns.kunde - 1]),
            message: `${sheetType === 'einnahmen' ? 'Kunde' : 'Lieferant'} fehlt.`,
        });
    } else if (sheetType === 'eigenbelege') {
        baseRules.push({
            check: r => stringUtils.isEmpty(r[columns.ausgelegtVon - 1]),
            message: 'Ausgelegt von fehlt.',
        });
        baseRules.push({
            check: r => stringUtils.isEmpty(r[columns.beschreibung - 1]),
            message: 'Beschreibung fehlt.',
        });
    }

    // Status-abhängige Regeln
    const zahlungsstatus = row[columns.zahlungsstatus - 1] ? row[columns.zahlungsstatus - 1].toString().trim().toLowerCase() : '';
    const isOpen = zahlungsstatus === 'offen';

    // Angepasste Bezeichnungen je nach Dokumenttyp
    const paidStatus = sheetType === 'eigenbelege' ? 'erstattet/teilerstattet' : 'bezahlt/teilbezahlt';
    const paymentType = sheetType === 'eigenbelege' ? 'Erstattungsart' : 'Zahlungsart';
    const paymentDate = sheetType === 'eigenbelege' ? 'Erstattungsdatum' : 'Zahlungsdatum';

    // Regeln für offene Zahlungen
    const openPaymentRules = [
        {
            check: r => !stringUtils.isEmpty(r[columns.zahlungsart - 1]),
            message: `${paymentType} darf bei offener Zahlung nicht gesetzt sein.`,
        },
        {
            check: r => !stringUtils.isEmpty(r[columns.zahlungsdatum - 1]),
            message: `${paymentDate} darf bei offener Zahlung nicht gesetzt sein.`,
        },
    ];

    // Regeln für bezahlte/erstattete Zahlungen
    const paidPaymentRules = [
        {
            check: r => stringUtils.isEmpty(r[columns.zahlungsart - 1]),
            message: `${paymentType} muss bei ${paidStatus} Zahlung gesetzt sein.`,
        },
        {
            check: r => stringUtils.isEmpty(r[columns.zahlungsdatum - 1]),
            message: `${paymentDate} muss bei ${paidStatus} Zahlung gesetzt sein.`,
        },
        {
            check: r => {
                if (stringUtils.isEmpty(r[columns.zahlungsdatum - 1])) return false; // Wird schon durch andere Regel geprüft

                const paymentDate = dateUtils.parseDate(r[columns.zahlungsdatum - 1]);
                return paymentDate ? paymentDate > new Date() : false;
            },
            message: `${paymentDate} darf nicht in der Zukunft liegen.`,
        },
        {
            check: r => {
                if (stringUtils.isEmpty(r[columns.zahlungsdatum - 1]) || stringUtils.isEmpty(r[columns.datum - 1])) return false;

                const paymentDate = dateUtils.parseDate(r[columns.zahlungsdatum - 1]);
                const documentDate = dateUtils.parseDate(r[columns.datum - 1]);
                return paymentDate && documentDate ? paymentDate < documentDate : false;
            },
            message: `${paymentDate} darf nicht vor dem ${sheetType === 'eigenbelege' ? 'Beleg' : 'Rechnungs'}datum liegen.`,
        },
    ];

    // Regeln basierend auf Zahlungsstatus zusammenstellen
    const paymentRules = isOpen ? openPaymentRules : paidPaymentRules;

    // Alle Regeln kombinieren und anwenden
    const rules = [...baseRules, ...paymentRules];
    validateRow(row, rowIndex, rules);

    return warnings;
}

/**
 * Validiert alle Zeilen in einem Sheet
 * @param {Array} data - Zeilen-Daten (ohne Header)
 * @param {string} sheetType - Typ des Sheets ('einnahmen', 'ausgaben' oder 'eigenbelege')
 * @param {Object} config - Die Konfiguration
 * @returns {Array<string>} - Array mit Warnungen
 */
function validateSheet(data, sheetType, config) {
    return data.reduce((warnings, row, index) => {
        // Nur nicht-leere Zeilen prüfen
        if (row.some(cell => cell !== '')) {
            const rowWarnings = validateDocumentRow(row, index + 2, sheetType, config);
            warnings.push(...rowWarnings);
        }
        return warnings;
    }, []);
}

/**
 * Validiert alle Sheets auf Fehler
 * @param {Sheet} revenueSheet - Das Einnahmen-Sheet
 * @param {Sheet} expenseSheet - Das Ausgaben-Sheet
 * @param {Sheet|null} bankSheet - Das Bankbewegungen-Sheet (optional)
 * @param {Sheet|null} eigenSheet - Das Eigenbelege-Sheet (optional)
 * @param {Object} config - Die Konfiguration
 * @returns {boolean} - True, wenn keine Fehler gefunden wurden
 */
function validateAllSheets(revenueSheet, expenseSheet, bankSheet = null, eigenSheet = null, config) {
    if (!revenueSheet || !expenseSheet) {
        SpreadsheetApp.getUi().alert('Fehler: Benötigte Sheets nicht gefunden!');
        return false;
    }

    try {
        // Warnungen für alle Sheets sammeln
        const allWarnings = [];

        // Einnahmen validieren (wenn Daten vorhanden)
        if (revenueSheet.getLastRow() > 1) {
            const revenueData = revenueSheet.getDataRange().getValues().slice(1); // Header überspringen
            const revenueWarnings = validateSheet(revenueData, 'einnahmen', config);
            if (revenueWarnings.length) {
                allWarnings.push("Fehler in 'Einnahmen':\n" + revenueWarnings.join('\n'));
            }
        }

        // Ausgaben validieren (wenn Daten vorhanden)
        if (expenseSheet.getLastRow() > 1) {
            const expenseData = expenseSheet.getDataRange().getValues().slice(1); // Header überspringen
            const expenseWarnings = validateSheet(expenseData, 'ausgaben', config);
            if (expenseWarnings.length) {
                allWarnings.push("Fehler in 'Ausgaben':\n" + expenseWarnings.join('\n'));
            }
        }

        // Eigenbelege validieren (wenn vorhanden und Daten vorhanden)
        if (eigenSheet && eigenSheet.getLastRow() > 1) {
            const eigenData = eigenSheet.getDataRange().getValues().slice(1); // Header überspringen
            const eigenWarnings = validateSheet(eigenData, 'eigenbelege', config);
            if (eigenWarnings.length) {
                allWarnings.push("Fehler in 'Eigenbelege':\n" + eigenWarnings.join('\n'));
            }
        }

        // Bankbewegungen validieren (wenn vorhanden)
        if (bankSheet) {
            const bankWarnings = bankValidator.validateBanking(bankSheet, config);
            if (bankWarnings.length) {
                allWarnings.push("Fehler in 'Bankbewegungen':\n" + bankWarnings.join('\n'));
            }
        }

        // Fehlermeldungen anzeigen, falls vorhanden
        if (allWarnings.length) {
            const ui = SpreadsheetApp.getUi();
            // Bei vielen Fehlern ggf. einschränken, um UI-Limits zu vermeiden
            const maxMsgLength = 1500; // Google Sheets Alert-Dialog hat Beschränkungen
            let alertMsg = allWarnings.join('\n\n');

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