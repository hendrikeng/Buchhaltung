// src/modules/bankReconciliationModule/index.js
import matchingHandler from './matchingHandler.js';
import formattingHandler from './formattingHandler.js';
import syncHandler from './syncHandler.js';
import globalCache from '../../utils/cacheUtils.js';
import stringUtils from '../../utils/stringUtils.js';

/**
 * Modul für den Bankabgleich
 * Führt den Abgleich zwischen Bankbewegungen und Ein-/Ausgaben durch
 */
const BankReconciliationModule = {
    /**
     * Cache leeren
     */
    clearCache() {
        console.log('Clearing cache for bank reconciliation');
        globalCache.clear();
    },

    /**
     * Hauptfunktion für den Bankabgleich
     * @param {Object} config - Die Konfiguration
     * @returns {boolean} - true bei erfolgreichem Abgleich
     */
    reconcileBank(config) {
        try {
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            const ui = SpreadsheetApp.getUi();

            // Cache zurücksetzen
            this.clearCache();

            // Bankbewegungen-Sheet finden
            const bankSheet = ss.getSheetByName('Bankbewegungen');
            if (!bankSheet) {
                ui.alert('Fehler: Das Bankbewegungen-Sheet wurde nicht gefunden.');
                return false;
            }

            // Basis-Daten zum Sheet ermitteln
            const lastRow = bankSheet.getLastRow();
            if (lastRow < 3) {
                ui.alert('Es sind nicht genügend Daten im Bankbewegungen-Sheet vorhanden.');
                return false;
            }

            const firstDataRow = 3; // Erste Datenzeile (nach Header-Zeile und Anfangssaldo)
            const numDataRows = lastRow - firstDataRow + 1;

            // Spalten-Konfiguration für Bankbewegungen
            const columns = config.bankbewegungen.columns;

            // Spaltenbuchstaben aus den Indizes generieren
            const columnLetters = {};
            Object.entries(columns).forEach(([key, index]) => {
                columnLetters[key] = stringUtils.getColumnLetter(index);
            });

            // Durchführung des eigentlichen Bankabgleichs
            console.log('Starting bank reconciliation...');

            // 1. Matching zwischen Bankbewegungen und Rechnungen
            const matchResults = matchingHandler.performBankReferenceMatching(
                ss, bankSheet, firstDataRow, numDataRows, lastRow, columns, config);

            // 2. Formatierung der Bankbewegungen basierend auf den Match-Ergebnissen
            formattingHandler.formatMatchedRows(
                bankSheet, firstDataRow, matchResults.matchInfo, columns);

            // 3. Match-Spalte formatieren
            if (columns.matchInfo) {
                formattingHandler.setMatchColumnFormatting(
                    bankSheet, columnLetters.matchInfo);
            }

            // 4. Aktualisierung der Zahlungsstatus in den Dokument-Sheets
            syncHandler.markPaidInvoices(ss, matchResults.bankZuordnungen, config);

            ui.alert('Bankabgleich erfolgreich durchgeführt!');
            return true;

        } catch (e) {
            console.error('Fehler beim Bankabgleich:', e);
            SpreadsheetApp.getUi().alert('Ein Fehler ist beim Bankabgleich aufgetreten: ' + e.toString());
            return false;
        }
    },
};

export default BankReconciliationModule;