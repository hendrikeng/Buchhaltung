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
     * Cache leeren mit gezielter Bereinigung
     */
    clearCache() {
        console.log('Clearing cache for bank reconciliation');
        // Gezieltes Löschen relevanter Cache-Bereiche
        globalCache.clear('computed');
        globalCache.clear('references');
    },

    /**
     * Hauptfunktion für den Bankabgleich mit optimierter Fehlerbehandlung und Performance
     * @param {Object} config - Die Konfiguration
     * @returns {boolean} - true bei erfolgreichem Abgleich
     */
    reconcileBank(config) {
        try {
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            const ui = SpreadsheetApp.getUi();

            // Cache zurücksetzen
            this.clearCache();

            // 1. Bankbewegungen-Sheet validieren
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

            console.log(`Bank reconciliation starting with ${lastRow} rows`);

            // 2. Parameter für den Abgleich definieren
            const firstDataRow = 3; // Nach Header-Zeile und Anfangssaldo
            const numDataRows = lastRow - firstDataRow + 1;

            // Spalten-Konfiguration für Bankbewegungen
            const columns = config.bankbewegungen.columns;

            // 3. Spaltenbuchstaben aus den Indizes generieren (einmalig)
            const columnLetters = {};
            Object.entries(columns).forEach(([key, index]) => {
                columnLetters[key] = stringUtils.getColumnLetter(index);
            });

            // 4. Bankabgleich starten - mit besserer Fehlerbehandlung
            console.log('Starting bank reconciliation process...');

            try {
                // 5. Matching zwischen Bankbewegungen und Rechnungen
                const matchResults = matchingHandler.performBankReferenceMatching(
                    ss, bankSheet, firstDataRow, numDataRows, lastRow, columns, config);

                if (!matchResults) {
                    throw new Error('Der Bankabgleich konnte keine Übereinstimmungen finden.');
                }

                console.log(`Found ${Object.keys(matchResults.bankZuordnungen).length} matches`);

                // 6. Formatierung der Bankbewegungen basierend auf den Match-Ergebnissen
                formattingHandler.formatMatchedRows(
                    bankSheet, firstDataRow, matchResults.matchInfo, columns);

                // 7. Match-Spalte formatieren
                if (columns.matchInfo) {
                    formattingHandler.setMatchColumnFormatting(
                        bankSheet, columnLetters.matchInfo);
                }

                // 8. Aktualisierung der Zahlungsstatus in den Dokument-Sheets
                syncHandler.markPaidInvoices(ss, matchResults.bankZuordnungen, config);

                ui.alert('Bankabgleich erfolgreich durchgeführt!');
                return true;
            } catch (reconciliationError) {
                console.error('Fehler beim Bankabgleich-Prozess:', reconciliationError);
                ui.alert(`Fehler beim Bankabgleich: ${reconciliationError.message}`);
                return false;
            }
        } catch (e) {
            console.error('Schwerwiegender Fehler beim Bankabgleich:', e);
            SpreadsheetApp.getUi().alert('Ein schwerwiegender Fehler ist beim Bankabgleich aufgetreten: ' + e.toString());
            return false;
        }
    },
};

export default BankReconciliationModule;