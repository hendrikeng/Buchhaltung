// src/modules/bankReconciliationModule/index.js
import matchingHandler from './matchingHandler.js';
import formattingHandler from './formattingHandler.js';
import syncHandler from './syncHandler.js';
import approvalHandler from './approvalHandler.js';
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

                // 6. Interaktive Genehmigung der Zuordnungen
                const approvalResults = approvalHandler.processApprovals(
                    ss, matchResults.bankZuordnungen, config);

                if (Object.keys(approvalResults.approvedMatches).length === 0) {
                    ui.alert('Bankabgleich abgebrochen', 'Es wurden keine Zuordnungen genehmigt.', ui.ButtonSet.OK);
                    return false;
                }

                // 7. Formatierung der Bankbewegungen basierend auf den genehmigten Match-Ergebnissen
                formattingHandler.formatMatchedRows(
                    bankSheet, firstDataRow, matchResults.matchInfo, columns);

                // 8. Match-Spalte formatieren
                if (columns.matchInfo) {
                    formattingHandler.setMatchColumnFormatting(
                        bankSheet, columnLetters.matchInfo);
                }

                // 9. Aktualisierung der Zahlungsstatus in den Dokument-Sheets (nur genehmigte)
                syncHandler.markPaidInvoices(ss, approvalResults.approvedMatches, config);

                ui.alert('Bankabgleich erfolgreich durchgeführt!',
                    `Es wurden ${Object.keys(approvalResults.approvedMatches).length} Zuordnungen verarbeitet.`,
                    ui.ButtonSet.OK);
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