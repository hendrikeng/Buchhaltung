// modules/bilanzModule/collector.js
import globalCache from '../../utils/cacheUtils.js';
import dataModel from './dataModel.js';
import calculator from './calculator.js';
import numberUtils from '../../utils/numberUtils.js';

/**
 * Sammelt Daten aus verschiedenen Sheets für die Bilanz mit optimierter Batch-Verarbeitung
 * @param {Object} config - Die Konfiguration
 * @returns {Object} Bilanz-Datenstruktur mit befüllten Werten
 */
function aggregateBilanzData(config) {
    try {
        // Prüfen ob Cache gültig ist
        if (globalCache.has('computed', 'bilanz')) {
            console.log('Using cached bilanz data');
            return globalCache.get('computed', 'bilanz');
        }

        console.log('Aggregating bilanz data...');
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const bilanzData = dataModel.createEmptyBilanz();

        // Spalten-Konfigurationen für die verschiedenen Sheets abrufen
        const bankCols = config.bankbewegungen.columns;
        const ausgabenCols = config.ausgaben.columns;
        const gesellschafterCols = config.gesellschafterkonto.columns;

        // 1. Banksaldo aus "Bankbewegungen" (Endsaldo) - optimiert mit gezielter Datenabfrage
        const bankSheet = ss.getSheetByName('Bankbewegungen');
        if (bankSheet) {
            const lastRow = bankSheet.getLastRow();
            if (lastRow >= 1) {
                // Optimierung: Nur die benötigten Zellen direkt abfragen
                const lastRowData = bankSheet.getRange(lastRow, 1, 1, bankCols.buchungstext).getValues()[0];
                const label = lastRowData[bankCols.buchungstext - 1]?.toString().toLowerCase() || '';

                if (label === 'endsaldo') {
                    bilanzData.aktiva.bankguthaben = numberUtils.parseCurrency(
                        bankSheet.getRange(lastRow, bankCols.saldo).getValue(),
                    );
                }
            }
        }

        // 2. Jahresüberschuss aus "BWA" - optimiert mit gezielter Suche
        const bwaSheet = ss.getSheetByName('BWA');
        if (bwaSheet) {
            const lastCol = bwaSheet.getLastColumn();
            // Optimierung: Nur in der letzten Spalte suchen, wo das Jahresergebnis steht
            const yearColIndex = lastCol;

            // Finde die Jahresüberschusszeile
            const searchRange = bwaSheet.getRange(1, 1, bwaSheet.getLastRow(), 1).getValues();
            let jahresueberschussRow = -1;

            for (let i = searchRange.length - 1; i >= 0; i--) {
                if (searchRange[i][0].toString().toLowerCase().includes('jahresüberschuss')) {
                    jahresueberschussRow = i + 1; // 1-basierter Index
                    break;
                }
            }

            if (jahresueberschussRow > 0) {
                bilanzData.passiva.jahresueberschuss = numberUtils.parseCurrency(
                    bwaSheet.getRange(jahresueberschussRow, yearColIndex).getValue(),
                );
            }
        }

        // 3. Stammkapital aus Konfiguration
        bilanzData.passiva.stammkapital = config.tax.stammkapital || 25000;

        // 4. Gesellschafterdarlehen aus dem Gesellschafterkonto-Sheet
        bilanzData.passiva.gesellschafterdarlehen = calculator.getDarlehensumme(ss, gesellschafterCols, config);

        // 5. Steuerrückstellungen aus dem Ausgaben-Sheet
        bilanzData.passiva.steuerrueckstellungen = calculator.getSteuerRueckstellungen(ss, ausgabenCols, config);

        // 6. Berechnung der Summen
        calculator.calculateBilanzSummen(bilanzData);

        // Daten im Cache speichern
        globalCache.set('computed', 'bilanz', bilanzData);
        console.log('Bilanz data aggregation complete');

        return bilanzData;
    } catch (e) {
        console.error('Fehler bei der Sammlung der Bilanzdaten:', e);
        return null;
    }
}

export default {
    aggregateBilanzData,
};