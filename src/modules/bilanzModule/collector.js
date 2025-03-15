// modules/bilanzModule/collector.js
import globalCache from '../../utils/cacheUtils.js';
import dataModel from './dataModel.js';
import calculator from './calculator.js';
import numberUtils from '../../utils/numberUtils.js';

/**
 * Sammelt Daten aus verschiedenen Sheets für die Bilanz
 * @param {Object} config - Die Konfiguration
 * @returns {Object} Bilanz-Datenstruktur mit befüllten Werten
 */
function aggregateBilanzData(config) {
    try {
        // Prüfen ob Cache gültig ist
        if (globalCache.has('computed', 'bilanz')) {
            return globalCache.get('computed', 'bilanz');
        }

        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const bilanzData = dataModel.createEmptyBilanz();

        // Spalten-Konfigurationen für die verschiedenen Sheets abrufen
        const bankCols = config.bankbewegungen.columns;
        const ausgabenCols = config.ausgaben.columns;
        const gesellschafterCols = config.gesellschafterkonto.columns;

        // 1. Banksaldo aus "Bankbewegungen" (Endsaldo)
        const bankSheet = ss.getSheetByName("Bankbewegungen");
        if (bankSheet) {
            const lastRow = bankSheet.getLastRow();
            if (lastRow >= 1) {
                const label = bankSheet.getRange(lastRow, bankCols.buchungstext).getValue().toString().toLowerCase();
                if (label === "endsaldo") {
                    bilanzData.aktiva.bankguthaben = numberUtils.parseCurrency(
                        bankSheet.getRange(lastRow, bankCols.saldo).getValue()
                    );
                }
            }
        }

        // 2. Jahresüberschuss aus "BWA" (Letzte Zeile, sofern dort "Jahresüberschuss" vorkommt)
        const bwaSheet = ss.getSheetByName("BWA");
        if (bwaSheet) {
            const data = bwaSheet.getDataRange().getValues();
            for (let i = data.length - 1; i >= 0; i--) {
                const row = data[i];
                if (row[0].toString().toLowerCase().includes("jahresüberschuss")) {
                    // Letzte Spalte enthält den Jahreswert
                    bilanzData.passiva.jahresueberschuss = numberUtils.parseCurrency(row[row.length - 1]);
                    break;
                }
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

        return bilanzData;
    } catch (e) {
        console.error("Fehler bei der Sammlung der Bilanzdaten:", e);
        return null;
    }
}

export default {
    aggregateBilanzData
};