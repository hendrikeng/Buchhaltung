// modules/bwaModule/collector.js
import globalCache from '../../utils/cacheUtils.js';
import dataModel from './dataModel.js';
import calculator from './calculator.js';

/**
 * Sammelt alle BWA-Daten aus den verschiedenen Sheets
 * @param {Object} config - Die Konfiguration
 * @returns {Object|null} BWA-Daten nach Monaten oder null bei Fehler
 */
function aggregateBWAData(config) {
    try {
        // Prüfen ob Cache gültig ist
        if (globalCache.has('computed', 'bwa')) {
            return globalCache.get('computed', 'bwa');
        }

        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const revenueSheet = ss.getSheetByName("Einnahmen");
        const expenseSheet = ss.getSheetByName("Ausgaben");
        const eigenSheet = ss.getSheetByName("Eigenbelege");
        const gesellschafterSheet = ss.getSheetByName("Gesellschafterkonto");
        const holdingSheet = ss.getSheetByName("Holding Transfers");

        if (!revenueSheet || !expenseSheet) {
            console.error("Fehlendes Blatt: 'Einnahmen' oder 'Ausgaben'");
            return null;
        }

        // BWA-Daten für alle Monate initialisieren
        const bwaData = Object.fromEntries(Array.from({length: 12}, (_, i) => [i + 1, dataModel.createEmptyBWA()]));

        // Daten aus den Sheets verarbeiten
        revenueSheet.getDataRange().getValues().slice(1).forEach(row => calculator.processRevenue(row, bwaData, config));
        expenseSheet.getDataRange().getValues().slice(1).forEach(row => calculator.processExpense(row, bwaData, config));

        if (eigenSheet) {
            eigenSheet.getDataRange().getValues().slice(1).forEach(row => calculator.processEigen(row, bwaData, config));
        }

        if (gesellschafterSheet) {
            gesellschafterSheet.getDataRange().getValues().slice(1).forEach(row => calculator.processGesellschafter(row, bwaData, config));
        }

        if (holdingSheet) {
            holdingSheet.getDataRange().getValues().slice(1).forEach(row => calculator.processHolding(row, bwaData, config));
        }

        // Gruppensummen und weitere Berechnungen
        calculator.calculateAggregates(bwaData, config);

        // Daten cachen
        globalCache.set('computed', 'bwa', bwaData);

        return bwaData;
    } catch (e) {
        console.error("Fehler bei der Aggregation der BWA-Daten:", e);
        return null;
    }
}

export default {
    aggregateBWAData
};