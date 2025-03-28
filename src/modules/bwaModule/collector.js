// modules/bwaModule/collector.js
import globalCache from '../../utils/cacheUtils.js';
import dataModel from './dataModel.js';
import calculator from './calculator.js';

/**
 * Sammelt alle BWA-Daten aus den verschiedenen Sheets mit optimierter Batch-Verarbeitung
 * @param {Object} config - Die Konfiguration
 * @returns {Object|null} BWA-Daten nach Monaten oder null bei Fehler
 */
function aggregateBWAData(config) {
    try {
        // Prüfen ob Cache gültig ist
        if (globalCache.has('computed', 'bwa')) {
            console.log('Using cached BWA data');
            return globalCache.get('computed', 'bwa');
        }

        console.log('Aggregating BWA data...');
        const ss = SpreadsheetApp.getActiveSpreadsheet();

        // Optimierte Struktur für Sheets und ihre Prozessoren
        const sheetProcessors = [
            { name: 'Einnahmen', type: 'einnahmen', processor: calculator.processRevenue, required: true },
            { name: 'Ausgaben', type: 'ausgaben', processor: calculator.processExpense, required: true },
            { name: 'Eigenbelege', type: 'eigenbelege', processor: calculator.processEigen, required: false },
            { name: 'Gesellschafterkonto', type: 'gesellschafterkonto', processor: calculator.processGesellschafter, required: false },
            { name: 'Holding Transfers', type: 'holdingTransfers', processor: calculator.processHolding, required: false },
        ];

        // Prüfe zuerst, ob alle erforderlichen Sheets existieren
        const sheets = {};

        for (const processor of sheetProcessors) {
            sheets[processor.type] = ss.getSheetByName(processor.name);

            if (processor.required && !sheets[processor.type]) {
                console.error(`Fehlendes Blatt: '${processor.name}'`);
                return null;
            }
        }

        // BWA-Daten für alle Monate initialisieren (für bessere Performance einmal)
        const bwaData = Object.fromEntries(
            Array.from({length: 12}, (_, i) => [i + 1, dataModel.createEmptyBWA()]),
        );

        // Optimierung: Alle Sheets in einem Durchgang verarbeiten
        for (const processor of sheetProcessors) {
            const sheet = sheets[processor.type];
            if (!sheet) continue;

            const lastRow = sheet.getLastRow();
            if (lastRow <= 1) continue; // Nur Header, keine Daten

            console.log(`Processing ${processor.name} sheet for BWA...`);

            // Optimierung: Daten in einem Batch laden für weniger API-Calls
            const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

            // Daten verarbeiten mit der passenden Prozessorfunktion
            for (const row of data) {
                processor.processor(row, bwaData, config);
            }
        }

        // Alle Aggregate in einem Durchgang berechnen
        calculator.calculateAggregates(bwaData, config);

        // Ergebnis cachen
        globalCache.set('computed', 'bwa', bwaData);
        console.log('BWA data aggregation complete');

        return bwaData;
    } catch (e) {
        console.error('Fehler bei der Aggregation der BWA-Daten:', e);
        return null;
    }
}

export default {
    aggregateBWAData,
};