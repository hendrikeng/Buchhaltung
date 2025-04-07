// modules/bwaModule/collector.js
import globalCache from '../../utils/cacheUtils.js';
import dataModel from './dataModel.js';
import calculator from './calculator.js';

/**
 * Collects all BWA data from various sheets with optimized batch processing
 * @param {Object} config - Configuration
 * @returns {Object|null} BWA data by month or null in case of error
 */
function aggregateBWAData(config) {
    try {
        // Check if cache is valid
        if (globalCache.has('computed', 'bwa')) {
            return globalCache.get('computed', 'bwa');
        }

        const ss = SpreadsheetApp.getActiveSpreadsheet();

        // Optimized structure for sheets and their processors
        const sheetProcessors = [
            { name: 'Einnahmen', type: 'einnahmen', processor: calculator.processRevenue, required: true },
            { name: 'Ausgaben', type: 'ausgaben', processor: calculator.processExpense, required: true },
            { name: 'Eigenbelege', type: 'eigenbelege', processor: calculator.processEigen, required: false },
            { name: 'Gesellschafterkonto', type: 'gesellschafterkonto', processor: calculator.processGesellschafter, required: false },
            { name: 'Holding Transfers', type: 'holdingTransfers', processor: calculator.processHolding, required: false },
        ];

        // First check if all required sheets exist
        const sheets = {};

        for (const processor of sheetProcessors) {
            sheets[processor.type] = ss.getSheetByName(processor.name);

            if (processor.required && !sheets[processor.type]) {
                console.error(`Missing sheet: '${processor.name}'`);
                return null;
            }
        }

        // Initialize BWA data for all months (once for better performance)
        const bwaData = Object.fromEntries(
            Array.from({length: 12}, (_, i) => [i + 1, dataModel.createEmptyBWA()]),
        );

        // Process all sheets in a single pass
        for (const processor of sheetProcessors) {
            const sheet = sheets[processor.type];
            if (!sheet) continue;

            const lastRow = sheet.getLastRow();
            if (lastRow <= 1) continue; // Only header, no data

            // Optimization: Load data in a batch for fewer API calls
            const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

            // Process data with the appropriate processor function
            for (const row of data) {
                processor.processor(row, bwaData, config);
            }
        }

        // Calculate all aggregates in a single pass
        calculator.calculateAggregates(bwaData, config);

        // Cache result
        globalCache.set('computed', 'bwa', bwaData);

        return bwaData;
    } catch (e) {
        console.error('Error aggregating BWA data:', e);
        return null;
    }
}

export default {
    aggregateBWAData,
};