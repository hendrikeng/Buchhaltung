// modules/bwaModule/collector.js
import globalCache from '../../utils/cacheUtils.js';
import dataModel from './dataModel.js';
import calculator from './calculator.js';

/**
 * Collects all BWA data from the various sheets with DATEV-compliant structure
 * @param {Object} config - Configuration
 * @returns {Object|null} BWA data by month or null on error
 */
function aggregateBWAData(config) {
    try {
        // Check if cache is valid
        if (globalCache.has('computed', 'bwa')) {
            console.log('Using cached BWA data');
            return globalCache.get('computed', 'bwa');
        }

        console.log('Aggregating BWA data with DATEV-compliant structure...');
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

        // Counter for processed rows
        const processedRows = {
            einnahmen: 0,
            ausgaben: 0,
            eigenbelege: 0,
            gesellschafterkonto: 0,
            holdingTransfers: 0,
        };

        // Process all sheets in one pass
        for (const processor of sheetProcessors) {
            const sheet = sheets[processor.type];
            if (!sheet) continue;

            const lastRow = sheet.getLastRow();
            if (lastRow <= 1) continue; // Only header, no data

            console.log(`Processing ${processor.name} sheet for BWA...`);

            // Optimization: load data in one batch for fewer API calls
            const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

            // Process data with appropriate processor function
            for (const row of data) {
                processor.processor(row, bwaData, config);
                processedRows[processor.type]++;
            }
        }

        console.log('Processed rows:', JSON.stringify(processedRows));

        // Calculate all aggregates in one pass
        calculator.calculateAggregates(bwaData, config);

        // Cache result
        globalCache.set('computed', 'bwa', bwaData);
        console.log('BWA data aggregation complete');

        return bwaData;
    } catch (e) {
        console.error('Error aggregating BWA data:', e);
        return null;
    }
}

export default {
    aggregateBWAData,
};