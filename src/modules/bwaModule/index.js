// modules/bwaModule/index.js
import dataModel from './dataModel.js';
import collector from './collector.js';
import calculator from './calculator.js';
import formatter from './formatter.js';
import globalCache from '../../utils/cacheUtils.js';

/**
 * Module for calculating business management analysis (BWA)
 */
const BWAModule = {
    /**
     * Clear cache with targeted invalidation
     */
    clearCache() {
        console.log('Clearing BWA module cache');
        globalCache.clear('bwa');
    },

    /**
     * Main function for calculating BWA with optimized error handling
     * @param {Object} config - Configuration
     * @returns {boolean} - true on success, false on error
     */
    calculateBWA(config) {
        try {
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            const ui = SpreadsheetApp.getUi();

            // Reset cache for current data
            this.clearCache();

            console.log('Starting BWA calculation...');
            ui.alert('BWA wird berechnet...', 'Bitte warten Sie, während die BWA berechnet wird.', ui.ButtonSet.OK);

            // Collect data
            const bwaData = collector.aggregateBWAData(config);
            if (!bwaData) {
                ui.alert('BWA-Daten konnten nicht generiert werden.');
                return false;
            }

            // Create or update BWA sheet
            const success = formatter.generateBWASheet(bwaData, ss, config);

            if (success) {
                ui.alert('BWA wurde aktualisiert!');
                return true;
            } else {
                ui.alert('Bei der Erstellung der BWA ist ein Fehler aufgetreten.');
                return false;
            }
        } catch (e) {
            console.error('Error calculating BWA:', e);
            SpreadsheetApp.getUi().alert('Fehler bei der BWA-Berechnung: ' + e.toString());
            return false;
        }
    },

    // Methods for testing and extended functionality
    _internal: {
        createEmptyBWA: dataModel.createEmptyBWA,
        processRevenue: calculator.processRevenue,
        processExpense: calculator.processExpense,
        processEigen: calculator.processEigen,
        aggregateBWAData: collector.aggregateBWAData,
    },
};

export default BWAModule;