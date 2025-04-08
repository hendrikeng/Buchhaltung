// modules/bilanzModule/index.js
import dataModel from './dataModel.js';
import collector from './collector.js';
import calculator from './calculator.js';
import formatter from './formatter.js';
import globalCache from '../../utils/cacheUtils.js';
import numberUtils from '../../utils/numberUtils.js';

/**
 * Module for creating a DATEV-compliant balance sheet
 * Creates a standardized balance sheet based on data from other sheets
 */
const BilanzModule = {
    /**
     * Clears cache with targeted invalidation
     */
    clearCache() {
        console.log('Clearing bilanz module cache');
        globalCache.clear('bilanz');
    },

    /**
     * Main function for creating the balance sheet
     * Collects data and creates a balance sheet with optimized error handling
     * @param {Object} config - Configuration
     * @returns {boolean} true on success, false on error
     */
    calculateBilanz(config) {
        try {
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            const ui = SpreadsheetApp.getUi();

            // Reset cache for current data
            this.clearCache();

            console.log('Starting DATEV-compliant balance sheet creation...');
            ui.alert('Bilanz wird erstellt...', 'Bitte warten Sie, während die Bilanz erstellt wird.', ui.ButtonSet.OK);

            // Collect balance sheet data
            const bilanzData = collector.aggregateBilanzData(config);
            if (!bilanzData) {
                ui.alert('Fehler: Bilanzdaten konnten nicht gesammelt werden.');
                return false;
            }

            // Create balance sheet
            const success = formatter.generateBilanzSheet(bilanzData, ss, config);

            // Validation: Check if assets = liabilities
            const aktivaSumme = bilanzData.aktiva.summe;
            const passivaSumme = bilanzData.passiva.summe;
            const differenz = Math.abs(aktivaSumme - passivaSumme);

            if (differenz > 0.01) {
                // If there's a difference, show warning but still create the balance sheet
                ui.alert(
                    'Bilanz ist nicht ausgeglichen',
                    `Die Bilanzsummen von Aktiva (${numberUtils.formatCurrency(aktivaSumme)}) und Passiva (${numberUtils.formatCurrency(passivaSumme)}) ` +
                    `stimmen nicht überein. Differenz: ${numberUtils.formatCurrency(differenz)}. ` +
                    'Bitte überprüfen Sie Ihre Buchhaltungsdaten.',
                    ui.ButtonSet.OK,
                );
            }

            if (success) {
                ui.alert('Die Bilanz wurde erfolgreich erstellt!');
                return true;
            } else {
                ui.alert('Bei der Erstellung der Bilanz ist ein Fehler aufgetreten.');
                return false;
            }
        } catch (e) {
            console.error('Fehler bei der Bilanzerstellung:', e);
            SpreadsheetApp.getUi().alert('Fehler bei der Bilanzerstellung: ' + e.toString());
            return false;
        }
    },

    // Methods for testing and extended functionality
    _internal: {
        createEmptyBilanz: dataModel.createEmptyBilanz,
        aggregateBilanzData: collector.aggregateBilanzData,
    },
};

export default BilanzModule;