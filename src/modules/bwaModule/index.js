// modules/bwaModule/index.js
import dataModel from './dataModel.js';
import collector from './collector.js';
import calculator from './calculator.js';
import formatter from './formatter.js';
import globalCache from '../../utils/cacheUtils.js';

/**
 * Modul zur Berechnung der Betriebswirtschaftlichen Auswertung (BWA)
 */
const BWAModule = {
    /**
     * Cache leeren mit gezielter Invalidierung
     */
    clearCache() {
        console.log('Clearing BWA module cache');
        globalCache.clear('bwa');
    },

    /**
     * Hauptfunktion zur Berechnung der BWA mit optimierter Fehlerbehandlung
     * @param {Object} config - Die Konfiguration
     * @returns {boolean} - true bei Erfolg, false bei Fehler
     */
    calculateBWA(config) {
        try {
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            const ui = SpreadsheetApp.getUi();

            // Cache zurücksetzen für aktuelle Daten
            this.clearCache();

            console.log('Starting BWA calculation...');
            ui.alert('BWA wird berechnet...', 'Bitte warten Sie, während die BWA berechnet wird.', ui.ButtonSet.OK);

            // Daten sammeln
            const bwaData = collector.aggregateBWAData(config);
            if (!bwaData) {
                ui.alert('BWA-Daten konnten nicht generiert werden.');
                return false;
            }

            // BWA-Sheet erstellen oder aktualisieren
            const success = formatter.generateBWASheet(bwaData, ss, config);

            if (success) {
                ui.alert('BWA wurde aktualisiert!');
                return true;
            } else {
                ui.alert('Bei der Erstellung der BWA ist ein Fehler aufgetreten.');
                return false;
            }
        } catch (e) {
            console.error('Fehler bei der BWA-Berechnung:', e);
            SpreadsheetApp.getUi().alert('Fehler bei der BWA-Berechnung: ' + e.toString());
            return false;
        }
    },

    // Methoden für Testzwecke und erweiterte Funktionalität
    _internal: {
        createEmptyBWA: dataModel.createEmptyBWA,
        processRevenue: calculator.processRevenue,
        processExpense: calculator.processExpense,
        processEigen: calculator.processEigen,
        aggregateBWAData: collector.aggregateBWAData,
    },
};

export default BWAModule;