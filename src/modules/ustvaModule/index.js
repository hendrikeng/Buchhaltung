// modules/ustvaModule/index.js
import dataModel from './dataModel.js';
import collector from './collector.js';
import calculator from './calculator.js';
import formatter from './formatter.js';
import globalCache from '../../utils/cacheUtils.js';

/**
 * Modul zur Berechnung der Umsatzsteuervoranmeldung (UStVA)
 * Unterstützt die Berechnung nach SKR04 für monatliche und quartalsweise Auswertungen
 */
const UStVAModule = {
    /**
     * Cache leeren mit gezielter Invalidierung
     */
    clearCache() {
        console.log('Clearing UStVA module cache');
        globalCache.clear('ustva');
    },

    /**
     * Hauptfunktion zur Berechnung der UStVA
     * Sammelt Daten aus allen relevanten Sheets und erstellt ein UStVA-Sheet
     * @param {Object} config - Die Konfiguration
     * @returns {boolean} - true bei Erfolg, false bei Fehler
     */
    calculateUStVA(config) {
        try {
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            const ui = SpreadsheetApp.getUi();

            // Cache zurücksetzen für aktuelle Daten
            this.clearCache();

            console.log('Starting UStVA calculation...');
            ui.alert('UStVA wird berechnet...', 'Bitte warten Sie, während die UStVA berechnet wird.', ui.ButtonSet.OK);

            // Daten sammeln
            const ustvaData = collector.collectUStVAData(config);
            if (!ustvaData) {
                ui.alert('Die UStVA konnte nicht berechnet werden. Bitte prüfen Sie die Fehleranzeige.');
                return false;
            }

            // UStVA-Sheet erstellen oder aktualisieren
            const success = formatter.generateUStVASheet(ustvaData, ss, config);

            if (success) {
                ui.alert('UStVA wurde erfolgreich aktualisiert!');
                return true;
            } else {
                ui.alert('Bei der Erstellung der UStVA ist ein Fehler aufgetreten.');
                return false;
            }
        } catch (e) {
            console.error('Fehler bei der UStVA-Berechnung:', e);
            SpreadsheetApp.getUi().alert('Fehler bei der UStVA-Berechnung: ' + e.toString());
            return false;
        }
    },

    // Methoden für Testzwecke und erweiterte Funktionalität
    _internal: {
        createEmptyUStVA: dataModel.createEmptyUStVA,
        processUStVARow: calculator.processUStVARow,
        formatUStVARow: formatter.formatUStVARow,
        aggregateUStVA: calculator.aggregateUStVA,
        collectUStVAData: collector.collectUStVAData,
    },
};

export default UStVAModule;