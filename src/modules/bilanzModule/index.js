// modules/bilanzModule/index.js
import dataModel from './dataModel.js';
import collector from './collector.js';
import calculator from './calculator.js';
import formatter from './formatter.js';
import globalCache from '../../utils/cacheUtils.js';
import numberUtils from '../../utils/numberUtils.js';

/**
 * Modul zur Erstellung einer Bilanz nach SKR04
 * Erstellt eine standardkonforme Bilanz basierend auf den Daten aus anderen Sheets
 */
const BilanzModule = {
    /**
     * Cache leeren mit gezielter Invalidierung
     */
    clearCache() {
        console.log('Clearing bilanz module cache');
        globalCache.clear('bilanz');
    },

    /**
     * Hauptfunktion zur Erstellung der Bilanz
     * Sammelt Daten und erstellt ein Bilanz-Sheet mit optimierter Fehlerbehandlung
     * @param {Object} config - Die Konfiguration
     * @returns {boolean} true bei Erfolg, false bei Fehler
     */
    calculateBilanz(config) {
        try {
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            const ui = SpreadsheetApp.getUi();

            // Cache zurücksetzen für aktuelle Daten
            this.clearCache();

            console.log('Starting bilanz calculation...');
            ui.alert('Bilanz wird erstellt...', 'Bitte warten Sie, während die Bilanz erstellt wird.', ui.ButtonSet.OK);

            // Bilanzdaten aggregieren
            const bilanzData = collector.aggregateBilanzData(config);
            if (!bilanzData) {
                ui.alert('Fehler: Bilanzdaten konnten nicht gesammelt werden.');
                return false;
            }

            // Bilanz-Sheet erstellen
            const success = formatter.generateBilanzSheet(bilanzData, ss, config);

            // Optimierung: Prüfung der Bilanzsummen nur wenn nötig
            // Prüfen, ob Aktiva = Passiva
            const aktivaSumme = bilanzData.aktiva.summeAktiva;
            const passivaSumme = bilanzData.passiva.summePassiva;
            const differenz = Math.abs(aktivaSumme - passivaSumme);

            if (differenz > 0.01) {
                // Bei Differenz die Bilanz trotzdem erstellen, aber warnen
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

    // Methoden für Testzwecke und erweiterte Funktionalität
    _internal: {
        createEmptyBilanz: dataModel.createEmptyBilanz,
        aggregateBilanzData: collector.aggregateBilanzData,
    },
};

export default BilanzModule;