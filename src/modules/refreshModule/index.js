// src/modules/refreshModule/index.js
import dataSheetHandler from './dataSheetHandler.js';
import bankSheetHandler from './bankSheetHandler.js';
import globalCache from '../../utils/cacheUtils.js';

/**
 * Modul zum Aktualisieren der Tabellenblätter und Neuberechnen von Formeln
 */
const RefreshModule = {
    /**
     * Cache zurücksetzen
     */
    clearCache() {
        console.log('Clearing cache');
        globalCache.clear();
    },

    /**
     * Aktualisiert das aktive Tabellenblatt
     * @param {Object} config - Die Konfiguration
     */
    refreshActiveSheet(config) {
        try {
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            const sheet = ss.getActiveSheet();
            const name = sheet.getName();
            const ui = SpreadsheetApp.getUi();

            console.log(`Refreshing active sheet: ${name}`);

            // Cache zurücksetzen
            this.clearCache();

            if (['Einnahmen', 'Ausgaben', 'Eigenbelege'].includes(name)) {
                dataSheetHandler.refreshDataSheet(sheet, name, config);
                ui.alert(`Das Blatt "${name}" wurde aktualisiert.`);
            } else if (name === 'Bankbewegungen') {
                bankSheetHandler.refreshBankSheet(sheet, config);
                ui.alert(`Das Blatt "${name}" wurde aktualisiert.`);
            } else if (name === 'Gesellschafterkonto') {
                dataSheetHandler.refreshDataSheet(sheet, name, config);
                ui.alert(`Das Blatt "${name}" wurde aktualisiert.`);
            } else if (name === 'Holding Transfers') {
                dataSheetHandler.refreshDataSheet(sheet, name, config);
                ui.alert(`Das Blatt "${name}" wurde aktualisiert.`);
            } else {
                ui.alert('Für dieses Blatt gibt es keine Refresh-Funktion.');
            }
        } catch (e) {
            console.error('Fehler beim Aktualisieren des aktiven Sheets:', e);
            SpreadsheetApp.getUi().alert('Ein Fehler ist beim Aktualisieren aufgetreten: ' + e.toString());
        }
    },

    /**
     * Aktualisiert alle relevanten Tabellenblätter
     * @param {Object} config - Die Konfiguration
     */
    refreshAllSheets(config) {
        try {
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            console.log('Refreshing all sheets');

            // Cache zurücksetzen
            this.clearCache();

            // Sheets in der richtigen Reihenfolge aktualisieren, um Abhängigkeiten zu berücksichtigen
            const refreshOrder = ['Einnahmen', 'Ausgaben', 'Eigenbelege', 'Gesellschafterkonto', 'Holding Transfers', 'Bankbewegungen'];

            for (const name of refreshOrder) {
                const sheet = ss.getSheetByName(name);
                if (!sheet) {
                    console.log(`Sheet ${name} not found, skipping`);
                    continue;
                }

                console.log(`Processing sheet: ${name}`);
                if (name === 'Bankbewegungen') {
                    bankSheetHandler.refreshBankSheet(sheet, config);
                } else {
                    dataSheetHandler.refreshDataSheet(sheet, name, config);
                }

                // Kurze Pause einfügen, um API-Limits zu vermeiden
                Utilities.sleep(100);
            }

            console.log('All sheets refreshed successfully');
        } catch (e) {
            console.error('Fehler beim Aktualisieren aller Sheets:', e);
            throw e; // Fehlermeldung weiterleiten, damit sie in der Hauptfunktion angezeigt wird
        }
    },
};

export default RefreshModule;