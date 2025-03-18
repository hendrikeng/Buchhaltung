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
        console.log('Clearing refresh module cache');
        globalCache.clear();
    },

    /**
     * Aktualisiert das aktive Tabellenblatt mit optimierter Fehlerbehandlung
     * @param {Object} config - Die Konfiguration
     * @returns {boolean} - true bei erfolgreicher Aktualisierung
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

            // Mapping für Sheet-Handler
            const sheetTypeMap = {
                'Einnahmen': () => dataSheetHandler.refreshDataSheet(sheet, name, config),
                'Ausgaben': () => dataSheetHandler.refreshDataSheet(sheet, name, config),
                'Eigenbelege': () => dataSheetHandler.refreshDataSheet(sheet, name, config),
                'Bankbewegungen': () => bankSheetHandler.refreshBankSheet(sheet, config),
                'Gesellschafterkonto': () => dataSheetHandler.refreshDataSheet(sheet, name, config),
                'Holding Transfers': () => dataSheetHandler.refreshDataSheet(sheet, name, config),
            };

            // Handler für den Sheet-Typ ausführen
            if (sheetTypeMap[name]) {
                const success = sheetTypeMap[name]();
                if (success) {
                    ui.alert(`Das Blatt "${name}" wurde aktualisiert.`);
                    return true;
                } else {
                    ui.alert(`Bei der Aktualisierung von "${name}" ist ein Fehler aufgetreten.`);
                    return false;
                }
            } else {
                ui.alert('Für dieses Blatt gibt es keine Refresh-Funktion.');
                return false;
            }
        } catch (e) {
            console.error('Fehler beim Aktualisieren des aktiven Sheets:', e);
            SpreadsheetApp.getUi().alert('Ein Fehler ist beim Aktualisieren aufgetreten: ' + e.toString());
            return false;
        }
    },

    /**
     * Aktualisiert alle relevanten Tabellenblätter mit optimierter Batch-Verarbeitung
     * @param {Object} config - Die Konfiguration
     * @returns {boolean} - true bei erfolgreicher Aktualisierung
     */
    refreshAllSheets(config) {
        try {
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            console.log('Refreshing all sheets');

            // Cache zurücksetzen
            this.clearCache();

            // Sheets in der richtigen Reihenfolge aktualisieren, um Abhängigkeiten zu berücksichtigen
            const refreshOrder = [
                { name: 'Einnahmen', handler: dataSheetHandler.refreshDataSheet },
                { name: 'Ausgaben', handler: dataSheetHandler.refreshDataSheet },
                { name: 'Eigenbelege', handler: dataSheetHandler.refreshDataSheet },
                { name: 'Gesellschafterkonto', handler: dataSheetHandler.refreshDataSheet },
                { name: 'Holding Transfers', handler: dataSheetHandler.refreshDataSheet },
                { name: 'Bankbewegungen', handler: bankSheetHandler.refreshBankSheet },
            ];

            // Status-Tracking
            let successCount = 0;
            let failureCount = 0;
            const failedSheets = [];

            // Alle Sheets in optimierter Reihenfolge aktualisieren
            for (const { name, handler } of refreshOrder) {
                const sheet = ss.getSheetByName(name);
                if (!sheet) {
                    console.log(`Sheet ${name} not found, skipping`);
                    continue;
                }

                console.log(`Processing sheet: ${name}`);
                try {
                    // Bank-Handler oder Daten-Handler aufrufen
                    const success = handler === bankSheetHandler.refreshBankSheet ?
                        handler(sheet, config) : handler(sheet, name, config);

                    if (success) {
                        successCount++;
                    } else {
                        failureCount++;
                        failedSheets.push(name);
                    }

                    // Kurze Pause einfügen, um API-Limits zu vermeiden
                    Utilities.sleep(100);
                } catch (sheetError) {
                    console.error(`Error refreshing sheet ${name}:`, sheetError);
                    failureCount++;
                    failedSheets.push(name);
                }
            }

            console.log(`Refresh completed: ${successCount} sheets updated, ${failureCount} failed`);

            // Zeige Fehlermeldung wenn nötig
            if (failureCount > 0) {
                SpreadsheetApp.getUi().alert(
                    `Einige Sheets konnten nicht aktualisiert werden: ${failedSheets.join(', ')}`,
                );
                return false;
            }

            return true;
        } catch (e) {
            console.error('Fehler beim Aktualisieren aller Sheets:', e);
            SpreadsheetApp.getUi().alert('Ein Fehler ist beim Aktualisieren aller Sheets aufgetreten: ' + e.toString());
            return false;
        }
    },
};

export default RefreshModule;