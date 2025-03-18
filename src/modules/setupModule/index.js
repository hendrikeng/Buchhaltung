// src/modules/setupModule/index.js
import sheetCreator from './sheetCreator.js';
import globalCache from '../../utils/cacheUtils.js';

/**
 * Modul zum Einrichten und Initialisieren der Buchhaltungstabelle
 */
const SetupModule = {
    /**
     * Cache leeren
     */
    clearCache() {
        globalCache.clear();
    },

    /**
     * Richtet die grundlegende Struktur für die Buchhaltung ein
     * @param {Object} config - Die Konfiguration
     * @returns {boolean} - true bei erfolgreicher Einrichtung
     */
    setupSpreadsheet(config) {
        try {
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            const ui = SpreadsheetApp.getUi();

            // Frage den Benutzer, ob er fortfahren möchte
            const result = ui.alert(
                'Buchhaltung einrichten',
                'Dies wird alle benötigten Sheets erstellen oder aktualisieren. Vorhandene Daten bleiben erhalten. Möchten Sie fortfahren?',
                ui.ButtonSet.YES_NO,
            );

            if (result !== ui.Button.YES) {
                return false;
            }

            // Cache zurücksetzen
            this.clearCache();

            // Alle notwendigen Sheets erstellen
            sheetCreator.createAllSheets(ss, config);

            // // Trigger einrichten (onOpen, onEdit)
            // setupTrigger();

            ui.alert(
                'Einrichtung abgeschlossen',
                'Die Buchhaltungstabelle wurde erfolgreich eingerichtet.',
                ui.ButtonSet.OK,
            );

            return true;
        } catch (e) {
            console.error('Fehler bei der Einrichtung der Buchhaltung:', e);
            SpreadsheetApp.getUi().alert('Ein Fehler ist bei der Einrichtung aufgetreten: ' + e.toString());
            return false;
        }
    },

    // Methoden für Testzwecke und erweiterte Funktionalität
    _internal: {
        sheetCreator,
    },
};

/**
 * Richtet die notwendigen Trigger für das Spreadsheet ein
 */
// function setupTrigger() {
//     const triggers = ScriptApp.getProjectTriggers();
//
//     // Prüfe, ob der onOpen Trigger bereits existiert
//     if (!triggers.some(t => t.getHandlerFunction() === 'onOpen')) {
//         ScriptApp.newTrigger('onOpen')
//             .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
//             .onOpen()
//             .create();
//     }
//
//     // Prüfe, ob der onEdit Trigger bereits existiert
//     if (!triggers.some(t => t.getHandlerFunction() === 'onEdit')) {
//         ScriptApp.newTrigger('onEdit')
//             .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
//             .onEdit()
//             .create();
//     }
// }

export default SetupModule;