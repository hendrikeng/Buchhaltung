// src/modules/setupModule/index.js
import sheetCreator from './sheetCreator.js';
import globalCache from '../../utils/cacheUtils.js';

/**
 * Modul zum Einrichten und Initialisieren der Buchhaltungstabelle
 * Optimierte Version mit besserer Fehlerbehandlung und Performance
 */
const SetupModule = {
    /**
     * Cache leeren mit gezielter Invalidierung
     */
    clearCache() {
        console.log('Clearing setup module cache');
        // Gezieltes Löschen aller relevanten Cache-Bereiche
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

            // Fortschrittsanzeige starten
            ui.alert(
                'Einrichtung gestartet',
                'Die Einrichtung wurde gestartet. Bitte warten Sie, bis der Vorgang abgeschlossen ist.',
                ui.ButtonSet.OK,
            );

            // Alle notwendigen Sheets erstellen
            const setupSuccess = sheetCreator.createAllSheets(ss, config);

            if (!setupSuccess) {
                ui.alert(
                    'Fehler bei der Einrichtung',
                    'Bei der Einrichtung der Buchhaltungstabelle ist ein Fehler aufgetreten.',
                    ui.ButtonSet.OK,
                );
                return false;
            }

            // Trigger einrichten (onOpen, onEdit)
            try {
                this.setupTriggers(ss);
            } catch (triggerError) {
                console.error('Fehler beim Einrichten der Trigger:', triggerError);
                // Trotzdem fortfahren, da dies nicht kritisch ist
            }

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

    /**
     * Richtet die notwendigen Trigger für das Spreadsheet ein
     * @param {Spreadsheet} ss - Das Spreadsheet
     */
    setupTriggers(ss) {
        const triggers = ScriptApp.getProjectTriggers();

        // Prüfe, ob der onOpen Trigger bereits existiert
        if (!triggers.some(t =>
            t.getHandlerFunction() === 'onOpen' &&
            t.getTriggerSourceId() === ss.getId(),
        )) {
            ScriptApp.newTrigger('onOpen')
                .forSpreadsheet(ss)
                .onOpen()
                .create();
        }

        // Prüfe, ob der onEdit Trigger bereits existiert
        if (!triggers.some(t =>
            t.getHandlerFunction() === 'onEdit' &&
            t.getTriggerSourceId() === ss.getId(),
        )) {
            ScriptApp.newTrigger('onEdit')
                .forSpreadsheet(ss)
                .onEdit()
                .create();
        }
    },

    // Methoden für Testzwecke und erweiterte Funktionalität
    _internal: {
        sheetCreator,
    },
};

export default SetupModule;