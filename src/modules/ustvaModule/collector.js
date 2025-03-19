// modules/ustvaModule/collector.js
import globalCache from '../../utils/cacheUtils.js';
import dataModel from './dataModel.js';
import calculator from './calculator.js';

/**
 * Erfasst alle UStVA-Daten aus den verschiedenen Sheets mit optimierter Batch-Verarbeitung
 * @param {Object} config - Die Konfiguration
 * @returns {Object|null} UStVA-Daten nach Monaten oder null bei Fehler
 */
function collectUStVAData(config) {
    // Prüfen, ob der Cache aktuelle Daten enthält
    if (globalCache.has('computed', 'ustva')) {
        return globalCache.get('computed', 'ustva');
    }

    try {
        console.log('Collecting UStVA data...');
        const ss = SpreadsheetApp.getActiveSpreadsheet();

        // Optimierte Struktur für Sheets und Verarbeitung
        const sheetConfig = [
            { name: 'Einnahmen', type: 'einnahmen', isIncome: true, isEigen: false, required: true },
            { name: 'Ausgaben', type: 'ausgaben', isIncome: false, isEigen: false, required: true },
            { name: 'Eigenbelege', type: 'eigenbelege', isIncome: false, isEigen: true, required: false },
        ];

        // Alle benötigten Sheets in einem Batch laden
        const sheets = {};

        for (const cfg of sheetConfig) {
            sheets[cfg.type] = ss.getSheetByName(cfg.name);

            if (cfg.required && !sheets[cfg.type]) {
                console.error(`Fehlendes Blatt: '${cfg.name}' nicht gefunden`);
                return null;
            }
        }

        // Leere UStVA-Datenstruktur für alle Monate erstellen
        const ustvaData = Object.fromEntries(
            Array.from({length: 12}, (_, i) => [i + 1, dataModel.createEmptyUStVA()]),
        );

        // Optimierung: Alle Daten in einem Batch laden und verarbeiten
        for (const cfg of sheetConfig) {
            const sheet = sheets[cfg.type];
            if (!sheet) continue;

            const lastRow = sheet.getLastRow();
            if (lastRow <= 1) continue; // Nur Header, keine Daten

            console.log(`Processing ${cfg.name} for UStVA...`);

            // Alle Daten in einem API-Call laden
            const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

            // Daten verarbeiten
            data.forEach(row => {
                calculator.processUStVARow(row, ustvaData, cfg.isIncome, cfg.isEigen, config);
            });
        }

        // Daten in Cache speichern
        globalCache.set('computed', 'ustva', ustvaData);
        console.log('UStVA data collection complete');

        return ustvaData;
    } catch (e) {
        console.error('Fehler beim Sammeln der UStVA-Daten:', e);
        return null;
    }
}

export default {
    collectUStVAData,
};