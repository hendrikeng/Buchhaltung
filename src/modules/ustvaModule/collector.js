// modules/ustvaModule/collector.js
import globalCache from '../../utils/cacheUtils.js';
import dataModel from './dataModel.js';
import calculator from './calculator.js';

/**
 * Erfasst alle UStVA-Daten aus den verschiedenen Sheets
 * @param {Object} config - Die Konfiguration
 * @returns {Object|null} UStVA-Daten nach Monaten oder null bei Fehler
 */
function collectUStVAData(config) {
    // Prüfen, ob der Cache aktuelle Daten enthält
    if (globalCache.has('computed', 'ustva')) {
        return globalCache.get('computed', 'ustva');
    }

    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();

        // Benötigte Sheets abrufen
        const revenueSheet = ss.getSheetByName("Einnahmen");
        const expenseSheet = ss.getSheetByName("Ausgaben");
        const eigenSheet = ss.getSheetByName("Eigenbelege");

        // Prüfen, ob die wichtigsten Sheets vorhanden sind
        if (!revenueSheet || !expenseSheet) {
            console.error("Fehlende Blätter: 'Einnahmen' oder 'Ausgaben' nicht gefunden");
            return null;
        }

        // Daten aus den Sheets laden
        const revenueData = revenueSheet.getDataRange().getValues();
        const expenseData = expenseSheet.getDataRange().getValues();
        const eigenData = eigenSheet ? eigenSheet.getDataRange().getValues() : [];

        // Leere UStVA-Datenstruktur für alle Monate erstellen
        const ustvaData = Object.fromEntries(
            Array.from({length: 12}, (_, i) => [i + 1, dataModel.createEmptyUStVA()])
        );

        // Daten für jede Zeile verarbeiten
        revenueData.slice(1).forEach(row => calculator.processUStVARow(row, ustvaData, true, false, config));
        expenseData.slice(1).forEach(row => calculator.processUStVARow(row, ustvaData, false, false, config));
        if (eigenData.length) {
            eigenData.slice(1).forEach(row => calculator.processUStVARow(row, ustvaData, false, true, config));
        }

        // Daten cachen
        globalCache.set('computed', 'ustva', ustvaData);

        return ustvaData;
    } catch (e) {
        console.error("Fehler beim Sammeln der UStVA-Daten:", e);
        return null;
    }
}

export default {
    collectUStVAData
};