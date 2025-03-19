// modules/bilanzModule/calculator.js
import numberUtils from '../../utils/numberUtils.js';

/**
 * Ermittelt die Summe der Gesellschafterdarlehen mit optimierter Batch-Verarbeitung
 * @param {Spreadsheet} ss - Das Spreadsheet
 * @param {Object} gesellschafterCols - Spaltenkonfiguration
 * @param {Object} config - Die Konfiguration
 * @returns {number} Summe der Gesellschafterdarlehen
 */
function getDarlehensumme(ss, gesellschafterCols, config) {
    let darlehenSumme = 0;

    const gesellschafterSheet = ss.getSheetByName('Gesellschafterkonto');
    if (!gesellschafterSheet) {
        return 0;
    }

    // Optimierung: Daten in einem Batch laden
    const lastRow = gesellschafterSheet.getLastRow();
    if (lastRow <= 1) return 0; // Nur Header, keine Daten

    // Wir benötigen nur die Kategorie- und Betrag-Spalten
    const columnIndices = [
        gesellschafterCols.kategorie - 1,
        gesellschafterCols.betrag - 1,
    ];
    const maxColumn = Math.max(...columnIndices) + 1;

    const data = gesellschafterSheet.getRange(2, 1, lastRow - 1, maxColumn).getValues();

    // Optimierung: Einmal durch die Daten gehen und Darlehenssumme berechnen
    for (const row of data) {
        // Prüfen, ob es sich um ein Gesellschafterdarlehen handelt
        const kategorie = row[gesellschafterCols.kategorie - 1];
        if (kategorie && kategorie.toString().toLowerCase() === 'gesellschafterdarlehen') {
            darlehenSumme += numberUtils.parseCurrency(row[gesellschafterCols.betrag - 1] || 0);
        }
    }

    return darlehenSumme;
}

/**
 * Ermittelt die Summe der Steuerrückstellungen mit optimierter Batch-Verarbeitung
 * @param {Spreadsheet} ss - Das Spreadsheet
 * @param {Object} ausgabenCols - Spaltenkonfiguration
 * @param {Object} config - Die Konfiguration
 * @returns {number} Summe der Steuerrückstellungen
 */
function getSteuerRueckstellungen(ss, ausgabenCols, config) {
    let steuerRueckstellungen = 0;

    const ausSheet = ss.getSheetByName('Ausgaben');
    if (!ausSheet) {
        return 0;
    }

    // Optimierung: Daten in einem Batch laden
    const lastRow = ausSheet.getLastRow();
    if (lastRow <= 1) return 0; // Nur Header, keine Daten

    // Wir benötigen nur die Kategorie- und Betrag-Spalten
    const columnIndices = [
        ausgabenCols.kategorie - 1,
        ausgabenCols.nettobetrag - 1,
    ];
    const maxColumn = Math.max(...columnIndices) + 1;

    const data = ausSheet.getRange(2, 1, lastRow - 1, maxColumn).getValues();

    // Optimierung: Array mit Steuerrückstellungs-Kategorien in Set konvertieren für O(1) Lookup
    const steuerKategorien = new Set([
        'Gewerbesteuerrückstellungen',
        'Körperschaftsteuer',
        'Solidaritätszuschlag',
        'Sonstige Steuerrückstellungen',
    ]);

    // Optimierung: Einmal durch die Daten gehen und Steuerrückstellungen berechnen
    for (const row of data) {
        const category = row[ausgabenCols.kategorie - 1]?.toString().trim() || '';
        if (steuerKategorien.has(category)) {
            steuerRueckstellungen += numberUtils.parseCurrency(row[ausgabenCols.nettobetrag - 1] || 0);
        }
    }

    return steuerRueckstellungen;
}

/**
 * Berechnet die Summen für die Bilanz mit optimierter Berechnungslogik
 * @param {Object} bilanzData - Die Bilanzdaten
 */
function calculateBilanzSummen(bilanzData) {
    const { aktiva, passiva } = bilanzData;

    // Optimierung: Zwischenwerte für mehrfach verwendete Berechnungen
    const anlagevermoegen = aktiva.sachanlagen +
        aktiva.immaterielleVermoegen +
        aktiva.finanzanlagen;

    const umlaufvermoegen = aktiva.bankguthaben +
        aktiva.kasse +
        aktiva.forderungenLuL +
        aktiva.vorraete;

    const eigenkapital = passiva.stammkapital +
        passiva.kapitalruecklagen +
        passiva.gewinnvortrag -
        passiva.verlustvortrag +
        passiva.jahresueberschuss;

    const verbindlichkeiten = passiva.bankdarlehen +
        passiva.gesellschafterdarlehen +
        passiva.verbindlichkeitenLuL +
        passiva.steuerrueckstellungen;

    // Summen für Aktiva
    aktiva.summeAnlagevermoegen = numberUtils.round(anlagevermoegen, 2);
    aktiva.summeUmlaufvermoegen = numberUtils.round(umlaufvermoegen, 2);
    aktiva.summeAktiva = numberUtils.round(
        aktiva.summeAnlagevermoegen +
        aktiva.summeUmlaufvermoegen +
        aktiva.rechnungsabgrenzung,
        2,
    );

    // Summen für Passiva
    passiva.summeEigenkapital = numberUtils.round(eigenkapital, 2);
    passiva.summeVerbindlichkeiten = numberUtils.round(verbindlichkeiten, 2);
    passiva.summePassiva = numberUtils.round(
        passiva.summeEigenkapital +
        passiva.summeVerbindlichkeiten +
        passiva.rechnungsabgrenzung,
        2,
    );
}

export default {
    getDarlehensumme,
    getSteuerRueckstellungen,
    calculateBilanzSummen,
};