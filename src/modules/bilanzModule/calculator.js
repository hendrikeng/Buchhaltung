// modules/bilanzModule/calculator.js
import numberUtils from '../../utils/numberUtils.js';

/**
 * Ermittelt die Summe der Gesellschafterdarlehen
 * @param {Spreadsheet} ss - Das Spreadsheet
 * @param {Object} gesellschafterCols - Spaltenkonfiguration
 * @param {Object} config - Die Konfiguration
 * @returns {number} Summe der Gesellschafterdarlehen
 */
function getDarlehensumme(ss, gesellschafterCols, config) {
    let darlehenSumme = 0;

    const gesellschafterSheet = ss.getSheetByName("Gesellschafterkonto");
    if (gesellschafterSheet) {
        const data = gesellschafterSheet.getDataRange().getValues();

        // Überschrift überspringen
        for (let i = 1; i < data.length; i++) {
            const row = data[i];
            // Prüfen, ob es sich um ein Gesellschafterdarlehen handelt
            if (row[gesellschafterCols.kategorie - 1] &&
                row[gesellschafterCols.kategorie - 1].toString().toLowerCase() === "gesellschafterdarlehen") {
                darlehenSumme += numberUtils.parseCurrency(row[gesellschafterCols.betrag - 1] || 0);
            }
        }
    }

    return darlehenSumme;
}

/**
 * Ermittelt die Summe der Steuerrückstellungen
 * @param {Spreadsheet} ss - Das Spreadsheet
 * @param {Object} ausgabenCols - Spaltenkonfiguration
 * @param {Object} config - Die Konfiguration
 * @returns {number} Summe der Steuerrückstellungen
 */
function getSteuerRueckstellungen(ss, ausgabenCols, config) {
    let steuerRueckstellungen = 0;

    const ausSheet = ss.getSheetByName("Ausgaben");
    if (ausSheet) {
        const data = ausSheet.getDataRange().getValues();

        // Array mit Steuerrückstellungs-Kategorien
        const steuerKategorien = [
            "Gewerbesteuerrückstellungen",
            "Körperschaftsteuer",
            "Solidaritätszuschlag",
            "Sonstige Steuerrückstellungen"
        ];

        // Überschrift überspringen
        for (let i = 1; i < data.length; i++) {
            const row = data[i];
            const category = row[ausgabenCols.kategorie - 1]?.toString().trim() || "";

            if (steuerKategorien.includes(category)) {
                steuerRueckstellungen += numberUtils.parseCurrency(row[ausgabenCols.nettobetrag - 1] || 0);
            }
        }
    }

    return steuerRueckstellungen;
}

/**
 * Berechnet die Summen für die Bilanz
 * @param {Object} bilanzData - Die Bilanzdaten
 */
function calculateBilanzSummen(bilanzData) {
    const { aktiva, passiva } = bilanzData;

    // Summen für Aktiva
    aktiva.summeAnlagevermoegen = numberUtils.round(
        aktiva.sachanlagen +
        aktiva.immaterielleVermoegen +
        aktiva.finanzanlagen,
        2
    );

    aktiva.summeUmlaufvermoegen = numberUtils.round(
        aktiva.bankguthaben +
        aktiva.kasse +
        aktiva.forderungenLuL +
        aktiva.vorraete,
        2
    );

    aktiva.summeAktiva = numberUtils.round(
        aktiva.summeAnlagevermoegen +
        aktiva.summeUmlaufvermoegen +
        aktiva.rechnungsabgrenzung,
        2
    );

    // Summen für Passiva
    passiva.summeEigenkapital = numberUtils.round(
        passiva.stammkapital +
        passiva.kapitalruecklagen +
        passiva.gewinnvortrag -
        passiva.verlustvortrag +
        passiva.jahresueberschuss,
        2
    );

    passiva.summeVerbindlichkeiten = numberUtils.round(
        passiva.bankdarlehen +
        passiva.gesellschafterdarlehen +
        passiva.verbindlichkeitenLuL +
        passiva.steuerrueckstellungen,
        2
    );

    passiva.summePassiva = numberUtils.round(
        passiva.summeEigenkapital +
        passiva.summeVerbindlichkeiten +
        passiva.rechnungsabgrenzung,
        2
    );
}

export default {
    getDarlehensumme,
    getSteuerRueckstellungen,
    calculateBilanzSummen
};