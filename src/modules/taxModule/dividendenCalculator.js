// src/modules/taxModule/dividendenCalculator.js
import numberUtils from '../../utils/numberUtils.js';

/**
 * Modul zur Berechnung der Steuer auf Dividenden und Beteiligungserträge für eine Holding GmbH
 */
const dividendenCalculator = {
    /**
     * Berechnet die Steuer auf Dividenden und Beteiligungserträge für eine Holding GmbH
     * @param {Object} bilanzData - Die Bilanzdaten
     * @param {Object} config - Die Konfiguration
     * @returns {Object} - Berechnete Dividendensteuer-Daten
     */
    calculateDividendensteuer(bilanzData, config) {
        // Prüfen, ob Dividendensteuer berechnet werden soll (nur für Holding GmbH)
        if (!config.tax.isHolding) {
            return {
                beteiligungsertraege: 0,
                steuerfreierAnteil: 0,
                steuerpflichtigerAnteil: 0,
                koerperschaftsteuer: 0,
                solidaritaetszuschlag: 0,
                gesamtsteuer: 0,
                details: {
                    hinweis: 'Dividendensteuerberechnung entfällt für operative GmbH.',
                },
            };
        }

        // Gewinnübertrag aus operativer GmbH extrahieren
        const holdingSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Holding Transfers');
        let beteiligungsertraege = 0;

        if (holdingSheet && holdingSheet.getLastRow() > 1) {
            const data = holdingSheet.getDataRange().getValues();
            const headers = data[0];

            // Finde Indizes der relevanten Spalten
            const kategorieIndex = headers.findIndex(h => h === 'Kategorie');
            const betragIndex = headers.findIndex(h => h === 'Betrag');

            if (kategorieIndex !== -1 && betragIndex !== -1) {
                // Durchlaufe alle Zeilen und summiere Gewinnüberträge
                for (let i = 1; i < data.length; i++) {
                    if (data[i][kategorieIndex] === 'Gewinnübertrag') {
                        beteiligungsertraege += numberUtils.parseCurrency(data[i][betragIndex]);
                    }
                }
            }
        }

        // Steuerfreier Anteil der Beteiligungserträge (standardmäßig 95%)
        const steuerfreierProzentsatz = config.tax.holding.gewinnUebertragSteuerfrei;
        const steuerfreierAnteil = beteiligungsertraege * (steuerfreierProzentsatz / 100);

        // Steuerpflichtiger Anteil der Beteiligungserträge (standardmäßig 5%)
        const steuerpflichtigerProzentsatz = config.tax.holding.gewinnUebertragSteuerpflichtig;
        const steuerpflichtigerAnteil = beteiligungsertraege * (steuerpflichtigerProzentsatz / 100);

        // Körperschaftsteuer auf den steuerpflichtigen Anteil (15%)
        const kstSatz = config.tax.holding.koerperschaftsteuer / 100;
        const koerperschaftsteuer = steuerpflichtigerAnteil * kstSatz;

        // Solidaritätszuschlag auf die Körperschaftsteuer (5,5%)
        const soliSatz = config.tax.holding.solidaritaetszuschlag / 100;
        const solidaritaetszuschlag = koerperschaftsteuer * soliSatz;

        // Gesamtsteuer auf Beteiligungserträge
        const gesamtsteuer = koerperschaftsteuer + solidaritaetszuschlag;

        // Effektiver Steuersatz auf Beteiligungserträge
        const effektiverSteuersatz = beteiligungsertraege > 0 ?
            (gesamtsteuer / beteiligungsertraege) * 100 : 0;

        // Details für Nachvollziehbarkeit zurückgeben
        return {
            beteiligungsertraege,
            steuerfreierAnteil,
            steuerpflichtigerAnteil,
            koerperschaftsteuer,
            solidaritaetszuschlag,
            gesamtsteuer,
            effektiverSteuersatz,
            details: {
                beteiligungsertraege,
                steuerfreierProzentsatz,
                steuerpflichtigerProzentsatz,
                kstSatz,
                soliSatz,
            },
        };
    },

    /**
     * Berechnet die Abgeltungsteuer für Ausschüttungen an Gesellschafter
     * @param {number} ausschuettungsbetrag - Höhe der Ausschüttung
     * @param {Object} config - Die Konfiguration
     * @returns {Object} - Berechnete Abgeltungsteuer-Daten
     */
    calculateAbgeltungsteuer(ausschuettungsbetrag, config) {
        // Abgeltungsteuersatz (25%)
        const abgeltungsteuerSatz = 0.25;

        // Kirchensteuer (falls konfiguriert, standardmäßig 0%)
        const kirchensteuerSatz = config.tax.kirchensteuer ? config.tax.kirchensteuer / 100 : 0;

        // Solidaritätszuschlag (5,5% auf die Abgeltungsteuer)
        const soliSatz = config.tax.holding.solidaritaetszuschlag / 100;

        // Abgeltungsteuer berechnen
        const abgeltungsteuer = ausschuettungsbetrag * abgeltungsteuerSatz;

        // Solidaritätszuschlag berechnen
        const solidaritaetszuschlag = abgeltungsteuer * soliSatz;

        // Kirchensteuer berechnen (falls zutreffend)
        const kirchensteuer = abgeltungsteuer * kirchensteuerSatz;

        // Gesamtsteuer
        const gesamtsteuer = abgeltungsteuer + solidaritaetszuschlag + kirchensteuer;

        // Netto-Ausschüttung
        const nettoAusschuettung = ausschuettungsbetrag - gesamtsteuer;

        return {
            ausschuettungsbetrag,
            abgeltungsteuer,
            solidaritaetszuschlag,
            kirchensteuer,
            gesamtsteuer,
            nettoAusschuettung,
            details: {
                abgeltungsteuerSatz,
                soliSatz,
                kirchensteuerSatz,
            },
        };
    },

    /**
     * Vergleicht Abgeltungsteuer mit der Einkommensteuer (Teileinkünfteverfahren)
     * @param {number} ausschuettungsbetrag - Höhe der Ausschüttung
     * @param {number} einkommensteuersatz - Persönlicher ESt-Satz des Gesellschafters
     * @param {Object} config - Die Konfiguration
     * @returns {Object} - Vergleichsdaten
     */
    compareAbgeltungWithEinkommensteuer(ausschuettungsbetrag, einkommensteuersatz, config) {
        // 1. Abgeltungsteuer berechnen
        const abgeltungResult = this.calculateAbgeltungsteuer(ausschuettungsbetrag, config);

        // 2. Teileinkünfteverfahren berechnen (60% steuerpflichtig, 40% steuerfrei)
        const steuerpflichtigerAnteil = ausschuettungsbetrag * 0.6;
        const einkommensteuer = steuerpflichtigerAnteil * (einkommensteuersatz / 100);

        // Solidaritätszuschlag auf die Einkommensteuer
        const soliSatz = config.tax.holding.solidaritaetszuschlag / 100;
        const solidaritaetszuschlag = einkommensteuer * soliSatz;

        // Kirchensteuer (falls konfiguriert)
        const kirchensteuerSatz = config.tax.kirchensteuer ? config.tax.kirchensteuer / 100 : 0;
        const kirchensteuer = einkommensteuer * kirchensteuerSatz;

        // Gesamtsteuer Teileinkünfteverfahren
        const gesamtsteuerTEV = einkommensteuer + solidaritaetszuschlag + kirchensteuer;

        // Netto-Ausschüttung Teileinkünfteverfahren
        const nettoAusschuettungTEV = ausschuettungsbetrag - gesamtsteuerTEV;

        // Steuerersparnis durch Teileinkünfteverfahren
        const steuerersparnis = abgeltungResult.gesamtsteuer - gesamtsteuerTEV;

        // Empfehlung basierend auf der günstigeren Methode
        const empfehlung = steuerersparnis > 0 ?
            'Teileinkünfteverfahren (Antrag auf Regelbesteuerung)' :
            'Abgeltungsteuer';

        return {
            abgeltungsteuer: {
                steuersatz: 25,
                steuer: abgeltungResult.gesamtsteuer,
                netto: abgeltungResult.nettoAusschuettung,
            },
            teileinkuenfteverfahren: {
                einkommensteuersatz: einkommensteuersatz,
                steuerpflichtigerAnteil: steuerpflichtigerAnteil,
                steuer: gesamtsteuerTEV,
                netto: nettoAusschuettungTEV,
            },
            steuerersparnis,
            empfehlung,
        };
    },
};

export default dividendenCalculator;