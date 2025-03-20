// src/modules/taxModule/index.js
import gewerbesteuerCalculator from './gewerbesteuerCalculator.js';
import koerperschaftsteuerCalculator from './koerperschaftsteuerCalculator.js';
import dividendenCalculator from './dividendenCalculator.js';
import optimierungsCalculator from './optimierungsCalculator.js';
import formatter from './formatter.js';

/**
 * Modul für Steuerberechnungen (Gewerbesteuer, Körperschaftsteuer, etc.)
 * Optimierte Implementierung für korrekte Steuerberechnung nach deutschem Steuerrecht
 */
const TaxModule = {
    /**
     * Berechnet die Gewerbesteuer für eine operative GmbH
     * @param {Object} bilanzData - Die Bilanzdaten
     * @param {Object} config - Die Konfiguration
     * @returns {Object} - Berechnete Gewerbesteuer-Daten
     */
    calculateGewerbesteuer(bilanzData, config) {
        return gewerbesteuerCalculator.calculateGewerbesteuer(bilanzData, config);
    },

    /**
     * Berechnet die Körperschaftsteuer für eine GmbH (Holding oder operativ)
     * @param {Object} bilanzData - Die Bilanzdaten
     * @param {Object} config - Die Konfiguration
     * @returns {Object} - Berechnete Körperschaftsteuer-Daten
     */
    calculateKoerperschaftsteuer(bilanzData, config) {
        return koerperschaftsteuerCalculator.calculateKoerperschaftsteuer(bilanzData, config);
    },

    /**
     * Berechnet die Steuer auf Dividenden und Beteiligungserträge für eine Holding GmbH
     * @param {Object} bilanzData - Die Bilanzdaten
     * @param {Object} config - Die Konfiguration
     * @returns {Object} - Berechnete Dividendensteuer-Daten
     */
    calculateDividendensteuer(bilanzData, config) {
        return dividendenCalculator.calculateDividendensteuer(bilanzData, config);
    },

    /**
     * Berechnet und vergleicht verschiedene Steueroptimierungsszenarien
     * @param {Object} bilanzData - Die Bilanzdaten
     * @param {Object} config - Die Konfiguration
     * @returns {Object} - Berechnete Optimierungsszenarien
     */
    calculateOptimierungsszenarien(bilanzData, config) {
        return optimierungsCalculator.calculateOptimierungsszenarien(bilanzData, config);
    },

    /**
     * Erstellt eine Steuerberechnung für alle relevanten Steuerarten
     * @param {Object} bilanzData - Die Bilanzdaten
     * @param {Object} config - Die Konfiguration
     * @returns {Object} - Berechnete Steuern mit Details
     */
    calculateAllTaxes(bilanzData, config) {
        // Ergebnis-Objekt für alle Steuerberechnungen
        const taxResults = {
            gewerbesteuer: null,
            koerperschaftsteuer: null,
            solidaritaetszuschlag: null,
            dividendensteuer: null,
            optimierung: null,
            gesamtsteuerbelastung: 0,
            effektiverSteuersatz: 0,
        };

        // 1. Körperschaftsteuer (für beide GmbH-Typen relevant)
        const kstResults = this.calculateKoerperschaftsteuer(bilanzData, config);
        taxResults.koerperschaftsteuer = kstResults.koerperschaftsteuer;
        taxResults.solidaritaetszuschlag = kstResults.solidaritaetszuschlag;

        // 2. Gewerbesteuer (nur für operative GmbH)
        if (!config.tax.isHolding) {
            const gewStResults = this.calculateGewerbesteuer(bilanzData, config);
            taxResults.gewerbesteuer = gewStResults.gewerbesteuer;
        }

        // 3. Dividenden und Beteiligungserträge (nur für Holding GmbH)
        if (config.tax.isHolding) {
            const divResults = this.calculateDividendensteuer(bilanzData, config);
            taxResults.dividendensteuer = divResults;
        }

        // 4. Optimierungsszenarien (für beide GmbH-Typen)
        taxResults.optimierung = this.calculateOptimierungsszenarien(bilanzData, config);

        // 5. Gesamtsteuerbelastung ermitteln
        taxResults.gesamtsteuerbelastung = (taxResults.koerperschaftsteuer || 0) +
            (taxResults.solidaritaetszuschlag || 0) +
            (taxResults.gewerbesteuer || 0) +
            (taxResults.dividendensteuer?.gesamtsteuer || 0);

        // 6. Effektiven Steuersatz berechnen
        const gewinnVorSteuern = bilanzData?.passiva?.jahresueberschuss || 0;
        if (gewinnVorSteuern > 0) {
            taxResults.effektiverSteuersatz = (taxResults.gesamtsteuerbelastung / gewinnVorSteuern) * 100;
        }

        return taxResults;
    },

    /**
     * Erstellt eine Steuerübersicht als Sheet
     * @param {Spreadsheet} ss - Das Spreadsheet
     * @param {Object} taxData - Die berechneten Steuerdaten
     * @param {Object} bilanzData - Die Bilanzdaten
     * @param {Object} config - Die Konfiguration
     * @returns {boolean} - true bei Erfolg, false bei Fehler
     */
    generateTaxReport(ss, taxData, bilanzData, config) {
        return formatter.generateTaxReport(ss, taxData, bilanzData, config);
    },
};

export default TaxModule;