// src/modules/taxModule/gewerbesteuerCalculator.js
import numberUtils from '../../utils/numberUtils.js';

/**
 * Modul zur Berechnung der Gewerbesteuer für eine operative GmbH
 */
const gewerbesteuerCalculator = {
    /**
     * Berechnet die Gewerbesteuer für eine operative GmbH
     * @param {Object} bilanzData - Die Bilanzdaten
     * @param {Object} config - Die Konfiguration
     * @returns {Object} - Berechnete Gewerbesteuer-Daten
     */
    calculateGewerbesteuer(bilanzData, config) {
        // Prüfen, ob Gewerbesteuer berechnet werden soll (nur für operative GmbH)
        if (config.tax.isHolding) {
            return { gewerbesteuer: 0, details: { hinweis: 'Holding-GmbH ist von der Gewerbesteuer befreit.' } };
        }

        // Gewinn aus Bilanz übernehmen
        const gewinnVorSteuern = bilanzData?.passiva?.jahresueberschuss || 0;

        // Hinzurechnungen ermitteln
        const hinzurechnungen = this._calculateHinzurechnungen(bilanzData, config);

        // Kürzungen ermitteln
        const kuerzungen = this._calculateKuerzungen(bilanzData, config);

        // Gewerbeertrag berechnen
        const gewerbeertrag = gewinnVorSteuern + hinzurechnungen.summe - kuerzungen.summe;

        // Freibetrag berücksichtigen (24.500 EUR für natürliche Personen und Personengesellschaften)
        // Bei Kapitalgesellschaften (GmbH) gibt es keinen Freibetrag
        const freibetrag = 0;
        const gewerbeertragNachFreibetrag = Math.max(0, gewerbeertrag - freibetrag);

        // Gewerbesteuermessbetrag berechnen (3,5%)
        const steuermesszahl = 3.5 / 100;
        const messbetrag = gewerbeertragNachFreibetrag * steuermesszahl;

        // Hebesatz aus Konfiguration holen (z.B. 400%)
        const hebesatz = (config.tax.operative.gewerbesteuer || 400) / 100;

        // Gewerbesteuer berechnen
        const gewerbesteuer = numberUtils.round(messbetrag * hebesatz, 2);

        // Details für Nachvollziehbarkeit zurückgeben
        return {
            gewerbesteuer,
            details: {
                gewinnVorSteuern,
                hinzurechnungen,
                kuerzungen,
                gewerbeertrag,
                freibetrag,
                gewerbeertragNachFreibetrag,
                steuermesszahl,
                messbetrag,
                hebesatz,
            },
        };
    },

    /**
     * Berechnet die Hinzurechnungen für die Gewerbesteuer
     * @param {Object} bilanzData - Die Bilanzdaten
     * @param {Object} config - Die Konfiguration
     * @returns {Object} - Berechnete Hinzurechnungen
     */
    _calculateHinzurechnungen(bilanzData, config) {
        // Ausgaben aus BWA holen (wenn verfügbar)
        const ausgabenSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Ausgaben');

        // Initialisiere Hinzurechnungen mit Standardwerten
        const hinzurechnungen = {
            miete: 0,
            zinsen: 0,
            leasing: 0,
            sonstige: 0,
            summe: 0,
        };

        // Wenn Ausgaben-Sheet existiert, extrahiere relevante Daten
        if (ausgabenSheet && ausgabenSheet.getLastRow() > 1) {
            const data = ausgabenSheet.getDataRange().getValues();
            const headers = data[0];

            // Finde Indizes der relevanten Spalten
            const kategorieIndex = headers.findIndex(h => h === 'Kategorie');
            const nettoIndex = headers.findIndex(h => h === 'Nettobetrag');

            if (kategorieIndex !== -1 && nettoIndex !== -1) {
                // Durchlaufe alle Ausgabenzeilen und summiere relevante Kategorien
                for (let i = 1; i < data.length; i++) {
                    const kategorie = data[i][kategorieIndex];
                    const betrag = numberUtils.parseCurrency(data[i][nettoIndex]);

                    // Hinzurechnung nach Kategorie
                    switch (kategorie) {
                    case 'Miete':
                        // 25% der Miete wird hinzugerechnet
                        hinzurechnungen.miete += betrag * 0.25;
                        break;
                    case 'Zinsen auf Bankdarlehen':
                    case 'Zinsen auf Gesellschafterdarlehen':
                        // 25% der Zinsen wird hinzugerechnet
                        hinzurechnungen.zinsen += betrag * 0.25;
                        break;
                    case 'Leasingkosten':
                        // 20% der Leasingkosten hinzurechnen
                        hinzurechnungen.leasing += betrag * 0.20;
                        break;
                    }
                }
            }
        }

        // Gesamtsumme der Hinzurechnungen berechnen
        hinzurechnungen.summe = hinzurechnungen.miete + hinzurechnungen.zinsen +
            hinzurechnungen.leasing + hinzurechnungen.sonstige;

        return hinzurechnungen;
    },

    /**
     * Berechnet die Kürzungen für die Gewerbesteuer
     * @param {Object} bilanzData - Die Bilanzdaten
     * @param {Object} config - Die Konfiguration
     * @returns {Object} - Berechnete Kürzungen
     */
    _calculateKuerzungen(bilanzData, config) {
        // Einnahmen aus BWA holen (wenn verfügbar)
        const einnahmenSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Einnahmen');

        // Initialisiere Kürzungen mit Standardwerten
        const kuerzungen = {
            grundbesitz: 0,
            dividenden: 0,
            sonstige: 0,
            summe: 0,
        };

        // Wenn Einnahmen-Sheet existiert, extrahiere relevante Daten
        if (einnahmenSheet && einnahmenSheet.getLastRow() > 1) {
            const data = einnahmenSheet.getDataRange().getValues();
            const headers = data[0];

            // Finde Indizes der relevanten Spalten
            const kategorieIndex = headers.findIndex(h => h === 'Kategorie');
            const nettoIndex = headers.findIndex(h => h === 'Nettobetrag');

            if (kategorieIndex !== -1 && nettoIndex !== -1) {
                // Durchlaufe alle Einnahmenzeilen und summiere relevante Kategorien
                for (let i = 1; i < data.length; i++) {
                    const kategorie = data[i][kategorieIndex];
                    const betrag = numberUtils.parseCurrency(data[i][nettoIndex]);

                    // Kürzungen nach Kategorie
                    if (kategorie === 'Erträge aus Vermietung/Verpachtung') {
                        kuerzungen.grundbesitz += betrag;
                    }
                }
            }
        }

        // Gesamtsumme der Kürzungen berechnen
        kuerzungen.summe = kuerzungen.grundbesitz + kuerzungen.dividenden + kuerzungen.sonstige;

        return kuerzungen;
    },
};

export default gewerbesteuerCalculator;