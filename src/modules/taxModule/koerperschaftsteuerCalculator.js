// src/modules/taxModule/koerperschaftsteuerCalculator.js
import numberUtils from '../../utils/numberUtils.js';

/**
 * Modul zur Berechnung der Körperschaftsteuer und des Solidaritätszuschlags
 */
const koerperschaftsteuerCalculator = {
    /**
     * Berechnet die Körperschaftsteuer und den Solidaritätszuschlag
     * @param {Object} bilanzData - Die Bilanzdaten
     * @param {Object} config - Die Konfiguration
     * @returns {Object} - Berechnete KSt und Soli
     */
    calculateKoerperschaftsteuer(bilanzData, config) {
        // Gewinn aus Bilanz übernehmen
        const gewinnVorSteuern = bilanzData?.passiva?.jahresueberschuss || 0;

        // Verlustvortrag berücksichtigen
        const verlustvortrag = bilanzData?.passiva?.verlustvortrag || 0;

        // Bemessungsgrundlage berechnen
        let bemessungsgrundlage = gewinnVorSteuern;

        // Verlustvortrag anwenden (max. 1 Mio. EUR + 60% des darüber hinausgehenden Betrags)
        const verlustvortragBasis = Math.min(verlustvortrag, bemessungsgrundlage);
        let anrechenbareVerluste = 0;

        if (verlustvortragBasis > 0) {
            // Bis 1 Mio. EUR können Verluste vollständig angerechnet werden
            const basisVerlustvortrag = Math.min(1000000, verlustvortragBasis);

            // Darüber hinaus nur 60% anrechenbar
            const zusatzVerlustvortrag = Math.max(0, verlustvortragBasis - 1000000) * 0.6;

            anrechenbareVerluste = basisVerlustvortrag + zusatzVerlustvortrag;
            bemessungsgrundlage -= anrechenbareVerluste;
        }

        // Bei Holdings müssen wir den steuerfreien Anteil (95%) von Beteiligungserträgen berücksichtigen
        if (config.tax.isHolding) {
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

            // 95% der Beteiligungserträge sind steuerfrei
            const steuerfreierAnteil = beteiligungsertraege * (config.tax.holding.gewinnUebertragSteuerfrei / 100);

            // Nur 5% sind steuerpflichtig (bereits in der Bemessungsgrundlage enthalten)
            bemessungsgrundlage -= steuerfreierAnteil;
        }

        // Körperschaftsteuersatz aus Konfiguration holen (standardmäßig 15%)
        const kstSatz = config.tax.isHolding ?
            config.tax.holding.koerperschaftsteuer / 100 :
            config.tax.operative.koerperschaftsteuer / 100;

        // Körperschaftsteuer berechnen
        const koerperschaftsteuer = numberUtils.round(Math.max(0, bemessungsgrundlage) * kstSatz, 2);

        // Solidaritätszuschlag berechnen (5,5% der KSt)
        const soliSatz = config.tax.isHolding ?
            config.tax.holding.solidaritaetszuschlag / 100 :
            config.tax.operative.solidaritaetszuschlag / 100;

        const solidaritaetszuschlag = numberUtils.round(koerperschaftsteuer * soliSatz, 2);

        return {
            koerperschaftsteuer,
            solidaritaetszuschlag,
            details: {
                gewinnVorSteuern,
                verlustvortrag,
                anrechenbareVerluste,
                bemessungsgrundlage,
                kstSatz,
                soliSatz,
            },
        };
    },
};

export default koerperschaftsteuerCalculator;