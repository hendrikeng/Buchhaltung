// modules/bilanzModule/dataModel.js

/**
 * Creates an empty balance sheet structure with DATEV-compliant account structure (SKR04)
 * @returns {Object} DATEV-compliant balance sheet data structure
 */
function createEmptyBilanz() {
    return {
        // Aktiva (Vermögenswerte)
        aktiva: {
            // A. Anlagevermögen
            anlagevermoegen: {
                // I. Immaterielle Vermögensgegenstände
                immaterielleVermoegensgegenstaende: {
                    konzessionen: 0,               // 0100-0199, Konzessionen, Lizenzen und ähnliche Rechte
                    geschaeftswert: 0,             // 0200-0299, Geschäfts- oder Firmenwert
                    geleisteteAnzahlungen: 0,      // 0300-0399, Geleistete Anzahlungen
                    summe: 0,                      // Zwischensumme
                },

                // II. Sachanlagen
                sachanlagen: {
                    grundstuecke: 0,               // 0400-0499, Grundstücke, grundstücksgleiche Rechte
                    technischeAnlagenMaschinen: 0, // 0500-0599, Technische Anlagen und Maschinen
                    andereAnlagen: 0,              // 0600-0699, Andere Anlagen, Betriebs- und Geschäftsausstattung
                    geleisteteAnzahlungen: 0,      // 0700-0799, Geleistete Anzahlungen und Anlagen im Bau
                    summe: 0,                      // Zwischensumme
                },

                // III. Finanzanlagen
                finanzanlagen: {
                    anteileVerbundeneUnternehmen: 0, // 0800-0899, Anteile an verbundenen Unternehmen
                    ausleihungenVerbundene: 0,       // 0900-0999, Ausleihungen an verbundene Unternehmen
                    beteiligungen: 0,                // 1000-1099, Beteiligungen
                    ausleihungenBeteiligungen: 0,    // 1100-1199, Ausleihungen an Unternehmen mit Beteiligungsverhältnis
                    summe: 0,                        // Zwischensumme
                },

                summe: 0,                           // Gesamtsumme Anlagevermögen
            },

            // B. Umlaufvermögen
            umlaufvermoegen: {
                // I. Vorräte
                vorraete: {
                    rohHilfsBetriebsstoffe: 0,     // 1300-1399, Roh-, Hilfs- und Betriebsstoffe
                    unfertErzeugnisse: 0,          // 1400-1499, Unfertige Erzeugnisse
                    fertErzeugnisse: 0,            // 1500-1599, Fertige Erzeugnisse und Waren
                    geleisteteAnzahlungen: 0,      // 1600-1699, Geleistete Anzahlungen
                    summe: 0,                      // Zwischensumme
                },

                // II. Forderungen und sonstige Vermögensgegenstände
                forderungen: {
                    forderungenLuL: 0,             // 1700-1799, Forderungen aus Lieferungen und Leistungen
                    forderungenVerbundene: 0,      // 1800-1899, Forderungen gegen verbundene Unternehmen
                    forderungenBeteiligungen: 0,   // 1900-1999, Forderungen gegen Unternehmen mit Beteiligungsverhältnis
                    sonstigeVermoegensgegenstaende: 0, // 2000-2099, Sonstige Vermögensgegenstände
                    summe: 0,                      // Zwischensumme
                },

                // III. Wertpapiere
                wertpapiere: {
                    anteileVerbundeneUnternehmen: 0, // 2100-2199, Anteile an verbundenen Unternehmen
                    sonstigeWertpapiere: 0,          // 2200-2299, Sonstige Wertpapiere
                    summe: 0,                        // Zwischensumme
                },

                // IV. Kassenbestand, Bundesbankguthaben, Guthaben bei Kreditinstituten
                liquideMittel: {
                    kassenbestand: 0,              // 1000-1099, Kasse
                    bankguthaben: 0,               // 1200-1299, Bankguthaben
                    summe: 0,                      // Zwischensumme
                },

                summe: 0,                           // Gesamtsumme Umlaufvermögen
            },

            // C. Rechnungsabgrenzungsposten
            rechnungsabgrenzungsposten: 0,          // 2300-2399, Aktive Rechnungsabgrenzungsposten

            // D. Aktive latente Steuern
            aktiveLatenteSteuern: 0,                // 2400-2499, Aktive latente Steuern

            // E. Aktiver Unterschiedsbetrag aus Vermögensverrechnung
            aktiverUnterschiedsbetrag: 0,           // 2500-2599, Aktiver Unterschiedsbetrag aus Vermögensverrechnung

            // Summe Aktiva
            summe: 0,                               // Bilanzsumme Aktiva
        },

        // Passiva (Kapital und Schulden)
        passiva: {
            // A. Eigenkapital
            eigenkapital: {
                gezeichnetesKapital: 0,            // 2700-2799, Gezeichnetes Kapital
                kapitalruecklage: 0,               // 2800-2899, Kapitalrücklage
                gewinnruecklagen: {
                    gesetzlicheRuecklage: 0,       // 2900-2949, Gesetzliche Rücklage
                    ruecklageAnteileGesellschaft: 0, // 2950-2959, Rücklage für Anteile an herrschendem Unternehmen
                    satzungsmaessigeRuecklagen: 0, // 2960-2969, Satzungsmäßige Rücklagen
                    andereGewinnruecklagen: 0,     // 2970-2979, Andere Gewinnrücklagen
                    summe: 0,                      // Zwischensumme
                },
                gewinnvortrag: 0,                 // 2980-2989, Gewinnvortrag
                jahresueberschuss: 0,             // 2990-2999, Jahresüberschuss/Jahresfehlbetrag
                summe: 0,                          // Gesamtsumme Eigenkapital
            },

            // B. Rückstellungen
            rueckstellungen: {
                pensionsaehnlicheRueckstellungen: 0, // 3000-3099, Rückstellungen für Pensionen und ähnliche Verpflichtungen
                steuerrueckstellungen: 0,         // 3100-3199, Steuerrückstellungen
                sonstigeRueckstellungen: 0,       // 3200-3299, Sonstige Rückstellungen
                summe: 0,                          // Gesamtsumme Rückstellungen
            },

            // C. Verbindlichkeiten
            verbindlichkeiten: {
                anleihen: 0,                      // 3300-3399, Anleihen
                verbindlichkeitenKreditinstitute: 0, // 3400-3499, Verbindlichkeiten gegenüber Kreditinstituten
                erhalteneAnzahlungen: 0,          // 3500-3599, Erhaltene Anzahlungen auf Bestellungen
                verbindlichkeitenLuL: 0,          // 3600-3699, Verbindlichkeiten aus Lieferungen und Leistungen
                verbindlichkeitenWechsel: 0,      // 3700-3799, Verbindlichkeiten aus Wechseln
                verbindlichkeitenVerbundene: 0,   // 3800-3899, Verbindlichkeiten gegenüber verbundenen Unternehmen
                verbindlichkeitenBeteiligungen: 0, // 3900-3999, Verbindlichkeiten gegenüber Unternehmen mit Beteiligungsverhältnis
                sonstigeVerbindlichkeiten: 0,     // 4000-4099, Sonstige Verbindlichkeiten
                summe: 0,                          // Gesamtsumme Verbindlichkeiten
            },

            // D. Rechnungsabgrenzungsposten
            rechnungsabgrenzungsposten: 0,         // 4100-4199, Passive Rechnungsabgrenzungsposten

            // E. Passive latente Steuern
            passiveLatenteSteuern: 0,              // 4200-4299, Passive latente Steuern

            // Summe Passiva
            summe: 0,                               // Bilanzsumme Passiva
        },
    };
}

export default {
    createEmptyBilanz,
};