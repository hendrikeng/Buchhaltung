// modules/bilanzModule/dataModel.js
/**
 * Erstellt eine leere Bilanz-Datenstruktur
 * @returns {Object} Leere Bilanz-Datenstruktur
 */
function createEmptyBilanz() {
    return {
        // Aktiva (Vermögenswerte)
        aktiva: {
            // Anlagevermögen
            sachanlagen: 0,                 // SKR04: 0400-0699, Sachanlagen
            immaterielleVermoegen: 0,       // SKR04: 0100-0199, Immaterielle Vermögensgegenstände
            finanzanlagen: 0,               // SKR04: 0700-0899, Finanzanlagen
            summeAnlagevermoegen: 0,        // Summe Anlagevermögen

            // Umlaufvermögen
            bankguthaben: 0,                // SKR04: 1200, Bank
            kasse: 0,                       // SKR04: 1210, Kasse
            forderungenLuL: 0,              // SKR04: 1300-1370, Forderungen aus Lieferungen und Leistungen
            vorraete: 0,                    // SKR04: 1400-1590, Vorräte
            summeUmlaufvermoegen: 0,        // Summe Umlaufvermögen

            // Rechnungsabgrenzung
            rechnungsabgrenzung: 0,         // SKR04: 1900-1990, Aktiver Rechnungsabgrenzungsposten

            // Gesamtsumme
            summeAktiva: 0                  // Summe aller Aktiva
        },

        // Passiva (Kapital und Schulden)
        passiva: {
            // Eigenkapital
            stammkapital: 0,                // SKR04: 2000, Gezeichnetes Kapital
            kapitalruecklagen: 0,           // SKR04: 2100, Kapitalrücklage
            gewinnvortrag: 0,               // SKR04: 2970, Gewinnvortrag
            verlustvortrag: 0,              // SKR04: 2978, Verlustvortrag
            jahresueberschuss: 0,           // Jahresüberschuss aus BWA
            summeEigenkapital: 0,           // Summe Eigenkapital

            // Verbindlichkeiten
            bankdarlehen: 0,                // SKR04: 3150, Verbindlichkeiten gegenüber Kreditinstituten
            gesellschafterdarlehen: 0,      // SKR04: 3300, Verbindlichkeiten gegenüber Gesellschaftern
            verbindlichkeitenLuL: 0,        // SKR04: 3200, Verbindlichkeiten aus Lieferungen und Leistungen
            steuerrueckstellungen: 0,       // SKR04: 3060, Steuerrückstellungen
            summeVerbindlichkeiten: 0,      // Summe Verbindlichkeiten

            // Rechnungsabgrenzung
            rechnungsabgrenzung: 0,         // SKR04: 3800-3990, Passiver Rechnungsabgrenzungsposten

            // Gesamtsumme
            summePassiva: 0                 // Summe aller Passiva
        }
    };
}

export default {
    createEmptyBilanz
};