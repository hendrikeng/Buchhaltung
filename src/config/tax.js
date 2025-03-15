/**
 * Steuerliche Einstellungen für die Buchhaltungsanwendung
 */
export default {
    defaultMwst: 19,
    allowedMwst: [0, 7, 19],  // Erlaubte MwSt-Sätze
    stammkapital: 25000,
    year: 2021,  // Geschäftsjahr
    isHolding: false, // true bei Holding

    // Holding-spezifische Steuersätze
    holding: {
        gewerbesteuer: 470,  // Angepasst an lokalen Hebesatz
        koerperschaftsteuer: 15,
        solidaritaetszuschlag: 5.5,
        gewinnUebertragSteuerfrei: 95,  // % der Beteiligungserträge steuerfrei
        gewinnUebertragSteuerpflichtig: 5,  // % der Beteiligungserträge steuerpflichtig
    },

    // Operative GmbH Steuersätze
    operative: {
        gewerbesteuer: 470,  // Angepasst an lokalen Hebesatz
        koerperschaftsteuer: 15,
        solidaritaetszuschlag: 5.5,
        gewinnUebertragSteuerfrei: 0,
        gewinnUebertragSteuerpflichtig: 100,
    },
};