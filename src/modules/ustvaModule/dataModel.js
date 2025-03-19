// modules/ustvaModule/dataModel.js

/**
 * Erstellt ein leeres UStVA-Datenobjekt mit Nullwerten
 * @returns {Object} Leere UStVA-Datenstruktur
 */
function createEmptyUStVA() {
    // Optimierung: Struktur für bessere Performance
    return {
        // Einnahmen
        steuerpflichtige_einnahmen: 0,
        steuerfreie_inland_einnahmen: 0,
        steuerfreie_ausland_einnahmen: 0,

        // Ausgaben
        steuerpflichtige_ausgaben: 0,
        steuerfreie_inland_ausgaben: 0,
        steuerfreie_ausland_ausgaben: 0,

        // Eigenbelege
        eigenbelege_steuerpflichtig: 0,
        eigenbelege_steuerfrei: 0,

        // Steuer-Spezifika
        nicht_abzugsfaehige_vst: 0,

        // Umsatzsteuer
        ust_7: 0,
        ust_19: 0,

        // Vorsteuer
        vst_7: 0,
        vst_19: 0,
    };
}

export default {
    createEmptyUStVA,
};