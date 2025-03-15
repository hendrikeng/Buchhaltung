// modules/ustvaModule/dataModel.js
/**
 * Erstellt ein leeres UStVA-Datenobjekt mit Nullwerten
 * @returns {Object} Leere UStVA-Datenstruktur
 */
function createEmptyUStVA() {
    return {
        steuerpflichtige_einnahmen: 0,
        steuerfreie_inland_einnahmen: 0,
        steuerfreie_ausland_einnahmen: 0,
        steuerpflichtige_ausgaben: 0,
        steuerfreie_inland_ausgaben: 0,
        steuerfreie_ausland_ausgaben: 0,
        eigenbelege_steuerpflichtig: 0,
        eigenbelege_steuerfrei: 0,
        nicht_abzugsfaehige_vst: 0,
        ust_7: 0,
        ust_19: 0,
        vst_7: 0,
        vst_19: 0
    };
}

export default {
    createEmptyUStVA
};