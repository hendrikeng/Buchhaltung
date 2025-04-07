// modules/bwaModule/dataModel.js (Optimized)
/**
 * Creates an empty BWA data object with zero values
 * Optimized structure for better performance
 * @returns {Object} Empty BWA data structure
 */
function createEmptyBWA() {
    // Initialize with zero values
    return {
        // Group 1: Business revenue (Income)
        umsatzerloese: 0,
        provisionserloese: 0,
        steuerfreieInlandEinnahmen: 0,
        steuerfreieAuslandEinnahmen: 0,
        sonstigeErtraege: 0,
        vermietung: 0,
        zuschuesse: 0,
        waehrungsgewinne: 0,
        anlagenabgaenge: 0,
        gesamtErloese: 0,

        // Group 2: Material costs & Purchases
        wareneinsatz: 0,
        fremdleistungen: 0,
        rohHilfsBetriebsstoffe: 0,
        gesamtWareneinsatz: 0,

        // Group 3: Operating expenses (Overhead costs)
        bruttoLoehne: 0,
        sozialeAbgaben: 0,
        sonstigePersonalkosten: 0,
        werbungMarketing: 0,
        reisekosten: 0,
        versicherungen: 0,
        telefonInternet: 0,
        buerokosten: 0,
        fortbildungskosten: 0,
        kfzKosten: 0,
        itKosten: 0, // Added for IT costs
        mieteNebenkosten: 0, // Added for rent and utilities
        sonstigeAufwendungen: 0,
        gesamtBetriebsausgaben: 0,

        // Group 4: Depreciation & Interest
        abschreibungenMaschinen: 0,
        abschreibungenBueromaterial: 0,
        abschreibungenImmateriell: 0,
        zinsenBank: 0,
        zinsenGesellschafter: 0,
        leasingkosten: 0,
        gesamtAbschreibungenZinsen: 0,

        // Group 5: Special items (Capital movements)
        eigenkapitalveraenderungen: 0,
        gesellschafterdarlehen: 0,
        ausschuettungen: 0,
        gesamtBesonderePosten: 0,

        // Group 6: Provisions
        steuerrueckstellungen: 0,
        rueckstellungenSonstige: 0,
        gesamtRueckstellungenTransfers: 0,

        // Group 7: EBIT
        ebit: 0,

        // Group 8: Taxes & VAT
        umsatzsteuer: 0,
        vorsteuer: 0,
        nichtAbzugsfaehigeVSt: 0,
        koerperschaftsteuer: 0,
        solidaritaetszuschlag: 0,
        gewerbesteuer: 0,
        gewerbesteuerRueckstellungen: 0,
        sonstigeSteuerrueckstellungen: 0,
        steuerlast: 0,

        // Group 9: Annual net profit/loss
        gewinnNachSteuern: 0,

        // Own receipts (for aggregation)
        eigenbelegeSteuerfrei: 0,
        eigenbelegeSteuerpflichtig: 0,
    };
}

export default {
    createEmptyBWA,
};