// modules/bwaModule/dataModel.js

/**
 * Creates an empty BWA data object with zero values
 * Follows DATEV standard structure for BWA (betriebswirtschaftliche Auswertung)
 * @returns {Object} Empty BWA data structure
 */
function createEmptyBWA() {
    return {
        // Group 1: Business income
        umsatzerloese: 0,            // Revenue from goods and services
        provisionserloese: 0,        // Commission income
        steuerfreieInlandEinnahmen: 0, // Tax-free domestic income
        steuerfreieAuslandEinnahmen: 0, // Tax-free foreign income
        innergemeinschaftlicheLieferungen: 0, // EU sales
        sonstigeErtraege: 0,         // Other operating income
        vermietung: 0,               // Rental income (if business-related)
        zuschuesse: 0,               // Subsidies
        waehrungsgewinne: 0,         // Foreign exchange gains
        anlagenabgaenge: 0,          // Gains from disposals of fixed assets
        steuerlicheKorrekturen: 0,   // Tax corrections (e.g., VAT refunds)

        // Group 2: Material expenses & Cost of goods
        wareneinsatz: 0,             // Cost of goods
        fremdleistungen: 0,          // External services
        rohHilfsBetriebsstoffe: 0,   // Raw materials and supplies

        // Group 3: Operating expenses
        bruttoLoehne: 0,             // Gross salaries and wages
        sozialeAbgaben: 0,           // Social security contributions
        sonstigePersonalkosten: 0,   // Other personnel expenses
        mieteNebenkosten: 0,         // Rent and ancillary costs
        werbungMarketing: 0,         // Advertising and marketing
        reisekosten: 0,              // Travel expenses
        versicherungen: 0,           // Insurance
        telefonInternet: 0,          // Phone and internet
        buerokosten: 0,              // Office expenses
        fortbildungskosten: 0,       // Training costs
        kfzKosten: 0,                // Vehicle expenses
        sonstigeAufwendungen: 0,     // Other operating expenses

        // Group 4: Depreciation & Interest
        abschreibungenMaschinen: 0,  // Depreciation on machinery
        abschreibungenBueromaterial: 0, // Depreciation on office equipment
        abschreibungenImmateriell: 0, // Depreciation on intangible assets
        zinsenBank: 0,               // Interest on bank loans
        zinsenGesellschafter: 0,     // Interest on shareholder loans
        leasingkosten: 0,            // Leasing costs

        // Group 5: EBIT
        ebit: 0,                     // Earnings before interest and taxes

        // Group 6: Net profit/loss
        gewinnNachSteuern: 0,        // Profit after taxes

        // Calculated sum fields (for aggregation)
        gesamtErloese: 0,                 // Total revenue
        gesamtWareneinsatz: 0,            // Total material costs
        gesamtBetriebsausgaben: 0,        // Total operating expenses
        gesamtAbschreibungenZinsen: 0,    // Total depreciation and interest

        // Internal tracking fields (not for display)
        nichtAbzugsfaehigeVSt: 0,      // Non-deductible VAT (for calculation)
        eigenbelegeSteuerfrei: 0,      // Tax-free own receipts (for calculation)
        eigenbelegeSteuerpflichtig: 0, // Taxable own receipts (for calculation)
    };
}

export default {
    createEmptyBWA,
};