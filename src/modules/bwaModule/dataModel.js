// modules/bwaModule/dataModel.js

/**
 * Creates an empty BWA data object with DATEV-compliant structure
 * @returns {Object} Empty BWA data structure
 */
function createEmptyBWA() {
    return {
        // 1. Betriebserlöse (Einnahmen)
        erloeseLieferungenLeistungen: 0,
        provisionserloese: 0,
        steuerfreieInlandErloese: 0,
        steuerfreieAuslandErloese: 0,
        innergemeinschaftlicheLieferungen: 0,
        sonstigeBetrieblicheErtraege: 0,
        ertraegeVermietungVerpachtung: 0,
        ertraegeZuschuesse: 0,
        ertraegeKursgewinne: 0,
        ertraegeAnlagenabgaenge: 0,
        betriebserloese_gesamt: 0,

        // 2. Materialaufwand & Wareneinsatz
        wareneinsatz: 0,
        bezogeneLeistungen: 0,
        rohHilfsBetriebsstoffe: 0,
        materialaufwand_gesamt: 0,

        // 3. Personalaufwand
        loehneGehaelter: 0,
        sozialeAbgaben: 0,
        sonstigePersonalkosten: 0,
        personalaufwand_gesamt: 0,

        // 4. Sonstige betriebliche Aufwendungen
        provisionszahlungenDritte: 0,
        itKosten: 0,
        mieteLeasing: 0,
        werbungMarketing: 0,
        reisekosten: 0,
        versicherungen: 0,
        telefonInternet: 0,
        buerokosten: 0,
        fortbildungskosten: 0,
        kfzKosten: 0,
        beitraegeAbgaben: 0,
        bewirtungskosten: 0,
        sonstigeBetrieblicheAufwendungen: 0,
        sonstigeAufwendungen_gesamt: 0,

        // 5. Abschreibungen und Zinsen
        abschreibungenSachanlagen: 0,
        abschreibungenImmaterielleVG: 0,
        zinsenBankdarlehen: 0,
        zinsenGesellschafterdarlehen: 0,
        leasingzinsen: 0,
        abschreibungenZinsen_gesamt: 0,

        // 6. Betriebsergebnis (EBIT)
        ebit: 0,

        // 7. Jahresergebnis
        jahresergebnis: 0,

        // Additional fields for calculation
        umsatzsteuer: 0,
        vorsteuer: 0,
        nichtAbzugsfaehigeVSt: 0,
        koerperschaftsteuer: 0,
        solidaritaetszuschlag: 0,
        gewerbesteuer: 0,
        steuerlast: 0,
    };
}

export default {
    createEmptyBWA,
};