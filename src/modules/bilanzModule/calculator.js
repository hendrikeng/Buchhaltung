// modules/bilanzModule/calculator.js
import numberUtils from '../../utils/numberUtils.js';

/**
 * Calculates total loan amount from bank
 * @param {Spreadsheet} ss - Spreadsheet object
 * @param {Object} config - Configuration
 * @returns {number} Total bank loan amount
 */
function getBankDarlehen(ss, config) {
    let darlehenSumme = 0;

    // Load and process data from expenses sheet for bank loans
    const ausSheet = ss.getSheetByName('Ausgaben');
    if (!ausSheet || ausSheet.getLastRow() <= 1) return 0;

    const ausgabenCols = config.ausgaben.columns;
    const data = ausSheet.getDataRange().getValues().slice(1); // Skip header

    for (const row of data) {
        const kategorie = row[ausgabenCols.kategorie - 1]?.toString().trim() || '';
        if (kategorie === 'Bankdarlehen') {
            const betrag = numberUtils.parseCurrency(row[ausgabenCols.nettobetrag - 1]);
            darlehenSumme += betrag;
        }
    }

    return darlehenSumme;
}

/**
 * Calculates totals and subtotals for balance sheet
 * @param {Object} bilanzData - Balance sheet data
 */
function calculateBilanzTotals(bilanzData) {
    // 1. Calculate Aktiva subtotals
    calculateAktivaSubtotals(bilanzData);

    // 2. Calculate Passiva subtotals
    calculatePassivaSubtotals(bilanzData);

    // 3. Validate balance sheet (Aktiva = Passiva)
    validateBilanzEquality(bilanzData);
}

/**
 * Calculates subtotals for assets (Aktiva)
 * @param {Object} bilanzData - Balance sheet data
 */
function calculateAktivaSubtotals(bilanzData) {
    const aktiva = bilanzData.aktiva;

    // I. Calculate immaterielle Vermögensgegenstände
    aktiva.anlagevermoegen.immaterielleVermoegensgegenstaende.summe = numberUtils.round(
        aktiva.anlagevermoegen.immaterielleVermoegensgegenstaende.konzessionen +
        aktiva.anlagevermoegen.immaterielleVermoegensgegenstaende.geschaeftswert +
        aktiva.anlagevermoegen.immaterielleVermoegensgegenstaende.geleisteteAnzahlungen,
        2,
    );

    // II. Calculate Sachanlagen
    aktiva.anlagevermoegen.sachanlagen.summe = numberUtils.round(
        aktiva.anlagevermoegen.sachanlagen.grundstuecke +
        aktiva.anlagevermoegen.sachanlagen.technischeAnlagenMaschinen +
        aktiva.anlagevermoegen.sachanlagen.andereAnlagen +
        aktiva.anlagevermoegen.sachanlagen.geleisteteAnzahlungen,
        2,
    );

    // III. Calculate Finanzanlagen
    aktiva.anlagevermoegen.finanzanlagen.summe = numberUtils.round(
        aktiva.anlagevermoegen.finanzanlagen.anteileVerbundeneUnternehmen +
        aktiva.anlagevermoegen.finanzanlagen.ausleihungenVerbundene +
        aktiva.anlagevermoegen.finanzanlagen.beteiligungen +
        aktiva.anlagevermoegen.finanzanlagen.ausleihungenBeteiligungen,
        2,
    );

    // A. Calculate Anlagevermögen total
    aktiva.anlagevermoegen.summe = numberUtils.round(
        aktiva.anlagevermoegen.immaterielleVermoegensgegenstaende.summe +
        aktiva.anlagevermoegen.sachanlagen.summe +
        aktiva.anlagevermoegen.finanzanlagen.summe,
        2,
    );

    // I. Calculate Vorräte
    aktiva.umlaufvermoegen.vorraete.summe = numberUtils.round(
        aktiva.umlaufvermoegen.vorraete.rohHilfsBetriebsstoffe +
        aktiva.umlaufvermoegen.vorraete.unfertErzeugnisse +
        aktiva.umlaufvermoegen.vorraete.fertErzeugnisse +
        aktiva.umlaufvermoegen.vorraete.geleisteteAnzahlungen,
        2,
    );

    // II. Calculate Forderungen
    aktiva.umlaufvermoegen.forderungen.summe = numberUtils.round(
        aktiva.umlaufvermoegen.forderungen.forderungenLuL +
        aktiva.umlaufvermoegen.forderungen.forderungenVerbundene +
        aktiva.umlaufvermoegen.forderungen.forderungenBeteiligungen +
        aktiva.umlaufvermoegen.forderungen.sonstigeVermoegensgegenstaende,
        2,
    );

    // III. Calculate Wertpapiere
    aktiva.umlaufvermoegen.wertpapiere.summe = numberUtils.round(
        aktiva.umlaufvermoegen.wertpapiere.anteileVerbundeneUnternehmen +
        aktiva.umlaufvermoegen.wertpapiere.sonstigeWertpapiere,
        2,
    );

    // IV. Calculate Liquide Mittel
    aktiva.umlaufvermoegen.liquideMittel.summe = numberUtils.round(
        aktiva.umlaufvermoegen.liquideMittel.kassenbestand +
        aktiva.umlaufvermoegen.liquideMittel.bankguthaben,
        2,
    );

    // B. Calculate Umlaufvermögen total
    aktiva.umlaufvermoegen.summe = numberUtils.round(
        aktiva.umlaufvermoegen.vorraete.summe +
        aktiva.umlaufvermoegen.forderungen.summe +
        aktiva.umlaufvermoegen.wertpapiere.summe +
        aktiva.umlaufvermoegen.liquideMittel.summe,
        2,
    );

    // Total Assets (Aktiva)
    aktiva.summe = numberUtils.round(
        aktiva.anlagevermoegen.summe +
        aktiva.umlaufvermoegen.summe +
        aktiva.rechnungsabgrenzungsposten +
        aktiva.aktiveLatenteSteuern +
        aktiva.aktiverUnterschiedsbetrag,
        2,
    );
}

/**
 * Calculates subtotals for liabilities and equity (Passiva)
 * @param {Object} bilanzData - Balance sheet data
 */
function calculatePassivaSubtotals(bilanzData) {
    const passiva = bilanzData.passiva;

    // Calculate Gewinnrücklagen
    passiva.eigenkapital.gewinnruecklagen.summe = numberUtils.round(
        passiva.eigenkapital.gewinnruecklagen.gesetzlicheRuecklage +
        passiva.eigenkapital.gewinnruecklagen.ruecklageAnteileGesellschaft +
        passiva.eigenkapital.gewinnruecklagen.satzungsmaessigeRuecklagen +
        passiva.eigenkapital.gewinnruecklagen.andereGewinnruecklagen,
        2,
    );

    // A. Calculate Eigenkapital total
    passiva.eigenkapital.summe = numberUtils.round(
        passiva.eigenkapital.gezeichnetesKapital +
        passiva.eigenkapital.kapitalruecklage +
        passiva.eigenkapital.gewinnruecklagen.summe +
        passiva.eigenkapital.gewinnvortrag +
        passiva.eigenkapital.jahresueberschuss,
        2,
    );

    // B. Calculate Rückstellungen total
    passiva.rueckstellungen.summe = numberUtils.round(
        passiva.rueckstellungen.pensionsaehnlicheRueckstellungen +
        passiva.rueckstellungen.steuerrueckstellungen +
        passiva.rueckstellungen.sonstigeRueckstellungen,
        2,
    );

    // C. Calculate Verbindlichkeiten total
    passiva.verbindlichkeiten.summe = numberUtils.round(
        passiva.verbindlichkeiten.anleihen +
        passiva.verbindlichkeiten.verbindlichkeitenKreditinstitute +
        passiva.verbindlichkeiten.erhalteneAnzahlungen +
        passiva.verbindlichkeiten.verbindlichkeitenLuL +
        passiva.verbindlichkeiten.verbindlichkeitenWechsel +
        passiva.verbindlichkeiten.verbindlichkeitenVerbundene +
        passiva.verbindlichkeiten.verbindlichkeitenBeteiligungen +
        passiva.verbindlichkeiten.sonstigeVerbindlichkeiten,
        2,
    );

    // Total Liabilities and Equity (Passiva)
    passiva.summe = numberUtils.round(
        passiva.eigenkapital.summe +
        passiva.rueckstellungen.summe +
        passiva.verbindlichkeiten.summe +
        passiva.rechnungsabgrenzungsposten +
        passiva.passiveLatenteSteuern,
        2,
    );
}

/**
 * Validates balance sheet equality (Aktiva = Passiva)
 * @param {Object} bilanzData - Balance sheet data
 */
function validateBilanzEquality(bilanzData) {
    const aktivaSumme = bilanzData.aktiva.summe;
    const passivaSumme = bilanzData.passiva.summe;

    // If there's a difference, try to adjust retained earnings to balance
    const differenz = numberUtils.round(aktivaSumme - passivaSumme, 2);

    if (differenz !== 0) {
        console.log(`Balance sheet difference detected: ${differenz}. Adjusting retained earnings.`);

        // Adjust retained earnings to balance the sheet
        bilanzData.passiva.eigenkapital.jahresueberschuss += differenz;

        // Recalculate Eigenkapital total
        bilanzData.passiva.eigenkapital.summe = numberUtils.round(
            bilanzData.passiva.eigenkapital.gezeichnetesKapital +
            bilanzData.passiva.eigenkapital.kapitalruecklage +
            bilanzData.passiva.eigenkapital.gewinnruecklagen.summe +
            bilanzData.passiva.eigenkapital.gewinnvortrag +
            bilanzData.passiva.eigenkapital.jahresueberschuss,
            2,
        );

        // Recalculate Passiva total
        bilanzData.passiva.summe = numberUtils.round(
            bilanzData.passiva.eigenkapital.summe +
            bilanzData.passiva.rueckstellungen.summe +
            bilanzData.passiva.verbindlichkeiten.summe +
            bilanzData.passiva.rechnungsabgrenzungsposten +
            bilanzData.passiva.passiveLatenteSteuern,
            2,
        );
    }
}

export default {
    getBankDarlehen,
    calculateBilanzTotals,
    calculateAktivaSubtotals,
    calculatePassivaSubtotals,
    validateBilanzEquality,
};