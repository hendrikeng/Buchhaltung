// modules/bilanzModule/formatter.js
import numberUtils from '../../utils/numberUtils.js';
import sheetUtils from '../../utils/sheetUtils.js';

/**
 * Creates a DATEV-compliant balance sheet
 * @param {Object} bilanzData - Balance sheet data
 * @param {Spreadsheet} ss - Spreadsheet
 * @param {Object} config - Configuration
 * @returns {boolean} Success status
 */
function generateBilanzSheet(bilanzData, ss, config) {
    try {
        console.log('Generating DATEV-compliant balance sheet...');
        const year = config.tax.year || new Date().getFullYear();

        // Create or clear balance sheet
        const bilanzSheet = sheetUtils.getOrCreateSheet(ss, 'Bilanz');
        bilanzSheet.clearContents();

        // Prepare data for assets (Aktiva) and liabilities/equity (Passiva)
        const aktivaData = createAktivaData(bilanzData, year);
        const passivaData = createPassivaData(bilanzData, year);

        // Write data to the sheet
        const totalRows = Math.max(aktivaData.length, passivaData.length);

        // Create a combined array with both sides
        const combinedData = [];
        for (let i = 0; i < totalRows; i++) {
            const row = Array(10).fill(''); // 10 columns total

            // Add Aktiva data if available
            if (i < aktivaData.length) {
                row[0] = aktivaData[i][0]; // Label
                row[1] = aktivaData[i][1]; // Value
            }

            // Add divider column
            row[4] = ''; // Empty divider column

            // Add Passiva data if available
            if (i < passivaData.length) {
                row[5] = passivaData[i][0]; // Label
                row[6] = passivaData[i][1]; // Value
            }

            combinedData.push(row);
        }

        // Write all data in one batch
        bilanzSheet.getRange(1, 1, combinedData.length, 10).setValues(combinedData);

        // Apply formatting
        applyBilanzFormatting(bilanzSheet, aktivaData.length, passivaData.length, year);

        // Activate the sheet
        ss.setActiveSheet(bilanzSheet);
        console.log('Balance sheet generated successfully');

        return true;
    } catch (e) {
        console.error('Error generating balance sheet:', e);
        return false;
    }
}

/**
 * Creates data array for assets (Aktiva)
 * @param {Object} bilanzData - Balance sheet data
 * @param {number} year - Fiscal year
 * @returns {Array} Data array for assets
 */
function createAktivaData(bilanzData, year) {
    const aktiva = bilanzData.aktiva;
    const data = [];

    // Header
    data.push([`Aktiva (Vermögenswerte) ${year}`, '']);
    data.push(['', '']);

    // A. Anlagevermögen
    data.push(['A. Anlagevermögen', '']);

    // I. Immaterielle Vermögensgegenstände
    data.push(['I. Immaterielle Vermögensgegenstände', '']);
    data.push(['1. Konzessionen, Lizenzen und ähnliche Rechte', formatValue(aktiva.anlagevermoegen.immaterielleVermoegensgegenstaende.konzessionen)]);
    data.push(['2. Geschäfts- oder Firmenwert', formatValue(aktiva.anlagevermoegen.immaterielleVermoegensgegenstaende.geschaeftswert)]);
    data.push(['3. Geleistete Anzahlungen', formatValue(aktiva.anlagevermoegen.immaterielleVermoegensgegenstaende.geleisteteAnzahlungen)]);
    data.push(['Summe I', formatValue(aktiva.anlagevermoegen.immaterielleVermoegensgegenstaende.summe)]);

    // II. Sachanlagen
    data.push(['II. Sachanlagen', '']);
    data.push(['1. Grundstücke und grundstücksgleiche Rechte', formatValue(aktiva.anlagevermoegen.sachanlagen.grundstuecke)]);
    data.push(['2. Technische Anlagen und Maschinen', formatValue(aktiva.anlagevermoegen.sachanlagen.technischeAnlagenMaschinen)]);
    data.push(['3. Andere Anlagen, Betriebs- und Geschäftsausstattung', formatValue(aktiva.anlagevermoegen.sachanlagen.andereAnlagen)]);
    data.push(['4. Geleistete Anzahlungen und Anlagen im Bau', formatValue(aktiva.anlagevermoegen.sachanlagen.geleisteteAnzahlungen)]);
    data.push(['Summe II', formatValue(aktiva.anlagevermoegen.sachanlagen.summe)]);

    // III. Finanzanlagen
    data.push(['III. Finanzanlagen', '']);
    data.push(['1. Anteile an verbundenen Unternehmen', formatValue(aktiva.anlagevermoegen.finanzanlagen.anteileVerbundeneUnternehmen)]);
    data.push(['2. Ausleihungen an verbundene Unternehmen', formatValue(aktiva.anlagevermoegen.finanzanlagen.ausleihungenVerbundene)]);
    data.push(['3. Beteiligungen', formatValue(aktiva.anlagevermoegen.finanzanlagen.beteiligungen)]);
    data.push(['4. Ausleihungen an Unternehmen mit Beteiligungsverhältnis', formatValue(aktiva.anlagevermoegen.finanzanlagen.ausleihungenBeteiligungen)]);
    data.push(['Summe III', formatValue(aktiva.anlagevermoegen.finanzanlagen.summe)]);

    data.push(['Summe A', formatValue(aktiva.anlagevermoegen.summe)]);
    data.push(['', '']);

    // B. Umlaufvermögen
    data.push(['B. Umlaufvermögen', '']);

    // I. Vorräte
    data.push(['I. Vorräte', '']);
    data.push(['1. Roh-, Hilfs- und Betriebsstoffe', formatValue(aktiva.umlaufvermoegen.vorraete.rohHilfsBetriebsstoffe)]);
    data.push(['2. Unfertige Erzeugnisse', formatValue(aktiva.umlaufvermoegen.vorraete.unfertErzeugnisse)]);
    data.push(['3. Fertige Erzeugnisse und Waren', formatValue(aktiva.umlaufvermoegen.vorraete.fertErzeugnisse)]);
    data.push(['4. Geleistete Anzahlungen', formatValue(aktiva.umlaufvermoegen.vorraete.geleisteteAnzahlungen)]);
    data.push(['Summe I', formatValue(aktiva.umlaufvermoegen.vorraete.summe)]);

    // II. Forderungen und sonstige Vermögensgegenstände
    data.push(['II. Forderungen und sonstige Vermögensgegenstände', '']);
    data.push(['1. Forderungen aus Lieferungen und Leistungen', formatValue(aktiva.umlaufvermoegen.forderungen.forderungenLuL)]);
    data.push(['2. Forderungen gegen verbundene Unternehmen', formatValue(aktiva.umlaufvermoegen.forderungen.forderungenVerbundene)]);
    data.push(['3. Forderungen gegen Unternehmen mit Beteiligungsverhältnis', formatValue(aktiva.umlaufvermoegen.forderungen.forderungenBeteiligungen)]);
    data.push(['4. Sonstige Vermögensgegenstände', formatValue(aktiva.umlaufvermoegen.forderungen.sonstigeVermoegensgegenstaende)]);
    data.push(['Summe II', formatValue(aktiva.umlaufvermoegen.forderungen.summe)]);

    // III. Wertpapiere
    data.push(['III. Wertpapiere', '']);
    data.push(['1. Anteile an verbundenen Unternehmen', formatValue(aktiva.umlaufvermoegen.wertpapiere.anteileVerbundeneUnternehmen)]);
    data.push(['2. Sonstige Wertpapiere', formatValue(aktiva.umlaufvermoegen.wertpapiere.sonstigeWertpapiere)]);
    data.push(['Summe III', formatValue(aktiva.umlaufvermoegen.wertpapiere.summe)]);

    // IV. Kassenbestand, Bankguthaben
    data.push(['IV. Kassenbestand, Bankguthaben', '']);
    data.push(['1. Kassenbestand', formatValue(aktiva.umlaufvermoegen.liquideMittel.kassenbestand)]);
    data.push(['2. Bankguthaben', formatValue(aktiva.umlaufvermoegen.liquideMittel.bankguthaben)]);
    data.push(['Summe IV', formatValue(aktiva.umlaufvermoegen.liquideMittel.summe)]);

    data.push(['Summe B', formatValue(aktiva.umlaufvermoegen.summe)]);
    data.push(['', '']);

    // C. Rechnungsabgrenzungsposten
    data.push(['C. Rechnungsabgrenzungsposten', formatValue(aktiva.rechnungsabgrenzungsposten)]);

    // D. Aktive latente Steuern
    data.push(['D. Aktive latente Steuern', formatValue(aktiva.aktiveLatenteSteuern)]);

    // E. Aktiver Unterschiedsbetrag aus Vermögensverrechnung
    data.push(['E. Aktiver Unterschiedsbetrag aus Vermögensverrechnung', formatValue(aktiva.aktiverUnterschiedsbetrag)]);

    data.push(['', '']);
    data.push(['Summe Aktiva', formatValue(aktiva.summe)]);

    return data;
}

/**
 * Creates data array for liabilities and equity (Passiva)
 * @param {Object} bilanzData - Balance sheet data
 * @param {number} year - Fiscal year
 * @returns {Array} Data array for liabilities and equity
 */
function createPassivaData(bilanzData, year) {
    const passiva = bilanzData.passiva;
    const data = [];

    // Header
    data.push([`Passiva (Kapital und Schulden) ${year}`, '']);
    data.push(['', '']);

    // A. Eigenkapital
    data.push(['A. Eigenkapital', '']);
    data.push(['I. Gezeichnetes Kapital', formatValue(passiva.eigenkapital.gezeichnetesKapital)]);
    data.push(['II. Kapitalrücklage', formatValue(passiva.eigenkapital.kapitalruecklage)]);

    // III. Gewinnrücklagen
    data.push(['III. Gewinnrücklagen', '']);
    data.push(['1. Gesetzliche Rücklage', formatValue(passiva.eigenkapital.gewinnruecklagen.gesetzlicheRuecklage)]);
    data.push(['2. Rücklage für Anteile an herrschendem Unternehmen', formatValue(passiva.eigenkapital.gewinnruecklagen.ruecklageAnteileGesellschaft)]);
    data.push(['3. Satzungsmäßige Rücklagen', formatValue(passiva.eigenkapital.gewinnruecklagen.satzungsmaessigeRuecklagen)]);
    data.push(['4. Andere Gewinnrücklagen', formatValue(passiva.eigenkapital.gewinnruecklagen.andereGewinnruecklagen)]);
    data.push(['Summe III', formatValue(passiva.eigenkapital.gewinnruecklagen.summe)]);

    data.push(['IV. Gewinnvortrag/Verlustvortrag', formatValue(passiva.eigenkapital.gewinnvortrag)]);
    data.push(['V. Jahresüberschuss/Jahresfehlbetrag', formatValue(passiva.eigenkapital.jahresueberschuss)]);
    data.push(['Summe A', formatValue(passiva.eigenkapital.summe)]);
    data.push(['', '']);

    // B. Rückstellungen
    data.push(['B. Rückstellungen', '']);
    data.push(['1. Rückstellungen für Pensionen und ähnliche Verpflichtungen', formatValue(passiva.rueckstellungen.pensionsaehnlicheRueckstellungen)]);
    data.push(['2. Steuerrückstellungen', formatValue(passiva.rueckstellungen.steuerrueckstellungen)]);
    data.push(['3. Sonstige Rückstellungen', formatValue(passiva.rueckstellungen.sonstigeRueckstellungen)]);
    data.push(['Summe B', formatValue(passiva.rueckstellungen.summe)]);
    data.push(['', '']);

    // C. Verbindlichkeiten
    data.push(['C. Verbindlichkeiten', '']);
    data.push(['1. Anleihen', formatValue(passiva.verbindlichkeiten.anleihen)]);
    data.push(['2. Verbindlichkeiten gegenüber Kreditinstituten', formatValue(passiva.verbindlichkeiten.verbindlichkeitenKreditinstitute)]);
    data.push(['3. Erhaltene Anzahlungen auf Bestellungen', formatValue(passiva.verbindlichkeiten.erhalteneAnzahlungen)]);
    data.push(['4. Verbindlichkeiten aus Lieferungen und Leistungen', formatValue(passiva.verbindlichkeiten.verbindlichkeitenLuL)]);
    data.push(['5. Verbindlichkeiten aus der Annahme gezogener Wechsel', formatValue(passiva.verbindlichkeiten.verbindlichkeitenWechsel)]);
    data.push(['6. Verbindlichkeiten gegenüber verbundenen Unternehmen', formatValue(passiva.verbindlichkeiten.verbindlichkeitenVerbundene)]);
    data.push(['7. Verbindlichkeiten gegenüber Unternehmen mit Beteiligungsverhältnis', formatValue(passiva.verbindlichkeiten.verbindlichkeitenBeteiligungen)]);
    data.push(['8. Sonstige Verbindlichkeiten', formatValue(passiva.verbindlichkeiten.sonstigeVerbindlichkeiten)]);
    data.push(['Summe C', formatValue(passiva.verbindlichkeiten.summe)]);
    data.push(['', '']);

    // D. Rechnungsabgrenzungsposten
    data.push(['D. Rechnungsabgrenzungsposten', formatValue(passiva.rechnungsabgrenzungsposten)]);

    // E. Passive latente Steuern
    data.push(['E. Passive latente Steuern', formatValue(passiva.passiveLatenteSteuern)]);

    data.push(['', '']);
    data.push(['Summe Passiva', formatValue(passiva.summe)]);

    return data;
}

/**
 * Formats a value for the balance sheet
 * @param {number} value - The value to format
 * @returns {number|string} Formatted value or empty string if zero
 */
function formatValue(value) {
    // Values exactly zero are displayed as empty strings
    if (value === 0) return '';
    return value;
}

/**
 * Applies formatting to the balance sheet
 * @param {Sheet} sheet - Balance sheet
 * @param {number} aktivaRows - Number of asset rows
 * @param {number} passivaRows - Number of liabilities and equity rows
 * @param {number} year - Fiscal year
 */
function applyBilanzFormatting(sheet, aktivaRows, passivaRows, year) {
    // 1. Format headers
    sheet.getRange(1, 1, 1, 2).merge().setHorizontalAlignment('center').setFontWeight('bold').setBackground('#d9e1f2');
    sheet.getRange(1, 6, 1, 2).merge().setHorizontalAlignment('center').setFontWeight('bold').setBackground('#d9e1f2');

    // 2. Format section headers
    const sectionHeaders = [
        // Aktiva sections
        {row: 3, col: 1, background: '#e7eff9'}, // A. Anlagevermögen
        {row: 4, col: 1, background: '#f2f2f2'}, // I. Immaterielle Vermögensgegenstände
        {row: 9, col: 1, background: '#f2f2f2'}, // II. Sachanlagen
        {row: 15, col: 1, background: '#f2f2f2'}, // III. Finanzanlagen
        {row: 22, col: 1, background: '#e7eff9'}, // Summe A
        {row: 24, col: 1, background: '#e7eff9'}, // B. Umlaufvermögen
        {row: 25, col: 1, background: '#f2f2f2'}, // I. Vorräte
        {row: 31, col: 1, background: '#f2f2f2'}, // II. Forderungen
        {row: 37, col: 1, background: '#f2f2f2'}, // III. Wertpapiere
        {row: 41, col: 1, background: '#f2f2f2'}, // IV. Kassenbestand
        {row: 45, col: 1, background: '#e7eff9'}, // Summe B
        {row: 47, col: 1, background: '#e7eff9'}, // C. Rechnungsabgrenzungsposten
        {row: 48, col: 1, background: '#e7eff9'}, // D. Aktive latente Steuern
        {row: 49, col: 1, background: '#e7eff9'}, // E. Aktiver Unterschiedsbetrag
        {row: 51, col: 1, background: '#d9e1f2'}, // Summe Aktiva

        // Passiva sections
        {row: 3, col: 6, background: '#e7eff9'}, // A. Eigenkapital
        {row: 4, col: 6, background: '#f2f2f2'}, // I. Gezeichnetes Kapital
        {row: 5, col: 6, background: '#f2f2f2'}, // II. Kapitalrücklage
        {row: 6, col: 6, background: '#f2f2f2'}, // III. Gewinnrücklagen
        {row: 12, col: 6, background: '#f2f2f2'}, // IV. Gewinnvortrag
        {row: 13, col: 6, background: '#f2f2f2'}, // V. Jahresüberschuss
        {row: 14, col: 6, background: '#e7eff9'}, // Summe A
        {row: 16, col: 6, background: '#e7eff9'}, // B. Rückstellungen
        {row: 20, col: 6, background: '#e7eff9'}, // Summe B
        {row: 22, col: 6, background: '#e7eff9'}, // C. Verbindlichkeiten
        {row: 31, col: 6, background: '#e7eff9'}, // Summe C
        {row: 33, col: 6, background: '#e7eff9'}, // D. Rechnungsabgrenzungsposten
        {row: 34, col: 6, background: '#e7eff9'}, // E. Passive latente Steuern
        {row: 36, col: 6, background: '#d9e1f2'}, // Summe Passiva
    ];

    // Apply section header formatting
    sectionHeaders.forEach(header => {
        try {
            const range = sheet.getRange(header.row, header.col, 1, 1);
            range.setBackground(header.background)
                .setFontWeight('bold');
        } catch (e) {
            // Skip if row doesn't exist
        }
    });

    // 3. Format sum rows
    const sumRows = [
        // Aktiva sum rows
        {row: 8, col: 1},  // Summe I (Immaterielle Vermögensgegenstände)
        {row: 14, col: 1}, // Summe II (Sachanlagen)
        {row: 21, col: 1}, // Summe III (Finanzanlagen)
        {row: 22, col: 1}, // Summe A (Anlagevermögen)
        {row: 30, col: 1}, // Summe I (Vorräte)
        {row: 36, col: 1}, // Summe II (Forderungen)
        {row: 40, col: 1}, // Summe III (Wertpapiere)
        {row: 44, col: 1}, // Summe IV (Kassenbestand)
        {row: 45, col: 1}, // Summe B (Umlaufvermögen)
        {row: 51, col: 1}, // Summe Aktiva

        // Passiva sum rows
        {row: 11, col: 6}, // Summe III (Gewinnrücklagen)
        {row: 14, col: 6}, // Summe A (Eigenkapital)
        {row: 20, col: 6}, // Summe B (Rückstellungen)
        {row: 31, col: 6}, // Summe C (Verbindlichkeiten)
        {row: 36, col: 6}, // Summe Passiva
    ];

    // Apply sum row formatting
    sumRows.forEach(sum => {
        try {
            const range = sheet.getRange(sum.row, sum.col, 1, 2);
            range.setFontWeight('bold');

            // Format final sums with different styling
            if (sum.row === 51 || sum.row === 36 && sum.col === 6) {
                range.setBackground('#d9e1f2');
            }
        } catch (e) {
            // Skip if row doesn't exist
        }
    });

    // 4. Format currency values
    sheet.getRange(1, 2, sheet.getLastRow(), 1).setNumberFormat('#,##0.00 €');
    sheet.getRange(1, 7, sheet.getLastRow(), 1).setNumberFormat('#,##0.00 €');

    // 5. Set column widths
    sheet.setColumnWidth(1, 360); // Aktiva labels
    sheet.setColumnWidth(2, 130); // Aktiva values
    sheet.setColumnWidth(3, 20);  // Spacing
    sheet.setColumnWidth(4, 20);  // Spacing
    sheet.setColumnWidth(5, 20);  // Spacing
    sheet.setColumnWidth(6, 360); // Passiva labels
    sheet.setColumnWidth(7, 130); // Passiva values

    // 6. Add vertical divider
    sheet.getRange(1, 4, sheet.getLastRow(), 1).setBorder(false, true, false, true, false, false);
}

export default {
    generateBilanzSheet,
    createAktivaData,
    createPassivaData,
    formatValue,
};