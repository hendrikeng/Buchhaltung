// modules/bwaModule/formatter.js
import numberUtils from '../../utils/numberUtils.js';
import sheetUtils from '../../utils/sheetUtils.js';

/**
 * Creates a header row for the BWA with month and quarterly columns
 * @param {Object} config - Configuration
 * @returns {Array} Header row
 */
function buildHeaderRow(config) {
    const headers = ['Kategorie'];

    // Build header with months and quarters
    for (let q = 0; q < 4; q++) {
        for (let m = q * 3; m < q * 3 + 3; m++) {
            headers.push(`${config.common.months[m]} (€)`);
        }
        headers.push(`Q${q + 1} (€)`);
    }
    headers.push('Jahr (€)');

    return headers;
}

/**
 * Creates an output row for a position with optimized calculation logic
 * @param {string} label - Position label
 * @param {function} getter - Function to get the value from BWA data
 * @param {Object} bwaData - BWA data
 * @returns {Array} Formatted row
 */
function buildOutputRow(label, getter, bwaData) {
    // Pre-allocate arrays for values
    const monthly = new Array(12).fill(0);
    const quarters = new Array(4).fill(0);
    let yearly = 0;

    // Get monthly values in one pass
    for (let m = 1; m <= 12; m++) {
        const val = getter(bwaData[m]) || 0;
        monthly[m-1] = val;
        yearly += val;

        // Count quarters directly
        quarters[Math.floor((m-1) / 3)] += val;
    }

    // Round all values for better display
    const roundedMonthly = monthly.map(val => numberUtils.round(val, 2));
    const roundedQuarters = quarters.map(val => numberUtils.round(val, 2));
    const roundedYearly = numberUtils.round(yearly, 2);

    // Build row with optimized structure
    // [Label, Jan, Feb, Mar, Q1, Apr, May, Jun, Q2, ...]
    return [
        label,
        ...roundedMonthly.slice(0, 3),
        roundedQuarters[0],
        ...roundedMonthly.slice(3, 6),
        roundedQuarters[1],
        ...roundedMonthly.slice(6, 9),
        roundedQuarters[2],
        ...roundedMonthly.slice(9, 12),
        roundedQuarters[3],
        roundedYearly,
    ];
}

/**
 * Creates the BWA sheet based on BWA data with DATEV-compliant structure
 * @param {Object} bwaData - BWA data by month
 * @param {Spreadsheet} ss - Spreadsheet
 * @param {Object} config - Configuration
 * @returns {boolean} true on success, false on error
 */
function generateBWASheet(bwaData, ss, config) {
    try {
        console.log('Generating DATEV-compliant BWA sheet...');

        // Define DATEV-compliant BWA structure based on bwa.json
        const bwaStructure = [
            {
                kategorie: '1. Betriebserlöse (Einnahmen)',
                positionen: [
                    {label: 'Erlöse aus Lieferungen und Leistungen', get: d => d.erloeseLieferungenLeistungen || 0},
                    {label: 'Provisionserlöse', get: d => d.provisionserloese || 0},
                    {label: 'Steuerfreie Inland-Erlöse', get: d => d.steuerfreieInlandErloese || 0},
                    {label: 'Steuerfreie Ausland-Erlöse', get: d => d.steuerfreieAuslandErloese || 0},
                    {label: 'Innergemeinschaftliche Lieferungen', get: d => d.innergemeinschaftlicheLieferungen || 0},
                    {label: 'Sonstige betriebliche Erträge', get: d => d.sonstigeBetrieblicheErtraege || 0},
                    {label: 'Erträge aus Vermietung und Verpachtung', get: d => d.ertraegeVermietungVerpachtung || 0},
                    {label: 'Erträge aus Zuschüssen', get: d => d.ertraegeZuschuesse || 0},
                    {label: 'Erträge aus Kursgewinnen', get: d => d.ertraegeKursgewinne || 0},
                    {label: 'Erträge aus Anlagenabgängen', get: d => d.ertraegeAnlagenabgaenge || 0},
                ],
                summe: {label: 'Betriebserlöse gesamt', get: d => d.betriebserloese_gesamt || 0},
            },
            {
                kategorie: '2. Materialaufwand & Wareneinsatz',
                positionen: [
                    {label: 'Wareneinsatz', get: d => d.wareneinsatz || 0},
                    {label: 'Bezogene Leistungen', get: d => d.bezogeneLeistungen || 0},
                    {label: 'Roh-, Hilfs- und Betriebsstoffe', get: d => d.rohHilfsBetriebsstoffe || 0},
                ],
                summe: {label: 'Materialaufwand gesamt', get: d => d.materialaufwand_gesamt || 0},
            },
            {
                kategorie: '3. Personalaufwand',
                positionen: [
                    {label: 'Löhne und Gehälter', get: d => d.loehneGehaelter || 0},
                    {label: 'Soziale Abgaben und Arbeitgeberanteile', get: d => d.sozialeAbgaben || 0},
                    {label: 'Sonstige Personalkosten', get: d => d.sonstigePersonalkosten || 0},
                ],
                summe: {label: 'Personalaufwand gesamt', get: d => d.personalaufwand_gesamt || 0},
            },
            {
                kategorie: '4. Sonstige betriebliche Aufwendungen',
                positionen: [
                    {label: 'Provisionszahlungen (an Dritte)', get: d => d.provisionszahlungenDritte || 0},
                    {label: 'IT-Kosten (Software, Cloud, Support)', get: d => d.itKosten || 0},
                    {label: 'Miete und Leasing', get: d => d.mieteLeasing || 0},
                    {label: 'Werbung und Marketing', get: d => d.werbungMarketing || 0},
                    {label: 'Reisekosten', get: d => d.reisekosten || 0},
                    {label: 'Versicherungen', get: d => d.versicherungen || 0},
                    {label: 'Telefon und Internet', get: d => d.telefonInternet || 0},
                    {label: 'Bürokosten', get: d => d.buerokosten || 0},
                    {label: 'Fortbildungskosten', get: d => d.fortbildungskosten || 0},
                    {label: 'Kfz-Kosten', get: d => d.kfzKosten || 0},
                    {label: 'Beiträge und Abgaben', get: d => d.beitraegeAbgaben || 0},
                    {label: 'Bewirtungskosten (abzugsfähig)', get: d => d.bewirtungskosten || 0},
                    {label: 'Sonstige betriebliche Aufwendungen', get: d => d.sonstigeBetrieblicheAufwendungen || 0},
                ],
                summe: {label: 'Sonstige betriebliche Aufwendungen gesamt', get: d => d.sonstigeAufwendungen_gesamt || 0},
            },
            {
                kategorie: '5. Abschreibungen und Zinsen',
                positionen: [
                    {label: 'Abschreibungen auf Sachanlagen', get: d => d.abschreibungenSachanlagen || 0},
                    {label: 'Abschreibungen auf immaterielle Vermögensgegenstände', get: d => d.abschreibungenImmaterielleVG || 0},
                    {label: 'Zinsaufwendungen aus Bankdarlehen', get: d => d.zinsenBankdarlehen || 0},
                    {label: 'Zinsaufwendungen aus Gesellschafterdarlehen', get: d => d.zinsenGesellschafterdarlehen || 0},
                    {label: 'Leasingzinsen', get: d => d.leasingzinsen || 0},
                ],
                summe: {label: 'Abschreibungen und Zinsen gesamt', get: d => d.abschreibungenZinsen_gesamt || 0},
            },
            {
                kategorie: '6. Betriebsergebnis (EBIT)',
                positionen: [
                    {label: 'Betriebsergebnis vor Steuern (EBIT)', get: d => d.ebit || 0},
                ],
            },
            {
                kategorie: '7. Jahresergebnis',
                positionen: [
                    {label: 'Jahresüberschuss / Jahresfehlbetrag', get: d => d.jahresergebnis || 0},
                ],
            },
        ];

        // Header row
        const headerRow = buildHeaderRow(config);

        // Calculate total number of rows needed
        let totalRows = 1; // Header row
        bwaStructure.forEach(group => {
            totalRows += 1; // Group header
            totalRows += group.positionen.length;
            if (group.summe) totalRows += 1; // Sum row
            totalRows += 1; // Empty line after group
        });

        // Pre-allocate output array
        const outputRows = new Array(totalRows);
        outputRows[0] = headerRow;

        // Build output rows
        let rowIndex = 1;

        for (const group of bwaStructure) {
            // Group heading
            outputRows[rowIndex++] = [
                group.kategorie,
                ...Array(headerRow.length - 1).fill(''),
            ];

            // Position rows
            for (const position of group.positionen) {
                outputRows[rowIndex++] = buildOutputRow(position.label, position.get, bwaData);
            }

            // Sum row if exists
            if (group.summe) {
                outputRows[rowIndex++] = buildOutputRow(group.summe.label, group.summe.get, bwaData);
            }

            // Empty line after group
            outputRows[rowIndex++] = Array(headerRow.length).fill('');
        }

        // Create or update BWA sheet
        const bwaSheet = sheetUtils.getOrCreateSheet(ss, 'BWA');
        bwaSheet.clearContents();

        // Write data in one batch API call
        bwaSheet.getRange(1, 1, rowIndex, outputRows[0].length).setValues(outputRows);

        // Apply formatting in optimized batches
        applyBwaFormatting(bwaSheet, bwaStructure, rowIndex);

        // Activate BWA sheet
        ss.setActiveSheet(bwaSheet);
        console.log('BWA sheet generated successfully');

        return true;
    } catch (e) {
        console.error('Error creating BWA:', e);
        return false;
    }
}

/**
 * Applies formatting to the BWA sheet with DATEV-compliant styling
 * @param {Sheet} sheet - Sheet to format
 * @param {Array} bwaStructure - BWA structure definition
 * @param {number} totalRows - Total number of rows
 */
function applyBwaFormatting(sheet, bwaStructure, totalRows) {
    const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
    headerRange.setFontWeight('bold').setBackground('#f3f3f3');

    // Track row positions for formatting
    let rowIndex = 2; // Start after header

    // Group headers and sums formatting
    for (const group of bwaStructure) {
        // Format group header
        const groupHeader = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn());
        groupHeader.setFontWeight('bold').setBackground('#eaeaea');
        rowIndex++;

        // Skip position rows (just track the count)
        rowIndex += group.positionen.length;

        // Format sum row if exists
        if (group.summe) {
            const sumRow = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn());
            sumRow.setFontWeight('bold').setBackground('#e6f2ff');
            rowIndex++;
        }

        // Skip empty line
        rowIndex++;
    }

    // Special formatting for EBIT and Jahresergebnis
    // Find rows for EBIT and Jahresergebnis
    let ebitRow = 0;
    let jahresergebnisRow = 0;

    rowIndex = 2;
    for (const group of bwaStructure) {
        if (group.kategorie === '6. Betriebsergebnis (EBIT)') {
            ebitRow = rowIndex + 1; // +1 for the first position row
        } else if (group.kategorie === '7. Jahresergebnis') {
            jahresergebnisRow = rowIndex + 1; // +1 for the first position row
        }

        // Move to next group
        rowIndex += 1 + group.positionen.length;
        if (group.summe) rowIndex += 1;
        rowIndex += 1; // Empty line
    }

    // Apply special formatting
    if (ebitRow > 0) {
        sheet.getRange(ebitRow, 1, 1, sheet.getLastColumn())
            .setFontWeight('bold')
            .setBackground('#d9edf7');
    }

    if (jahresergebnisRow > 0) {
        sheet.getRange(jahresergebnisRow, 1, 1, sheet.getLastColumn())
            .setFontWeight('bold')
            .setBackground('#dff0d8');
    }

    // Currency format for all numeric values
    sheet.getRange(2, 2, totalRows - 1, sheet.getLastColumn() - 1)
        .setNumberFormat('#,##0.00 €');

    // Adjust column widths
    sheet.autoResizeColumns(1, sheet.getLastColumn());

    // Make first column a bit wider for labels
    sheet.setColumnWidth(1, 360);
}

export default {
    buildHeaderRow,
    buildOutputRow,
    generateBWASheet,
    applyBwaFormatting,
};