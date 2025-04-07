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

    // Optimization: Build header with a for loop instead of implicit iteration
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
 * @param {Object} pos - Position with label and value function
 * @param {Object} bwaData - BWA data
 * @returns {Array} Formatted row
 */
function buildOutputRow(pos, bwaData) {
    // Optimization: Pre-allocate arrays for values
    const monthly = new Array(12).fill(0);
    const quarters = new Array(4).fill(0);
    let yearly = 0;

    // Optimization: Get monthly values in one pass
    for (let m = 1; m <= 12; m++) {
        const val = pos.get(bwaData[m]) || 0;
        monthly[m-1] = val;
        yearly += val;

        // Count quarters directly (optimized)
        quarters[Math.floor((m-1) / 3)] += val;
    }

    // Round all values for better display
    const roundedMonthly = monthly.map(val => numberUtils.round(val, 2));
    const roundedQuarters = quarters.map(val => numberUtils.round(val, 2));
    const roundedYearly = numberUtils.round(yearly, 2);

    // Build row with optimized structure
    // [Label, Jan, Feb, Mar, Q1, Apr, May, Jun, Q2, ...]
    return [
        pos.label,
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
 * Creates the BWA sheet based on BWA data with optimized batch processing
 * Following DATEV standards for BWA reporting
 * @param {Object} bwaData - BWA data by month
 * @param {Spreadsheet} ss - Spreadsheet
 * @param {Object} config - Configuration
 * @returns {boolean} true on success, false on error
 */
function generateBWASheet(bwaData, ss, config) {
    try {
        console.log('Generating BWA sheet...');

        // Define positions for the BWA following DATEV standards
        const positions = [
            // 1. Revenue and income
            {label: 'Erlöse aus Lieferungen und Leistungen', get: d => d.umsatzerloese || 0},
            {label: 'Provisionserlöse', get: d => d.provisionserloese || 0},
            {label: 'Steuerfreie Inland-Einnahmen', get: d => d.steuerfreieInlandEinnahmen || 0},
            {label: 'Steuerfreie Ausland-Einnahmen', get: d => d.steuerfreieAuslandEinnahmen || 0},
            {label: 'Innergemeinschaftliche Lieferungen', get: d => d.innergemeinschaftlicheLieferungen || 0},
            {label: 'Sonstige betriebliche Erträge', get: d => d.sonstigeErtraege || 0},
            {label: 'Erträge aus Vermietung/Verpachtung', get: d => d.vermietung || 0},
            {label: 'Erträge aus Zuschüssen', get: d => d.zuschuesse || 0},
            {label: 'Erträge aus Währungsgewinnen', get: d => d.waehrungsgewinne || 0},
            {label: 'Erträge aus Anlagenabgängen', get: d => d.anlagenabgaenge || 0},
            {label: 'Steuerliche Korrekturen', get: d => d.steuerlicheKorrekturen || 0},
            {label: 'Betriebserlöse gesamt', get: d => d.gesamtErloese || 0},

            // 2. Material expenses & Cost of goods
            {label: 'Wareneinsatz', get: d => d.wareneinsatz || 0},
            {label: 'Bezogene Leistungen', get: d => d.fremdleistungen || 0},
            {label: 'Roh-, Hilfs- & Betriebsstoffe', get: d => d.rohHilfsBetriebsstoffe || 0},
            {label: 'Materialaufwand gesamt', get: d => d.gesamtWareneinsatz || 0},

            // 3. Operating expenses
            {label: 'Bruttolöhne & Gehälter', get: d => d.bruttoLoehne || 0},
            {label: 'Soziale Abgaben & Arbeitgeberanteile', get: d => d.sozialeAbgaben || 0},
            {label: 'Sonstige Personalkosten', get: d => d.sonstigePersonalkosten || 0},
            {label: 'Miete & Nebenkosten', get: d => d.mieteNebenkosten || 0},
            {label: 'Werbung & Marketing', get: d => d.werbungMarketing || 0},
            {label: 'Reisekosten', get: d => d.reisekosten || 0},
            {label: 'Versicherungen', get: d => d.versicherungen || 0},
            {label: 'Telefon & Internet', get: d => d.telefonInternet || 0},
            {label: 'Bürokosten', get: d => d.buerokosten || 0},
            {label: 'Fortbildungskosten', get: d => d.fortbildungskosten || 0},
            {label: 'Kfz-Kosten', get: d => d.kfzKosten || 0},
            {label: 'Sonstige betriebliche Aufwendungen', get: d => d.sonstigeAufwendungen || 0},
            {label: 'Betriebsausgaben gesamt', get: d => d.gesamtBetriebsausgaben || 0},

            // 4. Depreciation & Interest
            {label: 'Abschreibungen Maschinen', get: d => d.abschreibungenMaschinen || 0},
            {label: 'Abschreibungen Büroausstattung', get: d => d.abschreibungenBueromaterial || 0},
            {label: 'Abschreibungen immaterielle Wirtschaftsgüter', get: d => d.abschreibungenImmateriell || 0},
            {label: 'Zinsen auf Bankdarlehen', get: d => d.zinsenBank || 0},
            {label: 'Zinsen auf Gesellschafterdarlehen', get: d => d.zinsenGesellschafter || 0},
            {label: 'Leasingkosten', get: d => d.leasingkosten || 0},
            {label: 'Abschreibungen & Zinsen gesamt', get: d => d.gesamtAbschreibungenZinsen || 0},

            // 5. EBIT
            {label: 'Betriebsergebnis (EBIT)', get: d => d.ebit || 0},

            // 6. Net profit/loss (if included in BWA)
            {label: 'Jahresüberschuss/-fehlbetrag', get: d => d.gewinnNachSteuern || 0},
        ];

        // Header row and group hierarchy in one batch
        const headerRow = buildHeaderRow(config);

        // Optimization: Group hierarchy with better data structure
        const bwaGruppen = [
            {titel: 'Betriebserlöse (Einnahmen)', count: 12},
            {titel: 'Materialaufwand & Wareneinsatz', count: 4},
            {titel: 'Betriebsausgaben (Sachkosten)', count: 13},
            {titel: 'Abschreibungen & Zinsen', count: 7},
            {titel: 'Betriebsergebnis (EBIT)', count: 1},
            {titel: 'Jahresergebnis', count: 1},
        ];

        // Optimization: Pre-allocate output array with space
        // (headers + positions + group headers + empty lines)
        const totalRows = 1 + positions.length + bwaGruppen.length + (bwaGruppen.length - 1);
        const outputRows = new Array(totalRows);

        // Header as first row
        outputRows[0] = headerRow;

        // Output with group hierarchy
        let rowIndex = 1;
        let posIndex = 0;

        for (let gruppenIndex = 0; gruppenIndex < bwaGruppen.length; gruppenIndex++) {
            const gruppe = bwaGruppen[gruppenIndex];

            // Group heading
            outputRows[rowIndex++] = [
                `${gruppenIndex + 1}. ${gruppe.titel}`,
                ...Array(headerRow.length - 1).fill(''),
            ];

            // Group positions
            for (let i = 0; i < gruppe.count; i++) {
                outputRows[rowIndex++] = buildOutputRow(positions[posIndex++], bwaData);
            }

            // Empty line after each group except the last
            if (gruppenIndex < bwaGruppen.length - 1) {
                outputRows[rowIndex++] = Array(headerRow.length).fill('');
            }
        }

        // Create or update BWA sheet
        const bwaSheet = sheetUtils.getOrCreateSheet(ss, 'BWA');
        bwaSheet.clearContents();

        // Optimization: Write data in one batch API call
        bwaSheet.getRange(1, 1, outputRows.length, outputRows[0].length).setValues(outputRows);

        // Apply formatting in optimized batches
        applyBwaFormatting(bwaSheet, headerRow.length, bwaGruppen, outputRows.length);

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
 * Applies formatting to the BWA sheet with optimized batch formatting
 * @param {Sheet} sheet - Sheet to format
 * @param {number} headerLength - Number of columns
 * @param {Array} bwaGruppen - Group hierarchy
 * @param {number} totalRows - Total number of rows
 */
function applyBwaFormatting(sheet, headerLength, bwaGruppen, totalRows) {
    // Optimization: Group formatting for batch application
    const formatGroups = {
        // Header formatting
        header: {
            range: sheet.getRange(1, 1, 1, headerLength),
            formats: {
                fontWeight: 'bold',
                background: '#f3f3f3',
            },
        },

        // Group title formatting
        groupTitles: [],

        // Summary row formatting
        summaryRows: [],

        // EBIT and annual surplus formatting
        highlight: [],
    };

    // Find group title rows
    let rowIndex = 2;
    for (const gruppe of bwaGruppen) {
        formatGroups.groupTitles.push(sheet.getRange(rowIndex, 1));
        rowIndex += gruppe.count + 1; // +1 for the empty line
    }

    // Define summary rows based on BWA rows
    const summenZeilen = [12, 16, 29, 36, 37, 38]; // Adjusted for new structure with steuerlicheKorrekturen
    summenZeilen.forEach(row => {
        if (row <= totalRows) {
            formatGroups.summaryRows.push(sheet.getRange(row, 1, 1, headerLength));
        }
    });

    // Highlight EBIT and annual surplus
    [37, 38].forEach(row => { // Adjusted for new structure
        if (row <= totalRows) {
            formatGroups.highlight.push(sheet.getRange(row, 1, 1, headerLength));
        }
    });

    // Apply optimized batch formatting

    // 1. Format header
    formatGroups.header.range
        .setFontWeight(formatGroups.header.formats.fontWeight)
        .setBackground(formatGroups.header.formats.background);

    // 2. Format group titles
    if (formatGroups.groupTitles.length > 0) {
        formatGroups.groupTitles.forEach(range => {
            range.setFontWeight('bold');
        });
    }

    // 3. Format summary rows
    if (formatGroups.summaryRows.length > 0) {
        formatGroups.summaryRows.forEach(range => {
            range.setBackground('#e6f2ff');
        });
    }

    // 4. Highlight EBIT and annual surplus
    if (formatGroups.highlight.length > 0) {
        formatGroups.highlight.forEach(range => {
            range.setFontWeight('bold');
        });
    }

    // 5. Currency format for all numeric values in one batch
    sheet.getRange(2, 2, totalRows - 1, headerLength - 1).setNumberFormat('#,##0.00 €');

    // 6. Adjust column widths
    sheet.autoResizeColumns(1, headerLength);
}

export default {
    buildHeaderRow,
    buildOutputRow,
    generateBWASheet,
    applyBwaFormatting,
};