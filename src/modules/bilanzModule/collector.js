// modules/bilanzModule/collector.js
import globalCache from '../../utils/cacheUtils.js';
import dataModel from './dataModel.js';
import calculator from './calculator.js';
import numberUtils from '../../utils/numberUtils.js';

/**
 * Collects data from various sheets for the balance sheet with DATEV-compliant structure
 * @param {Object} config - Configuration
 * @returns {Object} Balance sheet data structure with populated values
 */
function aggregateBilanzData(config) {
    try {
        // Check if cache is valid
        if (globalCache.has('computed', 'bilanz')) {
            console.log('Using cached bilanz data');
            return globalCache.get('computed', 'bilanz');
        }

        console.log('Aggregating DATEV-compliant balance sheet data...');
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const bilanzData = dataModel.createEmptyBilanz();

        // Get column configurations for the different sheets
        const bankCols = config.bankbewegungen.columns;
        const ausgabenCols = config.ausgaben.columns;
        const einnahmenCols = config.einnahmen.columns;
        const gesellschafterCols = config.gesellschafterkonto.columns;
        const holdingCols = config.holdingTransfers.columns;

        // I. Process Bank Accounts (1200-1299)
        processBankAccounts(ss, bilanzData, bankCols);

        // II. Process Fixed Assets (0100-0799)
        processFixedAssets(ss, bilanzData, ausgabenCols, config);

        // III. Process Financial Assets (0800-1199)
        processFinancialAssets(ss, bilanzData, config);

        // IV. Process Receivables (1700-2099)
        processReceivables(ss, bilanzData, einnahmenCols, config);

        // V. Process Payables (3300-4099)
        processPayables(ss, bilanzData, ausgabenCols, config);

        // VI. Process Equity (2700-2999)
        processEquity(ss, bilanzData, gesellschafterCols, holdingCols, config);

        // VII. Process Provisions (3000-3299)
        processProvisions(ss, bilanzData, ausgabenCols, config);

        // Calculate totals and subtotals for all sections
        calculator.calculateBilanzTotals(bilanzData);

        // Cache the results
        globalCache.set('computed', 'bilanz', bilanzData);
        console.log('Balance sheet data aggregation complete');

        return bilanzData;
    } catch (e) {
        console.error('Error collecting balance sheet data:', e);
        return null;
    }
}

/**
 * Processes bank accounts from bank sheet
 * @param {Spreadsheet} ss - Spreadsheet object
 * @param {Object} bilanzData - Balance sheet data
 * @param {Object} bankCols - Bank sheet column configuration
 */
function processBankAccounts(ss, bilanzData, bankCols) {
    const bankSheet = ss.getSheetByName('Bankbewegungen');
    if (!bankSheet) return;

    const lastRow = bankSheet.getLastRow();
    if (lastRow <= 1) return;

    try {
        // Find the last row (Endsaldo)
        const lastRowData = bankSheet.getRange(lastRow, 1, 1, bankCols.buchungstext).getValues()[0];
        const lastRowText = lastRowData[bankCols.buchungstext - 1]?.toString().trim().toLowerCase() || '';

        if (lastRowText === 'endsaldo') {
            // Get the Saldo value from the last row
            const saldo = numberUtils.parseCurrency(
                bankSheet.getRange(lastRow, bankCols.saldo).getValue(),
            );
            bilanzData.aktiva.umlaufvermoegen.liquideMittel.bankguthaben = saldo;
        }
    } catch (e) {
        console.error('Error processing bank accounts:', e);
    }
}

/**
 * Processes fixed assets from expense sheet using bilanzMapping
 * @param {Spreadsheet} ss - Spreadsheet object
 * @param {Object} bilanzData - Balance sheet data
 * @param {Object} ausgabenCols - Expense sheet column configuration
 * @param {Object} config - Configuration
 */
function processFixedAssets(ss, bilanzData, ausgabenCols, config) {
    const ausSheet = ss.getSheetByName('Ausgaben');
    if (!ausSheet || ausSheet.getLastRow() <= 1) return;

    // Load data in one batch
    const data = ausSheet.getDataRange().getValues().slice(1); // Skip header

    // Process rows using bilanzMapping from config
    for (const row of data) {
        const kategorie = row[ausgabenCols.kategorie - 1]?.toString().trim() || '';
        if (!kategorie) continue;

        // Get category configuration with bilanzMapping
        const categoryConfig = config.ausgaben.categories[kategorie];
        if (!categoryConfig || !categoryConfig.bilanzMapping) continue;

        // Skip categories that don't belong in fixed assets
        if (!categoryConfig.bilanzMapping.positiv?.includes('anlagevermoegen')) continue;

        // For fixed assets, we need the amount for the balance sheet
        const nettobetrag = numberUtils.parseCurrency(row[ausgabenCols.nettobetrag - 1]);
        if (nettobetrag <= 0) continue;

        // Apply the amount to the correct position in the balance sheet based on the mapping
        applyToBilanzPosition(bilanzData, categoryConfig.bilanzMapping.positiv, nettobetrag);
    }
}

/**
 * Processes financial assets using bilanzMapping with improved handling
 * @param {Spreadsheet} ss - Spreadsheet object
 * @param {Object} bilanzData - Balance sheet data
 * @param {Object} config - Configuration
 */
function processFinancialAssets(ss, bilanzData, config) {
    // Process Holding Transfers
    const holdingSheet = ss.getSheetByName('Holding Transfers');
    if (holdingSheet && holdingSheet.getLastRow() > 1) {
        const holdingCols = config.holdingTransfers.columns;
        const data = holdingSheet.getDataRange().getValues().slice(1); // Skip header
        const isHolding = config.tax.isHolding || false;

        // Group transfers by category for better handling
        const transfersByCategory = {};

        for (const row of data) {
            const kategorie = row[holdingCols.kategorie - 1]?.toString().trim() || '';
            if (!kategorie) continue;

            const betrag = numberUtils.parseCurrency(row[holdingCols.betrag - 1]);

            if (!transfersByCategory[kategorie]) {
                transfersByCategory[kategorie] = [];
            }

            transfersByCategory[kategorie].push({
                betrag: betrag,
                ziel: row[holdingCols.ziel - 1],
            });
        }

        // Process each category separately
        for (const [kategorie, transfers] of Object.entries(transfersByCategory)) {
            // Get category configuration with bilanzMapping
            const categoryConfig = config.holdingTransfers.categories[kategorie];
            if (!categoryConfig || !categoryConfig.bilanzMapping) continue;

            // Calculate net amount for this category
            let netAmount = 0;
            for (const transfer of transfers) {
                netAmount += transfer.betrag;
            }

            // Skip if net amount is zero
            if (netAmount === 0) continue;

            // Use the appropriate mapping based on amount sign and company type
            const mappingKey = netAmount > 0 ? 'positiv' : 'negativ';
            const bilanzPosition = categoryConfig.bilanzMapping[mappingKey];

            // Apply the amount to the correct position in the balance sheet
            if (bilanzPosition) {
                applyToBilanzPosition(bilanzData, bilanzPosition, Math.abs(netAmount));
            }
        }
    }

    // Note: Gesellschafterdarlehen werden jetzt in processEquity verarbeitet
    // um Doppelzählungen zu vermeiden und eine korrekte Saldierung sicherzustellen
}

/**
 * Processes receivables using bilanzMapping
 * @param {Spreadsheet} ss - Spreadsheet object
 * @param {Object} bilanzData - Balance sheet data
 * @param {Object} einnahmenCols - Revenue sheet column configuration
 * @param {Object} config - Configuration
 */
function processReceivables(ss, bilanzData, einnahmenCols, config) {
    const einnahmenSheet = ss.getSheetByName('Einnahmen');
    if (!einnahmenSheet || einnahmenSheet.getLastRow() <= 1) return;

    // Process all receivables
    const data = einnahmenSheet.getDataRange().getValues().slice(1); // Skip header

    // Group by category mapping
    const categoryTotals = {};

    for (const row of data) {
        const kategorie = row[einnahmenCols.kategorie - 1]?.toString().trim() || '';
        if (!kategorie) continue;

        // Get category configuration with bilanzMapping
        const categoryConfig = config.einnahmen.categories[kategorie];
        if (!categoryConfig || !categoryConfig.bilanzMapping) continue;

        // Check if invoice has unpaid amount
        const bruttobetrag = numberUtils.parseCurrency(row[einnahmenCols.bruttoBetrag - 1]);
        const bezahlt = numberUtils.parseCurrency(row[einnahmenCols.bezahlt - 1]);

        if (bezahlt < bruttobetrag) {
            const outstanding = bruttobetrag - bezahlt;
            const isNegative = bruttobetrag < 0; // Check if it's a credit note

            // Determine mapping key based on amount sign
            const mappingKey = isNegative ? 'negativ' : 'positiv';
            const bilanzPosition = categoryConfig.bilanzMapping[mappingKey];

            if (bilanzPosition) {
                if (!categoryTotals[bilanzPosition]) {
                    categoryTotals[bilanzPosition] = 0;
                }
                categoryTotals[bilanzPosition] += outstanding;
            }
        }
    }

    // Apply totals to balance sheet
    Object.entries(categoryTotals).forEach(([position, amount]) => {
        applyToBilanzPosition(bilanzData, position, amount);
    });
}

/**
 * Processes payables using bilanzMapping
 * @param {Spreadsheet} ss - Spreadsheet object
 * @param {Object} bilanzData - Balance sheet data
 * @param {Object} ausgabenCols - Expense sheet column configuration
 * @param {Object} config - Configuration
 */
function processPayables(ss, bilanzData, ausgabenCols, config) {
    // Process accounts payable from expenses
    const ausSheet = ss.getSheetByName('Ausgaben');
    if (ausSheet && ausSheet.getLastRow() > 1) {
        const data = ausSheet.getDataRange().getValues().slice(1); // Skip header

        // Group by category mapping
        const categoryTotals = {};

        for (const row of data) {
            const kategorie = row[ausgabenCols.kategorie - 1]?.toString().trim() || '';
            if (!kategorie) continue;

            // Get category configuration with bilanzMapping
            const categoryConfig = config.ausgaben.categories[kategorie];
            if (!categoryConfig || !categoryConfig.bilanzMapping) continue;

            // Check if invoice has unpaid amount
            const bruttobetrag = numberUtils.parseCurrency(row[ausgabenCols.bruttoBetrag - 1]);
            const bezahlt = numberUtils.parseCurrency(row[ausgabenCols.bezahlt - 1]);

            if (bezahlt < bruttobetrag) {
                const outstanding = bruttobetrag - bezahlt;

                // For payables, we use the positive mapping key (liabilities are positive in our model)
                const bilanzPosition = categoryConfig.bilanzMapping.positiv;

                if (bilanzPosition) {
                    if (!categoryTotals[bilanzPosition]) {
                        categoryTotals[bilanzPosition] = 0;
                    }
                    categoryTotals[bilanzPosition] += outstanding;
                }
            }
        }

        // Apply totals to balance sheet
        Object.entries(categoryTotals).forEach(([position, amount]) => {
            applyToBilanzPosition(bilanzData, position, amount);
        });
    }

    // Process own receipts (Eigenbelege) that are unpaid
    const eigenSheet = ss.getSheetByName('Eigenbelege');
    if (eigenSheet && eigenSheet.getLastRow() > 1) {
        const eigenCols = config.eigenbelege.columns;
        const data = eigenSheet.getDataRange().getValues().slice(1); // Skip header

        for (const row of data) {
            const kategorie = row[eigenCols.kategorie - 1]?.toString().trim() || '';
            if (!kategorie) continue;

            // Get category configuration with bilanzMapping
            const categoryConfig = config.eigenbelege.categories[kategorie];
            if (!categoryConfig || !categoryConfig.bilanzMapping) continue;

            // Check if receipt has unpaid amount
            const bruttobetrag = numberUtils.parseCurrency(row[eigenCols.bruttoBetrag - 1]);
            const bezahlt = numberUtils.parseCurrency(row[eigenCols.bezahlt - 1]);

            if (bezahlt < bruttobetrag) {
                const outstanding = bruttobetrag - bezahlt;

                // For payables, we use the positive mapping key
                const bilanzPosition = categoryConfig.bilanzMapping.positiv;

                // Apply to balance sheet
                applyToBilanzPosition(bilanzData, bilanzPosition, outstanding);
            }
        }
    }
}

/**
 * Processes equity using bilanzMapping with improved handling of transactions
 * @param {Spreadsheet} ss - Spreadsheet object
 * @param {Object} bilanzData - Balance sheet data
 * @param {Object} gesellschafterCols - Shareholder account column configuration
 * @param {Object} holdingCols - Holding transfers column configuration
 * @param {Object} config - Configuration
 */
function processEquity(ss, bilanzData, gesellschafterCols, holdingCols, config) {
    // Start with base share capital from configuration
    const basisStammkapital = config.tax.stammkapital || 25000;

    // Initialize with default values
    bilanzData.passiva.eigenkapital.gezeichnetesKapital = basisStammkapital;

    // Track capital changes specifically related to gezeichnetes Kapital (formal capital changes)
    let kapitalErhoehnungen = 0;
    let kapitalHerabsetzungen = 0;

    // Process transactions from Gesellschafterkonto
    const gesellschafterSheet = ss.getSheetByName('Gesellschafterkonto');
    if (gesellschafterSheet && gesellschafterSheet.getLastRow() > 1) {
        const data = gesellschafterSheet.getDataRange().getValues().slice(1); // Skip header

        // Group all transactions by category for better handling
        const transaktionenByKategorie = {};

        // Track balancing transactions keyed by their reference transaction
        const ausgleichsbuchungen = new Map();

        // First pass: Identify all balancing transactions
        for (const row of data) {
            const transaktionstyp = row[gesellschafterCols.transaktionstyp - 1]?.toString().trim() || '';
            if (transaktionstyp === 'Ausgleichsbuchung') {
                const bezugRef = row[gesellschafterCols.bezugstransaktion - 1]?.toString().trim() || '';
                if (bezugRef) {
                    const betrag = numberUtils.parseCurrency(row[gesellschafterCols.betrag - 1]);
                    if (!ausgleichsbuchungen.has(bezugRef)) {
                        ausgleichsbuchungen.set(bezugRef, 0);
                    }
                    // Add the amount to the running total for this reference
                    ausgleichsbuchungen.set(bezugRef, ausgleichsbuchungen.get(bezugRef) + Math.abs(betrag));
                }
            }
        }

        // Second pass: Process regular transactions
        for (const row of data) {
            const kategorie = row[gesellschafterCols.kategorie - 1]?.toString().trim() || '';
            if (!kategorie) continue;

            const transaktionstyp = row[gesellschafterCols.transaktionstyp - 1]?.toString().trim() || '';
            if (transaktionstyp === 'Ausgleichsbuchung') continue; // Skip balancing transactions

            const betrag = numberUtils.parseCurrency(row[gesellschafterCols.betrag - 1]);
            const referenz = row[gesellschafterCols.referenz - 1]?.toString().trim() || '';

            // Check if this transaction has been balanced by an Ausgleichsbuchung
            let effectiveBetrag = betrag;
            if (referenz && ausgleichsbuchungen.has(referenz)) {
                const ausgleichsAmount = ausgleichsbuchungen.get(referenz);
                // Calculate the remaining amount after balancing
                effectiveBetrag = betrag > 0 ?
                    Math.max(0, betrag - ausgleichsAmount) :
                    Math.min(0, betrag + ausgleichsAmount);
            }

            // Skip if fully balanced (effectiveBetrag is 0)
            if (effectiveBetrag === 0) continue;

            // Add to transaction list by category
            if (!transaktionenByKategorie[kategorie]) {
                transaktionenByKategorie[kategorie] = [];
            }

            transaktionenByKategorie[kategorie].push({
                betrag: effectiveBetrag,
                datum: row[gesellschafterCols.datum - 1],
                gesellschafter: row[gesellschafterCols.gesellschafter - 1],
                rowData: row,
                // Store original betrag for reference
                originalBetrag: betrag,
            });
        }

        // Process shareholder loans (Gesellschafterdarlehen)
        if (transaktionenByKategorie['Gesellschafterdarlehen']) {
            let darlehenSaldo = 0;

            // Sum all effective amounts
            for (const transaktion of transaktionenByKategorie['Gesellschafterdarlehen']) {
                darlehenSaldo += transaktion.betrag;
            }

            // For Gesellschafterdarlehen, negative means company gave loan to shareholder
            if (darlehenSaldo < 0) {
                // Company loaned money to shareholders (asset)
                bilanzData.aktiva.anlagevermoegen.finanzanlagen.ausleihungenVerbundene += Math.abs(darlehenSaldo);
            } else if (darlehenSaldo > 0) {
                // Shareholders loaned money to company (liability)
                bilanzData.passiva.verbindlichkeiten.sonstigeVerbindlichkeiten += darlehenSaldo;
            }
        }

        // Process formal capital increases
        if (transaktionenByKategorie['Kapitalerhöhung']) {
            for (const transaktion of transaktionenByKategorie['Kapitalerhöhung']) {
                kapitalErhoehnungen += transaktion.betrag;
            }
        }

        // Process formal capital decreases
        if (transaktionenByKategorie['Kapitalherabsetzung']) {
            for (const transaktion of transaktionenByKategorie['Kapitalherabsetzung']) {
                kapitalHerabsetzungen += Math.abs(transaktion.betrag);
            }
        }

        // Process other equity transactions using their bilanzMapping
        for (const [kategorie, transaktionen] of Object.entries(transaktionenByKategorie)) {
            // Skip categories we've already processed specifically
            if (kategorie === 'Gesellschafterdarlehen' ||
                kategorie === 'Kapitalerhöhung' ||
                kategorie === 'Kapitalherabsetzung') {
                continue;
            }

            // Get category configuration with bilanzMapping
            const categoryConfig = config.gesellschafterkonto.categories[kategorie];
            if (!categoryConfig || !categoryConfig.bilanzMapping) continue;

            // Calculate total amount for this category (already accounts for balancing)
            let gesamtBetrag = 0;
            for (const transaktion of transaktionen) {
                gesamtBetrag += transaktion.betrag;
            }

            // Skip if total amount is zero
            if (gesamtBetrag === 0) continue;

            // Determine mapping key based on amount sign
            const mappingKey = gesamtBetrag > 0 ? 'positiv' : 'negativ';
            const bilanzPosition = categoryConfig.bilanzMapping[mappingKey];

            // Apply to balance sheet if position is defined
            if (bilanzPosition) {
                applyToBilanzPosition(bilanzData, bilanzPosition, Math.abs(gesamtBetrag));
            }
        }
    }

    // Update gezeichnetes Kapital with formal changes
    bilanzData.passiva.eigenkapital.gezeichnetesKapital = basisStammkapital + kapitalErhoehnungen - kapitalHerabsetzungen;

    // Process retained earnings from BWA sheet
    const bwaSheet = ss.getSheetByName('BWA');
    if (bwaSheet) {
        try {
            // Find the rows for result data
            const lastCol = bwaSheet.getLastColumn();
            const yearColIndex = lastCol; // Year total in last column

            // Find net income row
            const rows = bwaSheet.getDataRange().getValues();
            let netIncomeRow = -1;

            for (let i = 0; i < rows.length; i++) {
                if (rows[i][0].toString().toLowerCase().includes('jahresüberschuss')) {
                    netIncomeRow = i + 1; // 1-based index
                    break;
                }
            }

            if (netIncomeRow > 0) {
                const netIncome = numberUtils.parseCurrency(
                    bwaSheet.getRange(netIncomeRow, yearColIndex).getValue(),
                );
                bilanzData.passiva.eigenkapital.jahresueberschuss = netIncome;
            }
        } catch (e) {
            console.error('Error processing equity from BWA:', e);
        }
    }
}

/**
 * Processes provisions using bilanzMapping
 * @param {Spreadsheet} ss - Spreadsheet object
 * @param {Object} bilanzData - Balance sheet data
 * @param {Object} ausgabenCols - Expense sheet column configuration
 * @param {Object} config - Configuration
 */
function processProvisions(ss, bilanzData, ausgabenCols, config) {
    // Process tax provisions from expenses
    const ausSheet = ss.getSheetByName('Ausgaben');
    if (!ausSheet || ausSheet.getLastRow() <= 1) return;

    const data = ausSheet.getDataRange().getValues().slice(1); // Skip header

    // Map for tax provision categories
    const taxProvisionCategories = new Set([
        'Gewerbesteuerrückstellungen',
        'Körperschaftsteuer',
        'Solidaritätszuschlag',
        'Sonstige Steuerrückstellungen',
    ]);

    // Group by bilanzMapping position
    const positionTotals = {};

    for (const row of data) {
        const kategorie = row[ausgabenCols.kategorie - 1]?.toString().trim() || '';
        if (!kategorie) continue;

        // Check if it's a provision category
        if (!kategorie.includes('Rückstellung') && !taxProvisionCategories.has(kategorie)) {
            continue;
        }

        // Get category configuration with bilanzMapping
        const categoryConfig = config.ausgaben.categories[kategorie];
        if (!categoryConfig || !categoryConfig.bilanzMapping) continue;

        const nettobetrag = numberUtils.parseCurrency(row[ausgabenCols.nettobetrag - 1]);
        if (nettobetrag <= 0) continue;

        // For provisions, we use the positive mapping key
        const bilanzPosition = categoryConfig.bilanzMapping.positiv;

        if (bilanzPosition) {
            if (!positionTotals[bilanzPosition]) {
                positionTotals[bilanzPosition] = 0;
            }
            positionTotals[bilanzPosition] += nettobetrag;
        }
    }

    // Apply totals to balance sheet
    Object.entries(positionTotals).forEach(([position, amount]) => {
        applyToBilanzPosition(bilanzData, position, amount);
    });
}

/**
 * Helper function to apply an amount to a specific position in the balance sheet structure
 * @param {Object} bilanzData - Balance sheet data object
 * @param {string} position - Dot-notation path to the position in the balance sheet
 * @param {number} amount - Amount to add to the position
 */
function applyToBilanzPosition(bilanzData, position, amount) {
    if (!position) return;

    // Navigate to the target property and add the amount
    const parts = position.split('.');
    let current = bilanzData;

    for (let i = 0; i < parts.length; i++) {
        if (i === parts.length - 1) {
            current[parts[i]] += amount;
        } else {
            current = current[parts[i]];
            if (!current) {
                console.warn(`Invalid bilanzMapping path: ${position}`);
                return;
            }
        }
    }
}

export default {
    aggregateBilanzData,
};