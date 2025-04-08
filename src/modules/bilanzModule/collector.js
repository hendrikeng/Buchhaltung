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
        processFinancialAssets(ss, bilanzData, einnahmenCols, config);

        // IV. Process Receivables (1700-2099)
        processReceivables(ss, bilanzData, einnahmenCols, config);

        // V. Process Payables (3300-4099)
        processPayables(ss, bilanzData, ausgabenCols, gesellschafterCols, config);

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
 * Processes fixed assets from expense sheet
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

    // DATEV account mapping for asset categories
    const assetCategories = {
        'Abschreibungen immaterielle Wirtschaftsgüter': {
            target: 'aktiva.anlagevermoegen.immaterielleVermoegensgegenstaende.konzessionen',
            kontoPrefix: '01', // SKR04 account prefix
        },
        'Abschreibungen Büroausstattung': {
            target: 'aktiva.anlagevermoegen.sachanlagen.andereAnlagen',
            kontoPrefix: '06', // SKR04 account prefix
        },
        'Abschreibungen Maschinen': {
            target: 'aktiva.anlagevermoegen.sachanlagen.technischeAnlagenMaschinen',
            kontoPrefix: '05', // SKR04 account prefix
        },
    };

    // Process rows
    for (const row of data) {
        const kategorie = row[ausgabenCols.kategorie - 1]?.toString().trim() || '';

        // Skip if not a fixed asset
        if (!assetCategories[kategorie]) continue;

        // For each asset type, we have expenses that actually represent assets
        // These should be in the balance sheet at purchase value minus accumulated depreciation
        const nettobetrag = numberUtils.parseCurrency(row[ausgabenCols.nettobetrag - 1]);

        // Skip if no amount
        if (nettobetrag <= 0) continue;

        // Get the target property in the balance sheet
        const targetPath = assetCategories[kategorie].target;
        const parts = targetPath.split('.');

        // Navigate to the target property and add the amount
        let current = bilanzData;
        for (let i = 0; i < parts.length; i++) {
            if (i === parts.length - 1) {
                current[parts[i]] += nettobetrag;
            } else {
                current = current[parts[i]];
            }
        }
    }
}

/**
 * Processes financial assets
 * @param {Spreadsheet} ss - Spreadsheet object
 * @param {Object} bilanzData - Balance sheet data
 * @param {Object} einnahmenCols - Revenue sheet column configuration
 * @param {Object} config - Configuration
 */
function processFinancialAssets(ss, bilanzData, einnahmenCols, config) {
    // Check if this is a holding company
    const isHolding = config.tax.isHolding || false;

    if (isHolding) {
        const holdingSheet = ss.getSheetByName('Holding Transfers');
        if (!holdingSheet || holdingSheet.getLastRow() <= 1) return;

        // For holding companies, investments in subsidiaries are financial assets
        // Process holdings data
        const data = holdingSheet.getDataRange().getValues().slice(1); // Skip header
        const holdingCols = config.holdingTransfers.columns;

        let totalInvestments = 0;

        for (const row of data) {
            const kategorie = row[holdingCols.kategorie - 1]?.toString().trim() || '';
            const betrag = numberUtils.parseCurrency(row[holdingCols.betrag - 1]);

            if (kategorie === 'Kapitalrückführung' && betrag < 0) {
                // Capital reductions decrease investments
                totalInvestments -= Math.abs(betrag);
            } else if (kategorie === 'Kapitaleinlage' && betrag > 0) {
                // Capital increases add to investments
                totalInvestments += betrag;
            }
        }

        // Set the investments value in the balance sheet
        bilanzData.aktiva.anlagevermoegen.finanzanlagen.anteileVerbundeneUnternehmen = totalInvestments;
    }

    // Process loans and similar financial assets
    const gesellschafterSheet = ss.getSheetByName('Gesellschafterkonto');
    if (!gesellschafterSheet || gesellschafterSheet.getLastRow() <= 1) return;

    const gesellschafterCols = config.gesellschafterkonto.columns;
    const data = gesellschafterSheet.getDataRange().getValues().slice(1); // Skip header

    for (const row of data) {
        const kategorie = row[gesellschafterCols.kategorie - 1]?.toString().trim() || '';
        const betrag = numberUtils.parseCurrency(row[gesellschafterCols.betrag - 1]);

        if (kategorie === 'Gesellschafterdarlehen' && betrag > 0) {
            // Loans to shareholders are financial assets
            bilanzData.aktiva.anlagevermoegen.finanzanlagen.ausleihungenVerbundene += betrag;
        }
    }
}

/**
 * Processes receivables
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

    let totalReceivables = 0;

    for (const row of data) {
        const nettobetrag = numberUtils.parseCurrency(row[einnahmenCols.nettobetrag - 1]);
        const bruttobetrag = numberUtils.parseCurrency(row[einnahmenCols.bruttoBetrag - 1]);
        const bezahlt = numberUtils.parseCurrency(row[einnahmenCols.bezahlt - 1]);

        // Check if invoice has unpaid amount
        if (bezahlt < bruttobetrag) {
            const outstanding = bruttobetrag - bezahlt;
            totalReceivables += outstanding;
        }
    }

    // Set receivables in balance sheet
    bilanzData.aktiva.umlaufvermoegen.forderungen.forderungenLuL = totalReceivables;
}

/**
 * Processes payables
 * @param {Spreadsheet} ss - Spreadsheet object
 * @param {Object} bilanzData - Balance sheet data
 * @param {Object} ausgabenCols - Expense sheet column configuration
 * @param {Object} gesellschafterCols - Shareholder account column configuration
 * @param {Object} config - Configuration
 */
function processPayables(ss, bilanzData, ausgabenCols, gesellschafterCols, config) {
    // Process accounts payable from expenses
    const ausSheet = ss.getSheetByName('Ausgaben');
    if (ausSheet && ausSheet.getLastRow() > 1) {
        const data = ausSheet.getDataRange().getValues().slice(1); // Skip header

        let totalPayables = 0;

        for (const row of data) {
            const bruttobetrag = numberUtils.parseCurrency(row[ausgabenCols.bruttoBetrag - 1]);
            const bezahlt = numberUtils.parseCurrency(row[ausgabenCols.bezahlt - 1]);

            // Check if invoice has unpaid amount
            if (bezahlt < bruttobetrag) {
                const outstanding = bruttobetrag - bezahlt;
                totalPayables += outstanding;
            }
        }

        // Set accounts payable in balance sheet
        bilanzData.passiva.verbindlichkeiten.verbindlichkeitenLuL = totalPayables;
    }

    // Process shareholder loans (liabilities)
    const gesellschafterSheet = ss.getSheetByName('Gesellschafterkonto');
    if (gesellschafterSheet && gesellschafterSheet.getLastRow() > 1) {
        const data = gesellschafterSheet.getDataRange().getValues().slice(1); // Skip header

        let shareholderLoans = 0;

        for (const row of data) {
            const kategorie = row[gesellschafterCols.kategorie - 1]?.toString().trim() || '';
            const betrag = numberUtils.parseCurrency(row[gesellschafterCols.betrag - 1]);

            if (kategorie === 'Gesellschafterdarlehen' && betrag < 0) {
                // Loans from shareholders are liabilities
                shareholderLoans += Math.abs(betrag);
            }
        }

        // Set shareholder loans in balance sheet
        bilanzData.passiva.verbindlichkeiten.sonstigeVerbindlichkeiten = shareholderLoans;
    }

    // Process bank loans
    const bankLoans = calculator.getBankDarlehen(ss, config);
    bilanzData.passiva.verbindlichkeiten.verbindlichkeitenKreditinstitute = bankLoans;
}

/**
 * Processes equity
 * @param {Spreadsheet} ss - Spreadsheet object
 * @param {Object} bilanzData - Balance sheet data
 * @param {Object} gesellschafterCols - Shareholder account column configuration
 * @param {Object} holdingCols - Holding transfers column configuration
 * @param {Object} config - Configuration
 */
function processEquity(ss, bilanzData, gesellschafterCols, holdingCols, config) {
    // Set share capital from configuration
    bilanzData.passiva.eigenkapital.gezeichnetesKapital = config.tax.stammkapital || 25000;

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

    // Process capital reserve and other equity items
    const gesellschafterSheet = ss.getSheetByName('Gesellschafterkonto');
    if (gesellschafterSheet && gesellschafterSheet.getLastRow() > 1) {
        const data = gesellschafterSheet.getDataRange().getValues().slice(1); // Skip header

        for (const row of data) {
            const kategorie = row[gesellschafterCols.kategorie - 1]?.toString().trim() || '';
            const betrag = numberUtils.parseCurrency(row[gesellschafterCols.betrag - 1]);

            if (kategorie === 'Kapitalrücklage' && betrag != 0) {
                bilanzData.passiva.eigenkapital.kapitalruecklage += betrag;
            } else if (kategorie === 'Gewinnvortrag' && betrag != 0) {
                bilanzData.passiva.eigenkapital.gewinnvortrag += betrag;
            } else if (kategorie === 'Andere Gewinnrücklagen' && betrag != 0) {
                bilanzData.passiva.eigenkapital.gewinnruecklagen.andereGewinnruecklagen += betrag;
            }
        }
    }
}

/**
 * Processes provisions
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

    let taxProvisions = 0;
    let otherProvisions = 0;

    for (const row of data) {
        const kategorie = row[ausgabenCols.kategorie - 1]?.toString().trim() || '';
        const nettobetrag = numberUtils.parseCurrency(row[ausgabenCols.nettobetrag - 1]);

        if (taxProvisionCategories.has(kategorie)) {
            taxProvisions += nettobetrag;
        } else if (kategorie.includes('Rückstellung')) {
            otherProvisions += nettobetrag;
        }
    }

    // Set provisions in balance sheet
    bilanzData.passiva.rueckstellungen.steuerrueckstellungen = taxProvisions;
    bilanzData.passiva.rueckstellungen.sonstigeRueckstellungen = otherProvisions;
}

export default {
    aggregateBilanzData,
};