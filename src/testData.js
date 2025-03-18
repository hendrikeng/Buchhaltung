/**
 * Test Data Generator for Buchhaltung App
 *
 * This script generates test data for the accounting application to
 * test all features including:
 * - Import module
 * - UStVA calculation
 * - BWA calculation
 * - Balance sheet generation
 * - Bank transaction matching
 * - Document validation
 *
 * Instructions:
 * 1. Copy this entire script into the Google Apps Script editor of your spreadsheet
 * 2. Run the 'onOpen' function once to add the menu
 * 3. Use the new menu "Test Data" -> "Generate Test Data" to create the test data
 * 4. Use the "Buchhaltung" menu to test the app functionality
 */

// Create menu when the spreadsheet is opened
function onOpen() {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('Test Data')
        .addItem('Generate Test Data', 'generateTestData')
        .addSeparator()
        .addItem('Clear All Data', 'clearAllData')
        .addToUi();
}

/**
 * Clears all data from all sheets
 */
function clearAllData() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    clearAllSheets(ss);
    SpreadsheetApp.getUi().alert('All data has been cleared!');
}

function generateTestData() {
    // Get active spreadsheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Clear existing data
    clearAllSheets(ss);

    // Generate test data for each sheet
    generateEinnahmenData(ss);
    generateAusgabenData(ss);
    generateEigenbelegeData(ss);
    generateGesellschafterkontoData(ss);
    generateHoldingTransfersData(ss);
    generateBankbewegungenData(ss);

    // Refresh all sheets to update calculations
    refreshSheets(ss);

    // Show completion message
    SpreadsheetApp.getUi().alert('Test data generation complete!');
}

/**
 * Clears all sheets from existing data (keeps headers)
 */
function clearAllSheets(ss) {
    const sheets = [
        'Einnahmen',
        'Ausgaben',
        'Eigenbelege',
        'Gesellschafterkonto',
        'Holding Transfers',
        'Bankbewegungen',
        'UStVA',
        'BWA',
        'Bilanz',
    ];

    for (const sheetName of sheets) {
        const sheet = ss.getSheetByName(sheetName);
        if (sheet) {
            if (sheet.getLastRow() > 1) {
                sheet.deleteRows(2, sheet.getLastRow() - 1);
            } else if (sheet.getLastRow() === 0) {
                // Empty sheet, do nothing
            } else {
                // Just clear content but keep header
                sheet.getRange(2, 1, sheet.getMaxRows() - 1, sheet.getMaxColumns()).clearContent();
            }
        }
    }
}

/**
 * Helper function to refresh sheets after data generation
 */
function refreshSheets(ss) {
    try {
        // Explicit formula calculation for all sheets
        ss.getSheets().forEach(sheet => {
            if (sheet.getLastRow() > 1) {
                sheet.getDataRange().activate();
                SpreadsheetApp.flush();
            }
        });

        // Short sleep to allow calculations to complete
        Utilities.sleep(1000);

    } catch (e) {
        console.error('Error refreshing sheets:', e);
    }
}

/**
 * Generate revenue data
 */
function generateEinnahmenData(ss) {
    const sheet = ss.getSheetByName('Einnahmen');
    if (!sheet) return;

    const testData = [
        // Standard invoices covering all categories with different VAT rates
        ['01.01.2021', 'RE-2021-001', 'Kunde A GmbH', 'Webdesign', 'Test Rechnung 1', 'Erlöse aus Lieferungen und Leistungen', '', 1000, 19, '', '', 1190, 0, '', 'Offen', '', '', '', 2021, new Date()],
        ['15.01.2021', 'RE-2021-002', 'Kunde B AG', 'SEO-Optimierung', '', 'Erlöse aus Lieferungen und Leistungen', '', 2500, 19, '', '', 2975, 2975, 0, 'Bezahlt', 'Überweisung', '20.01.2021', '✓ Bank: 20.01.2021', 2021, new Date()],
        ['02.02.2021', 'RE-2021-003', 'Kunde C Ltd', 'Beratung', '', 'Provisionserlöse', '', 1500, 19, '', '', 1785, 1785, 0, 'Bezahlt', 'Überweisung', '10.02.2021', '✓ Bank: 10.02.2021', 2021, new Date()],
        ['15.02.2021', 'RE-2021-004', 'Kunde D KG', 'Bücher', '', 'Erlöse aus Lieferungen und Leistungen', '', 500, 7, '', '', 535, 0, '', 'Offen', '', '', '', 2021, new Date()],
        ['01.03.2021', 'RE-2021-005', 'Kunde E', 'Beratung Ausland', '', 'Steuerfreie Inland-Einnahmen', '', 3000, 0, '', '', 3000, 3000, 0, 'Bezahlt', 'Überweisung', '15.03.2021', '✓ Bank: 15.03.2021', 2021, new Date()],
        ['15.03.2021', 'RE-2021-006', 'Kunde F', 'Vermietung Büroraum', '', 'Erträge aus Vermietung/Verpachtung', '', 800, 0, '', '', 800, 800, 0, 'Bezahlt', 'Überweisung', '01.04.2021', '✓ Bank: 01.04.2021', 2021, new Date()],

        // Partial payments to test matching algorithm
        ['01.04.2021', 'RE-2021-007', 'Kunde G', 'Softwareentwicklung', '', 'Erlöse aus Lieferungen und Leistungen', '', 5000, 19, '', '', 5950, 2975, 2500, 'Teilbezahlt', 'Überweisung', '15.04.2021', '✓ Bank: 15.04.2021', 2021, new Date()],

        // Credit notes to test gutschrift handling
        ['05.04.2021', 'GS-2021-001', 'Kunde A GmbH', 'Korrektur zu RE-2021-001', 'Gutschrift', 'Erlöse aus Lieferungen und Leistungen', '', -200, 19, '', '', -238, -238, 0, 'Bezahlt', 'Überweisung', '10.04.2021', '✓ Bank: 10.04.2021', 2021, new Date()],

        // Entries for each quarter to test quarterly reports
        ['15.04.2021', 'RE-2021-008', 'Kunde H GmbH', 'Schulung', '', 'Erlöse aus Lieferungen und Leistungen', '', 1200, 19, '', '', 1428, 1428, 0, 'Bezahlt', 'Überweisung', '20.04.2021', '✓ Bank: 20.04.2021', 2021, new Date()],
        ['03.05.2021', 'RE-2021-009', 'Kunde I', 'Beratung', '', 'Erlöse aus Lieferungen und Leistungen', '', 1800, 19, '', '', 2142, 0, '', 'Offen', '', '', '', 2021, new Date()],
        ['18.06.2021', 'RE-2021-010', 'Kunde J', 'Schulung', '', 'Erlöse aus Lieferungen und Leistungen', '', 2200, 19, '', '', 2618, 2618, 0, 'Bezahlt', 'Überweisung', '25.06.2021', '✓ Bank: 25.06.2021', 2021, new Date()],
        ['05.07.2021', 'RE-2021-011', 'Kunde K', 'Workshop', '', 'Erlöse aus Lieferungen und Leistungen', '', 3000, 19, '', '', 3570, 3570, 0, 'Bezahlt', 'Überweisung', '12.07.2021', '✓ Bank: 12.07.2021', 2021, new Date()],
        ['15.08.2021', 'RE-2021-012', 'Kunde L', 'Support', '', 'Erlöse aus Lieferungen und Leistungen', '', 1500, 19, '', '', 1785, 0, '', 'Offen', '', '', '', 2021, new Date()],
        ['20.09.2021', 'RE-2021-013', 'Kunde M', 'IT-Dienstleistung', '', 'Erlöse aus Lieferungen und Leistungen', '', 2800, 19, '', '', 3332, 3332, 0, 'Bezahlt', 'Überweisung', '01.10.2021', '✓ Bank: 01.10.2021', 2021, new Date()],
        ['10.10.2021', 'RE-2021-014', 'Kunde N', 'Marketing', '', 'Erlöse aus Lieferungen und Leistungen', '', 1800, 19, '', '', 2142, 2142, 0, 'Bezahlt', 'Überweisung', '20.10.2021', '✓ Bank: 20.10.2021', 2021, new Date()],
        ['05.11.2021', 'RE-2021-015', 'Kunde O', 'Webdesign', '', 'Erlöse aus Lieferungen und Leistungen', '', 2500, 19, '', '', 2975, 0, '', 'Offen', '', '', '', 2021, new Date()],
        ['15.12.2021', 'RE-2021-016', 'Kunde P', 'Beratung', '', 'Erlöse aus Lieferungen und Leistungen', '', 3300, 19, '', '', 3927, 3927, 0, 'Bezahlt', 'Überweisung', '20.12.2021', '✓ Bank: 20.12.2021', 2021, new Date()],
    ];

    // Add test data to sheet
    const range = sheet.getRange(2, 1, testData.length, testData[0].length);
    range.setValues(testData);
}

/**
 * Generate expense data
 */
function generateAusgabenData(ss) {
    const sheet = ss.getSheetByName('Ausgaben');
    if (!sheet) return;

    const testData = [
        // Standard expenses with different VAT rates and categories
        ['05.01.2021', 'RE-2021-A001', 'Lieferant A', 'Büromaterial', '', 'Bürokosten', '', 200, 19, '', '', 238, 238, 0, 'Bezahlt', 'Überweisung', '10.01.2021', '✓ Bank: 10.01.2021', 2021, new Date()],
        ['10.01.2021', 'RE-2021-A002', 'Lieferant B', 'Miete Büro', '', 'Miete', '', 1200, 0, '', '', 1200, 1200, 0, 'Bezahlt', 'Überweisung', '15.01.2021', '✓ Bank: 15.01.2021', 2021, new Date()],
        ['15.01.2021', 'RE-2021-A003', 'Lieferant C', 'Fachliteratur', '', 'Bürokosten', '', 80, 7, '', '', 85.60, 85.60, 0, 'Bezahlt', 'Überweisung', '20.01.2021', '✓ Bank: 20.01.2021', 2021, new Date()],
        ['20.01.2021', 'RE-2021-A004', 'Lieferant D', 'Software-Lizenz', '', 'IT-Kosten', '', 350, 19, '', '', 416.50, 416.50, 0, 'Bezahlt', 'Überweisung', '25.01.2021', '✓ Bank: 25.01.2021', 2021, new Date()],

        // Tax-related expenses for verifying tax calculations
        ['01.02.2021', 'RE-2021-A005', 'Steuerberater', 'Steuerberatung Q4/2020', '', 'Körperschaftsteuer', '', 800, 19, '', '', 952, 952, 0, 'Bezahlt', 'Überweisung', '10.02.2021', '✓ Bank: 10.02.2021', 2021, new Date()],
        ['05.02.2021', 'RE-2021-A006', 'Finanzamt', 'Gewerbesteuer Q4/2020', '', 'Gewerbesteuerrückstellungen', '', 1500, 0, '', '', 1500, 1500, 0, 'Bezahlt', 'Überweisung', '15.02.2021', '✓ Bank: 15.02.2021', 2021, new Date()],

        // Partial payments
        ['10.02.2021', 'RE-2021-A007', 'Lieferant E', 'Server Hardware', '', 'Abschreibungen Maschinen', '', 3000, 19, '', '', 3570, 1785, 1500, 'Teilbezahlt', 'Überweisung', '15.02.2021', '✓ Bank: 15.02.2021', 2021, new Date()],

        // Unpaid invoices
        ['15.02.2021', 'RE-2021-A008', 'Lieferant F', 'Marketing Dienstleistung', '', 'Marketing & Werbung', '', 2500, 19, '', '', 2975, 0, '', 'Offen', '', '', '', 2021, new Date()],

        // Quarterly expenses
        ['05.03.2021', 'RE-2021-A009', 'Lieferant G', 'Telefon & Internet Q1', '', 'Telefon & Internet', '', 300, 19, '', '', 357, 357, 0, 'Bezahlt', 'Überweisung', '10.03.2021', '✓ Bank: 10.03.2021', 2021, new Date()],
        ['10.04.2021', 'RE-2021-A010', 'Lieferant H', 'Versicherung', '', 'Versicherungen', '', 500, 19, '', '', 595, 595, 0, 'Bezahlt', 'Überweisung', '15.04.2021', '✓ Bank: 15.04.2021', 2021, new Date()],
        ['15.05.2021', 'RE-2021-A011', 'Lieferant I', 'Fachliteratur', '', 'Fortbildungskosten', '', 150, 7, '', '', 160.50, 160.50, 0, 'Bezahlt', 'Überweisung', '20.05.2021', '✓ Bank: 20.05.2021', 2021, new Date()],
        ['20.06.2021', 'RE-2021-A012', 'Lieferant J', 'Telefon & Internet Q2', '', 'Telefon & Internet', '', 300, 19, '', '', 357, 357, 0, 'Bezahlt', 'Überweisung', '25.06.2021', '✓ Bank: 25.06.2021', 2021, new Date()],
        ['05.07.2021', 'RE-2021-A013', 'Lieferant K', 'Büromaterial', '', 'Bürokosten', '', 180, 19, '', '', 214.20, 214.20, 0, 'Bezahlt', 'Überweisung', '10.07.2021', '✓ Bank: 10.07.2021', 2021, new Date()],
        ['10.08.2021', 'RE-2021-A014', 'Lieferant L', 'Software-Wartung', '', 'Abschreibungen immaterielle Wirtschaftsgüter', '', 400, 19, '', '', 476, 476, 0, 'Bezahlt', 'Überweisung', '15.08.2021', '✓ Bank: 15.08.2021', 2021, new Date()],
        ['15.09.2021', 'RE-2021-A015', 'Lieferant M', 'Telefon & Internet Q3', '', 'Telefon & Internet', '', 300, 19, '', '', 357, 357, 0, 'Bezahlt', 'Überweisung', '20.09.2021', '✓ Bank: 20.09.2021', 2021, new Date()],
        ['05.10.2021', 'RE-2021-A016', 'Lieferant N', 'Anwalt', '', 'Sonstige betriebliche Aufwendungen', '', 800, 19, '', '', 952, 952, 0, 'Bezahlt', 'Überweisung', '10.10.2021', '✓ Bank: 10.10.2021', 2021, new Date()],
        ['10.11.2021', 'RE-2021-A017', 'Lieferant O', 'Büromaterial', '', 'Bürokosten', '', 220, 19, '', '', 261.80, 261.80, 0, 'Bezahlt', 'Überweisung', '15.11.2021', '✓ Bank: 15.11.2021', 2021, new Date()],
        ['15.12.2021', 'RE-2021-A018', 'Lieferant P', 'Telefon & Internet Q4', '', 'Telefon & Internet', '', 300, 19, '', '', 357, 357, 0, 'Bezahlt', 'Überweisung', '20.12.2021', '✓ Bank: 20.12.2021', 2021, new Date()],

        // Different expense categories for BWA testing
        ['20.12.2021', 'RE-2021-A019', 'Lieferant Q', 'Personal', '', 'Bruttolöhne & Gehälter', '', 3000, 0, '', '', 3000, 3000, 0, 'Bezahlt', 'Überweisung', '20.12.2021', '✓ Bank: 20.12.2021', 2021, new Date()],
        ['21.12.2021', 'RE-2021-A020', 'Lieferant R', 'Sozialversicherung', '', 'Soziale Abgaben & Arbeitgeberanteile', '', 600, 0, '', '', 600, 600, 0, 'Bezahlt', 'Überweisung', '21.12.2021', '✓ Bank: 21.12.2021', 2021, new Date()],
    ];

    // Add test data to sheet
    const range = sheet.getRange(2, 1, testData.length, testData[0].length);
    range.setValues(testData);
}

/**
 * Generate Eigenbelege data
 */
function generateEigenbelegeData(ss) {
    const sheet = ss.getSheetByName('Eigenbelege');
    if (!sheet) return;

    const testData = [
        // Variety of Eigenbelege with different categories and tax rates
        ['10.01.2021', 'EB-2021-001', 'Shareholder 1', 'Taxi zur Kundenberatung', '', 'Reisekosten', '', 45, 7, '', '', 48.15, 48.15, 0, 'Erstattet', 'Überweisung', '20.01.2021', '✓ Bank: 20.01.2021', 2021, new Date()],
        ['20.01.2021', 'EB-2021-002', 'Shareholder 1', 'Bewirtung Kunde', '', 'Bewirtung', '', 120, 19, '', '', 142.80, 142.80, 0, 'Erstattet', 'Überweisung', '25.01.2021', '✓ Bank: 25.01.2021', 2021, new Date()],
        ['05.02.2021', 'EB-2021-003', 'Shareholder 2', 'Büromaterial', '', 'Bürokosten', '', 35, 19, '', '', 41.65, 41.65, 0, 'Erstattet', 'Überweisung', '15.02.2021', '✓ Bank: 15.02.2021', 2021, new Date()],
        ['15.02.2021', 'EB-2021-004', 'Shareholder 2', 'Parken Kundentermin', '', 'Reisekosten', '', 15, 19, '', '', 17.85, 17.85, 0, 'Erstattet', 'Überweisung', '20.02.2021', '✓ Bank: 20.02.2021', 2021, new Date()],

        // Partially reimbursed
        ['01.03.2021', 'EB-2021-005', 'Shareholder 1', 'Hotelübernachtung Messe', '', 'Reisekosten', '', 200, 7, '', '', 214, 100, 106.54, 'Teilerstattet', 'Überweisung', '10.03.2021', '✓ Bank: 10.03.2021', 2021, new Date()],

        // Not yet reimbursed
        ['10.03.2021', 'EB-2021-006', 'Shareholder 1', 'Geschäftsessen mit Partner', '', 'Bewirtung', '', 85, 19, '', '', 101.15, 0, '', 'Offen', '', '', '', 2021, new Date()],

        // Quarterly distribution
        ['05.04.2021', 'EB-2021-007', 'Shareholder 2', 'Betriebsausflug Team', '', 'Sonstige Personalkosten', '', 300, 19, '', '', 357, 357, 0, 'Erstattet', 'Überweisung', '15.04.2021', '✓ Bank: 15.04.2021', 2021, new Date()],
        ['15.05.2021', 'EB-2021-008', 'Shareholder 1', 'Tankbeleg Dienstreise', '', 'Reisekosten', '', 65, 19, '', '', 77.35, 77.35, 0, 'Erstattet', 'Überweisung', '25.05.2021', '✓ Bank: 25.05.2021', 2021, new Date()],
        ['10.06.2021', 'EB-2021-009', 'Shareholder 2', 'Bewirtung Kunden', '', 'Bewirtung', '', 95, 19, '', '', 113.05, 113.05, 0, 'Erstattet', 'Überweisung', '20.06.2021', '✓ Bank: 20.06.2021', 2021, new Date()],
        ['05.07.2021', 'EB-2021-010', 'Shareholder 1', 'Betriebsbedarf', '', 'Bürokosten', '', 50, 19, '', '', 59.50, 59.50, 0, 'Erstattet', 'Überweisung', '15.07.2021', '✓ Bank: 15.07.2021', 2021, new Date()],
        ['20.08.2021', 'EB-2021-011', 'Shareholder 2', 'Trinkgeld Restaurant', '', 'Trinkgeld', '', 15, 0, '', '', 15, 15, 0, 'Erstattet', 'Überweisung', '25.08.2021', '✓ Bank: 25.08.2021', 2021, new Date()],
        ['10.09.2021', 'EB-2021-012', 'Shareholder 1', 'Bahnticket Kundentermin', '', 'Reisekosten', '', 75, 7, '', '', 80.25, 80.25, 0, 'Erstattet', 'Überweisung', '20.09.2021', '✓ Bank: 20.09.2021', 2021, new Date()],
        ['05.10.2021', 'EB-2021-013', 'Shareholder 1', 'Geschäftskleidung', '', 'Kleidung', '', 180, 19, '', '', 214.20, 214.20, 0, 'Erstattet', 'Überweisung', '15.10.2021', '✓ Bank: 15.10.2021', 2021, new Date()],
        ['15.11.2021', 'EB-2021-014', 'Shareholder 2', 'Bewirtung Partner', '', 'Bewirtung', '', 110, 19, '', '', 130.90, 130.90, 0, 'Erstattet', 'Überweisung', '25.11.2021', '✓ Bank: 25.11.2021', 2021, new Date()],
        ['10.12.2021', 'EB-2021-015', 'Shareholder 1', 'Fachliteratur', '', 'Fortbildungskosten', '', 65, 7, '', '', 69.55, 69.55, 0, 'Erstattet', 'Überweisung', '20.12.2021', '✓ Bank: 20.12.2021', 2021, new Date()],
    ];

    // Add test data to sheet
    const range = sheet.getRange(2, 1, testData.length, testData[0].length);
    range.setValues(testData);
}

/**
 * Generate Gesellschafterkonto data
 */
function generateGesellschafterkontoData(ss) {
    const sheet = ss.getSheetByName('Gesellschafterkonto');
    if (!sheet) return;

    const testData = [
        // Initial investments and shareholder loans
        ['05.01.2021', 'GK-2021-001', 'Shareholder 1', 'Gesellschafterdarlehen', 'Gesellschafterdarlehen', '', 50000, 1, 'Überweisung', '05.01.2021', '✓ Bank: 05.01.2021', 2021, new Date()],
        ['05.01.2021', 'GK-2021-002', 'Shareholder 2', 'Gesellschafterdarlehen', 'Gesellschafterdarlehen', '', 50000, 1, 'Überweisung', '05.01.2021', '✓ Bank: 05.01.2021', 2021, new Date()],

        // Capital-related transactions during the year
        ['10.04.2021', 'GK-2021-003', 'Shareholder 1', 'Privatentnahme', 'Privatentnahme', '', -2000, 2, 'Überweisung', '10.04.2021', '✓ Bank: 10.04.2021', 2021, new Date()],
        ['15.07.2021', 'GK-2021-004', 'Shareholder 2', 'Privatentnahme', 'Privatentnahme', '', -2000, 3, 'Überweisung', '15.07.2021', '✓ Bank: 15.07.2021', 2021, new Date()],
        ['10.10.2021', 'GK-2021-005', 'Shareholder 1', 'Privatentnahme', 'Privatentnahme', '', -2000, 4, 'Überweisung', '10.10.2021', '✓ Bank: 10.10.2021', 2021, new Date()],
        ['15.12.2021', 'GK-2021-006', 'Shareholder 2', 'Privatentnahme', 'Privatentnahme', '', -2000, 4, 'Überweisung', '15.12.2021', '✓ Bank: 15.12.2021', 2021, new Date()],

        // Year-end distributions
        ['20.12.2021', 'GK-2021-007', 'Shareholder 1', 'Ausschüttung', 'Ausschüttungen', '', -10000, 4, 'Überweisung', '20.12.2021', '✓ Bank: 20.12.2021', 2021, new Date()],
        ['20.12.2021', 'GK-2021-008', 'Shareholder 2', 'Ausschüttung', 'Ausschüttungen', '', -10000, 4, 'Überweisung', '20.12.2021', '✓ Bank: 20.12.2021', 2021, new Date()],
    ];

    // Add test data to sheet
    const range = sheet.getRange(2, 1, testData.length, testData[0].length);
    range.setValues(testData);
}

/**
 * Generate Holding Transfers data
 */
function generateHoldingTransfersData(ss) {
    const sheet = ss.getSheetByName('Holding Transfers');
    if (!sheet) return;

    const testData = [
        ['20.06.2021', 'HT-2021-001', 'GmbH 1', 'Gewinnübertrag Q2', '', 'Gewinnübertrag', '', 15000, 2, 'Überweisung', '20.06.2021', '✓ Bank: 20.06.2021', 2021, new Date()],
        ['15.12.2021', 'HT-2021-002', 'GmbH 2', 'Gewinnübertrag Q4', '', 'Gewinnübertrag', '', 20000, 4, 'Überweisung', '15.12.2021', '✓ Bank: 15.12.2021', 2021, new Date()],
        ['20.12.2021', 'HT-2021-003', 'GmbH 2', 'Kapitalrückführung', '', 'Kapitalrückführung', '', 5000, 4, 'Überweisung', '20.12.2021', '✓ Bank: 20.12.2021', 2021, new Date()],
    ];

    // Add test data to sheet
    const range = sheet.getRange(2, 1, testData.length, testData[0].length);
    range.setValues(testData);
}

/**
 * Generate Bank Movements data
 */
function generateBankbewegungenData(ss) {
    const sheet = ss.getSheetByName('Bankbewegungen');
    if (!sheet) return;

    // First, clear the entire sheet and add a starting balance
    if (sheet.getLastRow() > 0) {
        sheet.clear();
    }

    // Add header row
    const header = ['Datum', 'Buchungstext', 'Betrag', 'Saldo', 'Transaktionstyp', 'Kategorie', 'Konto (Soll)', 'Gegenkonto (Haben)', 'Referenz', 'Verwendungszweck', 'Match-Information', 'Zeitstempel'];
    sheet.appendRow(header);

    // Add starting balance
    sheet.appendRow(['01.01.2021', 'Anfangssaldo', '', 50000, '', '', '', '', '', '', '', new Date()]);

    // Bank transaction data - these match with the other sheets to enable testing of the matching algorithm
    const testData = [
        // Revenue matching entries
        ['20.01.2021', 'Eingang: Kunde B AG', 2975, '', '', '', '', '', 'RE-2021-002', 'Rechnung RE-2021-002', '', new Date()],
        ['10.02.2021', 'Eingang: Kunde C Ltd', 1785, '', '', '', '', '', 'RE-2021-003', 'Beratung', '', new Date()],
        ['15.03.2021', 'Eingang: Kunde E', 3000, '', '', '', '', '', 'RE-2021-005', 'Beratung Ausland', '', new Date()],
        ['01.04.2021', 'Eingang: Kunde F', 800, '', '', '', '', '', 'RE-2021-006', 'Vermietung Büroraum', '', new Date()],
        ['15.04.2021', 'Eingang: Kunde G', 2975, '', '', '', '', '', 'RE-2021-007', 'Teilzahlung Softwareentwicklung', '', new Date()],
        ['10.04.2021', 'Ausgang: Kunde A GmbH', -238, '', '', '', '', '', 'GS-2021-001', 'Gutschrift zu RE-2021-001', '', new Date()],
        ['20.04.2021', 'Eingang: Kunde H GmbH', 1428, '', '', '', '', '', 'RE-2021-008', 'Schulung', '', new Date()],
        ['25.06.2021', 'Eingang: Kunde J', 2618, '', '', '', '', '', 'RE-2021-010', 'Schulung', '', new Date()],
        ['12.07.2021', 'Eingang: Kunde K', 3570, '', '', '', '', '', 'RE-2021-011', 'Workshop', '', new Date()],
        ['01.10.2021', 'Eingang: Kunde M', 3332, '', '', '', '', '', 'RE-2021-013', 'IT-Dienstleistung', '', new Date()],
        ['20.10.2021', 'Eingang: Kunde N', 2142, '', '', '', '', '', 'RE-2021-014', 'Marketing', '', new Date()],
        ['20.12.2021', 'Eingang: Kunde P', 3927, '', '', '', '', '', 'RE-2021-016', 'Beratung', '', new Date()],

        // Expense matching entries
        ['10.01.2021', 'Ausgang: Lieferant A', -238, '', '', '', '', '', 'RE-2021-A001', 'Büromaterial', '', new Date()],
        ['15.01.2021', 'Ausgang: Lieferant B', -1200, '', '', '', '', '', 'RE-2021-A002', 'Miete Büro', '', new Date()],
        ['20.01.2021', 'Ausgang: Lieferant C', -85.60, '', '', '', '', '', 'RE-2021-A003', 'Fachliteratur', '', new Date()],
        ['25.01.2021', 'Ausgang: Lieferant D', -416.50, '', '', '', '', '', 'RE-2021-A004', 'Software-Lizenz', '', new Date()],
        ['10.02.2021', 'Ausgang: Steuerberater', -952, '', '', '', '', '', 'RE-2021-A005', 'Steuerberatung Q4/2020', '', new Date()],
        ['15.02.2021', 'Ausgang: Finanzamt', -1500, '', '', '', '', '', 'RE-2021-A006', 'Gewerbesteuer Q4/2020', '', new Date()],
        ['15.02.2021', 'Ausgang: Lieferant E', -1785, '', '', '', '', '', 'RE-2021-A007', 'Teilzahlung Server Hardware', '', new Date()],
        ['10.03.2021', 'Ausgang: Lieferant G', -357, '', '', '', '', '', 'RE-2021-A009', 'Telefon & Internet Q1', '', new Date()],
        ['15.04.2021', 'Ausgang: Lieferant H', -595, '', '', '', '', '', 'RE-2021-A010', 'Versicherung', '', new Date()],
        ['20.05.2021', 'Ausgang: Lieferant I', -160.50, '', '', '', '', '', 'RE-2021-A011', 'Fachliteratur', '', new Date()],
        ['25.06.2021', 'Ausgang: Lieferant J', -357, '', '', '', '', '', 'RE-2021-A012', 'Telefon & Internet Q2', '', new Date()],
        ['10.07.2021', 'Ausgang: Lieferant K', -214.20, '', '', '', '', '', 'RE-2021-A013', 'Büromaterial', '', new Date()],
        ['15.08.2021', 'Ausgang: Lieferant L', -476, '', '', '', '', '', 'RE-2021-A014', 'Software-Wartung', '', new Date()],
        ['20.09.2021', 'Ausgang: Lieferant M', -357, '', '', '', '', '', 'RE-2021-A015', 'Telefon & Internet Q3', '', new Date()],
        ['10.10.2021', 'Ausgang: Lieferant N', -952, '', '', '', '', '', 'RE-2021-A016', 'Anwalt', '', new Date()],
        ['15.11.2021', 'Ausgang: Lieferant O', -261.80, '', '', '', '', '', 'RE-2021-A017', 'Büromaterial', '', new Date()],
        ['20.12.2021', 'Ausgang: Lieferant P', -357, '', '', '', '', '', 'RE-2021-A018', 'Telefon & Internet Q4', '', new Date()],
        ['20.12.2021', 'Ausgang: Lieferant Q', -3000, '', '', '', '', '', 'RE-2021-A019', 'Personal', '', new Date()],
        ['21.12.2021', 'Ausgang: Lieferant R', -600, '', '', '', '', '', 'RE-2021-A020', 'Sozialversicherung', '', new Date()],

        // Eigenbelege matching entries
        ['20.01.2021', 'Ausgang: Shareholder 1', -48.15, '', '', '', '', '', 'EB-2021-001', 'Erstattung Taxi', '', new Date()],
        ['25.01.2021', 'Ausgang: Shareholder 1', -142.80, '', '', '', '', '', 'EB-2021-002', 'Erstattung Bewirtung', '', new Date()],
        ['15.02.2021', 'Ausgang: Shareholder 2', -41.65, '', '', '', '', '', 'EB-2021-003', 'Erstattung Büromaterial', '', new Date()],
        ['20.02.2021', 'Ausgang: Shareholder 2', -17.85, '', '', '', '', '', 'EB-2021-004', 'Erstattung Parken', '', new Date()],
        ['10.03.2021', 'Ausgang: Shareholder 1', -100, '', '', '', '', '', 'EB-2021-005', 'Erstattung Hotel (teilweise)', '', new Date()],
        ['15.04.2021', 'Ausgang: Shareholder 2', -357, '', '', '', '', '', 'EB-2021-007', 'Erstattung Betriebsausflug', '', new Date()],
        ['25.05.2021', 'Ausgang: Shareholder 1', -77.35, '', '', '', '', '', 'EB-2021-008', 'Erstattung Tankbeleg', '', new Date()],
        ['20.06.2021', 'Ausgang: Shareholder 2', -113.05, '', '', '', '', '', 'EB-2021-009', 'Erstattung Bewirtung', '', new Date()],
        ['15.07.2021', 'Ausgang: Shareholder 1', -59.50, '', '', '', '', '', 'EB-2021-010', 'Erstattung Betriebsbedarf', '', new Date()],
        ['25.08.2021', 'Ausgang: Shareholder 2', -15, '', '', '', '', '', 'EB-2021-011', 'Erstattung Trinkgeld', '', new Date()],
        ['20.09.2021', 'Ausgang: Shareholder 1', -80.25, '', '', '', '', '', 'EB-2021-012', 'Erstattung Bahnticket', '', new Date()],
        ['15.10.2021', 'Ausgang: Shareholder 1', -214.20, '', '', '', '', '', 'EB-2021-013', 'Erstattung Kleidung', '', new Date()],
        ['25.11.2021', 'Ausgang: Shareholder 2', -130.90, '', '', '', '', '', 'EB-2021-014', 'Erstattung Bewirtung', '', new Date()],
        ['20.12.2021', 'Ausgang: Shareholder 1', -69.55, '', '', '', '', '', 'EB-2021-015', 'Erstattung Fachliteratur', '', new Date()],

        // Gesellschafterkonto matching entries
        ['05.01.2021', 'Eingang: Shareholder 1', 50000, '', '', '', '', '', 'GK-2021-001', 'Gesellschafterdarlehen', '', new Date()],
        ['05.01.2021', 'Eingang: Shareholder 2', 50000, '', '', '', '', '', 'GK-2021-002', 'Gesellschafterdarlehen', '', new Date()],
        ['10.04.2021', 'Ausgang: Shareholder 1', -2000, '', '', '', '', '', 'GK-2021-003', 'Privatentnahme', '', new Date()],
        ['15.07.2021', 'Ausgang: Shareholder 2', -2000, '', '', '', '', '', 'GK-2021-004', 'Privatentnahme', '', new Date()],
        ['10.10.2021', 'Ausgang: Shareholder 1', -2000, '', '', '', '', '', 'GK-2021-005', 'Privatentnahme', '', new Date()],
        ['15.12.2021', 'Ausgang: Shareholder 2', -2000, '', '', '', '', '', 'GK-2021-006', 'Privatentnahme', '', new Date()],
        ['20.12.2021', 'Ausgang: Shareholder 1', -10000, '', '', '', '', '', 'GK-2021-007', 'Ausschüttung', '', new Date()],
        ['20.12.2021', 'Ausgang: Shareholder 2', -10000, '', '', '', '', '', 'GK-2021-008', 'Ausschüttung', '', new Date()],

        // Holding Transfers matching entries
        ['20.06.2021', 'Eingang: GmbH 1', 15000, '', '', '', '', '', 'HT-2021-001', 'Gewinnübertrag Q2', '', new Date()],
        ['15.12.2021', 'Eingang: GmbH 2', 20000, '', '', '', '', '', 'HT-2021-002', 'Gewinnübertrag Q4', '', new Date()],
        ['20.12.2021', 'Eingang: GmbH 2', 5000, '', '', '', '', '', 'HT-2021-003', 'Kapitalrückführung', '', new Date()],
    ];

    // Add test data to sheet
    sheet.getRange(2, 1, testData.length, testData[0].length).setValues(testData);

    // Add end balance row
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, 1, 4).setValues([['31.12.2021', 'Endsaldo', '', '=D' + lastRow + '+C' + lastRow]]);
}