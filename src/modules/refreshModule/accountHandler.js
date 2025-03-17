// src/modules/refreshModule/accountHandler.js
import stringUtils from '../../utils/stringUtils.js';

/**
 * Updates booking accounts based on categories for any sheet type
 * @param {Sheet} sheet - The sheet to update
 * @param {string} sheetType - Type of sheet (einnahmen, ausgaben, bankbewegungen, etc.)
 * @param {Object} config - Configuration object
 * @param {boolean} isBankSheet - Whether this is a bank sheet (with soll/haben instead of single account)
 * @returns {boolean} - Success status
 */
function updateBookingAccounts(sheet, sheetType, config, isBankSheet = false) {
    try {
        const columns = config[sheetType].columns;

        // Exit early if missing required columns
        if (!columns.kategorie || (!columns.buchungskonto && !isBankSheet)) {
            console.log(`Missing required columns for ${sheetType}`);
            return false;
        }

        // For bank sheet, we need kontoSoll and kontoHaben
        if (isBankSheet && (!columns.kontoSoll || !columns.kontoHaben)) {
            console.log('Missing kontoSoll/kontoHaben columns for bank sheet');
            return false;
        }

        const lastRow = sheet.getLastRow();
        const numRows = lastRow - 1;

        if (numRows <= 0) return true;

        console.log(`Updating booking accounts for ${numRows} rows in ${sheetType}`);

        // Get category data
        const kategorieRange = sheet.getRange(2, columns.kategorie, numRows, 1);
        const kategorieValues = kategorieRange.getValues();

        // Get transaction type for bank sheets
        let transTypeValues = [];
        if (isBankSheet && columns.transaktionstyp) {
            const transTypeRange = sheet.getRange(2, columns.transaktionstyp, numRows, 1);
            transTypeValues = transTypeRange.getValues();
        }

        // For bank sheets, get kontoSoll and kontoHaben
        let kontoSollValues = [];
        let kontoHabenValues = [];
        if (isBankSheet) {
            // Get current kontoSoll and kontoHaben values
            kontoSollValues = sheet.getRange(2, columns.kontoSoll, numRows, 1).getValues();
            kontoHabenValues = sheet.getRange(2, columns.kontoHaben, numRows, 1).getValues();
        } else {
            // Get current buchungskonto values for data sheets
            kontoSollValues = sheet.getRange(2, columns.buchungskonto, numRows, 1).getValues();
        }

        // Create updated account values
        const updatedSollValues = [];
        const updatedHabenValues = isBankSheet ? [] : null;

        // Process each row
        for (let i = 0; i < numRows; i++) {
            const kategorie = kategorieValues[i][0] ? kategorieValues[i][0].toString().trim() : '';
            let sollKonto = kontoSollValues[i][0] ? kontoSollValues[i][0].toString() : '';
            let habenKonto = isBankSheet ?
                (kontoHabenValues[i][0] ? kontoHabenValues[i][0].toString() : '') :
                null;

            // Skip if no category
            if (!kategorie) {
                updatedSollValues.push([sollKonto]);
                if (isBankSheet) updatedHabenValues.push([habenKonto]);
                continue;
            }

            // Get relevant kontoMapping based on sheet type
            let kontoMapping;
            if (isBankSheet) {
                // For bank sheets, determine mapping based on transaction type
                const transType = transTypeValues[i][0] ? transTypeValues[i][0].toString().trim() : '';

                if (transType === 'Einnahme') {
                    kontoMapping = config.einnahmen.kontoMapping[kategorie];
                } else if (transType === 'Ausgabe') {
                    // Try multiple mappings in order
                    kontoMapping = config.ausgaben.kontoMapping[kategorie] ||
                        config.eigenbelege.kontoMapping[kategorie] ||
                        config.gesellschafterkonto.kontoMapping[kategorie] ||
                        config.holdingTransfers.kontoMapping[kategorie];
                }
            } else {
                // For data sheets, use the sheet's own mapping
                kontoMapping = config[sheetType].kontoMapping[kategorie];
            }

            // Update accounts if mapping exists
            if (kontoMapping) {
                if (isBankSheet) {
                    sollKonto = kontoMapping.soll || sollKonto;
                    habenKonto = kontoMapping.gegen || habenKonto;
                } else {
                    // For data sheets, set the appropriate buchungskonto
                    if (sheetType === 'einnahmen') {
                        // For revenue, we use the gegen account
                        sollKonto = kontoMapping.gegen || sollKonto;
                    } else {
                        // For expenses, we use the soll account
                        sollKonto = kontoMapping.soll || sollKonto;
                    }
                }
            }

            // Add to update arrays
            updatedSollValues.push([sollKonto]);
            if (isBankSheet) updatedHabenValues.push([habenKonto]);
        }

        // Update the sheet with the new values
        if (isBankSheet) {
            sheet.getRange(2, columns.kontoSoll, numRows, 1).setValues(updatedSollValues);
            sheet.getRange(2, columns.kontoHaben, numRows, 1).setValues(updatedHabenValues);
        } else {
            sheet.getRange(2, columns.buchungskonto, numRows, 1).setValues(updatedSollValues);
        }

        console.log(`Updated ${updatedSollValues.length} booking accounts in ${sheetType}`);
        return true;
    } catch (e) {
        console.error(`Error updating booking accounts for ${sheetType}:`, e);
        return false;
    }
}

/**
 * Checks if the Soll and Haben accounts are compatible with the category
 * and updates them if they don't match
 * @param {Object} row - Row data
 * @param {Object} category - Category information
 * @param {Object} mapping - Account mapping
 * @returns {Object} Updated accounts
 */
function validateAccountsForCategory(kategorie, sollKonto, habenKonto, mapping) {
    if (!mapping) return { sollKonto, habenKonto };

    // Check if accounts match the expected accounts for the category
    const expectedSoll = mapping.soll || '';
    const expectedHaben = mapping.gegen || '';

    // Return updated accounts if they don't match
    return {
        sollKonto: expectedSoll || sollKonto,
        habenKonto: expectedHaben || habenKonto,
    };
}

export default {
    updateBookingAccounts,
    validateAccountsForCategory,
};