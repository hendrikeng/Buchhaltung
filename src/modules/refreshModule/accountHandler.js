// src/modules/refreshModule/accountHandler.js
/**
 * Updates booking accounts based on categories for any sheet type with optimized batch operations
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

        if (numRows <= 0) return true; // No data to update

        console.log(`Updating booking accounts for ${numRows} rows in ${sheetType}`);

        // Optimization: Load all data in a single API call
        const requiredColumns = [columns.kategorie];

        if (isBankSheet) {
            // For bank sheets, include transaction type, kontoSoll and kontoHaben
            if (columns.transaktionstyp) requiredColumns.push(columns.transaktionstyp);
            requiredColumns.push(columns.kontoSoll, columns.kontoHaben);
        } else {
            // For data sheets, just include buchungskonto
            requiredColumns.push(columns.buchungskonto);
        }

        const maxCol = Math.max(...requiredColumns);
        const data = sheet.getRange(2, 1, numRows, maxCol).getValues();

        // Optimization: Build update arrays for batched writes
        const updates = {
            sollAccounts: [],
            habenAccounts: [],
            sollRows: [],
            habenRows: [],
        };

        // Process each row for updates
        data.forEach((row, i) => {
            const rowIndex = i + 2; // 1-based + header
            const kategorie = row[columns.kategorie - 1]?.toString().trim() || '';

            // Skip if no category
            if (!kategorie) return;

            // Get account mapping based on sheet type and transaction type
            const mapping = getAccountMapping(kategorie, row, columns, sheetType, config, isBankSheet);
            if (!mapping) return;

            // For bank sheets, update both Soll and Haben
            if (isBankSheet) {
                const currentSoll = row[columns.kontoSoll - 1]?.toString() || '';
                const currentHaben = row[columns.kontoHaben - 1]?.toString() || '';

                // Only update if accounts are different from current values
                if (mapping.soll && mapping.soll !== currentSoll) {
                    updates.sollAccounts.push(mapping.soll);
                    updates.sollRows.push(rowIndex);
                }

                if (mapping.gegen && mapping.gegen !== currentHaben) {
                    updates.habenAccounts.push(mapping.gegen);
                    updates.habenRows.push(rowIndex);
                }
            }
            // For data sheets, update buchungskonto
            else {
                const currentAccount = row[columns.buchungskonto - 1]?.toString() || '';
                const accountToUse = sheetType === 'einnahmen' ?
                    (mapping.gegen || '') : (mapping.soll || '');

                if (accountToUse && accountToUse !== currentAccount) {
                    updates.sollAccounts.push(accountToUse);
                    updates.sollRows.push(rowIndex);
                }
            }
        });

        // Apply updates in optimized batches
        applyAccountUpdates(sheet, updates, columns, isBankSheet);

        console.log(`Updated accounts in ${sheetType}: ${updates.sollRows.length + updates.habenRows.length} changes`);
        return true;
    } catch (e) {
        console.error(`Error updating booking accounts for ${sheetType}:`, e);
        return false;
    }
}

/**
 * Gets account mapping based on category and sheet type
 * @param {string} kategorie - The category
 * @param {Array} row - The row data
 * @param {Object} columns - Column configuration
 * @param {string} sheetType - Sheet type
 * @param {Object} config - Configuration
 * @param {boolean} isBankSheet - Whether this is a bank sheet
 * @returns {Object|null} Account mapping or null
 */
function getAccountMapping(kategorie, row, columns, sheetType, config, isBankSheet) {
    if (isBankSheet) {
        // For bank sheets, determine mapping based on transaction type
        const transType = columns.transaktionstyp ?
            (row[columns.transaktionstyp - 1]?.toString().trim() || '') : '';

        if (transType === 'Einnahme') {
            return config.einnahmen.kontoMapping[kategorie];
        } else if (transType === 'Ausgabe') {
            // Try multiple mappings in order of probability
            return config.ausgaben.kontoMapping[kategorie] ||
                config.eigenbelege.kontoMapping[kategorie] ||
                config.gesellschafterkonto.kontoMapping[kategorie] ||
                config.holdingTransfers.kontoMapping[kategorie];
        }
    } else {
        // For data sheets, use the sheet's own mapping
        return config[sheetType].kontoMapping[kategorie];
    }

    return null;
}

/**
 * Applies account updates in optimized batches
 * @param {Sheet} sheet - The sheet
 * @param {Object} updates - The updates
 * @param {Object} columns - Column configuration
 * @param {boolean} isBankSheet - Whether this is a bank sheet
 */
function applyAccountUpdates(sheet, updates, columns, isBankSheet) {
    // Group rows by account for more efficient updates
    const applyUpdatesToColumn = (rows, accounts, column) => {
        if (rows.length === 0) return;

        // Group by account for batch updates
        const accountGroups = {};
        rows.forEach((row, i) => {
            const account = accounts[i];
            if (!accountGroups[account]) {
                accountGroups[account] = [];
            }
            accountGroups[account].push(row);
        });

        // Apply updates by account group
        Object.entries(accountGroups).forEach(([account, groupRows]) => {
            try {
                // Find consecutive ranges to minimize API calls
                const ranges = findConsecutiveRanges(groupRows);

                ranges.forEach(range => {
                    const [startRow, endRow] = range;
                    const numRows = endRow - startRow + 1;
                    const values = Array(numRows).fill([account]);

                    sheet.getRange(startRow, column, numRows, 1).setValues(values);
                });
            } catch (e) {
                console.error(`Error updating account ${account} in column ${column}:`, e);

                // Fallback: Update rows individually
                groupRows.forEach(row => {
                    try {
                        sheet.getRange(row, column).setValue(account);
                    } catch (rowError) {
                        console.error(`Error updating row ${row}:`, rowError);
                    }
                });
            }
        });
    };

    // Apply Soll account updates
    if (isBankSheet) {
        applyUpdatesToColumn(updates.sollRows, updates.sollAccounts, columns.kontoSoll);
        applyUpdatesToColumn(updates.habenRows, updates.habenAccounts, columns.kontoHaben);
    } else {
        applyUpdatesToColumn(updates.sollRows, updates.sollAccounts, columns.buchungskonto);
    }
}

/**
 * Finds consecutive ranges in an array of row indices
 * @param {Array} rows - Row indices
 * @returns {Array} Array of [startRow, endRow] pairs
 */
function findConsecutiveRanges(rows) {
    if (!rows || rows.length === 0) return [];

    // Sort rows for range detection
    rows.sort((a, b) => a - b);

    const ranges = [];
    let currentStart = rows[0];
    let currentEnd = rows[0];

    for (let i = 1; i < rows.length; i++) {
        if (rows[i] === currentEnd + 1) {
            // Extend current range
            currentEnd = rows[i];
        } else {
            // Store current range and start a new one
            ranges.push([currentStart, currentEnd]);
            currentStart = currentEnd = rows[i];
        }
    }

    // Add the last range
    ranges.push([currentStart, currentEnd]);

    return ranges;
}

export default {
    updateBookingAccounts,
};