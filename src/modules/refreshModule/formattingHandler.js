// src/modules/refreshModule/formattingHandler.js
import sheetUtils from '../../utils/sheetUtils.js';
import cellValidator from '../validatorModule/cellValidator.js';

/**
 * Setzt bedingte Formatierung für eine Statusspalte
 * @param {Sheet} sheet - Das Sheet
 * @param {string} column - Spaltenbezeichnung (z.B. "L")
 * @param {Array<Object>} conditions - Bedingungen (value, background, fontColor)
 */
function setConditionalFormattingForStatusColumn(sheet, column, conditions) {
    sheetUtils.setConditionalFormattingForColumn(sheet, column, conditions);
}

/**
 * Setzt Dropdown-Validierungen für ein Sheet
 * @param {Sheet} sheet - Das Sheet
 * @param {string} sheetName - Name des Sheets
 * @param {number} numRows - Anzahl der Datenzeilen
 * @param {Object} columns - Spaltenkonfiguration
 * @param {Object} config - Die Konfiguration
 */
function setDropdownValidations(sheet, sheetName, numRows, columns, config) {
    if (sheetName === 'Einnahmen') {
        cellValidator.validateDropdown(
            sheet, 2, columns.kategorie, numRows, 1,
            Object.keys(config.einnahmen.categories),
        );
    } else if (sheetName === 'Ausgaben') {
        cellValidator.validateDropdown(
            sheet, 2, columns.kategorie, numRows, 1,
            Object.keys(config.ausgaben.categories),
        );
    } else if (sheetName === 'Eigenbelege') {
        cellValidator.validateDropdown(
            sheet, 2, columns.kategorie, numRows, 1,
            Object.keys(config.eigenbelege.categories),
        );

        // Dropdown für "Ausgelegt von" hinzufügen (merged aus shareholders und employees)
        const ausleger = [
            ...config.common.shareholders,
            ...config.common.employees,
        ];

        cellValidator.validateDropdown(
            sheet, 2, columns.ausgelegtVon, numRows, 1,
            ausleger,
        );
    } else if (sheetName === 'Gesellschafterkonto') {
        // Kategorie-Dropdown für Gesellschafterkonto
        cellValidator.validateDropdown(
            sheet, 2, columns.kategorie, numRows, 1,
            Object.keys(config.gesellschafterkonto.categories),
        );

        // Dropdown für Gesellschafter
        cellValidator.validateDropdown(
            sheet, 2, columns.gesellschafter, numRows, 1,
            config.common.shareholders,
        );
    } else if (sheetName === 'Holding Transfers') {
        // Kategorie-Dropdown für Holding Transfers
        cellValidator.validateDropdown(
            sheet, 2, columns.kategorie, numRows, 1,
            Object.keys(config.holdingTransfers.categories),
        );
        // Zielgesellschaft-Dropdown für Holding Transfers
        cellValidator.validateDropdown(
            sheet, 2, columns.ziel, numRows, 1,
            config.common.companies,
        );
    }

    // Zahlungsart-Dropdown für alle Blätter mit Zahlungsart-Spalte
    if (columns.zahlungsart) {
        cellValidator.validateDropdown(
            sheet, 2, columns.zahlungsart, numRows, 1,
            config.common.paymentType,
        );
    }
}

/**
 * Wendet Validierungen auf das Bankbewegungen-Sheet an
 * @param {Sheet} sheet - Das Sheet
 * @param {number} firstDataRow - Erste Datenzeile
 * @param {number} numDataRows - Anzahl der Datenzeilen
 * @param {Object} columns - Spaltenkonfiguration
 * @param {Object} config - Die Konfiguration
 */
function applyBankSheetValidations(sheet, firstDataRow, numDataRows, columns, config) {
    // Validierung für Transaktionstyp
    cellValidator.validateDropdown(
        sheet, firstDataRow, columns.transaktionstyp, numDataRows, 1,
        config.bankbewegungen.types,
    );

    // Validierung für Kategorie
    cellValidator.validateDropdown(
        sheet, firstDataRow, columns.kategorie, numDataRows, 1,
        config.bankbewegungen.categories,
    );

    // Konten für Dropdown-Validierung sammeln
    const allowedKontoSoll = new Set();
    const allowedGegenkonto = new Set();

    // Konten aus allen relevanten Mappings sammeln
    [
        config.einnahmen.kontoMapping,
        config.ausgaben.kontoMapping,
        config.eigenbelege.kontoMapping,
        config.gesellschafterkonto.kontoMapping,
        config.holdingTransfers.kontoMapping,
    ].forEach(mapping => {
        Object.values(mapping).forEach(m => {
            if (m.soll) allowedKontoSoll.add(m.soll);
            if (m.gegen) allowedGegenkonto.add(m.gegen);
        });
    });

    // Dropdown-Validierungen für Konten setzen
    cellValidator.validateDropdown(
        sheet, firstDataRow, columns.kontoSoll, numDataRows, 1,
        Array.from(allowedKontoSoll),
    );

    cellValidator.validateDropdown(
        sheet, firstDataRow, columns.kontoHaben, numDataRows, 1,
        Array.from(allowedGegenkonto),
    );
}

export default {
    setConditionalFormattingForStatusColumn,
    setDropdownValidations,
    applyBankSheetValidations,
};