// src/modules/refreshModule/formattingHandler.js
import sheetUtils from '../../utils/sheetUtils.js';
import cellValidator from '../validatorModule/cellValidator.js';

/**
 * Setzt bedingte Formatierung für eine Statusspalte mit optimierter API-Nutzung
 * @param {Sheet} sheet - Das Sheet
 * @param {string} column - Spaltenbezeichnung (z.B. "L")
 * @param {Array<Object>} conditions - Bedingungen (value, background, fontColor)
 */
function setConditionalFormattingForStatusColumn(sheet, column, conditions) {
    sheetUtils.setConditionalFormattingForColumn(sheet, column, conditions);
}

/**
 * Setzt Dropdown-Validierungen für ein Sheet mit Batch-Validierung
 * @param {Sheet} sheet - Das Sheet
 * @param {string} sheetName - Name des Sheets
 * @param {number} numRows - Anzahl der Datenzeilen
 * @param {Object} columns - Spaltenkonfiguration
 * @param {Object} config - Die Konfiguration
 */
function setDropdownValidations(sheet, sheetName, numRows, columns, config) {
    // Optimierung: Alle Validierungen in einem Objekt sammeln für weniger API-Calls
    const validations = [];

    // 1. Sheet-spezifische Validierungen
    if (sheetName === 'Einnahmen') {
        validations.push({
            column: columns.kategorie,
            values: Object.keys(config.einnahmen.categories),
        });
    } else if (sheetName === 'Ausgaben') {
        validations.push({
            column: columns.kategorie,
            values: Object.keys(config.ausgaben.categories),
        });
    } else if (sheetName === 'Eigenbelege') {
        validations.push({
            column: columns.kategorie,
            values: Object.keys(config.eigenbelege.categories),
        });

        // Dropdown für "Ausgelegt von"
        if (columns.ausgelegtVon) {
            validations.push({
                column: columns.ausgelegtVon,
                values: [...config.common.shareholders, ...config.common.employees],
            });
        }
    } else if (sheetName === 'Gesellschafterkonto') {
        // Kategorie-Dropdown für Gesellschafterkonto
        if (columns.kategorie) {
            validations.push({
                column: columns.kategorie,
                values: Object.keys(config.gesellschafterkonto.categories),
            });
        }

        // Dropdown für Gesellschafter
        if (columns.gesellschafter) {
            validations.push({
                column: columns.gesellschafter,
                values: config.common.shareholders,
            });
        }
    } else if (sheetName === 'Holding Transfers') {
        // Kategorie-Dropdown für Holding Transfers
        if (columns.kategorie) {
            validations.push({
                column: columns.kategorie,
                values: Object.keys(config.holdingTransfers.categories),
            });
        }

        // Zielgesellschaft-Dropdown für Holding Transfers
        if (columns.ziel) {
            validations.push({
                column: columns.ziel,
                values: config.common.companies,
            });
        }
    }

    // 2. Gemeinsame Validierungen für alle Sheets
    if (columns.zahlungsart) {
        validations.push({
            column: columns.zahlungsart,
            values: config.common.paymentType,
        });
    }

    // 4. Ausland-Dropdown für relevante Sheets
    if (columns.ausland) {
        validations.push({
            column: columns.ausland,
            values: config.common.auslandType,
        });
    }

    // 5. Alle Validierungen in effizienten Batches anwenden
    validations.forEach(validation => {
        if (validation.column && validation.values?.length > 0) {
            cellValidator.validateDropdown(
                sheet, 2, validation.column, numRows, 1,
                validation.values,
            );
        }
    });
}

/**
 * Wendet Validierungen auf das Bankbewegungen-Sheet an mit optimierter Batch-Verarbeitung
 * @param {Sheet} sheet - Das Sheet
 * @param {number} firstDataRow - Erste Datenzeile
 * @param {number} numDataRows - Anzahl der Datenzeilen
 * @param {Object} columns - Spaltenkonfiguration
 * @param {Object} config - Die Konfiguration
 */
function applyBankSheetValidations(sheet, firstDataRow, numDataRows, columns, config) {
    // Optimierung: Alle Validierungen in einem Array sammeln
    const validations = [];

    // 1. Validierung für Transaktionstyp
    if (columns.transaktionstyp) {
        validations.push({
            column: columns.transaktionstyp,
            values: config.bankbewegungen.types,
        });
    }

    // 2. Validierung für Kategorie
    if (columns.kategorie) {
        validations.push({
            column: columns.kategorie,
            values: config.bankbewegungen.categories,
        });
    }

    // 3. Konten für Dropdown-Validierung sammeln (nur einmal)
    const kontoCollections = collectAllowedAccounts(config);

    // 4. Validierungen für Konten
    if (columns.kontoSoll) {
        validations.push({
            column: columns.kontoSoll,
            values: Array.from(kontoCollections.allowedKontoSoll),
        });
    }

    if (columns.kontoHaben) {
        validations.push({
            column: columns.kontoHaben,
            values: Array.from(kontoCollections.allowedGegenkonto),
        });
    }

    // 5. Alle Validierungen in effizienten Batches anwenden
    validations.forEach(validation => {
        if (validation.column && validation.values?.length > 0) {
            cellValidator.validateDropdown(
                sheet, firstDataRow, validation.column, numDataRows, 1,
                validation.values,
            );
        }
    });
}

/**
 * Sammelt erlaubte Konten aus allen Mappings
 * @param {Object} config - Die Konfiguration
 * @returns {Object} Erlaubte Konten
 */
function collectAllowedAccounts(config) {
    const allowedKontoSoll = new Set();
    const allowedGegenkonto = new Set();

    // Optimierung: Alle Mappings in einem Array definieren
    const mappings = [
        config.einnahmen.kontoMapping,
        config.ausgaben.kontoMapping,
        config.eigenbelege.kontoMapping,
        config.gesellschafterkonto.kontoMapping,
        config.holdingTransfers.kontoMapping,
    ];

    // Sammle alle Konten aus den Mappings
    mappings.forEach(mapping => {
        Object.values(mapping).forEach(m => {
            if (m.soll) allowedKontoSoll.add(m.soll);
            if (m.gegen) allowedGegenkonto.add(m.gegen);
        });
    });

    return { allowedKontoSoll, allowedGegenkonto };
}

export default {
    setConditionalFormattingForStatusColumn,
    setDropdownValidations,
    applyBankSheetValidations,
};