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

/**
 * Formatiert Zeilen basierend auf dem Match-Typ
 */
function formatMatchedRows(sheet, firstDataRow, matchResults, columns) {
    // Performance-optimiertes Batch-Update vorbereiten
    const formatBatches = {
        'Einnahme': { rows: [], color: '#EAF1DD' },
        'Vollständige Zahlung (Einnahme)': { rows: [], color: '#C6EFCE' },
        'Teilzahlung (Einnahme)': { rows: [], color: '#FCE4D6' },
        'Ausgabe': { rows: [], color: '#FFCCCC' },
        'Vollständige Zahlung (Ausgabe)': { rows: [], color: '#FFC7CE' },
        'Teilzahlung (Ausgabe)': { rows: [], color: '#FCE4D6' },
        'Eigenbeleg': { rows: [], color: '#DDEBF7' },
        'Vollständige Zahlung (Eigenbeleg)': { rows: [], color: '#9BC2E6' },
        'Teilzahlung (Eigenbeleg)': { rows: [], color: '#FCE4D6' },
        'Gesellschafterkonto': { rows: [], color: '#E2EFDA' },
        'Holding Transfer': { rows: [], color: '#FFF2CC' },
        'Gutschrift': { rows: [], color: '#E6E0FF' },
        'Gesellschaftskonto/Holding': { rows: [], color: '#FFEB9C' },
    };

    // Zeilen nach Kategorien gruppieren
    matchResults.forEach((matchInfo, index) => {
        const rowIndex = firstDataRow + index;
        const matchText = (matchInfo && matchInfo[0]) ? matchInfo[0].toString() : '';

        if (!matchText) return; // Überspringe leere Matches

        if (matchText.includes('Einnahme')) {
            if (matchText.includes('Vollständige Zahlung')) {
                formatBatches['Vollständige Zahlung (Einnahme)'].rows.push(rowIndex);
            } else if (matchText.includes('Teilzahlung')) {
                formatBatches['Teilzahlung (Einnahme)'].rows.push(rowIndex);
            } else {
                formatBatches['Einnahme'].rows.push(rowIndex);
            }
        } else if (matchText.includes('Ausgabe')) {
            if (matchText.includes('Vollständige Zahlung')) {
                formatBatches['Vollständige Zahlung (Ausgabe)'].rows.push(rowIndex);
            } else if (matchText.includes('Teilzahlung')) {
                formatBatches['Teilzahlung (Ausgabe)'].rows.push(rowIndex);
            } else {
                formatBatches['Ausgabe'].rows.push(rowIndex);
            }
        } else if (matchText.includes('Eigenbeleg')) {
            if (matchText.includes('Vollständige Zahlung')) {
                formatBatches['Vollständige Zahlung (Eigenbeleg)'].rows.push(rowIndex);
            } else if (matchText.includes('Teilzahlung')) {
                formatBatches['Teilzahlung (Eigenbeleg)'].rows.push(rowIndex);
            } else {
                formatBatches['Eigenbeleg'].rows.push(rowIndex);
            }
        } else if (matchText.includes('Gesellschafterkonto')) {
            formatBatches['Gesellschafterkonto'].rows.push(rowIndex);
        } else if (matchText.includes('Holding Transfer')) {
            formatBatches['Holding Transfer'].rows.push(rowIndex);
        } else if (matchText.includes('Gutschrift')) {
            formatBatches['Gutschrift'].rows.push(rowIndex);
        } else if (matchText.includes('Gesellschaftskonto') || matchText.includes('Holding')) {
            formatBatches['Gesellschaftskonto/Holding'].rows.push(rowIndex);
        }
    });

    // Batch-Formatting anwenden
    Object.values(formatBatches).forEach(batch => {
        if (batch.rows.length > 0) {
            // Gruppen von maximal 20 Zeilen formatieren (API-Limits vermeiden)
            applyFormatBatches(sheet, batch.rows, batch.color, columns.matchInfo);
        }
    });
}

/**
 * Wendet Formatierung auf Batches von Zeilen an
 */
function applyFormatBatches(sheet, rows, color, maxColumn) {
    const chunkSize = 20;
    for (let i = 0; i < rows.length; i += chunkSize) {
        const chunk = rows.slice(i, i + chunkSize);
        chunk.forEach(rowIndex => {
            try {
                sheet.getRange(rowIndex, 1, 1, maxColumn)
                    .setBackground(color);
            } catch (e) {
                console.error(`Fehler beim Formatieren von Zeile ${rowIndex}:`, e);
            }
        });

        // Kurze Pause um API-Limits zu vermeiden
        if (i + chunkSize < rows.length) {
            Utilities.sleep(50);
        }
    }
}

/**
 * Setzt bedingte Formatierung für die Match-Spalte
 */
function setMatchColumnFormatting(sheet, columnLetter) {
    const conditions = [
        // Grundlegende Match-Typen
        {value: 'Einnahme', background: '#C6EFCE', fontColor: '#006100', pattern: 'beginsWith'},
        {value: 'Ausgabe', background: '#FFC7CE', fontColor: '#9C0006', pattern: 'beginsWith'},
        {value: 'Eigenbeleg', background: '#DDEBF7', fontColor: '#2F5597', pattern: 'beginsWith'},
        {value: 'Gesellschafterkonto', background: '#E2EFDA', fontColor: '#375623', pattern: 'beginsWith'},
        {value: 'Holding Transfer', background: '#FFF2CC', fontColor: '#7F6000', pattern: 'beginsWith'},
        {value: 'Gutschrift', background: '#E6E0FF', fontColor: '#5229A3', pattern: 'beginsWith'},

        // Zusätzliche Betragstypen
        {value: 'Vollständige Zahlung', background: '#C6EFCE', fontColor: '#006100', pattern: 'contains'},
        {value: 'Teilzahlung', background: '#FCE4D6', fontColor: '#974706', pattern: 'contains'},
        {value: 'Unsichere Zahlung', background: '#F8CBAD', fontColor: '#843C0C', pattern: 'contains'},
        {value: 'Vollständige Gutschrift', background: '#E6E0FF', fontColor: '#5229A3', pattern: 'contains'},
    ];

    // Bedingte Formatierung für die Match-Spalte setzen
    sheetUtils.setConditionalFormattingForColumn(sheet, columnLetter, conditions);
}

export default {
    setConditionalFormattingForStatusColumn,
    setDropdownValidations,
    applyBankSheetValidations,
    formatMatchedRows,
    applyFormatBatches,
    setMatchColumnFormatting,
};