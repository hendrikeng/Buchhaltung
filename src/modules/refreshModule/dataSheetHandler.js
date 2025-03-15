// modules/refreshModule/dataSheetHandler.js
import stringUtils from '../../utils/stringUtils.js';
import formattingHandler from './formattingHandler.js';

/**
 * Aktualisiert ein Datenblatt (Einnahmen, Ausgaben, Eigenbelege)
 * @param {Sheet} sheet - Das zu aktualisierende Sheet
 * @param {string} sheetName - Name des Sheets
 * @param {Object} config - Die Konfiguration
 * @returns {boolean} - true bei erfolgreicher Aktualisierung
 */
function refreshDataSheet(sheet, sheetName, config) {
    try {
        const lastRow = sheet.getLastRow();
        if (lastRow < 2) return true; // Keine Daten zum Aktualisieren

        const numRows = lastRow - 1;

        // Passende Spaltenkonfiguration für das entsprechende Sheet auswählen
        let columns;
        if (sheetName === "Einnahmen") {
            columns = config.einnahmen.columns;
        } else if (sheetName === "Ausgaben") {
            columns = config.ausgaben.columns;
        } else if (sheetName === "Eigenbelege") {
            columns = config.eigenbelege.columns;
        } else if (sheetName === "Gesellschafterkonto") {
            columns = config.gesellschafterkonto.columns;
        } else if (sheetName === "Holding Transfers") {
            columns = config.holdingTransfers.columns;
        } else {
            return false; // Unbekanntes Sheet
        }

        // Spaltenbuchstaben aus den Indizes generieren
        const columnLetters = {};
        Object.entries(columns).forEach(([key, index]) => {
            columnLetters[key] = stringUtils.getColumnLetter(index);
        });

        // Batch-Array für Formeln erstellen (effizienter als einzelne Range-Updates)
        const formulasBatch = {};

        // Standardformeln setzen
        setStandardFormulas(formulasBatch, columnLetters, numRows, sheetName);

        // Spezifische Formeln für bestimmte Sheet-Typen
        if (sheetName === "Gesellschafterkonto") {
            // Hier später spezifische Formeln für Gesellschafterkonto ergänzen
        } else if (sheetName === "Holding Transfers") {
            // Hier später spezifische Formeln für Holding Transfers ergänzen
        }

        // Formeln in Batches anwenden (weniger API-Calls)
        Object.entries(formulasBatch).forEach(([col, formulas]) => {
            sheet.getRange(2, Number(col), numRows, 1).setFormulas(formulas);
        });

        // Bezahlter Betrag - Leerzeichen durch 0 ersetzen für Berechnungen
        if (columns.bezahlt) {
            const bezahltRange = sheet.getRange(2, columns.bezahlt, numRows, 1);
            const bezahltValues = bezahltRange.getValues();
            const updatedBezahltValues = bezahltValues.map(
                ([val]) => [stringUtils.isEmpty(val) ? 0 : val]
            );
            bezahltRange.setValues(updatedBezahltValues);
        }

        // Dropdown-Validierungen je nach Sheet-Typ setzen
        formattingHandler.setDropdownValidations(sheet, sheetName, numRows, columns, config);

        // Bedingte Formatierung für Status-Spalte
        if (columns.zahlungsstatus) {
            setStatusFormatting(sheet, sheetName, columnLetters.zahlungsstatus);
        }

        // Spaltenbreiten automatisch anpassen
        sheet.autoResizeColumns(1, sheet.getLastColumn());

        return true;
    } catch (e) {
        console.error(`Fehler beim Aktualisieren von ${sheetName}:`, e);
        return false;
    }
}

/**
 * Setzt Standardformeln für ein Datenblatt
 * @param {Object} formulasBatch - Objekt für Batch-Formeln
 * @param {Object} columnLetters - Spaltenbuchstaben
 * @param {number} numRows - Anzahl der Datenzeilen
 * @param {string} sheetName - Name des Sheets
 */
function setStandardFormulas(formulasBatch, columnLetters, numRows, sheetName) {
    // Prüfen, ob die benötigten Spalten vorhanden sind
    const hasRequiredColumns = (
        columnLetters.mwstBetrag &&
        columnLetters.nettobetrag &&
        columnLetters.mwstSatz &&
        columnLetters.bruttoBetrag &&
        columnLetters.restbetragNetto &&
        columnLetters.bezahlt &&
        columnLetters.datum &&
        columnLetters.quartal
    );

    if (!hasRequiredColumns) return;

    // MwSt-Betrag
    formulasBatch[columnLetters.mwstBetrag] = Array.from(
        {length: numRows},
        (_, i) => [`=${columnLetters.nettobetrag}${i + 2}*${columnLetters.mwstSatz}${i + 2}/100`]
    );

    // Brutto-Betrag
    formulasBatch[columnLetters.bruttoBetrag] = Array.from(
        {length: numRows},
        (_, i) => [`=${columnLetters.nettobetrag}${i + 2}+${columnLetters.mwstBetrag}${i + 2}`]
    );

    // Bezahlter Betrag - für Teilzahlungen
    formulasBatch[columnLetters.restbetragNetto] = Array.from(
        {length: numRows},
        (_, i) => [`=(${columnLetters.bruttoBetrag}${i + 2}-${columnLetters.bezahlt}${i + 2})/(1+${columnLetters.mwstSatz}${i + 2}/100)`]
    );

    // Quartal
    formulasBatch[columnLetters.quartal] = Array.from(
        {length: numRows},
        (_, i) => [`=IF(${columnLetters.datum}${i + 2}="";"";ROUNDUP(MONTH(${columnLetters.datum}${i + 2})/3;0))`]
    );

    // Zahlungsstatus
    if (columnLetters.zahlungsstatus) {
        if (sheetName !== "Eigenbelege") {
            // Für Einnahmen und Ausgaben: Zahlungsstatus
            formulasBatch[columnLetters.zahlungsstatus] = Array.from(
                {length: numRows},
                (_, i) => [`=IF(VALUE(${columnLetters.bezahlt}${i + 2})=0;"Offen";IF(VALUE(${columnLetters.bezahlt}${i + 2})>=VALUE(${columnLetters.bruttoBetrag}${i + 2});"Bezahlt";"Teilbezahlt"))`]
            );
        } else {
            // Für Eigenbelege: Zahlungsstatus
            formulasBatch[columnLetters.zahlungsstatus] = Array.from(
                {length: numRows},
                (_, i) => [`=IF(VALUE(${columnLetters.bezahlt}${i + 2})=0;"Offen";IF(VALUE(${columnLetters.bezahlt}${i + 2})>=VALUE(${columnLetters.bruttoBetrag}${i + 2});"Erstattet";"Teilerstattet"))`]
            );
        }
    }
}

/**
 * Setzt bedingte Formatierung für die Status-Spalte
 * @param {Sheet} sheet - Das Sheet
 * @param {string} sheetName - Name des Sheets
 * @param {string} statusColumn - Statusspaltenbuchstabe
 */
function setStatusFormatting(sheet, sheetName, statusColumn) {
    if (sheetName !== "Eigenbelege") {
        // Für Einnahmen und Ausgaben: Zahlungsstatus
        formattingHandler.setConditionalFormattingForStatusColumn(sheet, statusColumn, [
            {value: "Offen", background: "#FFC7CE", fontColor: "#9C0006"},
            {value: "Teilbezahlt", background: "#FFEB9C", fontColor: "#9C6500"},
            {value: "Bezahlt", background: "#C6EFCE", fontColor: "#006100"}
        ]);
    } else {
        // Für Eigenbelege: Status
        formattingHandler.setConditionalFormattingForStatusColumn(sheet, statusColumn, [
            {value: "Offen", background: "#FFC7CE", fontColor: "#9C0006"},
            {value: "Teilerstattet", background: "#FFEB9C", fontColor: "#9C6500"},
            {value: "Erstattet", background: "#C6EFCE", fontColor: "#006100"}
        ]);
    }
}

export default {
    refreshDataSheet
};