// src/modules/refreshModule/dataSheetHandler.js

import stringUtils from '../../utils/stringUtils.js';
import formattingHandler from './formattingHandler.js';
import cellValidator from '../validatorModule/cellValidator.js';

/**
 * Aktualisiert ein Datenblatt (Einnahmen, Ausgaben, Eigenbelege)
 * @param {Sheet} sheet - Das zu aktualisierende Sheet
 * @param {string} sheetName - Name des Sheets
 * @param {Object} config - Die Konfiguration
 * @returns {boolean} - true bei erfolgreicher Aktualisierung
 */
function refreshDataSheet(sheet, sheetName, config) {
    try {
        console.log(`Refreshing ${sheetName} sheet`);
        const lastRow = sheet.getLastRow();
        if (lastRow < 2) return true; // Keine Daten zum Aktualisieren

        const numRows = lastRow - 1;

        // Passende Spaltenkonfiguration für das entsprechende Sheet auswählen
        let columns;
        let configKey;

        if (sheetName === 'Einnahmen') {
            configKey = 'einnahmen';
        } else if (sheetName === 'Ausgaben') {
            configKey = 'ausgaben';
        } else if (sheetName === 'Eigenbelege') {
            configKey = 'eigenbelege';
        } else if (sheetName === 'Gesellschafterkonto') {
            configKey = 'gesellschafterkonto';
        } else if (sheetName === 'Holding Transfers') {
            configKey = 'holdingTransfers';
        } else {
            return false; // Unbekanntes Sheet
        }

        columns = config[configKey].columns;
        const kontoMapping = config[configKey].kontoMapping;

        console.log(`Found ${numRows} rows to process with columns:`, columns);

        // Spaltenbuchstaben aus den Indizes generieren
        const columnLetters = {};
        Object.entries(columns).forEach(([key, index]) => {
            columnLetters[key] = stringUtils.getColumnLetter(index);
        });

        // Batch-Array für Formeln erstellen (effizienter als einzelne Range-Updates)
        const formulasBatch = {};

        // MwSt-Betrag
        if (columns.mwstBetrag && columns.nettobetrag && columns.mwstSatz) {
            formulasBatch[columns.mwstBetrag] = Array.from(
                {length: numRows},
                (_, i) => [`=${columnLetters.nettobetrag}${i + 2}*${columnLetters.mwstSatz}${i + 2}/100`],
            );
        }

        // Brutto-Betrag
        if (columns.bruttoBetrag && columns.nettobetrag && columns.mwstBetrag) {
            formulasBatch[columns.bruttoBetrag] = Array.from(
                {length: numRows},
                (_, i) => [`=${columnLetters.nettobetrag}${i + 2}+${columnLetters.mwstBetrag}${i + 2}`],
            );
        }

        // Restbetrag Netto für Teilzahlungen
        if (columns.restbetragNetto && columns.bruttoBetrag && columns.bezahlt && columns.mwstSatz) {
            formulasBatch[columns.restbetragNetto] = Array.from(
                {length: numRows},
                (_, i) => [`=(${columnLetters.bruttoBetrag}${i + 2}-${columnLetters.bezahlt}${i + 2})/(1+${columnLetters.mwstSatz}${i + 2}/100)`],
            );
        }

        // Quartal
        if (columns.quartal && columns.datum) {
            formulasBatch[columns.quartal] = Array.from(
                {length: numRows},
                (_, i) => [`=IF(${columnLetters.datum}${i + 2}="";"";ROUNDUP(MONTH(${columnLetters.datum}${i + 2})/3;0))`],
            );
        }

        // Zahlungsstatus
        if (columns.zahlungsstatus && columns.bezahlt && columns.bruttoBetrag) {
            if (sheetName !== 'Eigenbelege') {
                // Für Einnahmen und Ausgaben: Zahlungsstatus
                formulasBatch[columns.zahlungsstatus] = Array.from(
                    {length: numRows},
                    (_, i) => [`=IF(VALUE(${columnLetters.bezahlt}${i + 2})=0;"Offen";IF(VALUE(${columnLetters.bezahlt}${i + 2})>=VALUE(${columnLetters.bruttoBetrag}${i + 2});"Bezahlt";"Teilbezahlt"))`],
                );
            } else if (sheetName === 'Eigenbelege') {
                // Für Eigenbelege: Zahlungsstatus
                formulasBatch[columns.zahlungsstatus] = Array.from(
                    {length: numRows},
                    (_, i) => [`=IF(VALUE(${columnLetters.bezahlt}${i + 2})=0;"Offen";IF(VALUE(${columnLetters.bezahlt}${i + 2})>=VALUE(${columnLetters.bruttoBetrag}${i + 2});"Erstattet";"Teilerstattet"))`],
                );
            }
        }

        console.log('Formula batch keys:', Object.keys(formulasBatch));

        // Formeln in Batches anwenden (weniger API-Calls)
        Object.entries(formulasBatch).forEach(([col, formulas]) => {
            try {
                console.log(`Setting ${formulas.length} formulas for column ${col}`);
                sheet.getRange(2, Number(col), numRows, 1).setFormulas(formulas);
            } catch (e) {
                console.error(`Fehler beim Setzen der Formeln für Spalte ${col}:`, e);
            }
        });

        // Bezahlter Betrag - Leerzeichen durch 0 ersetzen für Berechnungen
        if (columns.bezahlt) {
            const bezahltRange = sheet.getRange(2, columns.bezahlt, numRows, 1);
            const bezahltValues = bezahltRange.getValues();
            const updatedBezahltValues = bezahltValues.map(
                ([val]) => [stringUtils.isEmpty(val) ? 0 : val],
            );
            bezahltRange.setValues(updatedBezahltValues);
        }

        // Dropdown-Validierungen je nach Sheet-Typ setzen
        console.log('Setting dropdown validations');
        formattingHandler.setDropdownValidations(sheet, sheetName, numRows, columns, config);

        // Bedingte Formatierung für Status-Spalte
        if (columns.zahlungsstatus) {
            console.log('Setting conditional formatting for status column');
            if (sheetName !== 'Eigenbelege') {
                // Für Einnahmen und Ausgaben: Zahlungsstatus
                formattingHandler.setConditionalFormattingForStatusColumn(sheet, columnLetters.zahlungsstatus, [
                    {value: 'Offen', background: '#FFC7CE', fontColor: '#9C0006'},
                    {value: 'Teilbezahlt', background: '#FFEB9C', fontColor: '#9C6500'},
                    {value: 'Bezahlt', background: '#C6EFCE', fontColor: '#006100'},
                ]);
            } else if (sheetName === 'Eigenbelege') {
                // Für Eigenbelege: Status
                formattingHandler.setConditionalFormattingForStatusColumn(sheet, columnLetters.zahlungsstatus, [
                    {value: 'Offen', background: '#FFC7CE', fontColor: '#9C0006'},
                    {value: 'Teilerstattet', background: '#FFEB9C', fontColor: '#9C6500'},
                    {value: 'Erstattet', background: '#C6EFCE', fontColor: '#006100'},
                ]);
            }
        }

        // NEU: Buchungskonto basierend auf Kategorie setzen
        updateBookingAccounts(sheet, configKey, config);

        // Spaltenbreiten automatisch anpassen
        sheet.autoResizeColumns(1, sheet.getLastColumn());

        console.log(`${sheetName} sheet refreshed successfully`);
        return true;
    } catch (e) {
        console.error(`Fehler beim Aktualisieren von ${sheetName}:`, e);
        return false;
    }
}

/**
 * Aktualisiert die Buchungskonten basierend auf der ausgewählten Kategorie
 * @param {Sheet} sheet - Das Sheet
 * @param {string} configKey - Konfigurationsschlüssel (einnahmen, ausgaben, etc.)
 * @param {Object} config - Die Konfiguration
 */
function updateBookingAccounts(sheet, configKey, config) {
    const columns = config[configKey].columns;
    const kontoMapping = config[configKey].kontoMapping;

    if (!columns.kategorie || !columns.buchungsKonto) {
        console.log(`Keine Kategorie- oder Buchungskonto-Spalten für ${configKey} gefunden.`);
        return;
    }

    try {
        const lastRow = sheet.getLastRow();
        const numRows = lastRow - 1;

        if (numRows <= 0) return;

        console.log(`Updating booking accounts for ${numRows} rows in ${configKey}`);

        // Kategorie-Daten holen
        const kategorieRange = sheet.getRange(2, columns.kategorie, numRows, 1);
        const kategorieValues = kategorieRange.getValues();

        // Aktuelle Buchungskonto-Daten holen
        const kontoRange = sheet.getRange(2, columns.buchungsKonto, numRows, 1);
        const kontoValues = kontoRange.getValues();

        // Neue Buchungskonto-Werte basierend auf Kategorien erstellen
        const updatedKontoValues = [];

        for (let i = 0; i < numRows; i++) {
            const kategorie = kategorieValues[i][0] ? kategorieValues[i][0].toString().trim() : '';
            let buchungskonto = kontoValues[i][0] ? kontoValues[i][0].toString() : '';

            // Wenn Kategorie vorhanden und ein Mapping existiert, setze das Konto
            if (kategorie && kontoMapping[kategorie]) {
                // Je nach Sheet-Typ unterschiedliche Konto-Informationen verwenden
                if (configKey === 'einnahmen') {
                    buchungskonto = kontoMapping[kategorie].gegen || '';
                } else {
                    buchungskonto = kontoMapping[kategorie].soll || '';
                }
            }

            updatedKontoValues.push([buchungskonto]);
        }

        // Konten in Batch aktualisieren
        if (updatedKontoValues.length > 0) {
            kontoRange.setValues(updatedKontoValues);
            console.log(`Updated ${updatedKontoValues.length} booking accounts in ${configKey}`);
        }
    } catch (e) {
        console.error(`Fehler beim Aktualisieren der Buchungskonten für ${configKey}:`, e);
    }
}

export default {
    refreshDataSheet,
    updateBookingAccounts,
};