// src/modules/refreshModule/bankSheetHandler.js
import formattingHandler from './formattingHandler.js';
import accountHandler from './accountHandler.js';
import stringUtils from '../../utils/stringUtils.js';
import numberUtils from '../../utils/numberUtils.js';

/**
 * Aktualisiert das Bankbewegungen-Sheet mit optimierter Batch-Verarbeitung
 * @param {Sheet} sheet - Das Bankbewegungen-Sheet
 * @param {Object} config - Die Konfiguration
 * @returns {boolean} - true bei erfolgreicher Aktualisierung
 */
function refreshBankSheet(sheet, config) {
    try {
        console.log('Refreshing bank sheet');
        const lastRow = sheet.getLastRow();
        if (lastRow < 3) return true; // Nicht genügend Daten zum Aktualisieren

        const firstDataRow = 3; // Erste Datenzeile (nach Header-Zeile)
        const numDataRows = lastRow - firstDataRow + 1;
        const transRows = lastRow - firstDataRow; // Anzahl der Transaktionszeilen ohne Endsaldo

        console.log(`Found ${numDataRows} rows to process in bank sheet`);

        // Bankbewegungen-Konfiguration für Spalten
        const columns = config.bankbewegungen.columns;

        // Spaltenbuchstaben einmalig generieren
        const columnLetters = {};
        Object.entries(columns).forEach(([key, index]) => {
            columnLetters[key] = stringUtils.getColumnLetter(index);
        });

        // Optimierung: Alle Daten in einem API-Call laden
        const allData = sheet.getRange(firstDataRow, 1, numDataRows, sheet.getLastColumn()).getValues();

        // Optimierung: Arrays für Batch-Updates
        const updates = {
            saldiFormeln: [],
            types: [],
            kontoSoll: [],
            kontoHaben: [],
        };

        // 1. Saldo-Formeln und Transaktionstypen in einem Durchgang erstellen
        for (let i = 0; i < numDataRows; i++) {
            const row = allData[i];
            const rowIndex = firstDataRow + i;

            // Saldo-Formel (außer für die letzte Zeile)
            if (i < numDataRows - 1) {
                updates.saldiFormeln.push([
                    `=${columnLetters.saldo}${rowIndex - 1}+${columnLetters.betrag}${rowIndex}`,
                ]);
            }

            // Transaktionstyp basierend auf Betrag
            const betrag = numberUtils.parseCurrency(row[columns.betrag - 1]);
            updates.types.push([betrag > 0 ? 'Einnahme' : betrag < 0 ? 'Ausgabe' : '']);
        }

        // 2. Batch-Updates durchführen

        // Saldo-Formeln setzen
        if (updates.saldiFormeln.length > 0) {
            sheet.getRange(firstDataRow, columns.saldo, updates.saldiFormeln.length, 1)
                .setFormulas(updates.saldiFormeln);
        }

        // Transaktionstypen setzen
        sheet.getRange(firstDataRow, columns.transaktionstyp, numDataRows, 1)
            .setValues(updates.types);

        // 3. Dropdown-Validierungen und bedingte Formatierung
        formattingHandler.applyBankSheetValidations(sheet, firstDataRow, numDataRows, columns, config);

        // 4. Bedingte Formatierung für Transaktionstyp
        formattingHandler.setConditionalFormattingForStatusColumn(sheet, columnLetters.transaktionstyp, [
            {value: 'Einnahme', background: '#C6EFCE', fontColor: '#006100'},
            {value: 'Ausgabe', background: '#FFC7CE', fontColor: '#9C0006'},
        ]);

        // 5. Buchungskonten aktualisieren
        accountHandler.updateBookingAccounts(sheet, 'bankbewegungen', config, true);

        // 6. Endsaldo-Zeile aktualisieren
        updateEndSaldoRow(sheet, lastRow, columns, columnLetters);

        // 7. Spaltenbreiten anpassen (nur bei Änderungen)
        sheet.autoResizeColumns(1, sheet.getLastColumn());

        console.log('Bank sheet refreshed successfully');
        return true;
    } catch (e) {
        console.error('Fehler beim Aktualisieren des Bankbewegungen-Sheets:', e);
        return false;
    }
}

/**
 * Aktualisiert die Endsaldo-Zeile im Bankbewegungen-Sheet
 * @param {Sheet} sheet - Das Sheet
 * @param {number} lastRow - Letzte Zeile
 * @param {Object} columns - Spaltenkonfiguration
 * @param {Object} columnLetters - Buchstaben für die Spalten
 */
function updateEndSaldoRow(sheet, lastRow, columns, columnLetters) {
    try {
        // Endsaldo-Zeile prüfen
        const lastRowData = sheet.getRange(lastRow, 1, 1, columns.buchungstext).getValues()[0];
        const lastRowText = lastRowData[columns.buchungstext - 1]?.toString().trim().toLowerCase() || '';

        let endSaldoDate = '';
        if (lastRow > 1) {
            const previousRowDate = sheet.getRange(lastRow - 1, columns.datum).getDisplayValue();
            endSaldoDate = previousRowDate || '';
        }

        const updates = {};

        if (lastRowText === 'endsaldo') {
            // Endsaldo-Zeile aktualisieren
            updates.date = endSaldoDate;
            updates.formula = `=${columnLetters.saldo}${lastRow - 1}`;
        } else {
            // Neue Endsaldo-Zeile erstellen
            const newRow = Array(sheet.getLastColumn()).fill('');
            newRow[columns.datum - 1] = endSaldoDate;
            newRow[columns.buchungstext - 1] = 'Endsaldo';
            sheet.appendRow(newRow);

            // Saldo-Formel setzen
            const newLastRow = lastRow + 1;
            updates.row = newLastRow;
            updates.formula = `=${columnLetters.saldo}${lastRow}`;
        }

        // Updates anwenden
        if (updates.date) {
            sheet.getRange(lastRow, columns.datum).setValue(updates.date);
        }

        if (updates.formula) {
            const targetRow = updates.row || lastRow;
            sheet.getRange(targetRow, columns.saldo).setFormula(updates.formula);
        }
    } catch (e) {
        console.error('Fehler beim Aktualisieren der Endsaldo-Zeile:', e);
    }
}

export default {
    refreshBankSheet,
};