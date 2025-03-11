// file: src/refreshModule.js
import Validator from "./validator.js";
import config from "./config.js";
import Helpers from "./helpers.js";

/**
 * Modul zum Aktualisieren der Tabellenblätter und Neuberechnen von Formeln
 */
const RefreshModule = (() => {
    /**
     * Aktualisiert ein Datenblatt (Einnahmen, Ausgaben, Eigenbelege)
     * @param {Sheet} sheet - Das zu aktualisierende Sheet
     * @returns {boolean} - true bei erfolgreicher Aktualisierung
     */
    const refreshDataSheet = sheet => {
        try {
            const lastRow = sheet.getLastRow();
            if (lastRow < 2) return true; // Keine Daten zum Aktualisieren

            const numRows = lastRow - 1;
            const name = sheet.getName();

            // Formeln für verschiedene Spalten setzen (ORIGINAL-FORMELN BEIBEHALTEN)
            const formulas = {
                // MwSt-Betrag (G) - KORRIGIERT: Originalformel verwenden
                7: row => `=E${row}*F${row}`,

                // Brutto-Betrag (H)
                8: row => `=E${row}+G${row}`,

                // Steuerbemessungsgrundlage - für Teilzahlungen (J)
                10: row => `=(H${row}-I${row})/(1+F${row})`,

                // Quartal (K)
                11: row => `=IF(A${row}="";"";ROUNDUP(MONTH(A${row})/3;0))`,

                // Zahlungsstatus (L)
                12: row => `=IF(VALUE(I${row})=0;"Offen";IF(VALUE(I${row})>=VALUE(H${row});"Bezahlt";"Teilbezahlt"))`
            };

            // Formeln für jede Spalte anwenden
            Object.entries(formulas).forEach(([col, formulaFn]) => {
                const formulasArray = Array.from({length: numRows}, (_, i) => [formulaFn(i + 2)]);
                sheet.getRange(2, Number(col), numRows, 1).setFormulas(formulasArray);
            });

            // Bezahlter Betrag (I) - Leerzeichen durch 0 ersetzen für Berechnungen
            const col9Range = sheet.getRange(2, 9, numRows, 1);
            const col9Values = col9Range.getValues().map(([val]) => (val === "" || val === null ? 0 : val));
            col9Range.setValues(col9Values.map(val => [val]));

            // Dropdown-Validierungen je nach Sheet-Typ setzen
            if (name === "Einnahmen") {
                Validator.validateDropdown(
                    sheet, 2, 3, numRows, 1,
                    Object.keys(config.einnahmen.categories)
                );
            } else if (name === "Ausgaben") {
                Validator.validateDropdown(
                    sheet, 2, 3, numRows, 1,
                    Object.keys(config.ausgaben.categories)
                );
            } else if (name === "Eigenbelege") {
                Validator.validateDropdown(
                    sheet, 2, 3, numRows, 1,
                    config.eigenbelege.category
                );

                // Für Eigenbelege: Status-Dropdown hinzufügen
                Validator.validateDropdown(
                    sheet, 2, 12, numRows, 1,
                    config.eigenbelege.status
                );

                // Bedingte Formatierung für Status-Spalte (nur für Eigenbelege)
                Helpers.setConditionalFormattingForColumn(sheet, "L", [
                    {value: "Offen", background: "#FFC7CE", fontColor: "#9C0006"},
                    {value: "Erstattet", background: "#FFEB9C", fontColor: "#9C6500"},
                    {value: "Gebucht", background: "#C6EFCE", fontColor: "#006100"}
                ]);
            }

            // Zahlungsart-Dropdown für alle Blätter
            Validator.validateDropdown(
                sheet, 2, 13, numRows, 1,
                config.common.paymentType
            );

            // Bedingte Formatierung für Zahlungsstatus-Spalte (für alle außer Eigenbelege)
            if (name !== "Eigenbelege") {
                Helpers.setConditionalFormattingForColumn(sheet, "L", [
                    {value: "Offen", background: "#FFC7CE", fontColor: "#9C0006"},
                    {value: "Teilbezahlt", background: "#FFEB9C", fontColor: "#9C6500"},
                    {value: "Bezahlt", background: "#C6EFCE", fontColor: "#006100"}
                ]);
            }

            // Spaltenbreiten automatisch anpassen
            sheet.autoResizeColumns(1, sheet.getLastColumn());

            return true;
        } catch (e) {
            console.error(`Fehler beim Aktualisieren von ${sheet.getName()}:`, e);
            return false;
        }
    };

    /**
     * Aktualisiert das Bankbewegungen-Sheet
     * @param {Sheet} sheet - Das Bankbewegungen-Sheet
     * @returns {boolean} - true bei erfolgreicher Aktualisierung
     */
    const refreshBankSheet = sheet => {
        try {
            const lastRow = sheet.getLastRow();
            if (lastRow < 3) return true; // Nicht genügend Daten zum Aktualisieren

            const firstDataRow = 3; // Erste Datenzeile (nach Header-Zeile)
            const numDataRows = lastRow - firstDataRow + 1;
            const transRows = lastRow - firstDataRow - 1; // Anzahl der Transaktionszeilen ohne die letzte Zeile

            // Saldo-Formeln setzen (jede Zeile verwendet den Saldo der vorherigen Zeile + aktuellen Betrag)
            if (transRows > 0) {
                sheet.getRange(firstDataRow, 4, transRows, 1).setFormulas(
                    Array.from({length: transRows}, (_, i) =>
                        [`=D${firstDataRow + i - 1}+C${firstDataRow + i}`]
                    )
                );
            }

            // Transaktionstyp basierend auf dem Betrag setzen (Einnahme/Ausgabe)
            const amounts = sheet.getRange(firstDataRow, 3, numDataRows, 1).getValues();
            const typeValues = amounts.map(([val]) => {
                const amt = parseFloat(val) || 0;
                return [amt > 0 ? "Einnahme" : amt < 0 ? "Ausgabe" : ""];
            });
            sheet.getRange(firstDataRow, 5, numDataRows, 1).setValues(typeValues);

            // Dropdown-Validierungen für Typ, Kategorie und Konten
            Validator.validateDropdown(
                sheet, firstDataRow, 5, numDataRows, 1,
                config.bank.type
            );

            Validator.validateDropdown(
                sheet, firstDataRow, 6, numDataRows, 1,
                config.bank.category
            );

            // Konten für Dropdown-Validierung sammeln (wie im Original)
            const allowedKontoSoll = Object.values(config.einnahmen.kontoMapping)
                .concat(Object.values(config.ausgaben.kontoMapping))
                .map(m => m.soll);

            const allowedGegenkonto = Object.values(config.einnahmen.kontoMapping)
                .concat(Object.values(config.ausgaben.kontoMapping))
                .map(m => m.gegen);

            // Dropdown-Validierungen für Konten setzen
            Validator.validateDropdown(
                sheet, firstDataRow, 7, numDataRows, 1,
                allowedKontoSoll
            );

            Validator.validateDropdown(
                sheet, firstDataRow, 8, numDataRows, 1,
                allowedGegenkonto
            );

            // Bedingte Formatierung für Transaktionstyp-Spalte
            Helpers.setConditionalFormattingForColumn(sheet, "E", [
                {value: "Einnahme", background: "#C6EFCE", fontColor: "#006100"},
                {value: "Ausgabe", background: "#FFC7CE", fontColor: "#9C0006"}
            ]);

            // Kontonummern basierend auf Kategorie automatisch zuweisen
            const dataRange = sheet.getRange(firstDataRow, 1, numDataRows, sheet.getLastColumn());
            const data = dataRange.getValues();

            data.forEach((row, i) => {
                const globalRow = i + firstDataRow;

                // Letzte Zeile (Endsaldo) überspringen
                const label = row[1] ? row[1].toString().trim().toLowerCase() : "";
                if (globalRow === lastRow && label === "endsaldo") return;

                const type = row[4];
                const category = row[5] || "";

                let mapping = null;

                // Mapping basierend auf Typ und Kategorie finden
                if (type === "Einnahme") {
                    mapping = config.einnahmen.kontoMapping[category];
                } else if (type === "Ausgabe") {
                    mapping = config.ausgaben.kontoMapping[category];
                }

                // Falls keine Zuordnung gefunden, Hinweis setzen
                if (!mapping) {
                    mapping = {soll: "Manuell prüfen", gegen: "Manuell prüfen"};
                }

                row[6] = mapping.soll;
                row[7] = mapping.gegen;
            });

            // Aktualisierte Daten zurückschreiben
            dataRange.setValues(data);

            // Endsaldo-Zeile prüfen und aktualisieren/hinzufügen
            const lastRowText = sheet.getRange(lastRow, 2).getValue().toString().trim().toLowerCase();
            const formattedDate = Utilities.formatDate(
                new Date(),
                Session.getScriptTimeZone(),
                "dd.MM.yyyy"
            );

            if (lastRowText === "endsaldo") {
                // Datum und Saldo-Formel aktualisieren
                sheet.getRange(lastRow, 1).setValue(formattedDate);
                sheet.getRange(lastRow, 4).setFormula(`=D${lastRow - 1}`);
            } else {
                // Endsaldo-Zeile hinzufügen
                sheet.appendRow([
                    formattedDate,
                    "Endsaldo",
                    "",
                    sheet.getRange(lastRow, 4).getValue(),
                    "",
                    "",
                    "",
                    "",
                    "",
                    "",
                    "",
                    ""
                ]);
            }

            // Spaltenbreiten automatisch anpassen
            sheet.autoResizeColumns(1, sheet.getLastColumn());

            return true;
        } catch (e) {
            console.error("Fehler beim Aktualisieren des Bankbewegungen-Sheets:", e);
            return false;
        }
    };

    /**
     * Aktualisiert das aktive Tabellenblatt
     */
    const refreshActiveSheet = () => {
        try {
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            const sheet = ss.getActiveSheet();
            const name = sheet.getName();
            const ui = SpreadsheetApp.getUi();

            if (["Einnahmen", "Ausgaben", "Eigenbelege"].includes(name)) {
                refreshDataSheet(sheet);
                ui.alert(`Das Blatt "${name}" wurde aktualisiert.`);
            } else if (name === "Bankbewegungen") {
                refreshBankSheet(sheet);
                ui.alert(`Das Blatt "${name}" wurde aktualisiert.`);
            } else {
                ui.alert("Für dieses Blatt gibt es keine Refresh-Funktion.");
            }
        } catch (e) {
            console.error("Fehler beim Aktualisieren des aktiven Sheets:", e);
            SpreadsheetApp.getUi().alert("Ein Fehler ist beim Aktualisieren aufgetreten: " + e.toString());
        }
    };

    /**
     * Aktualisiert alle relevanten Tabellenblätter
     */
    const refreshAllSheets = () => {
        try {
            const ss = SpreadsheetApp.getActiveSpreadsheet();

            ["Einnahmen", "Ausgaben", "Eigenbelege", "Bankbewegungen"].forEach(name => {
                const sheet = ss.getSheetByName(name);
                if (sheet) {
                    name === "Bankbewegungen" ? refreshBankSheet(sheet) : refreshDataSheet(sheet);
                }
            });
        } catch (e) {
            console.error("Fehler beim Aktualisieren aller Sheets:", e);
        }
    };

    // Öffentliche API des Moduls
    return {
        refreshActiveSheet,
        refreshAllSheets
    };
})();

export default RefreshModule;