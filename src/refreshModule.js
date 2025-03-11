import Validator from "./validator.js";
import config from "./config.js";
import Helpers from "./helpers.js";

const RefreshModule = (() => {
    const refreshDataSheet = sheet => {
        const lastRow = sheet.getLastRow();
        if (lastRow < 2) return;
        const numRows = lastRow - 1;

        // Ersetze den Formelteil für Spalte 12:
        Object.entries({
            7: row => `=E${row}*F${row}`,
            8: row => `=E${row}+G${row}`,
            10: row => `=(H${row}-I${row})/(1+F${row})`,
            11: row => `=IF(A${row}="";"";ROUNDUP(MONTH(A${row})/3;0))`,
            12: row => `=IF(VALUE(I${row})=0;"Offen";IF(VALUE(I${row})>=VALUE(H${row});"Bezahlt";"Teilbezahlt"))`  // Überflüssiges OR entfernt
        }).forEach(([col, formulaFn]) => {
            const formulas = Array.from({length: numRows}, (_, i) => [formulaFn(i + 2)]);
            sheet.getRange(2, Number(col), numRows, 1).setFormulas(formulas);
        });

        const col9Range = sheet.getRange(2, 9, numRows, 1);
        const col9Values = col9Range.getValues().map(([val]) => (val === "" || val === null ? 0 : val));
        col9Range.setValues(col9Values.map(val => [val]));

        const name = sheet.getName();
        if (name === "Einnahmen")
            Validator.validateDropdown(sheet, 2, 3, lastRow - 1, 1, Object.keys(config.einnahmen.categories));
        if (name === "Ausgaben")
            Validator.validateDropdown(sheet, 2, 3, lastRow - 1, 1, Object.keys(config.ausgaben.categories));
        if (name === "Eigenbelege")
            Validator.validateDropdown(sheet, 2, 3, lastRow - 1, 1, config.eigenbelege.category);
        Validator.validateDropdown(sheet, 2, 13, lastRow - 1, 1, config.common.paymentType);
        sheet.autoResizeColumns(1, sheet.getLastColumn());
    };

    const refreshBankSheet = sheet => {
        const lastRow = sheet.getLastRow();
        if (lastRow < 3) return;
        const firstDataRow = 3;
        const numDataRows = lastRow - firstDataRow + 1;
        const transRows = lastRow - firstDataRow - 1;

        if (transRows > 0) {
            sheet.getRange(firstDataRow, 4, transRows, 1).setFormulas(
                Array.from({length: transRows}, (_, i) => [`=D${firstDataRow + i - 1}+C${firstDataRow + i}`])
            );
        }

        const amounts = sheet.getRange(firstDataRow, 3, numDataRows, 1).getValues();
        const typeValues = amounts.map(([val]) => {
            const amt = parseFloat(val) || 0;
            return [amt > 0 ? "Einnahme" : amt < 0 ? "Ausgabe" : ""];
        });
        sheet.getRange(firstDataRow, 5, numDataRows, 1).setValues(typeValues);

        Validator.validateDropdown(sheet, firstDataRow, 5, numDataRows, 1, config.bank.type);
        Validator.validateDropdown(sheet, firstDataRow, 6, numDataRows, 1, config.bank.category);
        const allowedKontoSoll = Object.values(config.einnahmen.kontoMapping)
            .concat(Object.values(config.ausgaben.kontoMapping))
            .map(m => m.soll);
        const allowedGegenkonto = Object.values(config.einnahmen.kontoMapping)
            .concat(Object.values(config.ausgaben.kontoMapping))
            .map(m => m.gegen);
        Validator.validateDropdown(sheet, firstDataRow, 7, numDataRows, 1, allowedKontoSoll);
        Validator.validateDropdown(sheet, firstDataRow, 8, numDataRows, 1, allowedGegenkonto);

        Helpers.setConditionalFormattingForColumn(sheet, "E", [
            {value: "Einnahme", background: "#C6EFCE", fontColor: "#006100"},
            {value: "Ausgabe", background: "#FFC7CE", fontColor: "#9C0006"}
        ]);

        const dataRange = sheet.getRange(firstDataRow, 1, numDataRows, sheet.getLastColumn());
        const data = dataRange.getValues();
        data.forEach((row, i) => {
            const globalRow = i + firstDataRow;
            const label = row[1] ? row[1].toString().trim().toLowerCase() : "";
            if (globalRow === lastRow && label === "endsaldo") return;
            const type = row[4];
            const category = row[5] || "";
            let mapping = type === "Einnahme"
                ? config.einnahmen.kontoMapping[category]
                : type === "Ausgabe"
                    ? config.ausgaben.kontoMapping[category]
                    : null;
            if (!mapping) mapping = {soll: "Manuell prüfen", gegen: "Manuell prüfen"};
            row[6] = mapping.soll;
            row[7] = mapping.gegen;
        });
        dataRange.setValues(data);

        const lastRowText = sheet.getRange(lastRow, 2).getValue().toString().trim().toLowerCase();
        const formattedDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd.MM.yyyy");
        if (lastRowText === "endsaldo") {
            sheet.getRange(lastRow, 1).setValue(formattedDate);
            sheet.getRange(lastRow, 4).setFormula(`=D${lastRow - 1}`);
        } else {
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
        sheet.autoResizeColumns(1, sheet.getLastColumn());
    };

    const refreshActiveSheet = () => {
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
    };

    const refreshAllSheets = () => {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        ["Einnahmen", "Ausgaben", "Eigenbelege", "Bankbewegungen"].forEach(name => {
            const sheet = ss.getSheetByName(name);
            if (sheet) {
                name === "Bankbewegungen" ? refreshBankSheet(sheet) : refreshDataSheet(sheet);
            }
        });
    };

    return {refreshActiveSheet, refreshAllSheets};
})();

export default RefreshModule;