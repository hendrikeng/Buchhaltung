// =================== Zentrale Konfiguration ===================
const CategoryConfig = {
    einnahmen: {
        // Kategorien im Einnahmen‑Sheet – hier definieren wir den steuerlichen Typ.
        // Die Eigenschaft taxType gibt an, ob der Posten steuerpflichtig oder steuerfrei (Inland) ist.
        categories: {
            "Umsatzerlöse": { taxType: "steuerpflichtig" },
            "Provisionserlöse": { taxType: "steuerpflichtig" },
            "Sonstige betriebliche Erträge": { taxType: "steuerpflichtig" },
            "Privateinlage": { taxType: "steuerfrei_inland" },
            "Darlehen": { taxType: "steuerfrei_inland" },
            "Zinsen": { taxType: "steuerfrei_inland" }
        },
        kontoMapping: {
            "Umsatzerlöse": { soll: "1200 Bank", gegen: "4400 Umsatzerlöse" },
            "Provisionserlöse": { soll: "1200 Bank", gegen: "4420 Provisionserlöse" },
            "Sonstige betriebliche Erträge": { soll: "1200 Bank", gegen: "4490 Sonstige betriebliche Erträge" },
            "Privateinlage": { soll: "1200 Bank", gegen: "1800 Privateinlage" },
            "Darlehen": { soll: "1200 Bank", gegen: "3000 Darlehen" },
            "Zinsen": { soll: "1200 Bank", gegen: "2650 Zinserträge" }
        },
        bwaMapping: {
            "Umsatzerlöse": "umsatzerloese",
            "Provisionserlöse": "provisionserloese",
            "Sonstige betriebliche Erträge": "sonstigeErtraege",
            "Privateinlage": "privateinlage",
            "Darlehen": "darlehen",
            "Zinsen": "zinsen"
        }
    },
    ausgaben: {
        // Kategorien im Ausgaben‑Sheet – hier nur betriebliche Ausgaben, die über das Firmenkonto gezahlt werden.
        categories: {
            "Wareneinsatz": { taxType: "steuerpflichtig" },
            "Betriebskosten": { taxType: "steuerpflichtig" },
            "Marketing & Werbung": { taxType: "steuerpflichtig" },
            "Reisekosten": { taxType: "steuerpflichtig" },
            "Personalkosten": { taxType: "steuerpflichtig" },
            "Sonstige betriebliche Aufwendungen": { taxType: "steuerpflichtig" },
            // Beispiele für steuerfreie Ausgaben (im Inland und Ausland)
            "Miete": { taxType: "steuerfrei_inland" },
            "Versicherungen": { taxType: "steuerfrei_inland" },
            "Porto": { taxType: "steuerfrei_inland" },
            "Google Ads": { taxType: "steuerfrei_ausland" },
            "AWS": { taxType: "steuerfrei_ausland" },
            "Facebook Ads": { taxType: "steuerfrei_ausland" },
            // Bewirtung, wenn über Firmenkonto gezahlt (im Ausgaben‑Sheet)
            "Bewirtung": { taxType: "steuerpflichtig" }
        },
        kontoMapping: {
            "Wareneinsatz": { soll: "4900 Wareneinsatz", gegen: "1200 Bank" },
            "Betriebskosten": { soll: "4900 Betriebskosten", gegen: "1200 Bank" },
            "Marketing & Werbung": { soll: "4900 Marketing & Werbung", gegen: "1200 Bank" },
            "Reisekosten": { soll: "4900 Reisekosten", gegen: "1200 Bank" },
            "Personalkosten": { soll: "4900 Personalkosten", gegen: "1200 Bank" },
            "Sonstige betriebliche Aufwendungen": { soll: "4900 Sonstige betriebliche Aufwendungen", gegen: "1200 Bank" },
            "Miete": { soll: "4900 Betriebskosten", gegen: "1200 Bank" },
            "Versicherungen": { soll: "4900 Betriebskosten", gegen: "1200 Bank" },
            "Porto": { soll: "4900 Betriebskosten", gegen: "1200 Bank" },
            "Google Ads": { soll: "4900 Marketing & Werbung", gegen: "1200 Bank" },
            "AWS": { soll: "4900 Marketing & Werbung", gegen: "1200 Bank" },
            "Facebook Ads": { soll: "4900 Marketing & Werbung", gegen: "1200 Bank" },
            "Bewirtung": { soll: "4900 Bewirtung", gegen: "1200 Bank" }  // Beispielwerte
        },
        bwaMapping: {
            "Wareneinsatz": "wareneinsatz",
            "Betriebskosten": "betriebskosten",
            "Marketing & Werbung": "marketing",
            "Reisekosten": "reisen",
            "Personalkosten": "personalkosten",
            "Sonstige betriebliche Aufwendungen": "sonstigeAufwendungen",
            "Miete": "betriebskosten",
            "Versicherungen": "betriebskosten",
            "Porto": "betriebskosten",
            "Google Ads": "marketing",
            "AWS": "marketing",
            "Facebook Ads": "marketing",
            "Bewirtung": "bewirtung"
        }
    },
    bank: {
        // Bank-Bewegungen – hier kommen alle Transaktionen, die über das Firmenkonto laufen.
        category: [
            "Umsatzerlöse", "Provisionserlöse", "Sonstige betriebliche Erträge", "Wareneinsatz", "Betriebskosten",
            "Marketing & Werbung", "Reisekosten", "Personalkosten", "Sonstige betriebliche Aufwendungen",
            "Privateinlage", "Darlehen", "Eigenbeleg", "Privatentnahme", "Zinsen",
            "Gewinnübertrag", "Kapitalrückführung"
        ],
        type: ["Einnahme", "Ausgabe"],
        bwaMapping: {
            "Gewinnübertrag": "gewinnuebertrag",
            "Kapitalrückführung": "kapitalrueckfuehrung"
        }
    },
    gesellschafterkonto: {
        category: ["Privateinlage", "Privatentnahme", "Darlehen"],
        shareholder: ["Christopher Giebel", "Hendrik Werner"]
    },
    eigenbelege: {
        // Eigenbelege‑Sheet: Hier werden ausschließlich private Belege erfasst.
        // Kategorien ohne den Begriff "Eigenbeleg", da das Sheet schon eigenständig ist.
        category: ["Kleidung", "Trinkgeld", "Private Vorauslage", "Bürokosten", "Reisekosten", "Bewirtung", "Sonstiges"],
        // Mapping für Eigenbelege: Hier legen wir fest, ob der Posten steuerpflichtig ist.
        mapping: {
            "Kleidung": { taxType: "steuerpflichtig" },
            "Trinkgeld": { taxType: "steuerfrei" },
            "Private Vorauslage": { taxType: "steuerfrei" },
            "Bürokosten": { taxType: "steuerpflichtig" },
            "Reisekosten": { taxType: "steuerpflichtig" },
            // Bewirtung wird hier mit Sonderregel behandelt (70% abzugsfähig, 30% nicht)
            "Bewirtung": { taxType: "eigenbeleg", besonderheit: "bewirtung" },
            "Sonstiges": { taxType: "steuerpflichtig" }
        },
        status: ["Offen", "Erstattet", "Gebucht"]
    },
    holdingTransfers: {
        category: ["Gewinnübertrag", "Kapitalrückführung"]
    },
    common: {
        paymentType: ["Überweisung", "Bar", "Kreditkarte", "Paypal"]
    }
};

// =================== Modul: Helpers ===================
const Helpers = (() => {
    // Konvertiert einen Wert in ein Date-Objekt
    const parseDate = value => {
        if (value instanceof Date) return isNaN(value.getTime()) ? null : value;
        if (typeof value === "string") {
            const d = new Date(value);
            return isNaN(d.getTime()) ? null : d;
        }
        return null;
    };

    // Konvertiert einen Wert in einen Float und entfernt unerwünschte Zeichen
    const parseCurrency = value => {
        if (typeof value === "number") return value;
        const str = value.toString().replace(/[^\d,.-]/g, "").replace(",", ".");
        const num = parseFloat(str);
        return isNaN(num) ? 0 : num;
    };

    // Extrahiert den MwSt-Wert als Zahl aus einem String (z. B. "19%")
    const parseMwstRate = value => {
        if (typeof value === "number") return value < 1 ? value * 100 : value;
        const rate = parseFloat(value.toString().replace("%", "").replace(",", ".").trim());
        return isNaN(rate) ? 19 : rate;
    };

    // Sucht in einem Ordner nach einem Unterordner mit dem angegebenen Namen
    const getFolderByName = (parent, name) => {
        const folderIter = parent.getFoldersByName(name);
        return folderIter.hasNext() ? folderIter.next() : null;
    };

    // Extrahiert ein Datum aus einem Dateinamen im Format "RE-YYYY-MM-DD" als "TT.MM.JJJJ"
    const extractDateFromFilename = filename => {
        const nameWithoutExtension = filename.replace(/\.[^/.]+$/, "");
        const match = nameWithoutExtension.match(/RE-(\d{4}-\d{2}-\d{2})/);
        if (match?.[1]) {
            const [year, month, day] = match[1].split("-");
            return `${day}.${month}.${year}`;
        }
        return "";
    };

    // Setzt bedingte Formatierung für eine Spalte anhand von Bedingungen
    const setConditionalFormattingForColumn = (sheet, column, conditions) => {
        const lastRow = sheet.getLastRow();
        const range = sheet.getRange(`${column}2:${column}${lastRow}`);
        const rules = conditions.map(({ value, background, fontColor }) =>
            SpreadsheetApp.newConditionalFormatRule()
                .whenTextEqualTo(value)
                .setBackground(background)
                .setFontColor(fontColor)
                .setRanges([range])
                .build()
        );
        sheet.setConditionalFormatRules(rules);
    };

    return { parseDate, parseCurrency, parseMwstRate, getFolderByName, extractDateFromFilename, setConditionalFormattingForColumn };
})();

// =================== Modul: Validator ===================
const Validator = (() => {
    const isEmpty = v => v == null || v.toString().trim() === "";
    const isInvalidNumber = v => isEmpty(v) || isNaN(parseFloat(v.toString().trim()));

    // Setzt Dropdown-Validierung für einen Bereich
    const validateDropdown = (sheet, row, col, numRows, numCols, list) =>
        sheet.getRange(row, col, numRows, numCols).setDataValidation(
            SpreadsheetApp.newDataValidation().requireValueInList(list, true).build()
        );

    // Regelbasierte Validierung für Revenue-/Ausgaben-Zeilen
    const validateRevenueAndExpenses = (row, rowIndex) => {
        const warnings = [];
        const validateRow = (row, idx, rules) => {
            rules.forEach(({ check, message }) => {
                if (check(row)) warnings.push(`Zeile ${idx}: ${message}`);
            });
        };
        const baseRules = [
            { check: r => isEmpty(r[0]), message: "Rechnungsdatum fehlt." },
            { check: r => isEmpty(r[1]), message: "Rechnungsnummer fehlt." },
            { check: r => isEmpty(r[2]), message: "Kategorie fehlt." },
            { check: r => isEmpty(r[3]), message: "Kunde fehlt." },
            { check: r => isInvalidNumber(r[4]), message: "Nettobetrag fehlt oder ungültig." },
            {
                check: r => {
                    const mwstStr = r[5] == null ? "" : r[5].toString().trim();
                    return isEmpty(mwstStr) || isNaN(parseFloat(mwstStr.replace("%", "").replace(",", ".")));
                },
                message: "Mehrwertsteuer fehlt oder ungültig."
            }
        ];
        const status = row[11] ? row[11].toString().trim().toLowerCase() : "";
        const paymentRules = status === "offen"
            ? [
                { check: r => !isEmpty(r[12]), message: 'Zahlungsart darf bei offener Zahlung nicht gesetzt sein.' },
                { check: r => !isEmpty(r[13]), message: 'Zahlungsdatum darf bei offener Zahlung nicht gesetzt sein.' }
            ]
            : [
                { check: r => isEmpty(r[12]), message: 'Zahlungsart muss bei bezahlter/teilbezahlter Zahlung gesetzt sein.' },
                { check: r => isEmpty(r[13]), message: 'Zahlungsdatum muss bei bezahlter/teilbezahlter Zahlung gesetzt sein.' },
                {
                    check: r => {
                        if (isEmpty(r[13]) || isEmpty(r[0])) return false;
                        const paymentDate = Helpers.parseDate(r[13]);
                        return paymentDate ? paymentDate > new Date() : false;
                    },
                    message: "Zahlungsdatum darf nicht in der Zukunft liegen."
                },
                {
                    check: r => {
                        if (isEmpty(r[13]) || isEmpty(r[0])) return false;
                        const paymentDate = Helpers.parseDate(r[13]);
                        const invoiceDate = Helpers.parseDate(r[0]);
                        return paymentDate && invoiceDate ? paymentDate < invoiceDate : false;
                    },
                    message: "Zahlungsdatum darf nicht vor dem Rechnungsdatum liegen."
                }
            ];
        const rules = baseRules.concat(paymentRules);
        validateRow(row, rowIndex, rules);
        return warnings;
    };

    // Regelbasierte Validierung für das Bank-Sheet (ohne Mapping)
    const validateBanking = bankSheet => {
        const data = bankSheet.getDataRange().getValues();
        const warnings = [];
        const validateRow = (row, idx, rules) => {
            rules.forEach(({ check, message }) => {
                if (check(row)) warnings.push(`Zeile ${idx}: ${message}`);
            });
        };
        const headerFooterRules = [
            { check: r => isEmpty(r[0]), message: "Buchungsdatum fehlt." },
            { check: r => isEmpty(r[1]), message: "Buchungstext fehlt." },
            { check: r => !isEmpty(r[2]) || !isNaN(parseFloat(r[2].toString().trim())), message: "Betrag darf nicht gesetzt sein." },
            { check: r => isEmpty(r[3]) || isInvalidNumber(r[3]), message: "Saldo fehlt oder ungültig." },
            { check: r => !isEmpty(r[4]), message: "Typ darf nicht gesetzt sein." },
            { check: r => !isEmpty(r[5]), message: "Kategorie darf nicht gesetzt sein." },
            { check: r => !isEmpty(r[6]), message: "Konto (Soll) darf nicht gesetzt sein." },
            { check: r => !isEmpty(r[7]), message: "Gegenkonto (Haben) darf nicht gesetzt sein." }
        ];
        const dataRowRules = [
            { check: r => isEmpty(r[0]), message: "Buchungsdatum fehlt." },
            { check: r => isEmpty(r[1]), message: "Buchungstext fehlt." },
            { check: r => isEmpty(r[2]) || isInvalidNumber(r[2]), message: "Betrag fehlt oder ungültig." },
            { check: r => isEmpty(r[3]) || isInvalidNumber(r[3]), message: "Saldo fehlt oder ungültig." },
            { check: r => isEmpty(r[4]), message: "Typ fehlt." },
            { check: r => isEmpty(r[5]), message: "Kategorie fehlt." },
            { check: r => isEmpty(r[6]), message: "Konto (Soll) fehlt." },
            { check: r => isEmpty(r[7]), message: "Gegenkonto (Haben) fehlt." }
        ];
        data.forEach((row, i) => {
            const idx = i + 1;
            if (i === 1 || i === data.length - 1) {
                validateRow(row, idx, headerFooterRules);
            } else if (i > 1 && i < data.length - 1) {
                validateRow(row, idx, dataRowRules);
            }
        });
        return warnings;
    };

    const validateAllSheets = (revenueSheet, expenseSheet, bankSheet = null) => {
        const revData = revenueSheet.getDataRange().getValues().slice(1);
        const expData = expenseSheet.getDataRange().getValues().slice(1);
        const revenueWarnings = revData.reduce((acc, row, i) => acc.concat(validateRevenueAndExpenses(row, i + 2)), []);
        const expenseWarnings = expData.reduce((acc, row, i) => acc.concat(validateRevenueAndExpenses(row, i + 2)), []);
        const bankWarnings = bankSheet ? validateBanking(bankSheet) : [];
        const msgArr = [];
        revenueWarnings.length && msgArr.push("Fehler in 'Einnahmen':\n" + revenueWarnings.join("\n"));
        expenseWarnings.length && msgArr.push("Fehler in 'Ausgaben':\n" + expenseWarnings.join("\n"));
        bankWarnings.length && msgArr.push("Fehler in 'Bankbewegungen':\n" + bankWarnings.join("\n"));
        if (msgArr.length) {
            SpreadsheetApp.getUi().alert(msgArr.join("\n\n"));
            return false;
        }
        return true;
    };

    return { validateDropdown, validateRevenueAndExpenses, validateBanking, validateAllSheets };
})();

// =================== Modul: ImportModule ===================
const ImportModule = (() => {
    const importFilesFromFolder = (folder, importSheet, mainSheet, type, historySheet) => {
        const files = folder.getFiles();
        const getExistingFiles = (sheet, colIndex) =>
            new Set(sheet.getDataRange().getValues().slice(1).map(row => row[colIndex]));
        const existingMain = getExistingFiles(mainSheet, 16);
        const existingImport = getExistingFiles(importSheet, 0);
        const newMainRows = [];
        const newImportRows = [];
        const newHistoryRows = [];
        const timestamp = new Date();

        while (files.hasNext()) {
            const file = files.next();
            const baseName = file.getName().replace(/\.[^/.]+$/, "");
            const invoiceName = baseName.replace(/^[^ ]* /, "");
            const fileName = baseName;
            const invoiceDate = Helpers.extractDateFromFilename(fileName);
            const fileUrl = file.getUrl();
            let wasImported = false;

            if (!existingMain.has(fileName)) {
                newMainRows.push([
                    invoiceDate,
                    invoiceName,
                    "", "", "", "",
                    "",
                    "",
                    "",
                    "",
                    "",
                    "",
                    "", "", "",
                    timestamp,
                    fileName,
                    fileUrl
                ]);
                existingMain.add(fileName);
                wasImported = true;
            }
            if (!existingImport.has(fileName)) {
                newImportRows.push([fileName, fileUrl, fileName]);
                existingImport.add(fileName);
                wasImported = true;
            }
            wasImported && newHistoryRows.push([timestamp, type, fileName, fileUrl]);
        }
        if (newImportRows.length)
            importSheet.getRange(importSheet.getLastRow() + 1, 1, newImportRows.length, newImportRows[0].length)
                .setValues(newImportRows);
        if (newMainRows.length)
            mainSheet.getRange(mainSheet.getLastRow() + 1, 1, newMainRows.length, newMainRows[0].length)
                .setValues(newMainRows);
        if (newHistoryRows.length)
            historySheet.getRange(historySheet.getLastRow() + 1, 1, newHistoryRows.length, newHistoryRows[0].length)
                .setValues(newHistoryRows);
    };

    const importDriveFiles = () => {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const revenueMain = ss.getSheetByName("Einnahmen");
        const expenseMain = ss.getSheetByName("Ausgaben");
        const revenue = ss.getSheetByName("Rechnungen Einnahmen") || ss.insertSheet("Rechnungen Einnahmen");
        const expense = ss.getSheetByName("Rechnungen Ausgaben") || ss.insertSheet("Rechnungen Ausgaben");
        const history = ss.getSheetByName("Änderungshistorie") || ss.insertSheet("Änderungshistorie");

        if (revenue.getLastRow() === 0)
            revenue.appendRow(["Dateiname", "Link zur Datei", "Rechnungsnummer"]);
        if (expense.getLastRow() === 0)
            expense.appendRow(["Dateiname", "Link zur Datei", "Rechnungsnummer"]);
        if (history.getLastRow() === 0)
            history.appendRow(["Datum", "Rechnungstyp", "Dateiname", "Link zur Datei"]);

        const file = DriveApp.getFileById(ss.getId());
        const parents = file.getParents();
        const parentFolder = parents.hasNext() ? parents.next() : null;
        if (!parentFolder) {
            SpreadsheetApp.getUi().alert("Kein übergeordneter Ordner gefunden.");
            return;
        }

        const revenueFolder = Helpers.getFolderByName(parentFolder, "Einnahmen");
        const expenseFolder = Helpers.getFolderByName(parentFolder, "Ausgaben");

        revenueFolder
            ? importFilesFromFolder(revenueFolder, revenue, revenueMain, "Einnahme", history)
            : SpreadsheetApp.getUi().alert("Fehler: 'Einnahmen'-Ordner nicht gefunden.");
        expenseFolder
            ? importFilesFromFolder(expenseFolder, expense, expenseMain, "Ausgabe", history)
            : SpreadsheetApp.getUi().alert("Fehler: 'Ausgaben'-Ordner nicht gefunden.");
    };

    return { importDriveFiles };
})();

// =================== Modul: RefreshModule ===================
const RefreshModule = (() => {
    // Aktualisiert Datenblätter: Einnahmen, Ausgaben, Eigenbelege
    const refreshDataSheet = sheet => {
        const lastRow = sheet.getLastRow();
        if (lastRow < 2) return;
        const numRows = lastRow - 1;

        // Batch: Formeln in Spalten 7,8,10,11,12 setzen
        Object.entries({
            7: row => `=E${row}*F${row}`,                           // MwSt
            8: row => `=E${row}+G${row}`,                           // Brutto
            10: row => `=(H${row}-I${row})/(1+VALUE(F${row}))`,      // Restbetrag
            11: row => `=IF(A${row}=""; ""; ROUNDUP(MONTH(A${row})/3;0))`, // Quartal
            12: row => `=IF(OR(I${row}=""; I${row}=0); "Offen"; IF(I${row}>=H${row}; "Bezahlt"; "Teilbezahlt"))` // Zahlungsstatus
        }).forEach(([col, formulaFn]) => {
            const formulas = Array.from({ length: numRows }, (_, i) => [formulaFn(i + 2)]);
            sheet.getRange(2, Number(col), numRows, 1).setFormulas(formulas);
        });

        // Batch: Falls leer, 0 in Spalte 9 setzen
        const col9Range = sheet.getRange(2, 9, numRows, 1);
        const col9Values = col9Range.getValues().map(([val]) => (val === "" || val === null ? 0 : val));
        col9Range.setValues(col9Values.map(val => [val]));

        // Dropdown-Validierung für Kategorien und Zahlungsarten
        const name = sheet.getName();
        if (name === "Einnahmen")
            Validator.validateDropdown(
                sheet,
                2,
                3,
                lastRow - 1,
                1,
                CategoryConfig.einnahmen.categories ? Object.keys(CategoryConfig.einnahmen.categories) : CategoryConfig.einnahmen.category
            );
        if (name === "Ausgaben")
            Validator.validateDropdown(
                sheet,
                2,
                3,
                lastRow - 1,
                1,
                CategoryConfig.ausgaben.categories ? Object.keys(CategoryConfig.ausgaben.categories) : Object.keys(CategoryConfig.ausgaben.categories)
            );
        if (name === "Eigenbelege")
            Validator.validateDropdown(sheet, 2, 3, lastRow - 1, 1, CategoryConfig.eigenbelege.category);
        Validator.validateDropdown(sheet, 2, 13, lastRow - 1, 1, CategoryConfig.common.paymentType);
        sheet.autoResizeColumns(1, sheet.getLastColumn());
    };

    // Aktualisiert das Bank-Sheet (Formeln, Formatierung, Mapping, Endsaldo)
    const refreshBankSheet = sheet => {
        const lastRow = sheet.getLastRow();
        if (lastRow < 3) return;
        const firstDataRow = 3;
        const numDataRows = lastRow - firstDataRow + 1;
        const transRows = lastRow - firstDataRow - 1;

        // Batch: Formeln in Spalte D setzen
        if (transRows > 0) {
            sheet.getRange(firstDataRow, 4, transRows, 1).setFormulas(
                Array.from({ length: transRows }, (_, i) => [
                    `=D${firstDataRow + i - 1}+C${firstDataRow + i}`
                ])
            );
        }

        // Batch: Transaktionstyp in Spalte E anhand von Spalte C (Betrag) setzen
        const amounts = sheet.getRange(firstDataRow, 3, numDataRows, 1).getValues();
        const typeValues = amounts.map(([val]) => {
            const amt = parseFloat(val) || 0;
            return [amt > 0 ? "Einnahme" : amt < 0 ? "Ausgabe" : ""];
        });
        sheet.getRange(firstDataRow, 5, numDataRows, 1).setValues(typeValues);

        // Dropdown-Validierung in Bank-Sheet
        Validator.validateDropdown(sheet, firstDataRow, 5, numDataRows, 1, CategoryConfig.bank.type);
        Validator.validateDropdown(sheet, firstDataRow, 6, numDataRows, 1, CategoryConfig.bank.category);
        const allowedKontoSoll = Object.values(CategoryConfig.einnahmen.kontoMapping)
            .concat(Object.values(CategoryConfig.ausgaben.kontoMapping))
            .map(m => m.soll);
        const allowedGegenkonto = Object.values(CategoryConfig.einnahmen.kontoMapping)
            .concat(Object.values(CategoryConfig.ausgaben.kontoMapping))
            .map(m => m.gegen);
        Validator.validateDropdown(sheet, firstDataRow, 7, numDataRows, 1, allowedKontoSoll);
        Validator.validateDropdown(sheet, firstDataRow, 8, numDataRows, 1, allowedGegenkonto);

        Helpers.setConditionalFormattingForColumn(sheet, "E", [
            { value: "Einnahme", background: "#C6EFCE", fontColor: "#006100" },
            { value: "Ausgabe", background: "#FFC7CE", fontColor: "#9C0006" }
        ]);

        sheet.getRange(`A2:A${lastRow}`).setNumberFormat("DD.MM.YYYY");
        ["C", "D"].forEach(col =>
            sheet.getRange(`${col}2:${col}${lastRow}`).setNumberFormat("€#,##0.00;€-#,##0.00")
        );

        // Konto-Mapping: Mapping in Spalten G und H anhand von Typ (Spalte E) und Kategorie (Spalte F)
        const dataRange = sheet.getRange(firstDataRow, 1, numDataRows, sheet.getLastColumn());
        const data = dataRange.getValues();
        data.forEach((row, i) => {
            const globalRow = i + firstDataRow;
            const label = row[1] ? row[1].toString().trim().toLowerCase() : "";
            if (globalRow === lastRow && label === "endsaldo") return;
            const type = row[4];
            const category = row[5];
            let mapping = type === "Einnahme"
                ? CategoryConfig.einnahmen.kontoMapping[category]
                : type === "Ausgabe"
                    ? CategoryConfig.ausgaben.kontoMapping[category]
                    : null;
            if (!mapping) mapping = { soll: "Manuell prüfen", gegen: "Manuell prüfen" };
            row[6] = mapping.soll;
            row[7] = mapping.gegen;
        });
        dataRange.setValues(data);

        // Endsaldo-Zeile aktualisieren oder hinzufügen
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

    return { refreshActiveSheet, refreshAllSheets };
})();

// =================== Modul: UStVA-Berechnung ===================
const UStVACalculator = (() => {
    // Neues UStVA-Datenobjekt mit allen Feldern (inkl. neuer Spalten)
    const createEmptyUStVA = () => ({
        // Einnahmen
        steuerpflichtige_einnahmen: 0,
        steuerfreie_inland_einnahmen: 0,
        steuerfreie_ausland_einnahmen: 0,
        // Ausgaben
        steuerpflichtige_ausgaben: 0,
        steuerfreie_inland_ausgaben: 0,
        steuerfreie_ausland_ausgaben: 0,
        // Eigenbelege
        eigenbelege_steuerpflichtig: 0,
        eigenbelege_steuerfrei: 0,
        // Nicht abzugsfähige VSt (z. B. bei Bewirtung)
        nicht_abzugsfaehige_vst: 0,
        // Steuern
        ust_7: 0,
        ust_19: 0,
        vst_7: 0,
        vst_19: 0
    });

    /**
     * Verarbeitet eine Zeile aus Einnahmen, Ausgaben oder Eigenbelegen.
     * Erwartete Spalten (Index):
     *  - 4: Nettobetrag
     *  - 5: MwSt (%) als String (z. B. "19%")
     *  - 9: Restbetrag
     *  - 13: Zahlungsdatum
     *  - 2: Kategorie
     *
     * @param {Array} row - Datenzeile
     * @param {Object} ustvaData - Aggregiertes UStVA-Datenobjekt (monatlich)
     * @param {Boolean} isIncome - true für Einnahmen, false für Ausgaben
     * @param {Boolean} [isEigenbelegSheet=false] - true, wenn die Zeile aus dem Sheet "Eigenbelege" stammt
     */
    const processUStVARow = (row, ustvaData, isIncome, isEigenbelegSheet = false) => {
        const paymentDate = Helpers.parseDate(row[13]); // Zahlungsdatum (Spalte N)
        if (!paymentDate || paymentDate > new Date()) return;
        const month = paymentDate.getMonth() + 1;
        const netto = Helpers.parseCurrency(row[4]);
        const restNetto = Helpers.parseCurrency(row[9]) || 0;
        const gezahlt = Math.max(0, netto - restNetto);
        const mwstRate = Helpers.parseMwstRate(row[5]);
        const roundedRate = Math.round(mwstRate);
        const tax = gezahlt * (mwstRate / 100);
        const category = row[2] ? row[2].toString().trim() : "";

        if (isIncome) {
            // Einnahmen: Nutze das Mapping aus CategoryConfig.einnahmen.categories
            const taxInfo = CategoryConfig.einnahmen.categories[category] || { taxType: "steuerpflichtig" };
            if (taxInfo.taxType === "steuerfrei_inland") {
                ustvaData[month].steuerfreie_inland_einnahmen += gezahlt;
            } else if (taxInfo.taxType === "steuerfrei_ausland" || roundedRate === 0) {
                ustvaData[month].steuerfreie_ausland_einnahmen += gezahlt;
            } else {
                ustvaData[month].steuerpflichtige_einnahmen += gezahlt;
                ustvaData[month][`ust_${roundedRate}`] += tax;
            }
        } else {
            // Für Ausgaben und Eigenbelege
            if (isEigenbelegSheet) {
                // Verarbeitung im Eigenbelege‑Sheet: Nutze CategoryConfig.eigenbelege.mapping
                const eigenMapping = CategoryConfig.eigenbelege.mapping[category] || { taxType: "steuerpflichtig" };
                if (eigenMapping.taxType === "steuerfrei") {
                    ustvaData[month].eigenbelege_steuerfrei += gezahlt;
                } else if (eigenMapping.taxType === "eigenbeleg" && eigenMapping.besonderheit === "bewirtung") {
                    // Sonderfall Bewirtung: 70 % abzugsfähige VSt, 30 % nicht
                    ustvaData[month].eigenbelege_steuerpflichtig += gezahlt;
                    ustvaData[month][`vst_${roundedRate}`] += tax * 0.7;
                    ustvaData[month].nicht_abzugsfaehige_vst += tax * 0.3;
                } else {
                    ustvaData[month].eigenbelege_steuerpflichtig += gezahlt;
                    ustvaData[month][`vst_${roundedRate}`] += tax;
                }
            } else {
                // Verarbeitung im regulären Ausgaben‑Sheet: Nutze CategoryConfig.ausgaben.categories
                const taxInfo = CategoryConfig.ausgaben.categories[category] || { taxType: "steuerpflichtig" };
                if (taxInfo.taxType === "steuerfrei_inland") {
                    ustvaData[month].steuerfreie_inland_ausgaben += gezahlt;
                } else if (taxInfo.taxType === "steuerfrei_ausland") {
                    ustvaData[month].steuerfreie_ausland_ausgaben += gezahlt;
                } else {
                    ustvaData[month].steuerpflichtige_ausgaben += gezahlt;
                    ustvaData[month][`vst_${roundedRate}`] += tax;
                }
            }
        }
    };

    /**
     * Formatiert eine Zeile für die UStVA-Ausgabe.
     * USt-Zahlung = (USt 7% + USt 19%) - ((VSt 7% + VSt 19%) - Nicht abzugsfähige VSt)
     *
     * @param {String} label - Bezeichner (Monat, Quartal, Gesamtjahr)
     * @param {Object} data - Aggregierte UStVA-Daten
     * @returns {Array} - Array mit formatierten Werten
     */
    const formatUStVARow = (label, data) => {
        const ustZahlung = (data.ust_7 + data.ust_19) - ((data.vst_7 + data.vst_19) - data.nicht_abzugsfaehige_vst);
        const ergebnis =
            (data.steuerpflichtige_einnahmen + data.steuerfreie_inland_einnahmen + data.steuerfreie_ausland_einnahmen) -
            (data.steuerpflichtige_ausgaben + data.steuerfreie_inland_ausgaben + data.steuerfreie_ausland_ausgaben +
                data.eigenbelege_steuerpflichtig + data.eigenbelege_steuerfrei);
        return [
            label,
            data.steuerpflichtige_einnahmen,
            data.steuerfreie_inland_einnahmen,
            data.steuerfreie_ausland_einnahmen,
            data.steuerpflichtige_ausgaben,
            data.steuerfreie_inland_ausgaben,
            data.steuerfreie_ausland_ausgaben,
            data.eigenbelege_steuerpflichtig,
            data.eigenbelege_steuerfrei,
            data.nicht_abzugsfaehige_vst,
            data.ust_7,
            data.ust_19,
            data.vst_7,
            data.vst_19,
            ustZahlung,
            ergebnis
        ];
    };

    // Aggregiert UStVA-Daten über einen Zeitraum (z. B. Quartal oder Gesamtjahr)
    const aggregateUStVA = (ustvaData, start, end) => {
        const sum = createEmptyUStVA();
        for (let m = start; m <= end; m++) {
            for (const key in sum) {
                sum[key] += ustvaData[m][key] || 0;
            }
        }
        return sum;
    };

    const calculateUStVA = () => {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const revenueSheet = ss.getSheetByName("Einnahmen");
        const expenseSheet = ss.getSheetByName("Ausgaben");
        const eigenbelegeSheet = ss.getSheetByName("Eigenbelege");
        if (!revenueSheet || !expenseSheet) {
            SpreadsheetApp.getUi().alert("Fehlendes Blatt: 'Einnahmen' oder 'Ausgaben'");
            return;
        }
        if (!Validator.validateAllSheets(revenueSheet, expenseSheet)) return;

        // Daten laden (ohne Header) – jeweils einmal pro Blatt
        const revenueData = revenueSheet.getDataRange().getValues().slice(1);
        const expenseData = expenseSheet.getDataRange().getValues().slice(1);
        // Alle Zeilen aus dem Eigenbelege‑Sheet werden als Eigenbelege verarbeitet
        const eigenbelegeData = eigenbelegeSheet ? eigenbelegeSheet.getDataRange().getValues().slice(1) : [];
        const ustvaData = {};
        for (let m = 1; m <= 12; m++) {
            ustvaData[m] = createEmptyUStVA();
        }

        // Verarbeitung: Einnahmen, reguläre Ausgaben, und Eigenbelege (mit isEigenbelegSheet=true)
        revenueData.forEach(row => processUStVARow(row, ustvaData, true));
        expenseData.forEach(row => processUStVARow(row, ustvaData, false));
        eigenbelegeData.forEach(row => processUStVARow(row, ustvaData, false, true));

        // Batch: Ausgabezeilen sammeln
        const outputRows = [];
        outputRows.push([
            "Zeitraum",
            "Steuerpflichtige Einnahmen",
            "Steuerfreie Inland-Einnahmen",
            "Steuerfreie Ausland-Einnahmen",
            "Steuerpflichtige Ausgaben",
            "Steuerfreie Inland-Ausgaben",
            "Steuerfreie Ausland-Ausgaben",
            "Eigenbelege steuerpflichtig",
            "Eigenbelege steuerfrei",
            "Nicht abzugsfähige VSt (Bewirtung)",
            "USt 7%",
            "USt 19%",
            "VSt 7%",
            "VSt 19%",
            "USt-Zahlung",
            "Ergebnis"
        ]);
        const monthNames = [
            "Januar", "Februar", "März", "April", "Mai", "Juni",
            "Juli", "August", "September", "Oktober", "November", "Dezember"
        ];
        for (let m = 1; m <= 12; m++) {
            outputRows.push(formatUStVARow(monthNames[m - 1], ustvaData[m]));
            if (m % 3 === 0) {
                outputRows.push(formatUStVARow(`Quartal ${m / 3}`, aggregateUStVA(ustvaData, m - 2, m)));
            }
        }
        outputRows.push(formatUStVARow("Gesamtjahr", aggregateUStVA(ustvaData, 1, 12)));

        const ustvaSheet = ss.getSheetByName("UStVA") || ss.insertSheet("UStVA");
        ustvaSheet.clearContents();
        ustvaSheet.getRange(1, 1, outputRows.length, outputRows[0].length).setValues(outputRows);
        ustvaSheet.getRange(2, 2, ustvaSheet.getLastRow() - 1, 1).setNumberFormat("#,##0.00€");
        ustvaSheet.autoResizeColumns(1, outputRows[0].length);
        ss.setActiveSheet(ustvaSheet);
        SpreadsheetApp.getUi().alert("UStVA wurde aktualisiert!");
    };

    return { calculateUStVA };
})();

// =================== Modul: BWACalculator ===================
// BWACalculator bleibt unverändert, da du diesen neu machen möchtest.
const BWACalculator = (() => {
    // ... Dein BWACalculator-Code, unverändert.
    return { };
})();

// =================== Globale Funktionen ===================
const onOpen = () => {
    SpreadsheetApp.getUi()
        .createMenu("📂 Buchhaltung")
        .addItem("📥 Dateien importieren", "importDriveFiles")
        .addItem("🔄 Refresh Active Sheet", "refreshSheet")
        .addItem("📊 UStVA berechnen", "calculateUStVA")
        // .addItem("📈 BWA berechnen", "calculateBWA")
        .addToUi();
};

const onEdit = e => {
    const { range } = e;
    const sheet = range.getSheet();
    const name = sheet.getName();
    const mapping = {
        "Einnahmen": 16,
        "Ausgaben": 16,
        "Eigenbelege": 16,
        "Bankbewegungen": 11,
        "Gesellschafterkonto": 12,
        "Holding Transfers": 6
    };
    if (!(name in mapping)) return;
    if (range.getRow() === 1) return;
    const headerLen = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].length;
    if (range.getColumn() > headerLen) return;
    if (range.getColumn() === mapping[name]) return;
    const rowValues = sheet.getRange(range.getRow(), 1, 1, headerLen).getValues()[0];
    if (rowValues.every(cell => cell === "")) return;
    const ts = new Date();
    sheet.getRange(range.getRow(), mapping[name]).setValue(ts);
};

const setupTrigger = () => {
    const triggers = ScriptApp.getProjectTriggers();
    if (!triggers.some(t => t.getHandlerFunction() === "onOpen"))
        SpreadsheetApp.newTrigger("onOpen")
            .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
            .onOpen()
            .create();
};

const refreshSheet = () => RefreshModule.refreshActiveSheet();
const calculateUStVA = () => { RefreshModule.refreshAllSheets(); UStVACalculator.calculateUStVA(); };
// const calculateBWA = () => { RefreshModule.refreshAllSheets(); BWACalculator.calculateBWA(); };
const importDriveFiles = () => { ImportModule.importDriveFiles(); RefreshModule.refreshAllSheets(); };
