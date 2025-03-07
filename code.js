// =================== Zentrale Konfiguration ===================
// Konfiguration der erlaubten Kategorien, Konto- und BWA-Mappings
const CategoryConfig = {
    einnahmen: {
        category: [
            "Umsatzerlöse",
            "Provisionserlöse",
            "Sonstige betriebliche Erträge",
            "Privateinlage",
            "Darlehen",
            "Zinsen"
        ],
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
        category: [
            "Wareneinsatz",
            "Betriebskosten",
            "Marketing & Werbung",
            "Reisekosten",
            "Personalkosten",
            "Sonstige betriebliche Aufwendungen",
            "Eigenbeleg",
            "Privatentnahme",
            "Darlehen",
            "Zinsen"
        ],
        kontoMapping: {
            "Wareneinsatz": { soll: "4900 Wareneinsatz", gegen: "1200 Bank" },
            "Betriebskosten": { soll: "4900 Betriebskosten", gegen: "1200 Bank" },
            "Marketing & Werbung": { soll: "4900 Marketing & Werbung", gegen: "1200 Bank" },
            "Reisekosten": { soll: "4900 Reisekosten", gegen: "1200 Bank" },
            "Personalkosten": { soll: "4900 Personalkosten", gegen: "1200 Bank" },
            "Sonstige betriebliche Aufwendungen": { soll: "4900 Sonstige betriebliche Aufwendungen", gegen: "1200 Bank" },
            "Eigenbeleg": { soll: "1890 Eigenbeleg", gegen: "1200 Bank" },
            "Privatentnahme": { soll: "1800 Privatentnahme", gegen: "1200 Bank" },
            "Darlehen": { soll: "3000 Darlehen", gegen: "1200 Bank" },
            "Zinsen": { soll: "2100 Zinsaufwand", gegen: "1200 Bank" }
        },
        bwaMapping: {
            "Wareneinsatz": "wareneinsatz",
            "Betriebskosten": "betriebskosten",
            "Marketing & Werbung": "marketing",
            "Reisekosten": "reisen",
            "Personalkosten": "personalkosten",
            "Sonstige betriebliche Aufwendungen": "sonstigeAufwendungen",
            "Eigenbeleg": "eigenbeleg",
            "Privatentnahme": "privatentnahme",
            "Darlehen": "darlehen",
            "Zinsen": "zinsen"
        }
    },
    bank: {
        category: [
            "Umsatzerlöse",
            "Provisionserlöse",
            "Sonstige betriebliche Erträge",
            "Wareneinsatz",
            "Betriebskosten",
            "Marketing & Werbung",
            "Reisekosten",
            "Personalkosten",
            "Sonstige betriebliche Aufwendungen",
            "Privateinlage",
            "Darlehen",
            "Eigenbeleg",
            "Privatentnahme",
            "Zinsen",
            "Gewinnübertrag",
            "Kapitalrückführung"
        ],
        type: ["Einnahme", "Ausgabe"],
        bwaMapping: {
            "Gewinnübertrag": "gewinnuebertrag",
            "Kapitalrückführung": "kapitalrueckfuehrung"
        }
    },
    gesellschafterkonto: {
        category: ["Privateinlage", "Privatentnahme", "Darlehen"],
        shareholder: ["Christopher Giebel", "Hendrik Werner"],
    },
    eigenbelege: {
        category: ["Kleidung", "Trinkgeld", "Private Vorauslage", "Bürokosten", "Reisekosten", "Bewirtung", "Sonstiges"],
        status: ["Offen", "Erstattet", "Gebucht"]
    },
    holdingTransfers: {
        category: ["Gewinnübertrag", "Kapitalrückführung"],
    },
    common: {
        paymentType: ["Überweisung", "Bar", "Kreditkarte", "Paypal"],
    },

};

// =================== Modul: Helpers ===================
// Gemeinsame Hilfsfunktionen
const Helpers = (() => {
    // Konvertiert String oder Datum in ein Date-Objekt
    const parseDate = value => {
        if (value instanceof Date) return isNaN(value.getTime()) ? null : value;
        if (typeof value === "string") {
            const d = new Date(value);
            return isNaN(d.getTime()) ? null : d;
        }
        return null;
    };

    // Wandelt einen String (oder Zahl) in einen Float um und entfernt dabei unerwünschte Zeichen
    const parseCurrency = value => {
        if (typeof value === "number") return value;
        const str = value.toString().replace(/[^\d,.-]/g, "").replace(",", ".");
        const num = parseFloat(str);
        return isNaN(num) ? 0 : num;
    };

    // Extrahiert einen Mehrwertsteuersatz aus String oder Zahl
    const parseMwstRate = value => {
        if (typeof value === "number") return value < 1 ? value * 100 : value;
        const rate = parseFloat(value.toString().replace("%", "").replace(",", "."));
        return isNaN(rate) ? 19 : rate;
    };

    // Sucht in einem Ordner nach einem Unterordner mit dem angegebenen Namen
    const getFolderByName = (parent, name) => {
        const folderIter = parent.getFoldersByName(name);
        return folderIter.hasNext() ? folderIter.next() : null;
    };

    // Extrahiert das Datum aus einem Dateinamen im Format "RE-YYYY-MM-DD" und gibt es als "TT.MM.JJJJ" zurück
    const extractDateFromFilename = filename => {
        const nameWithoutExtension = filename.replace(/\.[^/.]+$/, "");
        const match = nameWithoutExtension.match(/RE-(\d{4}-\d{2}-\d{2})/);
        if (match?.[1]) {
            const [year, month, day] = match[1].split("-");
            return `${day}.${month}.${year}`;
        }
        return "";
    };

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
// Verantwortlich für die Validierung von Daten und Ausgabe von Warnmeldungen
const Validator = (() => {
    const isEmpty = v => v == null || v.toString().trim() === "";
    const isInvalidNumber = v => isEmpty(v) || isNaN(parseFloat(v.toString().trim()));

    // Setzt die Dropdown-Validierung auf einen bestimmten Bereich
    const validateDropdown = (sheet, row, col, numRows, numCols, list) =>
        sheet.getRange(row, col, numRows, numCols).setDataValidation(
            SpreadsheetApp.newDataValidation().requireValueInList(list, true).build()
        );

    // Validiert eine Zeile aus Revenue oder Ausgaben (unverändert)
    const validateRevenueAndExpenses = (row, rowIndex) => {
        const warnings = [];

        // Hilfsfunktion, die eine Reihe von Regeln auf eine Zeile anwendet
        const validateRow = (row, idx, rules) => {
            rules.forEach(({ check, message }) => {
                if (check(row)) warnings.push(`Zeile ${idx}: ${message}`);
            });
        };

        // Basisregeln für alle Zeilen (angenommene Spaltenreihenfolge)
        const baseRules = [
            { check: row => isEmpty(row[0]), message: "Rechnungsdatum fehlt." },
            { check: row => isEmpty(row[1]), message: "Rechnungsnummer fehlt." },
            { check: row => isEmpty(row[2]), message: "Kategorie fehlt." },
            { check: row => isEmpty(row[3]), message: "Kunde fehlt." },
            { check: row => isInvalidNumber(row[4]), message: "Nettobetrag fehlt oder ungültig." },
            {
                check: r => {
                    const mwstStr = row[5] == null ? "" : row[5].toString().trim();
                    return isEmpty(mwstStr) || isNaN(parseFloat(mwstStr.replace("%", "").replace(",", ".")));
                },
                message: "Mehrwertsteuer fehlt oder ungültig."
            }
        ];

        // Statusabhängige Regeln für Zahlungsangaben:
        // Wenn der Status "offen" ist, dürfen Zahlungsart und Zahlungsdatum nicht gesetzt sein.
        // Andernfalls müssen beide Felder gesetzt sein.
        const status = row[11] ? row[11].toString().trim().toLowerCase() : "";
        const paymentRules = status === "offen"
            ? [
                { check: row => !isEmpty(row[12]), message: 'Zahlungsart darf bei offener Zahlung nicht gesetzt sein.' },
                { check: row => !isEmpty(row[13]), message: 'Zahlungsdatum darf bei offener Zahlung nicht gesetzt sein.' }
            ]
            : [
                { check: row => isEmpty(row[12]), message: 'Zahlungsart muss bei bezahlter oder teilbezahlter Zahlung gesetzt sein.' },
                { check: row => isEmpty(row[13]), message: 'Zahlungsdatum muss bei bezahlter oder teilbezahlter Zahlung gesetzt sein.' }
            ];

        // Alle Regeln zusammenführen
        const rules = baseRules.concat(paymentRules);

        // Regeln abarbeiten
        validateRow(row, rowIndex, rules);
        return warnings;
    };


    // Validiert das Bank-Sheet (Mapping-Logik entfernt)
    const validateBanking = bankSheet => {
        const data = bankSheet.getDataRange().getValues();
        const warnings = [];
        const validateRow = (row, idx, rules) => {
            rules.forEach(({ check, message }) => {
                if (check(row)) warnings.push(`Zeile ${idx}: ${message}`);
            });
        };
        const headerFooterRules = [
            { check: row => isEmpty(row[0]), message: "Buchungsdatum fehlt." },
            { check: row => isEmpty(row[1]), message: "Buchungstext fehlt." },
            { check: row => !isEmpty(row[2]) || !isNaN(parseFloat(row[2].toString().trim())), message: "Betrag darf nicht gesetzt sein." },
            { check: row => isEmpty(row[3]) || isInvalidNumber(row[3]), message: "Saldo fehlt oder ungültig." },
            { check: row => !isEmpty(row[4]), message: "Typ darf nicht gesetzt sein." },
            { check: row => !isEmpty(row[5]), message: "Kategorie darf nicht gesetzt sein." },
            { check: row => !isEmpty(row[6]), message: "Konto (Soll) darf nicht gesetzt sein." },
            { check: row => !isEmpty(row[7]), message: "Gegenkonto (Haben) darf nicht gesetzt sein." }
        ];
        const dataRowRules = [
            { check: row => isEmpty(row[0]), message: "Buchungsdatum fehlt." },
            { check: row => isEmpty(row[1]), message: "Buchungstext fehlt." },
            { check: row => isEmpty(row[2]) || isInvalidNumber(row[2]), message: "Betrag fehlt oder ungültig." },
            { check: row => isEmpty(row[3]) || isInvalidNumber(row[3]), message: "Saldo fehlt oder ungültig." },
            { check: row => isEmpty(row[4]), message: "Typ fehlt." },
            { check: row => isEmpty(row[5]), message: "Kategorie fehlt." },
            { check: row => isEmpty(row[6]), message: "Konto (Soll) fehlt." },
            { check: row => isEmpty(row[7]), message: "Gegenkonto (Haben) fehlt." }
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

    // Aggregiert Warnungen aus allen relevanten Sheets und zeigt einen Alert, wenn Fehler vorhanden sind
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
// Importiert Dateien aus definierten Ordnern in die entsprechenden Sheets
const ImportModule = (() => {
    const importFilesFromFolder = (folder, importSheet, mainSheet, type, historySheet) => {
        const files = folder.getFiles();
        // Bestimme bereits importierte Dateien
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
                    invoiceDate,   // Spalte 1: Rechnungsdatum
                    invoiceName,   // Spalte 2: Rechnungsname
                    "", "", "", "", // Spalte 3-6: Platzhalter (Kategorie, Kunde, Nettobetrag, MwSt)
                    "",            // Spalte 7: MWSt (wird später per Refresh gesetzt)
                    "",            // Spalte 8: Brutto (wird später per Refresh gesetzt)
                    "",            // Spalte 9: (leer, sofern benötigt)
                    "",            // Spalte 10: Restbetrag (wird später per Refresh gesetzt)
                    "",            // Spalte 11: Quartal (wird später per Refresh gesetzt)
                    "",            // Spalte 12: Zahlungsstatus (wird später per Refresh gesetzt)
                    "", "", "",    // Spalte 13-15: weitere Platzhalter
                    timestamp,     // Spalte 16: Timestamp
                    fileName,      // Spalte 17: Dateiname
                    fileUrl        // Spalte 18: URL
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
        newImportRows.length &&
        importSheet.getRange(importSheet.getLastRow() + 1, 1, newImportRows.length, newImportRows[0].length)
            .setValues(newImportRows);
        newMainRows.length &&
        mainSheet.getRange(mainSheet.getLastRow() + 1, 1, newMainRows.length, newMainRows[0].length)
            .setValues(newMainRows);
        newHistoryRows.length &&
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

        // Initialisiere Header, falls die Sheets leer sind
        revenue.getLastRow() === 0 && revenue.appendRow(["Dateiname", "Link zur Datei", "Rechnungsnummer"]);
        expense.getLastRow() === 0 && expense.appendRow(["Dateiname", "Link zur Datei", "Rechnungsnummer"]);
        history.getLastRow() === 0 && history.appendRow(["Datum", "Rechnungstyp", "Dateiname", "Link zur Datei"]);

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
// Aktualisiert Datenblätter (Formeln, Formate, Validierungen) und das Bank-Sheet separat
const RefreshModule = (() => {
    // Aktualisiert Einnahmen/Ausgaben/Eigenbelege-Blätter
    const refreshDataSheet = sheet => {
        const lastRow = sheet.getLastRow();
        if (lastRow < 2) return;
        const numRows = lastRow - 1;

        // Setze Formeln in den angegebenen Spalten
        Object.entries({
            7: row => `=E${row}*F${row}`,                           // MwSt
            8: row => `=E${row}+G${row}`,                           // Brutto
            10: row => `=(H${row}-I${row})/(1+VALUE(F${row}))`,       // Restbetrag
            11: row => `=IF(A${row}=""; ""; ROUNDUP(MONTH(A${row})/3;0))`, // Quartal
            12: row => `=IF(OR(I${row}=""; I${row}=0); "Offen"; IF(I${row}>=H${row}; "Bezahlt"; "Teilbezahlt"))` // Zahlungsstatus
        }).forEach(([col, formulaFn]) => {
            const formulas = Array.from({ length: numRows }, (_, i) => [formulaFn(i + 2)]);
            sheet.getRange(2, Number(col), numRows, 1).setFormulas(formulas);
        });

        // Setze 0 in Spalte 9, falls leer (alle Zellen auf einmal abfragen und zurückschreiben)
        const col9Range = sheet.getRange(2, 9, numRows, 1);
        const col9Values = col9Range.getValues().map(([val]) => (val === "" || val === null ? 0 : val));
        col9Range.setValues(col9Values.map(val => [val]));

        // Data Validation für Dropdowns mit Short-Circuit If
        const name = sheet.getName();
        if (name === "Einnahmen") Validator.validateDropdown(sheet, 2, 3, lastRow - 1, 1, CategoryConfig.einnahmen.category);
        if (name === "Ausgaben") Validator.validateDropdown(sheet, 2, 3, lastRow - 1, 1, CategoryConfig.ausgaben.category);
        if (name === "Eigenbelege") Validator.validateDropdown(sheet, 2, 3, lastRow - 1, 1, CategoryConfig.eigenbelege.category);
        Validator.validateDropdown(sheet, 2, 13, lastRow - 1, 1, CategoryConfig.common.paymentType);
        sheet.autoResizeColumns(1, sheet.getLastColumn());
    };

    // Aktualisiert das Bank-Sheet (inkl. bedingter Formatierung und Endsaldo)
    const refreshBankSheet = sheet => {
        const lastRow = sheet.getLastRow();
        if (lastRow < 3) return;
        const firstDataRow = 3;
        const numDataRows = lastRow - firstDataRow + 1;
        const transRows = lastRow - firstDataRow - 1;

        // 1. Formeln für Transaktionszeilen (Spalte D) setzen
        if (transRows > 0) {
            sheet.getRange(firstDataRow, 4, transRows, 1).setFormulas(
                Array.from({ length: transRows }, (_, i) => [
                    `=D${firstDataRow + i - 1}+C${firstDataRow + i}`
                ])
            );
        }

        // 2. Batchweise Transaktionstyp (Spalte E) anhand von Spalte C (Betrag) setzen
        const amounts = sheet.getRange(firstDataRow, 3, numDataRows, 1).getValues();
        const typeValues = amounts.map(([val]) => {
            const amt = parseFloat(val) || 0;
            return [amt > 0 ? "Einnahme" : amt < 0 ? "Ausgabe" : ""];
        });
        sheet.getRange(firstDataRow, 5, numDataRows, 1).setValues(typeValues);

        // 3. Data Validations setzen
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

        // 4. Bedingte Formatierung für Spalte E (Transaktionstyp)
        Helpers.setConditionalFormattingForColumn(sheet, "E", [
            { value: "Einnahme", background: "#C6EFCE", fontColor: "#006100" },
            { value: "Ausgabe", background: "#FFC7CE", fontColor: "#9C0006" }
        ]);

        // 5. Zahlenformate: Spalte A (Datum) und Spalten C, D (Währungsformat)
        sheet.getRange(`A2:A${lastRow}`).setNumberFormat("DD.MM.YYYY");
        ["C", "D"].forEach(col =>
            sheet.getRange(`${col}2:${col}${lastRow}`).setNumberFormat("€#,##0.00;€-#,##0.00")
        );

        // 6. Konto‑Mapping: Werte in Spalten G (Soll) und H (Gegenkonto) anhand von Spalte E (Typ) und F (Kategorie) setzen
        const dataRange = sheet.getRange(firstDataRow, 1, numDataRows, sheet.getLastColumn());
        const data = dataRange.getValues();
        data.forEach((row, i) => {
            const globalRow = i + firstDataRow;
            // Überspringe Endsaldo-Zeile (falls in Spalte B "endsaldo" steht)
            const label = row[1] ? row[1].toString().trim().toLowerCase() : "";
            if (globalRow === lastRow && label === "endsaldo") return;
            const type = row[4];     // Spalte E
            const category = row[5]; // Spalte F
            let mapping = type === "Einnahme"
                ? CategoryConfig.einnahmen.kontoMapping[category]
                : type === "Ausgabe"
                    ? CategoryConfig.ausgaben.kontoMapping[category]
                    : null;
            if (!mapping) mapping = { soll: "Manuell prüfen", gegen: "Manuell prüfen" };
            row[6] = mapping.soll;  // Spalte G
            row[7] = mapping.gegen; // Spalte H
        });
        dataRange.setValues(data);

        // 7. Endsaldo-Zeile aktualisieren oder hinzufügen
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

    // TODO: refresh refreshGesellschafterkontoSheet
    const refreshGesellschafterkontoSheet = sheet => {
    };

    // TODO: refresh refreshHoldingTransfersSheet
    const refreshHoldingTransfersSheet = sheet => {
    };

    // Aktualisiert das aktuell aktive Blatt basierend auf dessen Namen
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

    // Aktualisiert alle relevanten Blätter
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

// =================== Modul: GuV-Berechnung ===================
// Berechnet die Gewinn- und Verlustrechnung (GuV)
const GuVCalculator = (() => {
    // Erzeugt ein leeres GuV-Objekt mit Standardwerten
    const createEmptyGuV = () => ({ einnahmen: 0, ausgaben: 0, ust_0: 0, ust_7: 0, ust_19: 0, vst_0: 0, vst_7: 0, vst_19: 0 });

    // Verarbeitet eine einzelne Zeile und aktualisiert die GuV-Daten
    const processGuVRow = (row, guvData, isIncome) => {
        const paymentDate = Helpers.parseDate(row[12]);
        if (!paymentDate || paymentDate > new Date()) return;
        const month = paymentDate.getMonth() + 1,
            netto = Helpers.parseCurrency(row[4]),
            restNetto = Helpers.parseCurrency(row[9]) || 0,
            gezahltNetto = Math.max(0, netto - restNetto),
            mwstRate = Helpers.parseMwstRate(row[5]),
            tax = gezahltNetto * (mwstRate / 100);
        isIncome
            ? (guvData[month].einnahmen += gezahltNetto, guvData[month][`ust_${Math.round(mwstRate)}`] += tax)
            : (guvData[month].ausgaben += gezahltNetto, guvData[month][`vst_${Math.round(mwstRate)}`] += tax);
    };

    // Formatiert eine Zeile der GuV-Ausgabe
    const formatGuVRow = (label, data) => {
        const ustZahlung = data.ust_19 - data.vst_19,
            ergebnis = data.einnahmen - data.ausgaben;
        return [label, data.einnahmen, data.ausgaben, data.ust_0, data.ust_7, data.ust_19, data.vst_0, data.vst_7, data.vst_19, ustZahlung, ergebnis];
    };

    // Aggregiert GuV-Daten über einen Zeitraum
    const aggregateGuV = (guvData, start, end) => {
        const sum = createEmptyGuV();
        for (let m = start; m <= end; m++) {
            for (const key in sum) sum[key] += guvData[m][key] || 0;
        }
        return sum;
    };

    // Hauptfunktion zur Berechnung der GuV und Ausgabe ins "GuV"-Sheet
    const calculateGuV = () => {
        const ss = SpreadsheetApp.getActiveSpreadsheet(),
            revenueSheet = ss.getSheetByName("Einnahmen"),
            expenseSheet = ss.getSheetByName("Ausgaben");
        if (!revenueSheet || !expenseSheet) {
            SpreadsheetApp.getUi().alert("Fehlendes Blatt: 'Einnahmen' oder 'Ausgaben'");
            return;
        }
        // Validierung der Daten in Revenue und Ausgaben
        if (!Validator.validateAllSheets(revenueSheet, expenseSheet)) return;
        const revenueData = revenueSheet.getDataRange().getValues().slice(1),
            expenseData = expenseSheet.getDataRange().getValues().slice(1),
            guvData = {};
        for (let m = 1; m <= 12; m++) guvData[m] = createEmptyGuV();
        revenueData.forEach(row => processGuVRow(row, guvData, true));
        expenseData.forEach(row => processGuVRow(row, guvData, false));
        const guvSheet = ss.getSheetByName("GuV") || ss.insertSheet("GuV");
        guvSheet.clearContents();
        guvSheet.appendRow(["Zeitraum", "Einnahmen (netto)", "Ausgaben (netto)", "USt 0%", "USt 7%", "USt 19%", "VSt 0%", "VSt 7%", "VSt 19%", "USt-Zahlung", "Ergebnis (Gewinn/Verlust)"]);
        const monthNames = ["Januar","Februar","März","April","Mai","Juni","Juli","August","September","Oktober","November","Dezember"];
        for (let m = 1; m <= 12; m++) {
            guvSheet.appendRow(formatGuVRow(monthNames[m - 1], guvData[m]));
            m % 3 === 0 && guvSheet.appendRow(formatGuVRow(`Quartal ${m / 3}`, aggregateGuV(guvData, m - 2, m)));
        }
        guvSheet.appendRow(formatGuVRow("Gesamtjahr", aggregateGuV(guvData, 1, 12)));
        guvSheet.getRange(2, 2, guvSheet.getLastRow() - 1, 1).setNumberFormat("#,##0.00€");
        guvSheet.autoResizeColumns(1, guvSheet.getLastColumn());
        SpreadsheetApp.getUi().alert("GuV wurde aktualisiert!");
    };

    return { calculateGuV };
})();

// =================== Modul: BWACalculator ===================
// Berechnet die betriebswirtschaftliche Auswertung (BWA)
const BWACalculator = (() => {
    const getBwaCategory = (category, isIncome, rowIndex, fehlendeKategorien, type = "operativ") => {
        const mapping = type === "bank"
            ? CategoryConfig.bank.bwaMapping
            : isIncome
                ? CategoryConfig.einnahmen.bwaMapping
                : CategoryConfig.ausgaben.bwaMapping;
        if (!category || !mapping[category]) {
            const fallback = isIncome ? "sonstigeErtraege" : "betriebskosten";
            fehlendeKategorien.push(`Zeile ${rowIndex}: Unbekannte Kategorie "${category || "N/A"}" → Verwende "${fallback}"`);
            return fallback;
        }
        return mapping[category];
    };

    const calculateBWA = () => {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const revenueSheet = ss.getSheetByName("Einnahmen");
        const expenseSheet = ss.getSheetByName("Ausgaben");
        const bankSheet = ss.getSheetByName("Bankbewegungen");
        const bwaSheet = ss.getSheetByName("BWA") || ss.insertSheet("BWA");

        if (!Validator.validateAllSheets(revenueSheet, expenseSheet, bankSheet)) return;

        const categories = {
            einnahmen: { umsatzerloese: 0, provisionserloese: 0, sonstigeErtraege: 0, privateinlage: 0, darlehen: 0, zinsen: 0 },
            ausgaben: { wareneinsatz: 0, betriebskosten: 0, marketing: 0, reisen: 0, personalkosten: 0, sonstigeAufwendungen: 0, eigenbeleg: 0, privatentnahme: 0, darlehen: 0, zinsen: 0 }
        };
        let totalEinnahmen = 0, totalAusgaben = 0, offeneForderungen = 0, offeneVerbindlichkeiten = 0, totalLiquiditaet = 0;
        const fehlendeKategorien = [];

        // Verarbeite Einnahmen-Daten
        const revenueData = revenueSheet.getDataRange().getValues().slice(1);
        revenueData.forEach((row, index) => {
            const nettoBetrag = parseFloat(row[4]) || 0,
                restBetrag = parseFloat(row[9]) || 0,
                bwaCat = getBwaCategory(row[2], true, index + 2, fehlendeKategorien);
            categories.einnahmen[bwaCat] += nettoBetrag;
            totalEinnahmen += nettoBetrag;
            offeneForderungen += restBetrag;
        });
        // Verarbeite Ausgaben-Daten
        const expenseData = expenseSheet.getDataRange().getValues().slice(1);
        expenseData.forEach((row, index) => {
            const nettoBetrag = parseFloat(row[4]) || 0,
                restBetrag = parseFloat(row[9]) || 0,
                bwaCat = getBwaCategory(row[2], false, index + 2, fehlendeKategorien);
            categories.ausgaben[bwaCat] = (categories.ausgaben[bwaCat] || 0) + nettoBetrag;
            totalAusgaben += nettoBetrag;
            offeneVerbindlichkeiten += restBetrag;
        });
        // Lese den letzten Saldo aus dem Bank-Sheet
        if (bankSheet) {
            const bankData = bankSheet.getDataRange().getValues();
            for (let i = 1; i < bankData.length - 1; i++) {
                totalLiquiditaet = parseFloat(bankData[i][3]) || 0;
            }
        }

        // Schreibe die BWA-Ergebnisse ins BWA-Sheet
        bwaSheet.clearContents();
        bwaSheet.appendRow(["Position", "Betrag (€)"]);
        bwaSheet.appendRow(["--- EINNAHMEN ---", ""]);
        Object.keys(categories.einnahmen).forEach(key => bwaSheet.appendRow([key, categories.einnahmen[key]]));
        bwaSheet.appendRow(["Gesamteinnahmen", totalEinnahmen]);
        bwaSheet.appendRow(["Erhaltene Einnahmen", totalEinnahmen - offeneForderungen]);
        bwaSheet.appendRow(["Offene Forderungen", offeneForderungen]);
        bwaSheet.appendRow(["--- AUSGABEN ---", ""]);
        Object.keys(categories.ausgaben).forEach(key => bwaSheet.appendRow([key, -categories.ausgaben[key]]));
        bwaSheet.appendRow(["Gesamtausgaben", -totalAusgaben]);
        bwaSheet.appendRow(["Offene Verbindlichkeiten", offeneVerbindlichkeiten]);
        const ergebnis = totalEinnahmen - totalAusgaben;
        bwaSheet.appendRow(["Betriebsergebnis", ergebnis]);
        bwaSheet.appendRow(["--- STEUERN ---", ""]);
        bwaSheet.appendRow(["Umsatzsteuer-Zahlung", 0]);
        bwaSheet.appendRow(["Vorsteuer", 0]);
        bwaSheet.appendRow(["Körperschaftsteuer", 0]);
        bwaSheet.appendRow(["Gewerbesteuer", 0]);
        bwaSheet.appendRow(["Ergebnis nach Steuern", ergebnis]);
        bwaSheet.appendRow(["--- FINANZIERUNG ---", ""]);
        bwaSheet.appendRow(["Eigenbeleg", 0]);
        bwaSheet.appendRow(["Privateinlage", 0]);
        bwaSheet.appendRow(["Privatentnahme", 0]);
        bwaSheet.appendRow(["Darlehen", 0]);
        bwaSheet.appendRow(["--- LIQUIDITÄT ---", ""]);
        bwaSheet.appendRow(["Kontostand (Bankbewegungen)", totalLiquiditaet]);
        bwaSheet.getRange("A1:B1").setFontWeight("bold");
        bwaSheet.getRange(2, 2, bwaSheet.getLastRow() - 1, 1).setNumberFormat("€#,##0.00;€-#,##0.00");
        bwaSheet.autoResizeColumns(1, bwaSheet.getLastColumn());

        fehlendeKategorien.length
            ? SpreadsheetApp.getUi().alert("Folgende Kategorien konnten nicht zugeordnet werden:\n" + fehlendeKategorien.join("\n"))
            : SpreadsheetApp.getUi().alert("BWA wurde erfolgreich berechnet und aktualisiert!");
    };

    return { calculateBWA };
})();


// =================== Globale Funktionen ===================
// Erzeugt das Menü und setzt Trigger
const onOpen = () => {
    SpreadsheetApp.getUi()
        .createMenu("📂 Buchhaltung")
        .addItem("📥 Dateien importieren", "importDriveFiles")
        .addItem("🔄 Refresh Active Sheet", "refreshSheet")
        .addItem("📊 GuV berechnen", "calculateGuV")
        .addItem("📈 BWA berechnen", "calculateBWA")
        .addToUi();
};

const onEdit = (e) => {
    const { range } = e;
    const sheet = range.getSheet();
    const name = sheet.getName();

    // Mapping: Sheetname -> Zielspalte für den Timestamp
    const mapping = {
        "Einnahmen": 16,         // Spalte P
        "Ausgaben": 16,          // Spalte P
        "Eigenbelege": 16,       // Spalte P
        "Bankbewegungen": 11,    // Spalte K
        "Gesellschafterkonto": 12, // Spalte L
        "Holding Transfers": 6   // Spalte F
    };

    // Trigger nur in den angegebenen Sheets ausführen
    if (!(name in mapping)) return;

    // Ermittle die Anzahl der Spalten in der Headerzeile (Zeile 1)
    const headerLen = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].length;
    // Bearbeite nur Zellen in den Spalten, die in der Headerzeile existieren
    if (range.getColumn() > headerLen) return;

    // Überspringe, falls in der Zielspalte (Timestamp-Spalte) editiert wurde
    if (range.getColumn() === mapping[name]) return;

    // Timestamp im deutschen Format erstellen
    const ts = new Date();
    // Timestamp in die entsprechende Spalte in der gleichen Zeile schreiben
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

// Wrapper-Funktionen für einfache Trigger-Aufrufe
const refreshSheet = () => RefreshModule.refreshActiveSheet();
const calculateGuV = () => { RefreshModule.refreshAllSheets(); GuVCalculator.calculateGuV(); };
const calculateBWA = () => { RefreshModule.refreshAllSheets(); BWACalculator.calculateBWA(); };
const importDriveFiles = () => { ImportModule.importDriveFiles(); RefreshModule.refreshAllSheets(); };
