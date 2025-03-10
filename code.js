// =================== Zentrale Konfiguration ===================
const CategoryConfig = {
    common: {
        paymentType: ["Überweisung", "Bar", "Kreditkarte", "Paypal"],
        months: ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember"],
        defaultMwst: 19
    },
    einnahmen: {
        categories: {
            "Erlöse aus Lieferungen und Leistungen": { taxType: "steuerpflichtig" },
            "Provisionserlöse": { taxType: "steuerpflichtig" },
            "Sonstige betriebliche Erträge": { taxType: "steuerpflichtig" },
            "Erträge aus Vermietung/Verpachtung": { taxType: "steuerfrei_inland" },
            "Erträge aus Zuschüssen": { taxType: "steuerpflichtig" },
            "Erträge aus Währungsgewinnen": { taxType: "steuerpflichtig" },
            "Erträge aus Anlagenabgängen": { taxType: "steuerpflichtig" },
            "Darlehen": { taxType: "steuerfrei_inland" }, // wird nicht in der BWA berücksichtigt
            "Zinsen": { taxType: "steuerfrei_inland" }    // wird nicht in der BWA berücksichtigt
        },
        kontoMapping: {
            "Erlöse aus Lieferungen und Leistungen": { soll: "1200 Bank", gegen: "4400 Erlöse aus L&L" },
            "Provisionserlöse": { soll: "1200 Bank", gegen: "4420 Provisionserlöse" },
            "Sonstige betriebliche Erträge": { soll: "1200 Bank", gegen: "4490 Sonstige betriebliche Erträge" },
            "Erträge aus Vermietung/Verpachtung": { soll: "1200 Bank", gegen: "4410 Vermietung/Verpachtung" },
            "Erträge aus Zuschüssen": { soll: "1200 Bank", gegen: "4430 Zuschüsse" },
            "Erträge aus Währungsgewinnen": { soll: "1200 Bank", gegen: "4440 Währungsgewinne" },
            "Erträge aus Anlagenabgängen": { soll: "1200 Bank", gegen: "4450 Anlagenabgänge" },
            "Darlehen": { soll: "1200 Bank", gegen: "3000 Darlehen" },
            "Zinsen": { soll: "1200 Bank", gegen: "2650 Zinserträge" }
        },
        // Für die BWA aggregieren wir nur die operativen Umsätze:
        bwaMapping: {
            "Erlöse aus Lieferungen und Leistungen": "umsatzerloese",
            "Provisionserlöse": "provisionserloese",
            "Sonstige betriebliche Erträge": "sonstigeErtraege",
            "Erträge aus Vermietung/Verpachtung": "vermietung",
            "Erträge aus Zuschüssen": "zuschuesse",
            "Erträge aus Währungsgewinnen": "waehrungsgewinne",
            "Erträge aus Anlagenabgängen": "anlagenabgaenge"
            // Darlehen & Zinsen werden nicht in die BWA einbezogen.
        }
    },
    ausgaben: {
        categories: {
            "Wareneinsatz": { taxType: "steuerpflichtig" },
            "Bezogene Leistungen": { taxType: "steuerpflichtig" },
            "Roh-, Hilfs- & Betriebsstoffe": { taxType: "steuerpflichtig" },
            "Betriebskosten": { taxType: "steuerpflichtig" },
            "Marketing & Werbung": { taxType: "steuerpflichtig" },
            "Reisekosten": { taxType: "steuerpflichtig" },
            "Personalkosten": { taxType: "steuerpflichtig" },
            "Bruttolöhne & Gehälter": { taxType: "steuerpflichtig" },
            "Soziale Abgaben & Arbeitgeberanteile": { taxType: "steuerpflichtig" },
            "Sonstige Personalkosten": { taxType: "steuerpflichtig" },
            "Sonstige betriebliche Aufwendungen": { taxType: "steuerpflichtig" },
            "Miete": { taxType: "steuerfrei_inland" },
            "Versicherungen": { taxType: "steuerfrei_inland" },
            "Porto": { taxType: "steuerfrei_inland" },
            "Google Ads": { taxType: "steuerfrei_ausland" },
            "AWS": { taxType: "steuerfrei_ausland" },
            "Facebook Ads": { taxType: "steuerfrei_ausland" },
            "Bewirtung": { taxType: "steuerpflichtig" },
            "Telefon & Internet": { taxType: "steuerpflichtig" },
            "Bürokosten": { taxType: "steuerpflichtig" },
            "Fortbildungskosten": { taxType: "steuerpflichtig" },
            "Abschreibungen Maschinen": { taxType: "steuerpflichtig" },
            "Abschreibungen Büroausstattung": { taxType: "steuerpflichtig" },
            "Abschreibungen immaterielle Wirtschaftsgüter": { taxType: "steuerpflichtig" },
            "Zinsen auf Bankdarlehen": { taxType: "steuerpflichtig" },
            "Zinsen auf Gesellschafterdarlehen": { taxType: "steuerpflichtig" },
            "Leasingkosten": { taxType: "steuerpflichtig" },
            "Gewerbesteuerrückstellungen": { taxType: "steuerpflichtig" },
            "Körperschaftsteuer": { taxType: "steuerpflichtig" },
            "Solidaritätszuschlag": { taxType: "steuerpflichtig" },
            "Sonstige Steuerrückstellungen": { taxType: "steuerpflichtig" }
        },
        kontoMapping: {
            "Wareneinsatz": { soll: "4900 Wareneinsatz", gegen: "1200 Bank" },
            "Bezogene Leistungen": { soll: "4900 Bezogene Leistungen", gegen: "1200 Bank" },
            "Roh-, Hilfs- & Betriebsstoffe": { soll: "4900 Roh-, Hilfs- & Betriebsstoffe", gegen: "1200 Bank" },
            "Betriebskosten": { soll: "4900 Betriebskosten", gegen: "1200 Bank" },
            "Marketing & Werbung": { soll: "4900 Marketing & Werbung", gegen: "1200 Bank" },
            "Reisekosten": { soll: "4900 Reisekosten", gegen: "1200 Bank" },
            "Bruttolöhne & Gehälter": { soll: "4900 Personalkosten", gegen: "1200 Bank" },
            "Soziale Abgaben & Arbeitgeberanteile": { soll: "4900 Personalkosten", gegen: "1200 Bank" },
            "Sonstige Personalkosten": { soll: "4900 Personalkosten", gegen: "1200 Bank" },
            "Sonstige betriebliche Aufwendungen": { soll: "4900 Sonstige betriebliche Aufwendungen", gegen: "1200 Bank" },
            "Miete": { soll: "4900 Miete & Nebenkosten", gegen: "1200 Bank" },
            "Versicherungen": { soll: "4900 Versicherungen", gegen: "1200 Bank" },
            "Porto": { soll: "4900 Betriebskosten", gegen: "1200 Bank" },
            "Google Ads": { soll: "4900 Marketing & Werbung", gegen: "1200 Bank" },
            "AWS": { soll: "4900 Marketing & Werbung", gegen: "1200 Bank" },
            "Facebook Ads": { soll: "4900 Marketing & Werbung", gegen: "1200 Bank" },
            "Bewirtung": { soll: "4900 Bewirtung", gegen: "1200 Bank" },
            "Telefon & Internet": { soll: "4900 Telefon & Internet", gegen: "1200 Bank" },
            "Bürokosten": { soll: "4900 Bürokosten", gegen: "1200 Bank" },
            "Fortbildungskosten": { soll: "4900 Fortbildungskosten", gegen: "1200 Bank" },
            "Abschreibungen Maschinen": { soll: "4900 Abschreibungen Maschinen", gegen: "1200 Bank" },
            "Abschreibungen Büroausstattung": { soll: "4900 Abschreibungen Büroausstattung", gegen: "1200 Bank" },
            "Abschreibungen immaterielle Wirtschaftsgüter": { soll: "4900 Abschreibungen Immateriell", gegen: "1200 Bank" },
            "Zinsen auf Bankdarlehen": { soll: "4900 Zinsen Bankdarlehen", gegen: "1200 Bank" },
            "Zinsen auf Gesellschafterdarlehen": { soll: "4900 Zinsen Gesellschafterdarlehen", gegen: "1200 Bank" },
            "Leasingkosten": { soll: "4900 Leasingkosten", gegen: "1200 Bank" },
            "Gewerbesteuerrückstellungen": { soll: "4900 Rückstellungen", gegen: "1200 Bank" },
            "Körperschaftsteuer": { soll: "4900 Körperschaftsteuer", gegen: "1200 Bank" },
            "Solidaritätszuschlag": { soll: "4900 Solidaritätszuschlag", gegen: "1200 Bank" },
            "Sonstige Steuerrückstellungen": { soll: "4900 Sonstige Steuerrückstellungen", gegen: "1200 Bank" }
        },
        bwaMapping: {
            // Gruppe 1:
            "Erlöse aus Lieferungen und Leistungen": "umsatzerloese",
            "Provisionserlöse": "provisionserloese",
            "Sonstige betriebliche Erträge": "sonstigeErtraege",
            "Erträge aus Vermietung/Verpachtung": "vermietung",
            "Erträge aus Zuschüssen": "zuschuesse",
            "Erträge aus Währungsgewinnen": "waehrungsgewinne",
            "Erträge aus Anlagenabgängen": "anlagenabgaenge",
            // Gruppe 2: (keine Anpassung – übernimmt die Felder direkt)
            // Gruppe 3:
            "Wareneinsatz": "wareneinsatz",
            "Bezogene Leistungen": "bezogeneLeistungen",
            "Roh-, Hilfs- & Betriebsstoffe": "rohHilfsBetriebsstoffe",
            "Betriebskosten": "betriebskosten",
            "Marketing & Werbung": "werbungMarketing",
            "Reisekosten": "reisekosten",
            "Bruttolöhne & Gehälter": "bruttoLoehne",
            "Soziale Abgaben & Arbeitgeberanteile": "sozialeAbgaben",
            "Sonstige Personalkosten": "sonstigePersonalkosten",
            "Sonstige betriebliche Aufwendungen": "sonstigeAufwendungen",
            "Miete": "mieteNebenkosten",
            "Versicherungen": "versicherungen",
            "Porto": "betriebskosten",
            "Telefon & Internet": "telefonInternet",
            "Bürokosten": "buerokosten",
            "Fortbildungskosten": "fortbildungskosten",
            // Gruppe 4:
            "Abschreibungen Maschinen": "abschreibungenMaschinen",
            "Abschreibungen Büroausstattung": "abschreibungenBueromaterial",
            "Abschreibungen immaterielle Wirtschaftsgüter": "abschreibungenImmateriell",
            "Zinsen auf Bankdarlehen": "zinsenBank",
            "Zinsen auf Gesellschafterdarlehen": "zinsenGesellschafter",
            "Leasingkosten": "leasingkosten",
            // Gruppe 5: Besondere Posten (Kapitalbewegungen)
            "Eigenkapitalveränderungen": "eigenkapitalveraenderungen",
            "Gesellschafterdarlehen": "gesellschafterdarlehen",
            "Ausschüttungen an Gesellschafter": "ausschuettungen",
            // Gruppe 6:
            "Gewerbesteuerrückstellungen": "gewerbesteuerRueckstellungen",
            // Gruppe 8:
            "Körperschaftsteuer": "koerperschaftsteuer",
            "Solidaritätszuschlag": "solidaritaetszuschlag",
            "Sonstige Steuerrückstellungen": "sonstigeSteuerrueckstellungen"
        }
    },
    bank: {
        category: [
            "Erlöse aus Lieferungen und Leistungen", "Provisionserlöse", "Sonstige betriebliche Erträge",
            "Wareneinsatz", "Betriebskosten", "Marketing & Werbung", "Reisekosten", "Personalkosten",
            "Sonstige betriebliche Aufwendungen", "Privateinlage", "Darlehen", "Eigenbeleg",
            "Privatentnahme", "Zinsen", "Gewinnübertrag", "Kapitalrückführung"
        ],
        type: ["Einnahme", "Ausgabe"]
    },
    gesellschafterkonto: {
        category: ["Privateinlage", "Privatentnahme", "Darlehen"],
        shareholder: ["Christopher Giebel", "Hendrik Werner"]
    },
    eigenbelege: {
        category: ["Kleidung", "Trinkgeld", "Private Vorauslage", "Bürokosten", "Reisekosten", "Bewirtung", "Sonstiges"],
        mapping: {
            "Kleidung": { taxType: "steuerpflichtig" },
            "Trinkgeld": { taxType: "steuerfrei" },
            "Private Vorauslage": { taxType: "steuerfrei" },
            "Bürokosten": { taxType: "steuerpflichtig" },
            "Reisekosten": { taxType: "steuerpflichtig" },
            "Bewirtung": { taxType: "eigenbeleg", besonderheit: "bewirtung" },
            "Sonstiges": { taxType: "steuerpflichtig" }
        },
        status: ["Offen", "Erstattet", "Gebucht"]
    },
    holdingTransfers: {
        category: ["Gewinnübertrag", "Kapitalrückführung"] // wird in der BWA ignoriert
    }
};

// =================== Modul: Helpers ===================
const Helpers = (() => {
    const parseDate = value => {
        if (value instanceof Date) return isNaN(value.getTime()) ? null : value;
        if (typeof value === "string") {
            const d = new Date(value);
            return isNaN(d.getTime()) ? null : d;
        }
        return null;
    };

    const parseCurrency = value => {
        if (typeof value === "number") return value;
        const str = value.toString().replace(/[^\d,.-]/g, "").replace(",", ".");
        const num = parseFloat(str);
        return isNaN(num) ? 0 : num;
    };

    const parseMwstRate = value => {
        if (typeof value === "number") return value < 1 ? value * 100 : value;
        const rate = parseFloat(value.toString().replace("%", "").replace(",", ".").trim());
        return isNaN(rate) ? CategoryConfig.common.defaultMwst : rate;
    };

    const getFolderByName = (parent, name) => {
        const folderIter = parent.getFoldersByName(name);
        return folderIter.hasNext() ? folderIter.next() : null;
    };

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
const Validator = (() => {
    const isEmpty = v => v == null || v.toString().trim() === "";
    const isInvalidNumber = v => isEmpty(v) || isNaN(parseFloat(v.toString().trim()));

    const validateDropdown = (sheet, row, col, numRows, numCols, list) =>
        sheet.getRange(row, col, numRows, numCols).setDataValidation(
            SpreadsheetApp.newDataValidation().requireValueInList(list, true).build()
        );

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
        const revData = revenueSheet.getDataRange().getValues();
        const expData = expenseSheet.getDataRange().getValues();
        const revenueWarnings = revData.length > 1 ? revData.slice(1).reduce((acc, row, i) => acc.concat(validateRevenueAndExpenses(row, i + 2)), []) : [];
        const expenseWarnings = expData.length > 1 ? expData.slice(1).reduce((acc, row, i) => acc.concat(validateRevenueAndExpenses(row, i + 2)), []) : [];
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
    const refreshDataSheet = sheet => {
        const lastRow = sheet.getLastRow();
        if (lastRow < 2) return;
        const numRows = lastRow - 1;

        Object.entries({
            7: row => `=E${row}*F${row}`,
            8: row => `=E${row}+G${row}`,
            10: row => `=(H${row}-I${row})/(1+VALUE(F${row}))`,
            11: row => `=IF(A${row}=""; ""; ROUNDUP(MONTH(A${row})/3;0))`,
            12: row => `=IF(OR(I${row}=""; I${row}=0); "Offen"; IF(I${row}>=H${row}; "Bezahlt"; "Teilbezahlt"))`
        }).forEach(([col, formulaFn]) => {
            const formulas = Array.from({ length: numRows }, (_, i) => [formulaFn(i + 2)]);
            sheet.getRange(2, Number(col), numRows, 1).setFormulas(formulas);
        });

        const col9Range = sheet.getRange(2, 9, numRows, 1);
        const col9Values = col9Range.getValues().map(([val]) => (val === "" || val === null ? 0 : val));
        col9Range.setValues(col9Values.map(val => [val]));

        const name = sheet.getName();
        if (name === "Einnahmen")
            Validator.validateDropdown(sheet, 2, 3, lastRow - 1, 1, Object.keys(CategoryConfig.einnahmen.categories));
        if (name === "Ausgaben")
            Validator.validateDropdown(sheet, 2, 3, lastRow - 1, 1, Object.keys(CategoryConfig.ausgaben.categories));
        if (name === "Eigenbelege")
            Validator.validateDropdown(sheet, 2, 3, lastRow - 1, 1, CategoryConfig.eigenbelege.category);
        Validator.validateDropdown(sheet, 2, 13, lastRow - 1, 1, CategoryConfig.common.paymentType);
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
                Array.from({ length: transRows }, (_, i) => [`=D${firstDataRow + i - 1}+C${firstDataRow + i}`])
            );
        }

        const amounts = sheet.getRange(firstDataRow, 3, numDataRows, 1).getValues();
        const typeValues = amounts.map(([val]) => {
            const amt = parseFloat(val) || 0;
            return [amt > 0 ? "Einnahme" : amt < 0 ? "Ausgabe" : ""];
        });
        sheet.getRange(firstDataRow, 5, numDataRows, 1).setValues(typeValues);

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
    const createEmptyUStVA = () => ({
        steuerpflichtige_einnahmen: 0,
        steuerfreie_inland_einnahmen: 0,
        steuerfreie_ausland_einnahmen: 0,
        steuerpflichtige_ausgaben: 0,
        steuerfreie_inland_ausgaben: 0,
        steuerfreie_ausland_ausgaben: 0,
        eigenbelege_steuerpflichtig: 0,
        eigenbelege_steuerfrei: 0,
        nicht_abzugsfaehige_vst: 0,
        ust_7: 0,
        ust_19: 0,
        vst_7: 0,
        vst_19: 0
    });

    const processUStVARow = (row, ustvaData, isIncome, isEigenbelegSheet = false) => {
        const paymentDate = Helpers.parseDate(row[13]);
        const month = paymentDate instanceof Date ? paymentDate.getMonth() + 1 : 0;
        if (!paymentDate || month === 0 || paymentDate > new Date()) return;
        const netto = Helpers.parseCurrency(row[4]);
        const restNetto = Helpers.parseCurrency(row[9]) || 0;
        const gezahlt = Math.max(0, netto - restNetto);
        const mwstRate = Helpers.parseMwstRate(row[5]);
        const roundedRate = Math.round(mwstRate);
        const tax = gezahlt * (mwstRate / 100);
        const category = row[2] ? row[2].toString().trim() : "";

        if (isIncome) {
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
            if (isEigenbelegSheet) {
                const eigenMapping = CategoryConfig.eigenbelege.mapping[category] || { taxType: "steuerpflichtig" };
                if (eigenMapping.taxType === "steuerfrei") {
                    ustvaData[month].eigenbelege_steuerfrei += gezahlt;
                } else if (eigenMapping.taxType === "eigenbeleg" && eigenMapping.besonderheit === "bewirtung") {
                    ustvaData[month].eigenbelege_steuerpflichtig += gezahlt;
                    ustvaData[month][`vst_${roundedRate}`] += tax * 0.7;
                    ustvaData[month].nicht_abzugsfaehige_vst += tax * 0.3;
                } else {
                    ustvaData[month].eigenbelege_steuerpflichtig += gezahlt;
                    ustvaData[month][`vst_${roundedRate}`] += tax;
                }
            } else {
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

        const revenueData = revenueSheet.getDataRange().getValues();
        const expenseData = expenseSheet.getDataRange().getValues();
        const eigenData = eigenbelegeSheet ? eigenbelegeSheet.getDataRange().getValues() : [];

        const ustvaData = {};
        for (let m = 1; m <= 12; m++) {
            ustvaData[m] = createEmptyUStVA();
        }

        const getMonthFromRow = row => {
            const paymentDate = Helpers.parseDate(row[13]);
            return paymentDate instanceof Date ? paymentDate.getMonth() + 1 : 0;
        };

        if (revenueData.length > 1) {
            revenueData.slice(1).forEach(row => {
                const m = getMonthFromRow(row);
                if (m > 0) processUStVARow(row, ustvaData, true);
            });
        }
        if (expenseData.length > 1) {
            expenseData.slice(1).forEach(row => {
                const m = getMonthFromRow(row);
                if (m > 0) processUStVARow(row, ustvaData, false);
            });
        }
        if (eigenData.length > 1) {
            eigenData.slice(1).forEach(row => {
                const m = getMonthFromRow(row);
                if (m > 0) processUStVARow(row, ustvaData, false, true);
            });
        }

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
        const monthNames = CategoryConfig.common.months;
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
        ustvaSheet.autoResizeColumns(1, outputRows[0].length);
        ss.setActiveSheet(ustvaSheet);
        SpreadsheetApp.getUi().alert("UStVA wurde aktualisiert!");
    };

    return { calculateUStVA };
})();

// =================== Modul: BWACalculator (Final) ===================
const BWACalculator = (() => {
    // Für die BWA werden nur betriebliche Einnahmen & Ausgaben aggregiert – Kapitalbewegungen (Privateinlage, Privatentnahme, Holding Transfers) werden ignoriert.
    // Anpassungen gemäß SKR04/Elster: NEU sind u.a. zusätzliche Kategorien in den Bereichen Einnahmen, Personalkosten, Zinsen sowie Steuern.
    const createEmptyBWA = () => ({
        // Gruppe 1: Betriebserlöse (Einnahmen)
        umsatzerloese: 0,         // Erlöse aus Lieferungen und Leistungen (inkl. Provisionserlöse)
        provisionserloese: 0,
        steuerfreieInlandEinnahmen: 0,
        steuerfreieAuslandEinnahmen: 0,
        sonstigeErtraege: 0,       // Sonstige betriebliche Erträge
        vermietung: 0,            // Erträge aus Vermietung/Verpachtung
        zuschuesse: 0,            // Erträge aus Zuschüssen
        waehrungsgewinne: 0,       // Erträge aus Währungsgewinnen
        anlagenabgaenge: 0,        // Erträge aus Anlagenabgängen
        gesamtErloese: 0,          // Aggregierte Betriebserlöse

        // Gruppe 2: Materialaufwand & Wareneinsatz
        wareneinsatz: 0,
        fremdleistungen: 0,        // Bezogene Leistungen
        rohHilfsBetriebsstoffe: 0,
        gesamtWareneinsatz: 0,

        // Gruppe 3: Betriebsausgaben (Sachkosten)
        bruttoLoehne: 0,
        sozialeAbgaben: 0,
        sonstigePersonalkosten: 0,
        werbungMarketing: 0,
        reisekosten: 0,
        versicherungen: 0,
        telefonInternet: 0,
        buerokosten: 0,
        fortbildungskosten: 0,
        kfzKosten: 0,
        sonstigeAufwendungen: 0,
        gesamtBetriebsausgaben: 0,

        // Gruppe 4: Abschreibungen & Zinsen
        abschreibungenMaschinen: 0,
        abschreibungenBueromaterial: 0,
        abschreibungenImmateriell: 0,
        zinsenBank: 0,
        zinsenGesellschafter: 0,
        leasingkosten: 0,
        gesamtAbschreibungenZinsen: 0,

        // Gruppe 5: Besondere Posten (Kapitalbewegungen)
        eigenkapitalveraenderungen: 0,
        gesellschafterdarlehen: 0,
        ausschuettungen: 0,

        // Gruppe 6: Rückstellungen
        steuerrueckstellungen: 0,
        rueckstellungenSonstige: 0,
        gesamtRueckstellungenTransfers: 0,

        // Gruppe 7: EBIT
        ebit: 0,

        // Gruppe 8: Steuern & Vorsteuer (nur Informationswerte für Elster)
        umsatzsteuer: 0,
        vorsteuer: 0,
        nichtAbzugsfaehigeVSt: 0,
        koerperschaftsteuer: 0,
        solidaritaetszuschlag: 0,
        gewerbesteuer: 0,
        gewerbesteuerRueckstellungen: 0,
        sonstigeSteuerrueckstellungen: 0,
        steuerlast: 0,

        // Gruppe 9: Jahresüberschuss/-fehlbetrag
        gewinnNachSteuern: 0
    });

    const aggregateBWAData = () => {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const revenueSheet = ss.getSheetByName("Einnahmen");
        const expenseSheet = ss.getSheetByName("Ausgaben");
        const eigenbelegeSheet = ss.getSheetByName("Eigenbelege");
        if (!revenueSheet || !expenseSheet) {
            SpreadsheetApp.getUi().alert("Fehlendes Blatt: 'Einnahmen' oder 'Ausgaben'");
            return null;
        }
        const bwaData = {};
        for (let m = 1; m <= 12; m++) {
            bwaData[m] = createEmptyBWA();
        }

        const getMonthFromRow = row => {
            const paymentDate = Helpers.parseDate(row[13]);
            return paymentDate instanceof Date ? paymentDate.getMonth() + 1 : 0;
        };

        // Aggregation der Einnahmen – Kapitalbewegungen (Darlehen, Zinsen) werden ignoriert
        const revenueData = revenueSheet.getDataRange().getValues();
        if (revenueData.length > 1) {
            revenueData.slice(1).forEach(row => {
                const m = getMonthFromRow(row);
                if (!m) return;
                const amount = Helpers.parseCurrency(row[4]);
                const category = row[2] ? row[2].toString().trim() : "";
                if (category === "Gewinnvortrag" || category === "Verlustvortrag" || category === "Gewinnvortrag/Verlustvortrag") {
                    bwaData[m].gewinnVerlustVortrag += amount;
                    return;
                }
                if (category === "Sonstige betriebliche Erträge") {
                    bwaData[m].sonstigeErtraege = (bwaData[m].sonstigeErtraege || 0) + amount;
                    return;
                }
                if (category === "Erträge aus Vermietung/Verpachtung") {
                    bwaData[m].vermietung += amount;
                    return;
                }
                if (category === "Erträge aus Zuschüssen") {
                    bwaData[m].zuschuesse += amount;
                    return;
                }
                if (category === "Erträge aus Währungsgewinnen") {
                    bwaData[m].waehrungsgewinne += amount;
                    return;
                }
                if (category === "Erträge aus Anlagenabgängen") {
                    bwaData[m].anlagenabgaenge += amount;
                    return;
                }
                const mapping = CategoryConfig.einnahmen.bwaMapping[category];
                if (mapping === "umsatzerloese" || mapping === "provisionserloese") {
                    bwaData[m][mapping] += amount;
                } else if (Helpers.parseMwstRate(row[5]) === 0) {
                    bwaData[m].steuerfreieInlandEinnahmen += amount;
                } else {
                    bwaData[m].umsatzerloese += amount;
                }
            });
        }
        for (let m = 1; m <= 12; m++) {
            const d = bwaData[m];
            d.gesamtErloese = d.umsatzerloese + d.provisionserloese + d.steuerfreieInlandEinnahmen + d.steuerfreieAuslandEinnahmen + (d.sonstigeErtraege || 0) + d.vermietung + d.zuschuesse + d.waehrungsgewinne + d.anlagenabgaenge;
        }

        // Aggregation der Ausgaben – Kapitalbewegungen ignorieren
        const expenseData = expenseSheet.getDataRange().getValues();
        if (expenseData.length > 1) {
            expenseData.slice(1).forEach(row => {
                const m = getMonthFromRow(row);
                if (!m) return;
                const amount = Helpers.parseCurrency(row[4]);
                const category = row[2] ? row[2].toString().trim() : "";
                if (category === "Privatentnahme" || category === "Privateinlage" || category === "Holding Transfers") return;
                if (category === "Gewinnvortrag" || category === "Verlustvortrag" || category === "Gewinnvortrag/Verlustvortrag") {
                    bwaData[m].gewinnVerlustVortrag += amount;
                    return;
                }
                if (category === "Bruttolöhne & Gehälter") {
                    bwaData[m].bruttoLoehne += amount;
                    return;
                }
                if (category === "Soziale Abgaben & Arbeitgeberanteile") {
                    bwaData[m].sozialeAbgaben += amount;
                    return;
                }
                if (category === "Sonstige Personalkosten") {
                    bwaData[m].sonstigePersonalkosten += amount;
                    return;
                }
                if (category === "Gewerbesteuerrückstellungen") {
                    bwaData[m].gewerbesteuerRueckstellungen += amount;
                    return;
                }
                if (category === "Telefon & Internet") {
                    bwaData[m].telefonInternet += amount;
                    return;
                }
                if (category === "Bürokosten") {
                    bwaData[m].buerokosten += amount;
                    return;
                }
                if (category === "Fortbildungskosten") {
                    bwaData[m].fortbildungskosten += amount;
                    return;
                }
                const mapping = CategoryConfig.ausgaben.bwaMapping[category];
                switch (mapping) {
                    case "wareneinsatz":
                        bwaData[m].wareneinsatz += amount;
                        break;
                    case "bezogeneLeistungen":
                        bwaData[m].fremdleistungen += amount;
                        break;
                    case "rohHilfsBetriebsstoffe":
                        bwaData[m].rohHilfsBetriebsstoffe += amount;
                        break;
                    case "betriebskosten":
                        bwaData[m].sonstigeAufwendungen += amount;
                        break;
                    case "werbungMarketing":
                        bwaData[m].werbungMarketing += amount;
                        break;
                    case "reisekosten":
                        bwaData[m].reisekosten += amount;
                        break;
                    case "versicherungen":
                        bwaData[m].versicherungen += amount;
                        break;
                    case "kfzKosten":
                        bwaData[m].kfzKosten += amount;
                        break;
                    case "abschreibungenMaschinen":
                        bwaData[m].abschreibungenMaschinen += amount;
                        break;
                    case "abschreibungenBueromaterial":
                        bwaData[m].abschreibungenBueromaterial += amount;
                        break;
                    case "zinsenBank":
                        bwaData[m].zinsenBank += amount;
                        break;
                    case "zinsenGesellschafter":
                        bwaData[m].zinsenGesellschafter += amount;
                        break;
                    case "leasingkosten":
                        bwaData[m].leasingkosten += amount;
                        break;
                    default:
                        bwaData[m].sonstigeAufwendungen += amount;
                }
            });
        }

        // Aggregation der Eigenbelege
        if (eigenbelegeSheet) {
            const eigenData = eigenbelegeSheet.getDataRange().getValues();
            if (eigenData.length > 1) {
                eigenData.slice(1).forEach(row => {
                    const m = getMonthFromRow(row);
                    if (!m) return;
                    const amount = Helpers.parseCurrency(row[4]);
                    const category = row[2] ? row[2].toString().trim() : "";
                    const eigenMapping = CategoryConfig.eigenbelege.mapping[category] || { taxType: "steuerpflichtig" };
                    if (eigenMapping.taxType === "steuerfrei") {
                        bwaData[m].eigenbelegeSteuerfrei += amount;
                    } else {
                        bwaData[m].eigenbelegeSteuerpflichtig += amount;
                    }
                });
            }
        }

        // Gruppensummen berechnen
        for (let m = 1; m <= 12; m++) {
            const d = bwaData[m];
            d.gesamtWareneinsatz = d.wareneinsatz + d.fremdleistungen + d.rohHilfsBetriebsstoffe;
            d.gesamtBetriebsausgaben = d.bruttoLoehne + d.sozialeAbgaben + d.sonstigePersonalkosten + d.werbungMarketing + d.reisekosten + d.versicherungen + d.telefonInternet + d.buerokosten + d.fortbildungskosten + d.kfzKosten + d.sonstigeAufwendungen;
            d.gesamtAbschreibungenZinsen = d.abschreibungenMaschinen + d.abschreibungenBueromaterial + d.abschreibungenImmateriell + d.zinsenBank + d.zinsenGesellschafter + d.leasingkosten;
            d.gesamtBesonderePosten = d.eigenkapitalveraenderungen + d.gesellschafterdarlehen + d.ausschuettungen;
            d.gesamtRueckstellungenTransfers = d.steuerrueckstellungen + d.rueckstellungenSonstige;
            d.ebit = d.gesamtErloese - (d.gesamtWareneinsatz + d.gesamtBetriebsausgaben + d.gesamtAbschreibungenZinsen + d.gesamtBesonderePosten);
            d.steuerlast = (d.umsatzsteuer - d.vorsteuer) + d.koerperschaftsteuer + d.solidaritaetszuschlag + d.gewerbesteuer + d.gewerbesteuerRueckstellungen + d.sonstigeSteuerrueckstellungen;
            d.gewinnNachSteuern = d.ebit - d.steuerlast;
        }
        return bwaData;
    };

    // --- Ausgabe der BWA mit Gruppen-Headern ---
    const calculateBWA = () => {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const bwaData = aggregateBWAData();
        if (!bwaData) return;

        // Positionsdefinition – nur betriebliche Positionen (Kapitalbewegungen werden entfernt)
        // Gruppe 1: Betriebserlöse (Einnahmen)
        const positions = [
            { label: "Erlöse aus Lieferungen und Leistungen", get: md => md.umsatzerloese || 0 },
            { label: "Sonstige betriebliche Erträge", get: md => md.sonstigeErtraege || 0 },
            { label: "Erträge aus Vermietung/Verpachtung", get: md => md.vermietung || 0 },
            { label: "Erträge aus Zuschüssen", get: md => md.zuschuesse || 0 },
            { label: "Erträge aus Währungsgewinnen", get: md => md.waehrungsgewinne || 0 },
            { label: "Erträge aus Anlagenabgängen", get: md => md.anlagenabgaenge || 0 },
            { label: "Betriebserlöse", get: md => md.gesamtErloese || 0 },
            // Gruppe 2: Materialaufwand & Wareneinsatz
            { label: "Wareneinsatz", get: md => md.wareneinsatz || 0 },
            { label: "Bezogene Leistungen", get: md => md.fremdleistungen || 0 },
            { label: "Roh-, Hilfs- & Betriebsstoffe", get: md => md.rohHilfsBetriebsstoffe || 0 },
            { label: "Gesamtkosten Material & Fremdleistungen", get: md => md.gesamtWareneinsatz || 0 },
            // Gruppe 3: Betriebsausgaben (Sachkosten)
            { label: "Bruttolöhne & Gehälter", get: md => md.bruttoLoehne || 0 },
            { label: "Soziale Abgaben & Arbeitgeberanteile", get: md => md.sozialeAbgaben || 0 },
            { label: "Sonstige Personalkosten", get: md => md.sonstigePersonalkosten || 0 },
            { label: "Werbung & Marketing", get: md => md.werbungMarketing || 0 },
            { label: "Reisekosten", get: md => md.reisekosten || 0 },
            { label: "Versicherungen", get: md => md.versicherungen || 0 },
            { label: "Telefon & Internet", get: md => md.telefonInternet || 0 },
            { label: "Bürokosten", get: md => md.buerokosten || 0 },
            { label: "Fortbildungskosten", get: md => md.fortbildungskosten || 0 },
            { label: "Kfz-Kosten", get: md => md.kfzKosten || 0 },
            { label: "Sonstige betriebliche Aufwendungen", get: md => md.sonstigeAufwendungen || 0 },
            // Gruppe 4: Abschreibungen & Zinsen
            { label: "Abschreibungen Maschinen", get: md => md.abschreibungenMaschinen || 0 },
            { label: "Abschreibungen Büroausstattung", get: md => md.abschreibungenBueromaterial || 0 },
            { label: "Abschreibungen immaterielle Wirtschaftsgüter", get: md => md.abschreibungenImmateriell || 0 },
            { label: "Zinsen auf Bankdarlehen", get: md => md.zinsenBank || 0 },
            { label: "Zinsen auf Gesellschafterdarlehen", get: md => md.zinsenGesellschafter || 0 },
            { label: "Leasingkosten", get: md => md.leasingkosten || 0 },
            { label: "Gesamt Abschreibungen & Zinsen", get: md => md.gesamtAbschreibungenZinsen || 0 },
            // Gruppe 5: Besondere Posten
            { label: "Eigenkapitalveränderungen", get: md => md.eigenkapitalveraenderungen || 0 },
            { label: "Gesellschafterdarlehen", get: md => md.gesellschafterdarlehen || 0 },
            { label: "Ausschüttungen an Gesellschafter", get: md => md.ausschuettungen || 0 },
            // Gruppe 6: Rückstellungen
            { label: "Steuerrückstellungen", get: md => md.steuerrueckstellungen || 0 },
            { label: "Rückstellungen sonstige", get: md => md.rueckstellungenSonstige || 0 },
            // Gruppe 7: EBIT
            { label: "Betriebsergebnis vor Steuern (EBIT)", get: md => md.ebit || 0 },
            // Gruppe 8: Steuern & Vorsteuer
            { label: "Umsatzsteuer (abzuführen)", get: md => md.umsatzsteuer || 0 },
            { label: "Vorsteuer", get: md => md.vorsteuer || 0 },
            { label: "Nicht abzugsfähige VSt (Bewirtung)", get: md => md.nichtAbzugsfaehigeVSt || 0 },
            { label: "Körperschaftsteuer", get: md => md.koerperschaftsteuer || 0 },
            { label: "Solidaritätszuschlag", get: md => md.solidaritaetszuschlag || 0 },
            { label: "Gewerbesteuer", get: md => md.gewerbesteuer || 0 },
            { label: "Gesamtsteueraufwand", get: md => md.steuerlast || 0 },
            // Gruppe 9: Jahresüberschuss/-fehlbetrag
            { label: "Jahresüberschuss/-fehlbetrag", get: md => md.gewinnNachSteuern || 0 }
        ];

        const groupHeaders = [
            { header: "1️⃣ Betriebserlöse (Einnahmen)", count: 7 },
            { header: "2️⃣ Materialaufwand & Wareneinsatz", count: 4 },
            { header: "3️⃣ Betriebsausgaben (Sachkosten)", count: 10 },
            { header: "4️⃣ Abschreibungen & Zinsen", count: 7 },
            { header: "5️⃣ Besondere Posten", count: 3 },
            { header: "6️⃣ Rückstellungen", count: 2 },
            { header: "7️⃣ Betriebsergebnis vor Steuern (EBIT)", count: 1 },
            { header: "8️⃣ Steuern & Vorsteuer", count: 7 },
            { header: "9️⃣ Jahresüberschuss/-fehlbetrag", count: 1 }
        ];

        const headerRow = [
            "Kategorie",
            "Januar (€)",
            "Februar (€)",
            "März (€)",
            "Q1 (€)",
            "April (€)",
            "Mai (€)",
            "Juni (€)",
            "Q2 (€)",
            "Juli (€)",
            "August (€)",
            "September (€)",
            "Q3 (€)",
            "Oktober (€)",
            "November (€)",
            "Dezember (€)",
            "Q4 (€)",
            "Jahr (€)"
        ];
        let outputRows = [];
        outputRows.push(headerRow);

        const buildOutputRow = pos => {
            const row = [pos.label];
            let quarterly = [0, 0, 0, 0];
            let yearly = 0;
            for (let m = 1; m <= 12; m++) {
                const val = pos.get(bwaData[m]) || 0;
                row.push(val);
                yearly += val;
                quarterly[Math.floor((m - 1) / 3)] += val;
            }
            row.splice(4, 0, quarterly[0]);
            row.splice(8, 0, quarterly[1]);
            row.splice(12, 0, quarterly[2]);
            row.splice(16, 0, quarterly[3]);
            row.push(yearly);
            return row;
        };

        let posIndex = 0;
        groupHeaders.forEach(group => {
            let groupRow = Array(headerRow.length).fill("");
            groupRow[0] = group.header;
            outputRows.push(groupRow);
            for (let i = 0; i < group.count; i++) {
                outputRows.push(buildOutputRow(positions[posIndex]));
                posIndex++;
            }
        });

        const bwaSheet = ss.getSheetByName("BWA") || ss.insertSheet("BWA");
        bwaSheet.clearContents();
        bwaSheet.getRange(1, 1, outputRows.length, outputRows[0].length).setValues(outputRows);
        bwaSheet.autoResizeColumns(1, outputRows[0].length);
        SpreadsheetApp.getUi().alert("BWA wurde aktualisiert!");
    };

    return { calculateBWA };
})();

// =================== Globale Funktionen ===================
const onOpen = () => {
    SpreadsheetApp.getUi()
        .createMenu("📂 Buchhaltung")
        .addItem("📥 Dateien importieren", "importDriveFiles")
        .addItem("🔄 Refresh Active Sheet", "refreshSheet")
        .addItem("📊 UStVA berechnen", "calculateUStVA")
        .addItem("📈 BWA berechnen", "calculateBWA")
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
const calculateBWA = () => { RefreshModule.refreshAllSheets(); BWACalculator.calculateBWA(); };
const importDriveFiles = () => { ImportModule.importDriveFiles(); RefreshModule.refreshAllSheets(); };
