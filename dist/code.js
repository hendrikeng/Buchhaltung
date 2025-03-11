/* Bundled Code for Google Apps Script */
'use strict';

// src/helpers.js

const Helpers = {
    parseDate(value) {
        if (value instanceof Date) return isNaN(value.getTime()) ? null : value;
        if (typeof value === "string") {
            const d = new Date(value);
            return isNaN(d.getTime()) ? null : d;
        }
        return null;
    },

    parseCurrency(value) {
        if (typeof value === "number") return value;
        const str = value.toString().replace(/[^\d,.-]/g, "").replace(",", ".");
        const num = parseFloat(str);
        return isNaN(num) ? 0 : num;
    },

    parseMwstRate(value) {
        // Hier kannst du zum Beispiel einen Default-Wert festlegen, falls config.tax.defaultMwst nicht verfügbar ist.
        if (typeof value === "number") return value < 1 ? value * 100 : value;
        const rate = parseFloat(value.toString().replace("%", "").replace(",", ".").trim());
        return isNaN(rate) ? 19 : rate;
    },

    getFolderByName(parent, name) {
        const folderIter = parent.getFoldersByName(name);
        return folderIter.hasNext() ? folderIter.next() : null;
    },

    extractDateFromFilename(filename) {
        const nameWithoutExtension = filename.replace(/\.[^/.]+$/, "");

        // Verschiedene Formate erkennen

        // Format: RE-YYYY-MM-DD oder ähnliches mit Trennzeichen
        let match = nameWithoutExtension.match(/[^0-9](\d{4}[-_.\/]\d{2}[-_.\/]\d{2})[^0-9]/);
        if (match?.[1]) {
            const dateParts = match[1].split(/[-_.\/]/);
            if (dateParts.length === 3) {
                const [year, month, day] = dateParts;
                return `${day}.${month}.${year}`;
            }
        }

        // Format: YYYY-MM-DD am Anfang oder Ende
        match = nameWithoutExtension.match(/(^|[^0-9])(\d{4}[-_.\/]\d{2}[-_.\/]\d{2})($|[^0-9])/);
        if (match?.[2]) {
            const dateParts = match[2].split(/[-_.\/]/);
            if (dateParts.length === 3) {
                const [year, month, day] = dateParts;
                return `${day}.${month}.${year}`;
            }
        }

        // Format: DD.MM.YYYY im Dateinamen
        match = nameWithoutExtension.match(/(\d{2}[.]\d{2}[.]\d{4})/);
        if (match?.[1]) {
            return match[1];
        }

        return "";
    },

    setConditionalFormattingForColumn(sheet, column, conditions) {
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
    },

    getMonthFromRow(row, colIndex = 13) {
        const d = Helpers.parseDate(row[colIndex]);

        // Auf das Jahr aus der Konfiguration prüfen oder das aktuelle Jahr verwenden
        const targetYear = config.tax?.year || new Date().getFullYear();

        // Wenn kein Datum oder das Jahr nicht übereinstimmt
        if (!d || d.getFullYear() !== targetYear) return 0;

        return d.getMonth() + 1;
    }
};

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
        const ui = SpreadsheetApp.getUi();

        try {
            const revenueMain = ss.getSheetByName("Einnahmen");
            const expenseMain = ss.getSheetByName("Ausgaben");

            if (!revenueMain || !expenseMain) {
                ui.alert("Fehler: Die Sheets 'Einnahmen' oder 'Ausgaben' existieren nicht!");
                return;
            }

            const revenue = ss.getSheetByName("Rechnungen Einnahmen") || ss.insertSheet("Rechnungen Einnahmen");
            const expense = ss.getSheetByName("Rechnungen Ausgaben") || ss.insertSheet("Rechnungen Ausgaben");
            const history = ss.getSheetByName("Änderungshistorie") || ss.insertSheet("Änderungshistorie");

            // Header-Zeilen initialisieren falls nötig
            if (revenue.getLastRow() === 0)
                revenue.appendRow(["Dateiname", "Link zur Datei", "Rechnungsnummer"]);
            if (expense.getLastRow() === 0)
                expense.appendRow(["Dateiname", "Link zur Datei", "Rechnungsnummer"]);
            if (history.getLastRow() === 0)
                history.appendRow(["Datum", "Rechnungstyp", "Dateiname", "Link zur Datei"]);

            // Auf übergeordneten Ordner zugreifen
            let parentFolder;
            try {
                const file = DriveApp.getFileById(ss.getId());
                const parents = file.getParents();
                parentFolder = parents.hasNext() ? parents.next() : null;
                if (!parentFolder) {
                    ui.alert("Fehler: Kein übergeordneter Ordner gefunden.");
                    return;
                }
            } catch (e) {
                ui.alert("Fehler beim Zugriff auf Google Drive: " + e.toString());
                return;
            }

            // Unterordner für Einnahmen und Ausgaben finden
            let revenueFolder, expenseFolder;

            try {
                revenueFolder = Helpers.getFolderByName(parentFolder, "Einnahmen");
                if (!revenueFolder) {
                    const createFolder = ui.alert(
                        "Der Ordner 'Einnahmen' existiert nicht. Soll er erstellt werden?",
                        ui.ButtonSet.YES_NO
                    );
                    if (createFolder === ui.Button.YES) {
                        revenueFolder = parentFolder.createFolder("Einnahmen");
                    }
                }
            } catch (e) {
                ui.alert("Fehler beim Zugriff auf den Einnahmen-Ordner: " + e.toString());
            }

            try {
                expenseFolder = Helpers.getFolderByName(parentFolder, "Ausgaben");
                if (!expenseFolder) {
                    const createFolder = ui.alert(
                        "Der Ordner 'Ausgaben' existiert nicht. Soll er erstellt werden?",
                        ui.ButtonSet.YES_NO
                    );
                    if (createFolder === ui.Button.YES) {
                        expenseFolder = parentFolder.createFolder("Ausgaben");
                    }
                }
            } catch (e) {
                ui.alert("Fehler beim Zugriff auf den Ausgaben-Ordner: " + e.toString());
            }

            // Import durchführen wenn Ordner existieren
            let importCount = 0;

            if (revenueFolder) {
                try {
                    importFilesFromFolder(revenueFolder, revenue, revenueMain, "Einnahme", history);
                    importCount++;
                } catch (e) {
                    ui.alert("Fehler beim Import der Einnahmen: " + e.toString());
                }
            }

            if (expenseFolder) {
                try {
                    importFilesFromFolder(expenseFolder, expense, expenseMain, "Ausgabe", history);
                    importCount++;
                } catch (e) {
                    ui.alert("Fehler beim Import der Ausgaben: " + e.toString());
                }
            }

            if (importCount === 0) {
                ui.alert("Es wurden keine Dateien importiert. Bitte überprüfe die Ordnerstruktur.");
            } else {
                ui.alert("Import abgeschlossen.");
            }
        } catch (e) {
            ui.alert("Ein unerwarteter Fehler ist aufgetreten: " + e.toString());
        }
    };

    return {importDriveFiles};
})();

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
            rules.forEach(({check, message}) => {
                if (check(row)) warnings.push(`Zeile ${idx}: ${message}`);
            });
        };
        const baseRules = [
            {check: r => isEmpty(r[0]), message: "Rechnungsdatum fehlt."},
            {check: r => isEmpty(r[1]), message: "Rechnungsnummer fehlt."},
            {check: r => isEmpty(r[2]), message: "Kategorie fehlt."},
            {check: r => isEmpty(r[3]), message: "Kunde fehlt."},
            {check: r => isInvalidNumber(r[4]), message: "Nettobetrag fehlt oder ungültig."},
            {
                check: r => {
                    const mwstStr = r[5] == null ? "" : r[5].toString().trim();
                    if (isEmpty(mwstStr)) return false; // Wird schon durch andere Regel geprüft
                    const mwst = parseFloat(mwstStr.replace("%", "").replace(",", "."));
                    if (isNaN(mwst)) return false; // Wird schon durch andere Regel geprüft
                    // Prüfe auf erlaubte MwSt-Sätze in Deutschland
                    return ![0, 7, 19].includes(Math.round(mwst));
                },
                message: "Ungültiger MwSt-Satz. Erlaubt sind: 0%, 7%, 19%."
            }
        ];
        const status = row[11] ? row[11].toString().trim().toLowerCase() : "";
        const paymentRules = status === "offen"
            ? [
                {check: r => !isEmpty(r[12]), message: 'Zahlungsart darf bei offener Zahlung nicht gesetzt sein.'},
                {check: r => !isEmpty(r[13]), message: 'Zahlungsdatum darf bei offener Zahlung nicht gesetzt sein.'}
            ]
            : [
                {
                    check: r => isEmpty(r[12]),
                    message: 'Zahlungsart muss bei bezahlter/teilbezahlter Zahlung gesetzt sein.'
                },
                {
                    check: r => isEmpty(r[13]),
                    message: 'Zahlungsdatum muss bei bezahlter/teilbezahlter Zahlung gesetzt sein.'
                },
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
            rules.forEach(({check, message}) => {
                if (check(row)) warnings.push(`Zeile ${idx}: ${message}`);
            });
        };
        const headerFooterRules = [
            {check: r => isEmpty(r[0]), message: "Buchungsdatum fehlt."},
            {check: r => isEmpty(r[1]), message: "Buchungstext fehlt."},
            {
                check: r => !isEmpty(r[2]) || !isNaN(parseFloat(r[2].toString().trim())),
                message: "Betrag darf nicht gesetzt sein."
            },
            {check: r => isEmpty(r[3]) || isInvalidNumber(r[3]), message: "Saldo fehlt oder ungültig."},
            {check: r => !isEmpty(r[4]), message: "Typ darf nicht gesetzt sein."},
            {check: r => !isEmpty(r[5]), message: "Kategorie darf nicht gesetzt sein."},
            {check: r => !isEmpty(r[6]), message: "Konto (Soll) darf nicht gesetzt sein."},
            {check: r => !isEmpty(r[7]), message: "Gegenkonto (Haben) darf nicht gesetzt sein."}
        ];
        const dataRowRules = [
            {check: r => isEmpty(r[0]), message: "Buchungsdatum fehlt."},
            {check: r => isEmpty(r[1]), message: "Buchungstext fehlt."},
            {check: r => isEmpty(r[2]) || isInvalidNumber(r[2]), message: "Betrag fehlt oder ungültig."},
            {check: r => isEmpty(r[3]) || isInvalidNumber(r[3]), message: "Saldo fehlt oder ungültig."},
            {check: r => isEmpty(r[4]), message: "Typ fehlt."},
            {check: r => isEmpty(r[5]), message: "Kategorie fehlt."},
            {check: r => isEmpty(r[6]), message: "Konto (Soll) fehlt."},
            {check: r => isEmpty(r[7]), message: "Gegenkonto (Haben) fehlt."}
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

    return {validateDropdown, validateRevenueAndExpenses, validateBanking, validateAllSheets};
})();

/**
 * Konfiguration für die Buchhaltungsanwendung
 * Unterstützt die Buchhaltung für Holding und operative GmbH nach SKR04
 */
const config$1 = {
    // Allgemeine Einstellungen
    common: {
        paymentType: ["Überweisung", "Bar", "Kreditkarte", "Paypal", "Lastschrift"],
        months: ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember"],
        currentYear: new Date().getFullYear()
    },

    // Steuerliche Einstellungen
    tax: {
        defaultMwst: 19,
        allowedMwst: [0, 7, 19],  // Erlaubte MwSt-Sätze
        stammkapital: 25000,
        year: 2021,  // Geschäftsjahr
        isHolding: false, // true bei Holding

        // Holding-spezifische Steuersätze
        holding: {
            gewerbesteuer: 16.45,  // Angepasst an lokalen Hebesatz
            koerperschaftsteuer: 15,
            solidaritaetszuschlag: 5.5,
            gewinnUebertragSteuerfrei: 95,  // % der Beteiligungserträge steuerfrei
            gewinnUebertragSteuerpflichtig: 5  // % der Beteiligungserträge steuerpflichtig
        },

        // Operative GmbH Steuersätze
        operative: {
            gewerbesteuer: 16.45,  // Angepasst an lokalen Hebesatz
            koerperschaftsteuer: 15,
            solidaritaetszuschlag: 5.5,
            gewinnUebertragSteuerfrei: 0,
            gewinnUebertragSteuerpflichtig: 100
        }
    },

    // Einnahmen-Konfiguration
    einnahmen: {
        // Kategorien mit Steuertyp
        categories: {
            "Erlöse aus Lieferungen und Leistungen": {taxType: "steuerpflichtig"},
            "Provisionserlöse": {taxType: "steuerpflichtig"},
            "Sonstige betriebliche Erträge": {taxType: "steuerpflichtig"},
            "Erträge aus Vermietung/Verpachtung": {taxType: "steuerfrei_inland"},
            "Erträge aus Zuschüssen": {taxType: "steuerpflichtig"},
            "Erträge aus Währungsgewinnen": {taxType: "steuerpflichtig"},
            "Erträge aus Anlagenabgängen": {taxType: "steuerpflichtig"},
            "Darlehen": {taxType: "steuerfrei_inland"},
            "Zinsen": {taxType: "steuerfrei_inland"},
            "Gewinnvortrag": {taxType: "steuerfrei_inland"},
            "Verlustvortrag": {taxType: "steuerfrei_inland"}
        },

        // SKR04-konforme Konten-Zuordnung
        kontoMapping: {
            "Erlöse aus Lieferungen und Leistungen": {soll: "1200", gegen: "4400", mwst: "1776"},
            "Provisionserlöse": {soll: "1200", gegen: "4120", mwst: "1776"},
            "Sonstige betriebliche Erträge": {soll: "1200", gegen: "4830", mwst: "1776"},
            "Erträge aus Vermietung/Verpachtung": {soll: "1200", gegen: "4180"},
            "Erträge aus Zuschüssen": {soll: "1200", gegen: "4190", mwst: "1776"},
            "Erträge aus Währungsgewinnen": {soll: "1200", gegen: "4840"},
            "Erträge aus Anlagenabgängen": {soll: "1200", gegen: "4855"},
            "Darlehen": {soll: "1200", gegen: "3100"},
            "Zinsen": {soll: "1200", gegen: "4130"},
            "Gewinnvortrag": {soll: "1200", gegen: "2970"},
            "Verlustvortrag": {soll: "1200", gegen: "2978"}
        },

        // BWA-Mapping für Einnahmen-Kategorien
        bwaMapping: {
            "Erlöse aus Lieferungen und Leistungen": "umsatzerloese",
            "Provisionserlöse": "provisionserloese",
            "Sonstige betriebliche Erträge": "sonstigeErtraege",
            "Erträge aus Vermietung/Verpachtung": "vermietung",
            "Erträge aus Zuschüssen": "zuschuesse",
            "Erträge aus Währungsgewinnen": "waehrungsgewinne",
            "Erträge aus Anlagenabgängen": "anlagenabgaenge",
            "Darlehen": "sonstigeErtraege",
            "Zinsen": "sonstigeErtraege"
        }
    },

    // Ausgaben-Konfiguration
    ausgaben: {
        // Kategorien mit Steuertyp
        categories: {
            // Materialaufwand & Wareneinsatz
            "Wareneinsatz": {taxType: "steuerpflichtig", group: "material"},
            "Bezogene Leistungen": {taxType: "steuerpflichtig", group: "material"},
            "Roh-, Hilfs- & Betriebsstoffe": {taxType: "steuerpflichtig", group: "material"},

            // Personalkosten
            "Bruttolöhne & Gehälter": {taxType: "steuerfrei_inland", group: "personal"},
            "Soziale Abgaben & Arbeitgeberanteile": {taxType: "steuerfrei_inland", group: "personal"},
            "Sonstige Personalkosten": {taxType: "steuerpflichtig", group: "personal"},

            // Raumkosten
            "Miete": {taxType: "steuerfrei_inland", group: "raum"},
            "Nebenkosten": {taxType: "steuerpflichtig", group: "raum"},

            // Betriebskosten
            "Betriebskosten": {taxType: "steuerpflichtig", group: "betrieb"},
            "Marketing & Werbung": {taxType: "steuerpflichtig", group: "betrieb"},
            "Reisekosten": {taxType: "steuerpflichtig", group: "betrieb"},
            "Versicherungen": {taxType: "steuerfrei_inland", group: "betrieb"},
            "Porto": {taxType: "steuerfrei_inland", group: "betrieb"},
            "Google Ads": {taxType: "steuerfrei_ausland", group: "betrieb"},
            "AWS": {taxType: "steuerfrei_ausland", group: "betrieb"},
            "Facebook Ads": {taxType: "steuerfrei_ausland", group: "betrieb"},
            "Bewirtung": {taxType: "steuerpflichtig", group: "betrieb", besonderheit: "bewirtung"},
            "Telefon & Internet": {taxType: "steuerpflichtig", group: "betrieb"},
            "Bürokosten": {taxType: "steuerpflichtig", group: "betrieb"},
            "Fortbildungskosten": {taxType: "steuerpflichtig", group: "betrieb"},

            // Abschreibungen & Zinsen
            "Abschreibungen Maschinen": {taxType: "steuerpflichtig", group: "abschreibung"},
            "Abschreibungen Büroausstattung": {taxType: "steuerpflichtig", group: "abschreibung"},
            "Abschreibungen immaterielle Wirtschaftsgüter": {taxType: "steuerpflichtig", group: "abschreibung"},
            "Zinsen auf Bankdarlehen": {taxType: "steuerpflichtig", group: "zinsen"},
            "Zinsen auf Gesellschafterdarlehen": {taxType: "steuerpflichtig", group: "zinsen"},
            "Leasingkosten": {taxType: "steuerpflichtig", group: "abschreibung"},

            // Steuern & Rückstellungen
            "Gewerbesteuerrückstellungen": {taxType: "steuerfrei_inland", group: "steuer"},
            "Körperschaftsteuer": {taxType: "steuerfrei_inland", group: "steuer"},
            "Solidaritätszuschlag": {taxType: "steuerfrei_inland", group: "steuer"},
            "Sonstige Steuerrückstellungen": {taxType: "steuerfrei_inland", group: "steuer"},

            // Sonstige Aufwendungen
            "Sonstige betriebliche Aufwendungen": {taxType: "steuerpflichtig", group: "sonstige"}
        },

        // SKR04-konforme Konten-Zuordnung
        kontoMapping: {
            "Wareneinsatz": {soll: "5000", gegen: "1200", vorsteuer: "1576"},
            "Bezogene Leistungen": {soll: "5300", gegen: "1200", vorsteuer: "1576"},
            "Roh-, Hilfs- & Betriebsstoffe": {soll: "5400", gegen: "1200", vorsteuer: "1576"},
            "Bruttolöhne & Gehälter": {soll: "6000", gegen: "1200"},
            "Soziale Abgaben & Arbeitgeberanteile": {soll: "6010", gegen: "1200"},
            "Sonstige Personalkosten": {soll: "6020", gegen: "1200", vorsteuer: "1576"},
            "Miete": {soll: "6310", gegen: "1200"},
            "Nebenkosten": {soll: "6320", gegen: "1200", vorsteuer: "1576"},
            "Betriebskosten": {soll: "6300", gegen: "1200", vorsteuer: "1576"},
            "Marketing & Werbung": {soll: "6600", gegen: "1200", vorsteuer: "1576"},
            "Reisekosten": {soll: "6650", gegen: "1200", vorsteuer: "1576"},
            "Versicherungen": {soll: "6400", gegen: "1200"},
            "Porto": {soll: "6810", gegen: "1200"},
            "Google Ads": {soll: "6600", gegen: "1200"},
            "AWS": {soll: "6500", gegen: "1200"},
            "Facebook Ads": {soll: "6600", gegen: "1200"},
            "Bewirtung": {soll: "6670", gegen: "1200", vorsteuer: "1576"},
            "Telefon & Internet": {soll: "6805", gegen: "1200", vorsteuer: "1576"},
            "Bürokosten": {soll: "6815", gegen: "1200", vorsteuer: "1576"},
            "Fortbildungskosten": {soll: "6830", gegen: "1200", vorsteuer: "1576"},
            "Abschreibungen Maschinen": {soll: "6200", gegen: "1200"},
            "Abschreibungen Büroausstattung": {soll: "6210", gegen: "1200"},
            "Abschreibungen immaterielle Wirtschaftsgüter": {soll: "6220", gegen: "1200"},
            "Zinsen auf Bankdarlehen": {soll: "7300", gegen: "1200"},
            "Zinsen auf Gesellschafterdarlehen": {soll: "7310", gegen: "1200"},
            "Leasingkosten": {soll: "6240", gegen: "1200", vorsteuer: "1576"},
            "Gewerbesteuerrückstellungen": {soll: "7610", gegen: "1200"},
            "Körperschaftsteuer": {soll: "7600", gegen: "1200"},
            "Solidaritätszuschlag": {soll: "7620", gegen: "1200"},
            "Sonstige Steuerrückstellungen": {soll: "7630", gegen: "1200"},
            "Sonstige betriebliche Aufwendungen": {soll: "6800", gegen: "1200", vorsteuer: "1576"}
        },

        // BWA-Mapping für Ausgaben-Kategorien
        bwaMapping: {
            "Wareneinsatz": "wareneinsatz",
            "Bezogene Leistungen": "fremdleistungen",
            "Roh-, Hilfs- & Betriebsstoffe": "rohHilfsBetriebsstoffe",
            "Betriebskosten": "sonstigeAufwendungen",
            "Marketing & Werbung": "werbungMarketing",
            "Reisekosten": "reisekosten",
            "Bruttolöhne & Gehälter": "bruttoLoehne",
            "Soziale Abgaben & Arbeitgeberanteile": "sozialeAbgaben",
            "Sonstige Personalkosten": "sonstigePersonalkosten",
            "Sonstige betriebliche Aufwendungen": "sonstigeAufwendungen",
            "Miete": "mieteNebenkosten",
            "Nebenkosten": "mieteNebenkosten",
            "Versicherungen": "versicherungen",
            "Porto": "buerokosten",
            "Telefon & Internet": "telefonInternet",
            "Bürokosten": "buerokosten",
            "Fortbildungskosten": "fortbildungskosten",
            "Abschreibungen Maschinen": "abschreibungenMaschinen",
            "Abschreibungen Büroausstattung": "abschreibungenBueromaterial",
            "Abschreibungen immaterielle Wirtschaftsgüter": "abschreibungenImmateriell",
            "Zinsen auf Bankdarlehen": "zinsenBank",
            "Zinsen auf Gesellschafterdarlehen": "zinsenGesellschafter",
            "Leasingkosten": "leasingkosten",
            "Google Ads": "werbungMarketing",
            "AWS": "sonstigeAufwendungen",
            "Facebook Ads": "werbungMarketing",
            "Bewirtung": "sonstigeAufwendungen",
            "Gewerbesteuerrückstellungen": "gewerbesteuerRueckstellungen",
            "Körperschaftsteuer": "koerperschaftsteuer",
            "Solidaritätszuschlag": "solidaritaetszuschlag",
            "Sonstige Steuerrückstellungen": "steuerrueckstellungen"
        }
    },

    // Bankbewegungen-Konfiguration
    bank: {
        // Kombinierte Liste aller Kategorien (automatisch generiert)
        category: [], // Wird dynamisch befüllt

        // Typen von Bankbewegungen
        type: ["Einnahme", "Ausgabe", "Interne Buchung"],

        // Standard-Bankkonto
        defaultAccount: "1200"
    },

    // Gesellschafterkonto-Konfiguration
    gesellschafterkonto: {
        category: ["Gesellschafterdarlehen", "Ausschüttungen", "Kapitalrückführung", "Privatentnahme", "Privateinlage"],
        shareholder: ["Christopher Giebel", "Hendrik Werner"]
    },

    // Eigenbelege-Konfiguration
    eigenbelege: {
        category: ["Kleidung", "Trinkgeld", "Private Vorauslage", "Bürokosten", "Reisekosten", "Bewirtung", "Sonstiges"],
        mapping: {
            "Kleidung": {taxType: "steuerpflichtig"},
            "Trinkgeld": {taxType: "steuerfrei"},
            "Private Vorauslage": {taxType: "steuerfrei"},
            "Bürokosten": {taxType: "steuerpflichtig"},
            "Reisekosten": {taxType: "steuerpflichtig"},
            "Bewirtung": {taxType: "eigenbeleg", besonderheit: "bewirtung"},
            "Sonstiges": {taxType: "steuerpflichtig"}
        },
        status: ["Offen", "Erstattet", "Gebucht"]
    },

    // Holding Transfers-Konfiguration
    holdingTransfers: {
        category: ["Gewinnübertrag", "Kapitalrückführung"]
    },

    // Kontenplan SKR04 (Auszug der wichtigsten Konten)
    kontenplan: {
        // Bestandskonten (Klasse 0-3)
        "0000": "Eröffnungsbilanz",
        "0100": "Immaterielle Vermögensgegenstände",
        "0400": "Grundstücke ohne Bauten",
        "0500": "Bauten auf eigenen Grundstücken",
        "0650": "Technische Anlagen und Maschinen",
        "0700": "Andere Anlagen, Betriebs- und Geschäftsausstattung",
        "1200": "Bank",
        "1210": "Kasse",
        "1300": "Forderungen aus Lieferungen und Leistungen",
        "1576": "Vorsteuer 19%",
        "1577": "Vorsteuer 7%",
        "1590": "Umsatzsteuer-Vorauszahlungen",
        "1600": "Forderungen gegen Gesellschafter",
        "1776": "Umsatzsteuer 19%",
        "1777": "Umsatzsteuer 7%",
        "1790": "Umsatzsteuer-Vorauszahlungen",
        "2000": "Gezeichnetes Kapital",
        "2100": "Kapitalrücklage",
        "2970": "Gewinnvortrag vor Verwendung",
        "2978": "Verlustvortrag vor Verwendung",
        "3100": "Darlehen",
        "3150": "Verbindlichkeiten gegenüber Kreditinstituten",
        "3200": "Verbindlichkeiten aus Lieferungen und Leistungen",
        "3300": "Verbindlichkeiten gegenüber Gesellschaftern",

        // Erlöskonten (Klasse 4)
        "4000": "Umsatzerlöse",
        "4120": "Provisionserlöse",
        "4130": "Zinsertrag",
        "4180": "Erträge aus Vermietung und Verpachtung",
        "4190": "Erträge aus Zuschüssen und Zulagen",
        "4400": "Erlöse aus Lieferungen und Leistungen",
        "4830": "Sonstige betriebliche Erträge",
        "4840": "Erträge aus Währungsumrechnung",
        "4855": "Erträge aus Anlagenabgang",

        // Aufwandskonten (Klasse 5-7)
        "5000": "Wareneinsatz",
        "5300": "Bezogene Leistungen",
        "5400": "Roh-, Hilfs- und Betriebsstoffe",
        "6000": "Löhne und Gehälter",
        "6010": "Gesetzliche soziale Aufwendungen",
        "6020": "Sonstige Personalkosten",
        "6200": "Abschreibungen auf Sachanlagen",
        "6210": "Abschreibungen auf Büroausstattung",
        "6220": "Abschreibungen auf immaterielle Vermögensgegenstände",
        "6240": "Leasingkosten",
        "6300": "Betriebskosten",
        "6310": "Miete",
        "6320": "Nebenkosten",
        "6400": "Versicherungen",
        "6500": "IT-Kosten",
        "6600": "Werbekosten",
        "6650": "Reisekosten",
        "6670": "Bewirtungskosten",
        "6800": "Sonstige betriebliche Aufwendungen",
        "6805": "Telefon und Internet",
        "6810": "Porto",
        "6815": "Bürobedarf",
        "6830": "Fortbildungskosten",
        "7300": "Zinsaufwendungen für Bankdarlehen",
        "7310": "Zinsaufwendungen für Gesellschafterdarlehen",
        "7600": "Körperschaftsteuer",
        "7610": "Gewerbesteuer",
        "7620": "Solidaritätszuschlag",
        "7630": "Sonstige Steuern",

        // Abschlusskonten (Klasse 8-9)
        "8000": "Eröffnungsbilanzkonto",
        "8400": "Abschlussbilanzkonto",
        "8999": "Gewinn- und Verlustkonto"
    }
};

// Bankkategorien dynamisch aus den Einnahmen- und Ausgaben-Kategorien befüllen
config$1.bank.category = [
    ...Object.keys(config$1.einnahmen.categories),
    ...Object.keys(config$1.ausgaben.categories),
    ...config$1.gesellschafterkonto.category,
    ...config$1.holdingTransfers.category,
    ...config$1.eigenbelege.category
];

// Duplikate aus den Kategorien entfernen
config$1.bank.category = [...new Set(config$1.bank.category)];

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
            Validator.validateDropdown(sheet, 2, 3, lastRow - 1, 1, Object.keys(config$1.einnahmen.categories));
        if (name === "Ausgaben")
            Validator.validateDropdown(sheet, 2, 3, lastRow - 1, 1, Object.keys(config$1.ausgaben.categories));
        if (name === "Eigenbelege")
            Validator.validateDropdown(sheet, 2, 3, lastRow - 1, 1, config$1.eigenbelege.category);
        Validator.validateDropdown(sheet, 2, 13, lastRow - 1, 1, config$1.common.paymentType);
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

        Validator.validateDropdown(sheet, firstDataRow, 5, numDataRows, 1, config$1.bank.type);
        Validator.validateDropdown(sheet, firstDataRow, 6, numDataRows, 1, config$1.bank.category);
        const allowedKontoSoll = Object.values(config$1.einnahmen.kontoMapping)
            .concat(Object.values(config$1.ausgaben.kontoMapping))
            .map(m => m.soll);
        const allowedGegenkonto = Object.values(config$1.einnahmen.kontoMapping)
            .concat(Object.values(config$1.ausgaben.kontoMapping))
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
                ? config$1.einnahmen.kontoMapping[category]
                : type === "Ausgabe"
                    ? config$1.ausgaben.kontoMapping[category]
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

    // Hier entfällt die eigene getMonthFromRow – stattdessen nutzen wir Helpers.getMonthFromRow:
    const processUStVARow = (row, data, isIncome, isEigen = false) => {
        const paymentDate = Helpers.parseDate(row[13]);
        if (!paymentDate || paymentDate > new Date()) return;
        const month = Helpers.getMonthFromRow(row);
        if (!month) return;
        const netto = Helpers.parseCurrency(row[4]);
        const restNetto = Helpers.parseCurrency(row[9]) || 0;
        const gezahlt = netto - restNetto; // Negative Werte bleiben erhalten
        const mwstRate = Helpers.parseMwstRate(row[5]);
        const roundedRate = Math.round(mwstRate);
        const tax = gezahlt * (mwstRate / 100);
        const category = row[2]?.toString().trim() || "";

        if (isIncome) {
            const catCfg = config$1.einnahmen.categories[category] ?? {};
            const taxType = catCfg.taxType ?? "steuerpflichtig";
            if (taxType === "steuerfrei_inland") {
                data[month].steuerfreie_inland_einnahmen += gezahlt;
            } else if (taxType === "steuerfrei_ausland" || !roundedRate) {
                data[month].steuerfreie_ausland_einnahmen += gezahlt;
            } else {
                data[month].steuerpflichtige_einnahmen += gezahlt;
                data[month][`ust_${roundedRate}`] += tax;
            }
        } else {
            if (isEigen) {
                const eigenCfg = config$1.eigenbelege.mapping[category] ?? {};
                const taxType = eigenCfg.taxType ?? "steuerpflichtig";
                if (taxType === "steuerfrei") {
                    data[month].eigenbelege_steuerfrei += gezahlt;
                } else if (taxType === "eigenbeleg" && eigenCfg.besonderheit === "bewirtung") {
                    data[month].eigenbelege_steuerpflichtig += gezahlt;
                    data[month][`vst_${roundedRate}`] += tax * 0.7;
                    data[month].nicht_abzugsfaehige_vst += tax * 0.3;
                } else {
                    data[month].eigenbelege_steuerpflichtig += gezahlt;
                    data[month][`vst_${roundedRate}`] += tax;
                }
            } else {
                const catCfg = config$1.ausgaben.categories[category] ?? {};
                const taxType = catCfg.taxType ?? "steuerpflichtig";
                if (taxType === "steuerfrei_inland") {
                    data[month].steuerfreie_inland_ausgaben += gezahlt;
                } else if (taxType === "steuerfrei_ausland") {
                    data[month].steuerfreie_ausland_ausgaben += gezahlt;
                } else {
                    data[month].steuerpflichtige_ausgaben += gezahlt;
                    data[month][`vst_${roundedRate}`] += tax;
                }
            }
        }
    };

    const formatUStVARow = (label, d) => {
        const ustZahlung = (d.ust_7 + d.ust_19) - ((d.vst_7 + d.vst_19) - d.nicht_abzugsfaehige_vst);
        const ergebnis = (d.steuerpflichtige_einnahmen + d.steuerfreie_inland_einnahmen + d.steuerfreie_ausland_einnahmen) -
            (d.steuerpflichtige_ausgaben + d.steuerfreie_inland_ausgaben + d.steuerfreie_ausland_ausgaben +
                d.eigenbelege_steuerpflichtig + d.eigenbelege_steuerfrei);
        return [
            label,
            d.steuerpflichtige_einnahmen,
            d.steuerfreie_inland_einnahmen,
            d.steuerfreie_ausland_einnahmen,
            d.steuerpflichtige_ausgaben,
            d.steuerfreie_inland_ausgaben,
            d.steuerfreie_ausland_ausgaben,
            d.eigenbelege_steuerpflichtig,
            d.eigenbelege_steuerfrei,
            d.nicht_abzugsfaehige_vst,
            d.ust_7,
            d.ust_19,
            d.vst_7,
            d.vst_19,
            ustZahlung,
            ergebnis
        ];
    };

    const aggregateUStVA = (data, start, end) => {
        const sum = createEmptyUStVA();
        for (let m = start; m <= end; m++) {
            for (const key in sum) {
                sum[key] += data[m][key] || 0;
            }
        }
        return sum;
    };

    const calculateUStVA = () => {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const revenueSheet = ss.getSheetByName("Einnahmen");
        const expenseSheet = ss.getSheetByName("Ausgaben");
        const eigenSheet = ss.getSheetByName("Eigenbelege");
        if (!revenueSheet || !expenseSheet) {
            SpreadsheetApp.getUi().alert("Fehlendes Blatt: 'Einnahmen' oder 'Ausgaben'");
            return;
        }
        if (!Validator.validateAllSheets(revenueSheet, expenseSheet)) return;

        const revenueData = revenueSheet.getDataRange().getValues();
        const expenseData = expenseSheet.getDataRange().getValues();
        const eigenData = eigenSheet ? eigenSheet.getDataRange().getValues() : [];
        const ustvaData = Object.fromEntries(Array.from({length: 12}, (_, i) => [i + 1, createEmptyUStVA()]));

        const processRows = (data, isIncome, isEigen = false) =>
            data.slice(1).forEach(row => {
                const m = Helpers.getMonthFromRow(row);
                if (m) processUStVARow(row, ustvaData, isIncome, isEigen);
            });
        processRows(revenueData, true);
        processRows(expenseData, false);
        if (eigenData.length) processRows(eigenData, false, true);

        const outputRows = [
            [
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
            ]
        ];
        config$1.common.months.forEach((name, i) => {
            outputRows.push(formatUStVARow(name, ustvaData[i + 1]));
            if ((i + 1) % 3 === 0) {
                outputRows.push(formatUStVARow(`Quartal ${(i + 1) / 3}`, aggregateUStVA(ustvaData, i - 1, i + 1)));
            }
        });
        outputRows.push(formatUStVARow("Gesamtjahr", aggregateUStVA(ustvaData, 1, 12)));

        const ustvaSheet = ss.getSheetByName("UStVA") || ss.insertSheet("UStVA");
        ustvaSheet.clearContents();
        ustvaSheet.getRange(1, 1, outputRows.length, outputRows[0].length).setValues(outputRows);
        ustvaSheet.autoResizeColumns(1, outputRows[0].length);
        ss.setActiveSheet(ustvaSheet);
        SpreadsheetApp.getUi().alert("UStVA wurde aktualisiert!");
    };

    return {calculateUStVA};
})();

const BWACalculator = (() => {
    const createEmptyBWA = () => ({
        // Gruppe 1: Betriebserlöse (Einnahmen)
        umsatzerloese: 0,
        provisionserloese: 0,
        steuerfreieInlandEinnahmen: 0,
        steuerfreieAuslandEinnahmen: 0,
        sonstigeErtraege: 0,
        vermietung: 0,
        zuschuesse: 0,
        waehrungsgewinne: 0,
        anlagenabgaenge: 0,
        gesamtErloese: 0,
        // Gruppe 2: Materialaufwand & Wareneinsatz
        wareneinsatz: 0,
        fremdleistungen: 0,
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
        // Gruppe 8: Steuern & Vorsteuer
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
        gewinnNachSteuern: 0,
        // Eigenbelege (zur Aggregation)
        eigenbelegeSteuerfrei: 0,
        eigenbelegeSteuerpflichtig: 0
    });

    // Hier nutzen wir die zentrale Helpers.getMonthFromRow-Funktion:
    const getMonthFromRow = row => {
        return Helpers.getMonthFromRow(row);
    };

    const aggregateBWAData = () => {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const revenueSheet = ss.getSheetByName("Einnahmen");
        const expenseSheet = ss.getSheetByName("Ausgaben");
        const eigenSheet = ss.getSheetByName("Eigenbelege");
        if (!revenueSheet || !expenseSheet) {
            SpreadsheetApp.getUi().alert("Fehlendes Blatt: 'Einnahmen' oder 'Ausgaben'");
            return null;
        }
        const bwaData = Object.fromEntries(Array.from({length: 12}, (_, i) => [i + 1, createEmptyBWA()]));

        const processRevenue = row => {
            const m = getMonthFromRow(row);
            if (!m) return;
            const amount = Helpers.parseCurrency(row[4]);
            const category = row[2]?.toString().trim() || "";
            if (["Gewinnvortrag", "Verlustvortrag", "Gewinnvortrag/Verlustvortrag"].includes(category)) return;
            if (category === "Sonstige betriebliche Erträge") return void (bwaData[m].sonstigeErtraege += amount);
            if (category === "Erträge aus Vermietung/Verpachtung") return void (bwaData[m].vermietung += amount);
            if (category === "Erträge aus Zuschüssen") return void (bwaData[m].zuschuesse += amount);
            if (category === "Erträge aus Währungsgewinnen") return void (bwaData[m].waehrungsgewinne += amount);
            if (category === "Erträge aus Anlagenabgängen") return void (bwaData[m].anlagenabgaenge += amount);
            const mapping = config$1.einnahmen.bwaMapping[category];
            if (["umsatzerloese", "provisionserloese"].includes(mapping)) {
                bwaData[m][mapping] += amount;
            } else if (Helpers.parseMwstRate(row[5]) === 0) {
                bwaData[m].steuerfreieInlandEinnahmen += amount;
            } else {
                bwaData[m].umsatzerloese += amount;
            }
        };

        const processExpense = row => {
            const m = getMonthFromRow(row);
            if (!m) return;
            const amount = Helpers.parseCurrency(row[4]);
            const category = row[2]?.toString().trim() || "";
            if (["Privatentnahme", "Privateinlage", "Holding Transfers",
                "Gewinnvortrag", "Verlustvortrag", "Gewinnvortrag/Verlustvortrag"].includes(category)) return;
            if (category === "Bruttolöhne & Gehälter") return void (bwaData[m].bruttoLoehne += amount);
            if (category === "Soziale Abgaben & Arbeitgeberanteile") return void (bwaData[m].sozialeAbgaben += amount);
            if (category === "Sonstige Personalkosten") return void (bwaData[m].sonstigePersonalkosten += amount);
            if (category === "Gewerbesteuerrückstellungen") return void (bwaData[m].gewerbesteuerRueckstellungen += amount);
            if (category === "Telefon & Internet") return void (bwaData[m].telefonInternet += amount);
            if (category === "Bürokosten") return void (bwaData[m].buerokosten += amount);
            if (category === "Fortbildungskosten") return void (bwaData[m].fortbildungskosten += amount);
            const mapping = config$1.ausgaben.bwaMapping[category];
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
        };

        const processEigen = row => {
            const m = getMonthFromRow(row);
            if (!m) return;
            const amount = Helpers.parseCurrency(row[4]);
            const category = row[2]?.toString().trim() || "";
            const eigenCfg = config$1.eigenbelege.mapping[category] ?? {};
            const taxType = eigenCfg.taxType ?? "steuerpflichtig";
            if (taxType === "steuerfrei") {
                bwaData[m].eigenbelegeSteuerfrei += amount;
            } else {
                bwaData[m].eigenbelegeSteuerpflichtig += amount;
            }
        };

        revenueSheet.getDataRange().getValues().slice(1).forEach(processRevenue);
        expenseSheet.getDataRange().getValues().slice(1).forEach(processExpense);
        if (eigenSheet) eigenSheet.getDataRange().getValues().slice(1).forEach(processEigen);

        // Gruppensummen und weitere Berechnungen
        for (let m = 1; m <= 12; m++) {
            const d = bwaData[m];
            d.gesamtErloese = d.umsatzerloese + d.provisionserloese + d.steuerfreieInlandEinnahmen + d.steuerfreieAuslandEinnahmen +
                d.sonstigeErtraege + d.vermietung + d.zuschuesse + d.waehrungsgewinne + d.anlagenabgaenge;
            d.gesamtWareneinsatz = d.wareneinsatz + d.fremdleistungen + d.rohHilfsBetriebsstoffe;
            d.gesamtBetriebsausgaben = d.bruttoLoehne + d.sozialeAbgaben + d.sonstigePersonalkosten +
                d.werbungMarketing + d.reisekosten + d.versicherungen + d.telefonInternet +
                d.buerokosten + d.fortbildungskosten + d.kfzKosten + d.sonstigeAufwendungen;
            d.gesamtAbschreibungenZinsen = d.abschreibungenMaschinen + d.abschreibungenBueromaterial +
                d.abschreibungenImmateriell + d.zinsenBank + d.zinsenGesellschafter + d.leasingkosten;
            d.gesamtBesonderePosten = d.eigenkapitalveraenderungen + d.gesellschafterdarlehen + d.ausschuettungen;
            d.gesamtRueckstellungenTransfers = d.steuerrueckstellungen + d.rueckstellungenSonstige;
            d.ebit = d.gesamtErloese - (d.gesamtWareneinsatz + d.gesamtBetriebsausgaben + d.gesamtAbschreibungenZinsen + d.gesamtBesonderePosten);
            d.gewerbesteuer = d.ebit * (config$1.tax.operative.gewerbesteuer / 100);
            if (config$1.tax.isHolding) {
                d.koerperschaftsteuer = d.ebit * (config$1.tax.holding.koerperschaftsteuer / 100) * (config$1.tax.holding.gewinnUebertragSteuerpflichtig / 100);
                d.solidaritaetszuschlag = d.ebit * (config$1.tax.holding.solidaritaetszuschlag / 100) * (config$1.tax.holding.gewinnUebertragSteuerpflichtig / 100);
            } else {
                d.koerperschaftsteuer = d.ebit * (config$1.tax.operative.koerperschaftsteuer / 100);
                d.solidaritaetszuschlag = d.ebit * (config$1.tax.operative.solidaritaetszuschlag / 100);
            }
            d.steuerlast = d.koerperschaftsteuer + d.solidaritaetszuschlag + d.gewerbesteuer;
            d.gewinnNachSteuern = d.ebit - d.steuerlast;
        }
        return bwaData;
    };

    // Erzeugt den Header (2 Spalten: Bezeichnung und Wert) mit Monats- und Quartalsspalten
    const buildHeaderRow = () => {
        const headers = ["Kategorie"];
        for (let q = 0; q < 4; q++) {
            for (let m = q * 3; m < q * 3 + 3; m++) {
                headers.push(`${config$1.common.months[m]} (€)`);
            }
            headers.push(`Q${q + 1} (€)`);
        }
        headers.push("Jahr (€)");
        return headers;
    };

    // Hier wird für jede Position eine Zeile mit Monats-, Quartals- und Jahreswerten aufgebaut:
    const calculateBWA = () => {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const bwaData = aggregateBWAData();
        if (!bwaData) return;

        const positions = [
            {label: "Erlöse aus Lieferungen und Leistungen", get: d => d.umsatzerloese || 0},
            {label: "Sonstige betriebliche Erträge", get: d => d.sonstigeErtraege || 0},
            {label: "Erträge aus Vermietung/Verpachtung", get: d => d.vermietung || 0},
            {label: "Erträge aus Zuschüssen", get: d => d.zuschuesse || 0},
            {label: "Erträge aus Währungsgewinnen", get: d => d.waehrungsgewinne || 0},
            {label: "Erträge aus Anlagenabgängen", get: d => d.anlagenabgaenge || 0},
            {label: "Betriebserlöse", get: d => d.gesamtErloese || 0},
            {label: "Wareneinsatz", get: d => d.wareneinsatz || 0},
            {label: "Bezogene Leistungen", get: d => d.fremdleistungen || 0},
            {label: "Roh-, Hilfs- & Betriebsstoffe", get: d => d.rohHilfsBetriebsstoffe || 0},
            {label: "Gesamtkosten Material & Fremdleistungen", get: d => d.gesamtWareneinsatz || 0},
            {label: "Bruttolöhne & Gehälter", get: d => d.bruttoLoehne || 0},
            {label: "Soziale Abgaben & Arbeitgeberanteile", get: d => d.sozialeAbgaben || 0},
            {label: "Sonstige Personalkosten", get: d => d.sonstigePersonalkosten || 0},
            {label: "Werbung & Marketing", get: d => d.werbungMarketing || 0},
            {label: "Reisekosten", get: d => d.reisekosten || 0},
            {label: "Versicherungen", get: d => d.versicherungen || 0},
            {label: "Telefon & Internet", get: d => d.telefonInternet || 0},
            {label: "Bürokosten", get: d => d.buerokosten || 0},
            {label: "Fortbildungskosten", get: d => d.fortbildungskosten || 0},
            {label: "Kfz-Kosten", get: d => d.kfzKosten || 0},
            {label: "Sonstige betriebliche Aufwendungen", get: d => d.sonstigeAufwendungen || 0},
            {label: "Abschreibungen Maschinen", get: d => d.abschreibungenMaschinen || 0},
            {label: "Abschreibungen Büroausstattung", get: d => d.abschreibungenBueromaterial || 0},
            {label: "Abschreibungen immaterielle Wirtschaftsgüter", get: d => d.abschreibungenImmateriell || 0},
            {label: "Zinsen auf Bankdarlehen", get: d => d.zinsenBank || 0},
            {label: "Zinsen auf Gesellschafterdarlehen", get: d => d.zinsenGesellschafter || 0},
            {label: "Leasingkosten", get: d => d.leasingkosten || 0},
            {label: "Gesamt Abschreibungen & Zinsen", get: d => d.gesamtAbschreibungenZinsen || 0},
            {label: "Eigenkapitalveränderungen", get: d => d.eigenkapitalveraenderungen || 0},
            {label: "Gesellschafterdarlehen", get: d => d.gesellschafterdarlehen || 0},
            {label: "Ausschüttungen an Gesellschafter", get: d => d.ausschuettungen || 0},
            {label: "Steuerrückstellungen", get: d => d.steuerrueckstellungen || 0},
            {label: "Rückstellungen sonstige", get: d => d.rueckstellungenSonstige || 0},
            {label: "Betriebsergebnis vor Steuern (EBIT)", get: d => d.ebit || 0},
            {label: "Umsatzsteuer (abzuführen)", get: d => d.umsatzsteuer || 0},
            {label: "Vorsteuer", get: d => d.vorsteuer || 0},
            {label: "Nicht abzugsfähige VSt (Bewirtung)", get: d => d.nichtAbzugsfaehigeVSt || 0},
            {label: "Körperschaftsteuer", get: d => d.koerperschaftsteuer || 0},
            {label: "Solidaritätszuschlag", get: d => d.solidaritaetszuschlag || 0},
            {label: "Gewerbesteuer", get: d => d.gewerbesteuer || 0},
            {label: "Gesamtsteueraufwand", get: d => d.steuerlast || 0},
            {label: "Jahresüberschuss/-fehlbetrag", get: d => d.gewinnNachSteuern || 0}
        ];

        const headerRow = buildHeaderRow();
        const outputRows = [headerRow];

        // Für jede Position wird eine Zeile mit Monats-, Quartals- und Jahreswerten aufgebaut:
        const buildOutputRow = pos => {
            const monthly = [];
            let yearly = 0;
            for (let m = 1; m <= 12; m++) {
                const val = pos.get(bwaData[m]) || 0;
                monthly.push(val);
                yearly += val;
            }
            const quarters = [0, 0, 0, 0];
            for (let i = 0; i < 12; i++) {
                quarters[Math.floor(i / 3)] += monthly[i];
            }
            return [pos.label, ...monthly.slice(0, 3), quarters[0], ...monthly.slice(3, 6), quarters[1],
                ...monthly.slice(6, 9), quarters[2], ...monthly.slice(9, 12), quarters[3], yearly];
        };

        let posIndex = 0;
        for (const {header, count} of [
            {header: "1️⃣ Betriebserlöse (Einnahmen)", count: 7},
            {header: "2️⃣ Materialaufwand & Wareneinsatz", count: 4},
            {header: "3️⃣ Betriebsausgaben (Sachkosten)", count: 10},
            {header: "4️⃣ Abschreibungen & Zinsen", count: 7},
            {header: "5️⃣ Besondere Posten", count: 3},
            {header: "6️⃣ Rückstellungen", count: 2},
            {header: "7️⃣ Betriebsergebnis vor Steuern (EBIT)", count: 1},
            {header: "8️⃣ Steuern & Vorsteuer", count: 7},
            {header: "9️⃣ Jahresüberschuss/-fehlbetrag", count: 1}
        ]) {
            outputRows.push([header, ...Array(headerRow.length - 1).fill("")]);
            for (let i = 0; i < count; i++) {
                outputRows.push(buildOutputRow(positions[posIndex++]));
            }
        }

        const bwaSheet = ss.getSheetByName("BWA") || ss.insertSheet("BWA");
        bwaSheet.clearContents();
        bwaSheet.getRange(1, 1, outputRows.length, outputRows[0].length).setValues(outputRows);
        bwaSheet.autoResizeColumns(1, outputRows[0].length);
        SpreadsheetApp.getUi().alert("BWA wurde aktualisiert!");
    };

    return {calculateBWA};
})();

// imports TEST
// import calculateBilanz from "./calculateBilanz.js";

// =================== Globale Funktionen ===================
const onOpen = () => {
    SpreadsheetApp.getUi()
        .createMenu("📂 Buchhaltung")
        .addItem("📥 Dateien importieren", "importDriveFiles")
        .addItem("🔄 Refresh Active Sheet", "refreshSheet")
        .addItem("📊 UStVA berechnen", "calculateUStVA")
        .addItem("📈 BWA berechnen", "calculateBWA")
        .addItem("📝 Bilanz erstellen", "calculateBilanz")
        .addToUi();
};

const onEdit = e => {
    const {range} = e;
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

    // Prüfen, ob wir dieses Sheet bearbeiten sollen
    if (!(name in mapping)) return;

    // Header-Zeile ignorieren
    if (range.getRow() === 1) return;

    // Spalte für Zeitstempel
    const timestampCol = mapping[name];

    // Prüfen, ob die bearbeitete Spalte bereits die Zeitstempel-Spalte ist
    if (range.getColumn() === timestampCol ||
        (range.getNumColumns() > 1 && range.getColumn() <= timestampCol &&
            range.getColumn() + range.getNumColumns() > timestampCol)) {
        return; // Vermeide Endlosschleife
    }

    // Multi-Zellen-Editierung unterstützen
    const numRows = range.getNumRows();
    const timestampValues = [];
    const now = new Date();

    // Für jede bearbeitete Zeile
    for (let i = 0; i < numRows; i++) {
        const rowIndex = range.getRow() + i;
        const headerLen = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].length;

        // Prüfen, ob die Zeile leer ist
        const rowValues = sheet.getRange(rowIndex, 1, 1, headerLen).getValues()[0];
        if (rowValues.every(cell => cell === "")) continue;

        // Zeitstempel für diese Zeile hinzufügen
        timestampValues.push([now]);
    }

    // Wenn es Zeitstempel zu setzen gibt
    if (timestampValues.length > 0) {
        sheet.getRange(range.getRow(), timestampCol, timestampValues.length, 1)
            .setValues(timestampValues);
    }
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
const calculateUStVA = () => {
    RefreshModule.refreshAllSheets();
    UStVACalculator.calculateUStVA();
};
const calculateBWA = () => {
    RefreshModule.refreshAllSheets();
    BWACalculator.calculateBWA();
};
const importDriveFiles = () => {
    ImportModule.importDriveFiles();
    RefreshModule.refreshAllSheets();
};
