// === Hinweis für Entwickler ===
// Bitte aktualisiere diese Kommentare, wenn du Änderungen am Code vornimmst,
// damit sie stets den aktuellen Stand widerspiegeln.


// Zentrale Kategorien-Konfiguration – Single Source of Truth
const CategoryConfig = {
    einnahmen: {
        allowed: ["Umsatzerlöse", "Provisionserlöse", "Sonstige betriebliche Erträge"],
        bwaMapping: {
            "Umsatzerlöse": "umsatzerloese",
            "Provisionserlöse": "provisionserloese",
            "Sonstige betriebliche Erträge": "sonstigeErtraege"
        }
    },
    ausgaben: {
        allowed: ["Wareneinsatz", "Betriebskosten", "Marketing & Werbung", "Reisekosten", "Personalkosten", "Sonstige betriebliche Aufwendungen"],
        bwaMapping: {
            "Wareneinsatz": "wareneinsatz",
            "Betriebskosten": "betriebskosten",
            "Marketing & Werbung": "marketing",
            "Reisekosten": "reisen",
            "Personalkosten": "personalkosten",
            "Sonstige betriebliche Aufwendungen": "sonstigeAufwendungen"
        }
    },
    bank: {
        allowed: [
            "Umsatzerlöse", "Provisionserlöse", "Sonstige betriebliche Erträge",
            "Wareneinsatz", "Betriebskosten", "Marketing & Werbung", "Reisekosten", "Personalkosten", "Sonstige betriebliche Aufwendungen",
            "Eigenbeleg", "Privateinlage", "Privatentnahme", "Darlehen"
        ],
        // Im Bankbereich sollen in Spalte E (Typ) nur "Einnahme" (bei positiven Beträgen)
        // und "Ausgabe" (bei negativen Beträgen) zulässig sein.
        typeAllowed: ["Einnahme", "Ausgabe"],
        bwaMapping: {
            "Umsatzerlöse": "umsatzerloese",
            "Provisionserlöse": "provisionserloese",
            "Sonstige betriebliche Erträge": "sonstigeErtraege",
            "Wareneinsatz": "wareneinsatz",
            "Betriebskosten": "betriebskosten",
            "Marketing & Werbung": "marketing",
            "Reisekosten": "reisen",
            "Personalkosten": "personalkosten",
            "Sonstige betriebliche Aufwendungen": "sonstigeAufwendungen",
            "Eigenbeleg": "eigenbeleg",
            "Privateinlage": "privateinlage",
            "Privatentnahme": "privatentnahme",
            "Darlehen": "darlehen"
        }
    }
};


// ================= Modul: Validator =================
const Validator = (() => {
    const isEmpty = value => value === undefined || value === null || value.toString().trim() === "";

    // Validiert eine Zeile aus dem GuV-/BWA-Kontext (Einnahmen/Ausgaben).
    // Erwartete Felder (Index):
    // A: Rechnungsdatum, B: Rechnungsnummer, C: Kategorie, D: Kunde,
    // E: Nettobetrag, F: Mehrwertsteuer, L: Zahlungsstatus, M: Zahlungsdatum.
    const validateRevenueAndExpenses = (row, rowIndex) => {
        const warnings = [];
        isEmpty(row[0]) && warnings.push(`Zeile ${rowIndex} (E/A): Rechnungsdatum fehlt.`);
        isEmpty(row[1]) && warnings.push(`Zeile ${rowIndex} (E/A): Rechnungsnummer fehlt.`);
        isEmpty(row[2]) && warnings.push(`Zeile ${rowIndex} (E/A): Kategorie fehlt.`);
        isEmpty(row[3]) && warnings.push(`Zeile ${rowIndex} (E/A): Kunde fehlt.`);
        (isEmpty(row[4]) || isNaN(parseFloat(row[4].toString().trim()))) &&
        warnings.push(`Zeile ${rowIndex} (E/A): Nettobetrag fehlt oder ungültig.`);
        // 0% Mehrwertsteuer soll erlaubt sein:
        const mwstStr = row[5] == null ? "" : row[5].toString().trim();
        (isEmpty(mwstStr) || isNaN(parseFloat(mwstStr.replace("%", "")))) &&
        warnings.push(`Zeile ${rowIndex} (E/A): Mehrwertsteuer fehlt oder ungültig.`);
        const status = row[11] ? row[11].toString().trim().toLowerCase() : "";
        const zahlungsdatum = row[12] ? row[12].toString().trim() : "";
        status === "offen"
            ? !isEmpty(zahlungsdatum) && warnings.push(`Zeile ${rowIndex} (E/A): Zahlungsdatum darf nicht gesetzt sein, wenn "offen".`)
            : isEmpty(zahlungsdatum) && warnings.push(`Zeile ${rowIndex} (E/A): Zahlungsdatum muss gesetzt sein, wenn bezahlt/teilbezahlt.`);
        return warnings;
    };

    // Validiert eine Zeile aus dem Bankbewegungskontext.
    // Erwartete Felder (Index):
    // A: Buchungsdatum, B: Buchungstext, C: Betrag, D: Saldo, E: Typ, F: Kategorie.
    // Alle Zeilen werden validiert – auch Zeile 2 (Anfangssaldo) und letzte Zeile (Endsaldo).
    const validateBankSheet = (row, rowIndex, totalRows) => {
        const warnings = [];
        if (rowIndex === 2 || rowIndex === totalRows) {
            isEmpty(row[1]) && warnings.push(`Zeile ${rowIndex} (Bank): Buchungstext fehlt.`);
            (isEmpty(row[3]) || isNaN(parseFloat(row[3].toString().trim()))) &&
            warnings.push(`Zeile ${rowIndex} (Bank): Saldo fehlt oder ungültig.`);
        } else {
            isEmpty(row[0]) && warnings.push(`Zeile ${rowIndex} (Bank): Buchungsdatum fehlt.`);
            isEmpty(row[1]) && warnings.push(`Zeile ${rowIndex} (Bank): Buchungstext fehlt.`);
            (isEmpty(row[2]) || isNaN(parseFloat(row[2].toString().trim()))) &&
            warnings.push(`Zeile ${rowIndex} (Bank): Betrag fehlt oder ungültig.`);
            (isEmpty(row[3]) || isNaN(parseFloat(row[3].toString().trim()))) &&
            warnings.push(`Zeile ${rowIndex} (Bank): Saldo fehlt oder ungültig.`);
            isEmpty(row[4]) && warnings.push(`Zeile ${rowIndex} (Bank): Typ fehlt.`);
            isEmpty(row[5]) && warnings.push(`Zeile ${rowIndex} (Bank): Kategorie fehlt.`);
        }
        return warnings;
    };

    return { validateRevenueAndExpenses, validateBankSheet };
})();


// ================= Modul: Buchhaltung =================
const Buchhaltung = (() => {
    const setupTrigger = () => {
        const triggers = ScriptApp.getProjectTriggers();
        if (!triggers.some(t => t.getHandlerFunction() === "onOpen"))
            ScriptApp.newTrigger("onOpen")
                .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
                .onOpen()
                .create();
    };

    const onOpen = () =>
        SpreadsheetApp.getUi()
            .createMenu("📂 Buchhaltung")
            .addItem("📥 Rechnungen importieren", "importDriveFiles")
            .addItem("🔄 Aktualisieren (Formeln & Formatierung)", "refreshSheets")
            .addItem("📊 GuV berechnen", "calculateGuV")
            .addItem("📈 BWA berechnen", "calculateBWA")
            .addToUi();

    const importDriveFiles = () => {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheets = {
            revenueMain: ss.getSheetByName("Einnahmen"),
            expenseMain: ss.getSheetByName("Ausgaben"),
            revenue: ss.getSheetByName("Rechnungen Einnahmen") || ss.insertSheet("Rechnungen Einnahmen"),
            expense: ss.getSheetByName("Rechnungen Ausgaben") || ss.insertSheet("Rechnungen Ausgaben"),
            history: ss.getSheetByName("Änderungshistorie") || ss.insertSheet("Änderungshistorie")
        };
        if (sheets.history.getLastRow() === 0)
            sheets.history.appendRow(["Datum", "Rechnungstyp", "Dateiname", "Link zur Datei"]);
        if (sheets.revenue.getLastRow() === 0)
            sheets.revenue.appendRow(["Dateiname", "Link zur Datei", "Rechnungsnummer"]);
        if (sheets.expense.getLastRow() === 0)
            sheets.expense.appendRow(["Dateiname", "Link zur Datei", "Rechnungsnummer"]);

        const file = DriveApp.getFileById(ss.getId());
        const parentFolder = file.getParents()?.hasNext() ? file.getParents().next() : null;
        if (!parentFolder) {
            SpreadsheetApp.getUi().alert("Kein übergeordneter Ordner gefunden.");
            return;
        }
        const revenueFolder = getFolderByName(parentFolder, "Einnahmen");
        const expenseFolder = getFolderByName(parentFolder, "Ausgaben");
        revenueFolder
            ? importFilesFromFolder(revenueFolder, sheets.revenue, sheets.revenueMain, "Einnahme", sheets.history)
            : SpreadsheetApp.getUi().alert("Fehler: 'Einnahmen'-Ordner nicht gefunden.");
        expenseFolder
            ? importFilesFromFolder(expenseFolder, sheets.expense, sheets.expenseMain, "Ausgabe", sheets.history)
            : SpreadsheetApp.getUi().alert("Fehler: 'Ausgaben'-Ordner nicht gefunden.");
        refreshSheets();
    };

    const importFilesFromFolder = (folder, importSheet, mainSheet, type, historySheet) => {
        const files = folder.getFiles();
        const getExistingFiles = (sheet, colIndex) =>
            new Set(sheet.getDataRange().getValues().slice(1).map(row => row[colIndex]));
        const existingMain = getExistingFiles(mainSheet, 1);
        const existingImport = getExistingFiles(importSheet, 0);
        const newMainRows = [];
        const newImportRows = [];
        const newHistoryRows = [];
        const timestamp = new Date();
        while (files.hasNext()) {
            const file = files.next();
            const fileName = file.getName().replace(/\.[^/.]+$/, "");
            const fileUrl = file.getUrl();
            let wasImported = false;
            if (!existingMain.has(fileName)) {
                const rowIndex = mainSheet.getLastRow() + newMainRows.length + 1;
                newMainRows.push([
                    timestamp, fileName, "", "", "", "",
                    `=E${rowIndex}*F${rowIndex}`,
                    `=E${rowIndex}+G${rowIndex}`, "",
                    `=E${rowIndex}-(I${rowIndex}-G${rowIndex})`,
                    `=IF(A${rowIndex}=""; ""; ROUNDUP(MONTH(A${rowIndex})/3;0))`,
                    `=IF(OR(I${rowIndex}=""; I${rowIndex}=0); "Offen"; IF(I${rowIndex}>=H${rowIndex}; "Bezahlt"; "Teilbezahlt"))`,
                    "", fileName, fileUrl, timestamp
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
        newImportRows.length > 0 &&
        importSheet.getRange(importSheet.getLastRow() + 1, 1, newImportRows.length, newImportRows[0].length).setValues(newImportRows);
        newMainRows.length > 0 &&
        mainSheet.getRange(mainSheet.getLastRow() + 1, 1, newMainRows.length, newMainRows[0].length).setValues(newMainRows);
        newHistoryRows.length > 0 &&
        historySheet.getRange(historySheet.getLastRow() + 1, 1, newHistoryRows.length, newHistoryRows[0].length).setValues(newHistoryRows);
    };

    const getFolderByName = (parent, name) => {
        const folderIter = parent.getFoldersByName(name);
        return folderIter.hasNext() ? folderIter.next() : null;
    };

    // Refresh für Einnahmen/Ausgaben: Setzt Formeln, Formatierung, Standardwert in Spalte I (Bezahlt) und Dropdown für Kategorie.
    // Die Validierung erfolgt in calculateGuV.
    const refreshSheet = sheet => {
        const lastRow = sheet.getLastRow();
        if (lastRow < 2) return;
        const numRows = lastRow - 1;
        const formulas = Array.from({ length: numRows }, (_, i) => {
            const row = i + 2;
            return {
                mwst: `=E${row}*F${row}`,
                brutto: `=E${row}+G${row}`,
                rest: `=(H${row}-I${row})/(1+VALUE(F${row}))`,
                quartal: `=IF(A${row}=""; ""; ROUNDUP(MONTH(A${row})/3;0))`,
                status: `=IF(OR(I${row}=""; I${row}=0); "Offen"; IF(I${row}>=H${row}; "Bezahlt"; "Teilbezahlt"))`
            };
        });
        sheet.getRange(2, 7, numRows, 1).setFormulas(formulas.map(f => [f.mwst]));
        sheet.getRange(2, 8, numRows, 1).setFormulas(formulas.map(f => [f.brutto]));
        sheet.getRange(2, 10, numRows, 1).setFormulas(formulas.map(f => [f.rest]));
        sheet.getRange(2, 11, numRows, 1).setFormulas(formulas.map(f => [f.quartal]));
        sheet.getRange(2, 12, numRows, 1).setFormulas(formulas.map(f => [f.status]));

        const dateFormat = "DD.MM.YYYY";
        const currencyFormat = "€#,##0.00;€-#,##0.00";
        sheet.getRange(`A2:A${lastRow}`).setNumberFormat(dateFormat);
        sheet.getRange(`M2:M${lastRow}`).setNumberFormat(dateFormat);
        sheet.getRange(`P2:P${lastRow}`).setNumberFormat(dateFormat);
        sheet.getRange(`E2:E${lastRow}`).setNumberFormat(currencyFormat);
        sheet.getRange(`G2:G${lastRow}`).setNumberFormat(currencyFormat);
        sheet.getRange(`H2:H${lastRow}`).setNumberFormat(currencyFormat);
        sheet.getRange(`I2:I${lastRow}`).setNumberFormat(currencyFormat);
        sheet.getRange(`J2:J${lastRow}`).setNumberFormat(currencyFormat);
        sheet.getRange(`F2:F${lastRow}`).setNumberFormat("0.00%");

        for (let r = 2; r <= lastRow; r++) {
            const cell = sheet.getRange(r, 9);
            (cell.getValue() === "" || cell.getValue() === null) && cell.setValue(0);
        }

        const sheetName = sheet.getName();
        if (sheetName === "Einnahmen") {
            const rule = SpreadsheetApp.newDataValidation().requireValueInList(CategoryConfig.einnahmen.allowed, true).build();
            sheet.getRange(2, 3, lastRow - 1, 1).setDataValidation(rule);
        } else if (sheetName === "Ausgaben") {
            const rule = SpreadsheetApp.newDataValidation().requireValueInList(CategoryConfig.ausgaben.allowed, true).build();
            sheet.getRange(2, 3, lastRow - 1, 1).setDataValidation(rule);
        }

        sheet.autoResizeColumns(1, sheet.getLastColumn());
    };

    // Refresh für Bankbewegungen:
    // - Zeile 2 (Anfangssaldo) und letzte Zeile (Endsaldo) werden voll validiert.
    // - Für Transaktionszeilen (Zeile 3 bis zur letzten Zeile) wird in Spalte E (Typ) automatisch "Einnahme" (bei positiv)
    //   bzw. "Ausgabe" (bei negativ) gesetzt; Dropdowns für Typ (Spalte E) und Kategorie (Spalte F) werden gesetzt.
    // - Anschließend wird geprüft, ob die letzte Zeile bereits als Endsaldo existiert; falls ja, wird sie aktualisiert,
    //   andernfalls wird eine neue Endsaldo-Zeile angehängt.
    const refreshBankSheet = sheet => {
        const lastRow = sheet.getLastRow();
        if (lastRow < 3) return;
        const firstDataRow = 3;
        const numTransRows = lastRow - 2;
        const transRows = numTransRows - 1; // Alle Transaktionszeilen außer der Endsaldo-Zeile
        if (transRows > 0) {
            const saldoFormulas = Array.from({ length: transRows }, (_, i) => {
                const row = firstDataRow + i;
                return [`=D${row - 1}+C${row}`];
            });
            sheet.getRange(firstDataRow, 4, transRows, 1).setFormulas(saldoFormulas);
        }
        for (let row = firstDataRow; row <= lastRow; row++) {
            const amount = parseFloat(sheet.getRange(row, 3).getValue()) || 0;
            const typeCell = sheet.getRange(row, 5);
            amount > 0 ? typeCell.setValue("Einnahme") : amount < 0 ? typeCell.setValue("Ausgabe") : typeCell.clearContent();
        }
        const typeRule = SpreadsheetApp.newDataValidation().requireValueInList(CategoryConfig.bank.typeAllowed, true).build();
        sheet.getRange(firstDataRow, 5, lastRow - firstDataRow + 1, 1).setDataValidation(typeRule);
        const catRule = SpreadsheetApp.newDataValidation().requireValueInList(CategoryConfig.bank.allowed, true).build();
        sheet.getRange(firstDataRow, 6, lastRow - firstDataRow + 1, 1).setDataValidation(catRule);
        const dateFormat = "DD.MM.YYYY";
        const currencyFormat = "€#,##0.00;€-#,##0.00";
        sheet.getRange("A2:A" + lastRow).setNumberFormat(dateFormat);
        sheet.getRange("C2:C" + lastRow).setNumberFormat(currencyFormat);
        sheet.getRange("D2:D" + lastRow).setNumberFormat(currencyFormat);
        sheet.autoResizeColumns(1, sheet.getLastColumn());

        // Endsaldo-Zeile prüfen: Wenn letzte Zeile bereits "endsaldo" ist, aktualisieren; sonst anhängen.
        const lastRowText = sheet.getRange(lastRow, 2).getValue().toString().trim().toLowerCase();
        const formattedDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd.MM.yyyy");
        if (lastRowText === "endsaldo") {
            sheet.getRange(lastRow, 1).setValue(formattedDate);
            sheet.getRange(lastRow, 4).setFormula(`=D${lastRow - 1}`);
        } else {
            sheet.appendRow([formattedDate, "Endsaldo", "", sheet.getRange(lastRow, 4).getValue(), "", "", "", "", "", "", "", ""]);
        }
    };

    const refreshSheets = () => {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const revenueSheet = ss.getSheetByName("Einnahmen");
        const expenseSheet = ss.getSheetByName("Ausgaben");
        const bankSheet = ss.getSheetByName("Bankbewegungen");
        if (!revenueSheet || !expenseSheet)
            SpreadsheetApp.getUi().alert("Eines der Blätter 'Einnahmen' oder 'Ausgaben' wurde nicht gefunden.");
        else (refreshSheet(revenueSheet), refreshSheet(expenseSheet));
        if (!bankSheet)
            SpreadsheetApp.getUi().alert("Fehlendes Blatt: 'Bankbewegungen'");
        else refreshBankSheet(bankSheet);
        SpreadsheetApp.getUi().alert("Alle relevanten Sheets wurden erfolgreich aktualisiert!");
    };

    return { setupTrigger, onOpen, importDriveFiles, refreshSheets, refreshSheet };
})();


// ================= Modul: GuV-Berechnung =================
const GuVCalculator = (() => {
    const parseMwstRate = value => {
        let rate = parseFloat(value.toString().replace("%", "").trim());
        return isNaN(rate) ? 19 : rate;
    };

    const createEmptyGuV = () => ({
        einnahmen: 0,
        ausgaben: 0,
        ust_0: 0,
        ust_7: 0,
        ust_19: 0,
        vst_0: 0,
        vst_7: 0,
        vst_19: 0
    });

    const parseDate = value => {
        const d = typeof value === "string" ? new Date(value) : (value instanceof Date ? value : null);
        return d && !isNaN(d.getTime()) ? d : null;
    };

    const parseCurrency = value =>
        parseFloat(value.toString().replace(/[^\d,.-]/g, "").replace(",", ".")) || 0;

    const processGuVRow = (row, guvData, isIncome) => {
        const paymentDate = parseDate(row[12]);
        if (!paymentDate || paymentDate > new Date()) return;
        const month = paymentDate.getMonth() + 1;
        const netto = parseCurrency(row[4]);
        const restNetto = parseCurrency(row[9]) || 0;
        const gezahltNetto = Math.max(0, netto - restNetto);
        const mwstRate = parseMwstRate(row[5]);
        const tax = gezahltNetto * (mwstRate / 100);
        if (isIncome) {
            guvData[month].einnahmen += gezahltNetto;
            guvData[month][`ust_${Math.round(mwstRate)}`] += tax;
        } else {
            guvData[month].ausgaben += gezahltNetto;
            guvData[month][`vst_${Math.round(mwstRate)}`] += tax;
        }
    };

    const formatGuVRow = (label, data) => {
        const ustZahlung = data.ust_19 - data.vst_19;
        const ergebnis = data.einnahmen - data.ausgaben;
        return [
            label,
            data.einnahmen,
            data.ausgaben,
            data.ust_0,
            data.ust_7,
            data.ust_19,
            data.vst_0,
            data.vst_7,
            data.vst_19,
            ustZahlung,
            ergebnis
        ];
    };

    const aggregateGuV = (guvData, start, end) => {
        const sum = createEmptyGuV();
        for (let m = start; m <= end; m++) {
            for (const key in sum) {
                sum[key] += guvData[m][key] || 0;
            }
        }
        return sum;
    };

    const calculateGuV = () => {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const revenueSheet = ss.getSheetByName("Einnahmen");
        const expenseSheet = ss.getSheetByName("Ausgaben");
        if (!revenueSheet || !expenseSheet) {
            SpreadsheetApp.getUi().alert("Fehlendes Blatt: 'Einnahmen' oder 'Ausgaben'");
            return;
        }
        let revenueWarnings = [];
        revenueSheet.getDataRange().getValues().slice(1).forEach((row, index) => {
            revenueWarnings = revenueWarnings.concat(Validator.validateRevenueAndExpenses(row, index + 2));
        });
        let expenseWarnings = [];
        expenseSheet.getDataRange().getValues().slice(1).forEach((row, index) => {
            expenseWarnings = expenseWarnings.concat(Validator.validateRevenueAndExpenses(row, index + 2));
        });
        let msg = "";
        revenueWarnings.length > 0 && (msg += "Fehler in 'Einnahmen':\n" + revenueWarnings.join("\n") + "\n\n");
        expenseWarnings.length > 0 && (msg += "Fehler in 'Ausgaben':\n" + expenseWarnings.join("\n"));
        if (msg !== "") {
            SpreadsheetApp.getUi().alert(msg);
            return;
        }
        const revenueData = revenueSheet.getDataRange().getValues().slice(1);
        const expenseData = expenseSheet.getDataRange().getValues().slice(1);
        const guvData = {};
        for (let m = 1; m <= 12; m++) {
            guvData[m] = createEmptyGuV();
        }
        revenueData.forEach(row => processGuVRow(row, guvData, true));
        expenseData.forEach(row => processGuVRow(row, guvData, false));
        const guvSheet = ss.getSheetByName("GuV") || ss.insertSheet("GuV");
        guvSheet.clearContents();
        guvSheet.appendRow([
            "Zeitraum",
            "Einnahmen (netto)",
            "Ausgaben (netto)",
            "USt 0%",
            "USt 7%",
            "USt 19%",
            "VSt 0%",
            "VSt 7%",
            "VSt 19%",
            "USt-Zahlung",
            "Ergebnis (Gewinn/Verlust)"
        ]);
        const monthNames = [
            "Januar", "Februar", "März", "April", "Mai", "Juni",
            "Juli", "August", "September", "Oktober", "November", "Dezember"
        ];
        for (let m = 1; m <= 12; m++) {
            guvSheet.appendRow(formatGuVRow(monthNames[m - 1], guvData[m]));
            m % 3 === 0 && guvSheet.appendRow(formatGuVRow(`Quartal ${m / 3}`, aggregateGuV(guvData, m - 2, m)));
        }
        const yearTotal = aggregateGuV(guvData, 1, 12);
        guvSheet.appendRow(formatGuVRow("Gesamtjahr", yearTotal));
        guvSheet.getRange(2, 2, guvSheet.getLastRow() - 1, 9)
            .setNumberFormat("#,##0.00€");
        guvSheet.autoResizeColumns(1, guvSheet.getLastColumn());
        SpreadsheetApp.getUi().alert("GuV wurde aktualisiert!");
    };

    return { calculateGuV };
})();


// ================= Modul: BWACalculator =================
const BWACalculator = (() => {
    // Ermittelt die interne BWA-Kategorie mithilfe der zentralen Konfiguration.
    function getBwaCategory(category, isIncome, rowIndex, fehlendeKategorien, type = "operativ") {
        const mapping = type === "bank"
            ? CategoryConfig.bank.bwaMapping
            : (isIncome ? CategoryConfig.einnahmen.bwaMapping : CategoryConfig.ausgaben.bwaMapping);
        if (!category || !mapping[category]) {
            const fallback = isIncome ? "sonstigeErtraege" : "betriebskosten";
            fehlendeKategorien.push(`Zeile ${rowIndex}: Unbekannte Kategorie "${category || "N/A"}" → Verwende "${fallback}"`);
            return fallback;
        }
        return mapping[category];
    }

    // Berechnet die BWA, nachdem Einnahmen, Ausgaben und Bankbewegungen validiert wurden.
    // Es werden alle Zeilen validiert – auch Zeile 2 (Anfangssaldo) und die letzte Zeile (Endsaldo).
    // Treten in irgendeiner Zeile Fehler auf, wird die Berechnung abgebrochen.
    function calculateBWA() {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const revenueSheet = ss.getSheetByName("Einnahmen");
        const expenseSheet = ss.getSheetByName("Ausgaben");
        const bankSheet = ss.getSheetByName("Bankbewegungen");
        const bwaSheet = ss.getSheetByName("BWA") || ss.insertSheet("BWA");

        let revenueWarnings = [];
        revenueSheet.getDataRange().getValues().slice(1).forEach((row, index) => {
            revenueWarnings = revenueWarnings.concat(Validator.validateRevenueAndExpenses(row, index + 2));
        });
        let expenseWarnings = [];
        expenseSheet.getDataRange().getValues().slice(1).forEach((row, index) => {
            expenseWarnings = expenseWarnings.concat(Validator.validateRevenueAndExpenses(row, index + 2));
        });
        let bankWarnings = [];
        if (bankSheet) {
            const bankData = bankSheet.getDataRange().getValues();
            // Validierung aller Zeilen im Banksheet (inkl. Zeile 2 und letzte Zeile)
            for (let i = 1; i < bankData.length; i++) {
                bankWarnings = bankWarnings.concat(Validator.validateBankSheet(bankData[i], i + 1, bankData.length));
            }
        }
        if (revenueWarnings.length > 0 || expenseWarnings.length > 0 || bankWarnings.length > 0) {
            const msg = "Fehler in 'Einnahmen':\n" + revenueWarnings.join("\n") +
                "\n\nFehler in 'Ausgaben':\n" + expenseWarnings.join("\n") +
                (bankSheet ? ("\n\nFehler in 'Bankbewegungen':\n" + bankWarnings.join("\n")) : "");
            SpreadsheetApp.getUi().alert(msg);
            return;
        }

        const categories = {
            einnahmen: { umsatzerloese: 0, provisionserloese: 0, sonstigeErtraege: 0 },
            ausgaben: { wareneinsatz: 0, betriebskosten: 0, marketing: 0, reisen: 0, personalkosten: 0, sonstigeAufwendungen: 0 },
            bank: { eigenbeleg: 0, privateinlage: 0, privatentnahme: 0, darlehen: 0 }
        };
        let totalEinnahmen = 0, totalAusgaben = 0;
        let offeneForderungen = 0, offeneVerbindlichkeiten = 0;
        let totalLiquiditaet = 0;
        const fehlendeKategorien = [];

        const revenueData = revenueSheet.getDataRange().getValues().slice(1);
        revenueData.forEach((row, index) => {
            const kategorieUser = row[2];
            const nettoBetrag = parseFloat(row[4]) || 0;
            const restBetrag = parseFloat(row[9]) || 0;
            const bwaCat = getBwaCategory(kategorieUser, true, index + 2, fehlendeKategorien);
            categories.einnahmen[bwaCat] += nettoBetrag;
            totalEinnahmen += nettoBetrag;
            offeneForderungen += restBetrag;
        });

        const expenseData = expenseSheet.getDataRange().getValues().slice(1);
        expenseData.forEach((row, index) => {
            const kategorieUser = row[2];
            const nettoBetrag = parseFloat(row[4]) || 0;
            const restBetrag = parseFloat(row[9]) || 0;
            const bwaCat = getBwaCategory(kategorieUser, false, index + 2, fehlendeKategorien);
            categories.ausgaben[bwaCat] = (categories.ausgaben[bwaCat] || 0) + nettoBetrag;
            totalAusgaben += nettoBetrag;
            offeneVerbindlichkeiten += restBetrag;
        });

        if (bankSheet) {
            const bankData = bankSheet.getDataRange().getValues();
            // Überspringe Zeile 2 (Anfangssaldo) und Endsaldo-Zeile (letzte Zeile)
            for (let i = 1; i < bankData.length - 1; i++) {
                const row = bankData[i];
                const saldo = parseFloat(row[3]) || 0;
                totalLiquiditaet = saldo;
            }
        }

        // Aufbau des BWA-Sheets
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
        const ergebnisNachSteuern = ergebnis;
        bwaSheet.appendRow(["Ergebnis nach Steuern", ergebnisNachSteuern]);
        bwaSheet.appendRow(["--- FINANZIERUNG ---", ""]);
        bwaSheet.appendRow(["Eigenbeleg", 0]);
        bwaSheet.appendRow(["Privateinlage", 0]);
        bwaSheet.appendRow(["Privatentnahme", 0]);
        bwaSheet.appendRow(["Darlehen", 0]);
        bwaSheet.appendRow(["--- LIQUIDITÄT ---", ""]);
        bwaSheet.appendRow(["Kontostand (Bankbewegungen)", totalLiquiditaet]);

        bwaSheet.getRange("A1:B1").setFontWeight("bold");
        bwaSheet.getRange(2, 2, bwaSheet.getLastRow() - 1, 1)
            .setNumberFormat("€#,##0.00;€-#,##0.00");
        bwaSheet.autoResizeColumns(1, 2);

        fehlendeKategorien.length > 0
            ? SpreadsheetApp.getUi().alert("Folgende Kategorien konnten nicht zugeordnet werden:\n" + fehlendeKategorien.join("\n"))
            : SpreadsheetApp.getUi().alert("BWA wurde erfolgreich berechnet und aktualisiert!");
    }

    return { calculateBWA };
})();


// ================= Globale Funktionen =================
const setupTrigger = () => Buchhaltung.setupTrigger();
const onOpen = () => Buchhaltung.onOpen();
const importDriveFiles = () => Buchhaltung.importDriveFiles();
const refreshSheets = () => Buchhaltung.refreshSheets();
const calculateGuV = () => GuVCalculator.calculateGuV();
const calculateBWA = () => BWACalculator.calculateBWA();
