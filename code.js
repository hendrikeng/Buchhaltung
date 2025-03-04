// =================== Zentrale Konfiguration ===================
// Definiert erlaubte Kategorien, Konto- und BWA-Mappings
const CategoryConfig = {
    einnahmen: {
        allowed: [
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
        allowed: [
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
        allowed: [
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
            "Zinsen"
        ],
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
            "Privateinlage": "privateinlage",
            "Darlehen": "darlehen",
            "Eigenbeleg": "eigenbeleg",
            "Privatentnahme": "privatentnahme",
            "Zinsen": "zinsen"
        }
    }
};

// =================== Modul: Helpers ===================
const Helpers = (() => {
    // Erstellt Dropdown-Validierung für ein Array
    const createDropdownValidation = list =>
        SpreadsheetApp.newDataValidation().requireValueInList(list, true).build();

    // Wandelt einen String oder Date in ein Date-Objekt um
    const parseDate = value => {
        const d = typeof value === "string" ? new Date(value) : value instanceof Date ? value : null;
        return d && !isNaN(d.getTime()) ? d : null;
    };

    // Parst einen Währungswert aus einem String
    const parseCurrency = value =>
        parseFloat(value.toString().replace(/[^\d,.-]/g, "").replace(",", ".")) || 0;

    // Parst einen Mehrwertsteuersatz aus String oder Zahl
    const parseMwstRate = value => {
        let rate =
            typeof value === "number"
                ? value < 1
                    ? value * 100
                    : value
                : parseFloat(value.toString().replace("%", "").replace(",", "."));
        return isNaN(rate) ? 19 : rate;
    };

    // Sucht in einem Ordner nach einem Unterordner mit dem gegebenen Namen
    const getFolderByName = (parent, name) => {
        const folderIter = parent.getFoldersByName(name);
        return folderIter.hasNext() ? folderIter.next() : null;
    };

    return { createDropdownValidation, parseDate, parseCurrency, parseMwstRate, getFolderByName };
})();

// =================== Modul: Validator ===================
const Validator = (() => {
    // Prüft, ob ein Wert leer ist
    const isEmpty = v => v == null || v.toString().trim() === "";

    // Validierung für Revenue & Expenses
    const validateRevenueAndExpenses = (row, rowIndex) => {
        const warnings = [];
        isEmpty(row[0]) && warnings.push(`Zeile ${rowIndex} (E/A): Rechnungsdatum fehlt.`);
        isEmpty(row[1]) && warnings.push(`Zeile ${rowIndex} (E/A): Rechnungsnummer fehlt.`);
        isEmpty(row[2]) && warnings.push(`Zeile ${rowIndex} (E/A): Kategorie fehlt.`);
        isEmpty(row[3]) && warnings.push(`Zeile ${rowIndex} (E/A): Kunde fehlt.`);
        (isEmpty(row[4]) || isNaN(parseFloat(row[4].toString().trim()))) &&
        warnings.push(`Zeile ${rowIndex} (E/A): Nettobetrag fehlt oder ungültig.`);
        const mwstStr = row[5] == null ? "" : row[5].toString().trim();
        (isEmpty(mwstStr) || isNaN(parseFloat(mwstStr.replace("%", "").replace(",", ".")))) &&
        warnings.push(`Zeile ${rowIndex} (E/A): Mehrwertsteuer fehlt oder ungültig.`);
        const status = row[11] ? row[11].toString().trim().toLowerCase() : "";
        const zahlungsdatum = row[12] ? row[12].toString().trim() : "";
        status === "offen"
            ? !isEmpty(zahlungsdatum) && warnings.push(`Zeile ${rowIndex} (E/A): Zahlungsdatum darf nicht gesetzt sein, wenn "offen".`)
            : isEmpty(zahlungsdatum) && warnings.push(`Zeile ${rowIndex} (E/A): Zahlungsdatum muss gesetzt sein, wenn bezahlt/teilbezahlt.`);
        return warnings;
    };

    // Validiert das Bank-Sheet und führt direkt das Mapping durch
    const validateBanking = bankSheet => {
        const data = bankSheet.getDataRange().getValues();
        let warnings = [];
        data.forEach((row, i) => {
            const idx = i + 1;
            if (i === 0 || i === data.length - 1) {
                isEmpty(row[0]) && warnings.push(`Zeile ${idx} (Bank): Buchungsdatum fehlt.`);
                isEmpty(row[1]) && warnings.push(`Zeile ${idx} (Bank): Buchungstext fehlt.`);
                (isEmpty(row[3]) || isNaN(parseFloat(row[3].toString().trim()))) &&
                warnings.push(`Zeile ${idx} (Bank): Saldo fehlt oder ungültig.`);
            } else {
                isEmpty(row[0]) && warnings.push(`Zeile ${idx} (Bank): Buchungsdatum fehlt.`);
                isEmpty(row[1]) && warnings.push(`Zeile ${idx} (Bank): Buchungstext fehlt.`);
                (isEmpty(row[2]) || isNaN(parseFloat(row[2].toString().trim()))) &&
                warnings.push(`Zeile ${idx} (Bank): Betrag fehlt oder ungültig.`);
                (isEmpty(row[3]) || isNaN(parseFloat(row[3].toString().trim()))) &&
                warnings.push(`Zeile ${idx} (Bank): Saldo fehlt oder ungültig.`);
                isEmpty(row[4]) && warnings.push(`Zeile ${idx} (Bank): Typ fehlt.`);
                isEmpty(row[5]) && warnings.push(`Zeile ${idx} (Bank): Kategorie fehlt.`);
            }
            // Konto-Mapping integrieren (ab Zeile 2, wenn Header in Zeile 1)
            if (i > 0) {
                const type = row[4],
                    cat = row[5];
                let mapping =
                    type === "Einnahme"
                        ? CategoryConfig.einnahmen.kontoMapping[cat]
                        : type === "Ausgabe"
                            ? CategoryConfig.ausgaben.kontoMapping[cat]
                            : null;
                !mapping && warnings.push(`Zeile ${idx} (Bank): Kein Konto-Mapping für Kategorie "${cat || "N/A"}" gefunden – bitte manuell zuordnen!`);
                mapping = mapping || { soll: "Manuell prüfen", gegen: "Manuell prüfen" };
                row[6] = mapping.soll;
                row[7] = mapping.gegen;
            }
        });
        // Aktualisiert das Sheet mit den Mapping-Werten
        bankSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
        return warnings;
    };

    // Aggregiert alle Warnungen aus Revenue, Ausgaben und optional Bank und gibt ggf. den Alert aus.
    // Gibt true zurück, wenn keine Fehler vorliegen, sonst false.
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

    return { validateRevenueAndExpenses, validateBanking, validateAllSheets };
})();

// =================== Modul: ImportModule ===================
const ImportModule = (() => {
    // Importiert Dateien aus einem Ordner in die entsprechenden Sheets
    const importFilesFromFolder = (folder, importSheet, mainSheet, type, historySheet) => {
        const files = folder.getFiles();
        // Ermittelt bereits importierte Dateien
        const getExistingFiles = (sheet, colIndex) =>
            new Set(sheet.getDataRange().getValues().slice(1).map(row => row[colIndex]));
        const existingMain = getExistingFiles(mainSheet, 1);
        const existingImport = getExistingFiles(importSheet, 0);
        const newMainRows = [];
        const newImportRows = [];
        const newHistoryRows = [];
        const timestamp = new Date();
        // Iteriert über alle Dateien im Ordner
        while (files.hasNext()) {
            const file = files.next();
            const fileName = file.getName().replace(/\.[^/.]+$/, "");
            const fileUrl = file.getUrl();
            let wasImported = false;
            if (!existingMain.has(fileName)) {
                const rowIndex = mainSheet.getLastRow() + newMainRows.length + 1;
                newMainRows.push([
                    timestamp,
                    fileName,
                    "",
                    "",
                    "",
                    "",
                    `=E${rowIndex}*F${rowIndex}`,
                    `=E${rowIndex}+G${rowIndex}`,
                    "",
                    `=E${rowIndex}-(I${rowIndex}-G${rowIndex})`,
                    `=IF(A${rowIndex}=""; ""; ROUNDUP(MONTH(A${rowIndex})/3;0))`,
                    `=IF(OR(I${rowIndex}=""; I${rowIndex}=0); "Offen"; IF(I${rowIndex}>=H${rowIndex}; "Bezahlt"; "Teilbezahlt"))`,
                    "",
                    fileName,
                    fileUrl,
                    timestamp
                ]);
                existingMain.add(fileName);
                wasImported = true;
            }
            if (!existingImport.has(fileName)) {
                newImportRows.push([fileName, fileUrl, fileName]);
                existingImport.add(fileName);
                wasImported = true;
            }
            // Fügt bei neu importierten Dateien einen History-Eintrag hinzu
            wasImported && newHistoryRows.push([timestamp, type, fileName, fileUrl]);
        }
        // Schreibt neue Zeilen in die jeweiligen Sheets
        if (newImportRows.length > 0)
            importSheet
                .getRange(importSheet.getLastRow() + 1, 1, newImportRows.length, newImportRows[0].length)
                .setValues(newImportRows);
        if (newMainRows.length > 0)
            mainSheet
                .getRange(mainSheet.getLastRow() + 1, 1, newMainRows.length, newMainRows[0].length)
                .setValues(newMainRows);
        if (newHistoryRows.length > 0)
            historySheet
                .getRange(historySheet.getLastRow() + 1, 1, newHistoryRows.length, newHistoryRows[0].length)
                .setValues(newHistoryRows);
    };

    // Startet den Importvorgang aus den Ordnern "Einnahmen" und "Ausgaben"
    const importDriveFiles = () => {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheets = {
            revenueMain: ss.getSheetByName("Einnahmen"),
            expenseMain: ss.getSheetByName("Ausgaben"),
            revenue: ss.getSheetByName("Rechnungen Einnahmen") || ss.insertSheet("Rechnungen Einnahmen"),
            expense: ss.getSheetByName("Rechnungen Ausgaben") || ss.insertSheet("Rechnungen Ausgaben"),
            history: ss.getSheetByName("Änderungshistorie") || ss.insertSheet("Änderungshistorie")
        };
        // Initialisiert Header, falls Sheets leer sind
        if (sheets.history.getLastRow() === 0)
            sheets.history.appendRow(["Datum", "Rechnungstyp", "Dateiname", "Link zur Datei"]);
        if (sheets.revenue.getLastRow() === 0)
            sheets.revenue.appendRow(["Dateiname", "Link zur Datei", "Rechnungsnummer"]);
        if (sheets.expense.getLastRow() === 0)
            sheets.expense.appendRow(["Dateiname", "Link zur Datei", "Rechnungsnummer"]);

        // Holt den übergeordneten Ordner der aktuellen Tabelle
        const file = DriveApp.getFileById(ss.getId());
        const parentFolder = file.getParents()?.hasNext() ? file.getParents().next() : null;
        if (!parentFolder) {
            SpreadsheetApp.getUi().alert("Kein übergeordneter Ordner gefunden.");
            return;
        }
        // Sucht gezielt nach den Ordnern "Einnahmen" und "Ausgaben"
        const revenueFolder = Helpers.getFolderByName(parentFolder, "Einnahmen");
        const expenseFolder = Helpers.getFolderByName(parentFolder, "Ausgaben");
        if (revenueFolder)
            importFilesFromFolder(revenueFolder, sheets.revenue, sheets.revenueMain, "Einnahme", sheets.history);
        else
            SpreadsheetApp.getUi().alert("Fehler: 'Einnahmen'-Ordner nicht gefunden.");
        if (expenseFolder)
            importFilesFromFolder(expenseFolder, sheets.expense, sheets.expenseMain, "Ausgabe", sheets.history);
        else
            SpreadsheetApp.getUi().alert("Fehler: 'Ausgaben'-Ordner nicht gefunden.");
    };

    return { importDriveFiles };
})();

// =================== Modul: RefreshModule ===================
const RefreshModule = (() => {
    // Aktualisiert Datenblätter (Einnahmen/Ausgaben): Formeln, Formate, Data Validation
    const refreshDataSheet = sheet => {
        const lastRow = sheet.getLastRow();
        if (lastRow < 2) return;
        const numRows = lastRow - 1;
        // Erstellt Formeln für jede Zeile
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
        // Setzt Formeln in die entsprechenden Spalten
        sheet.getRange(2, 7, numRows, 1).setFormulas(formulas.map(f => [f.mwst]));
        sheet.getRange(2, 8, numRows, 1).setFormulas(formulas.map(f => [f.brutto]));
        sheet.getRange(2, 10, numRows, 1).setFormulas(formulas.map(f => [f.rest]));
        sheet.getRange(2, 11, numRows, 1).setFormulas(formulas.map(f => [f.quartal]));
        sheet.getRange(2, 12, numRows, 1).setFormulas(formulas.map(f => [f.status]));

        // Setzt Datums- und Währungsformate
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

        // Setzt in Spalte 9 den Wert 0, falls leer
        for (let r = 2; r <= lastRow; r++) {
            const cell = sheet.getRange(r, 9);
            if (cell.getValue() === "" || cell.getValue() === null)
                cell.setValue(0);
        }

        // Setzt Data Validation für Kategorien je nach Blattname
        const sheetName = sheet.getName();
        if (sheetName === "Einnahmen") {
            sheet.getRange(2, 3, lastRow - 1, 1).setDataValidation(
                Helpers.createDropdownValidation(CategoryConfig.einnahmen.allowed)
            );
        } else if (sheetName === "Ausgaben") {
            sheet.getRange(2, 3, lastRow - 1, 1).setDataValidation(
                Helpers.createDropdownValidation(CategoryConfig.ausgaben.allowed)
            );
        }

        sheet.autoResizeColumns(1, sheet.getLastColumn());
    };

    // Aktualisiert das Bank-Sheet: Formeln, Formate, Data Validation und bedingte Formatierung
    const refreshBankSheet = sheet => {
        const lastRow = sheet.getLastRow();
        if (lastRow < 3) return;
        const firstDataRow = 3;
        const numTransRows = lastRow - 2;
        const transRows = numTransRows - 1;
        if (transRows > 0) {
            // Berechnet den neuen Saldo für jede Transaktion
            const saldoFormulas = Array.from({ length: transRows }, (_, i) => {
                const row = firstDataRow + i;
                return [`=D${row - 1}+C${row}`];
            });
            sheet.getRange(firstDataRow, 4, transRows, 1).setFormulas(saldoFormulas);
        }
        // Bestimmt den Transaktionstyp anhand des Betrags
        for (let row = firstDataRow; row <= lastRow; row++) {
            const amount = parseFloat(sheet.getRange(row, 3).getValue()) || 0;
            const typeCell = sheet.getRange(row, 5);
            if (amount > 0) typeCell.setValue("Einnahme");
            else if (amount < 0) typeCell.setValue("Ausgabe");
            else typeCell.clearContent();
        }
        // Setzt Data Validation für Typ und Kategorie
        sheet.getRange(firstDataRow, 5, lastRow - firstDataRow + 1, 1).setDataValidation(
            Helpers.createDropdownValidation(CategoryConfig.bank.typeAllowed)
        );
        sheet.getRange(firstDataRow, 6, lastRow - firstDataRow + 1, 1).setDataValidation(
            Helpers.createDropdownValidation(CategoryConfig.bank.allowed)
        );
        // Ermittelt erlaubte Konten aus den Mappings
        const allowedKontoSoll = Object.values(CategoryConfig.einnahmen.kontoMapping)
            .concat(Object.values(CategoryConfig.ausgaben.kontoMapping))
            .map(m => m.soll);
        const allowedGegenkonto = Object.values(CategoryConfig.einnahmen.kontoMapping)
            .concat(Object.values(CategoryConfig.ausgaben.kontoMapping))
            .map(m => m.gegen);
        sheet.getRange(firstDataRow, 7, lastRow - firstDataRow + 1, 1).setDataValidation(
            Helpers.createDropdownValidation(allowedKontoSoll)
        );
        sheet.getRange(firstDataRow, 8, lastRow - firstDataRow + 1, 1).setDataValidation(
            Helpers.createDropdownValidation(allowedGegenkonto)
        );

        // Setzt bedingte Formatierung für Typ (Einnahme/Ausgabe)
        const ruleEinnahme = SpreadsheetApp.newConditionalFormatRule()
            .whenTextEqualTo("Einnahme")
            .setBackground("#C6EFCE")
            .setFontColor("#006100")
            .setRanges([sheet.getRange(`E2:E${lastRow}`)])
            .build();
        const ruleAusgabe = SpreadsheetApp.newConditionalFormatRule()
            .whenTextEqualTo("Ausgabe")
            .setBackground("#FFC7CE")
            .setFontColor("#9C0006")
            .setRanges([sheet.getRange(`E2:E${lastRow}`)])
            .build();
        sheet.setConditionalFormatRules([ruleEinnahme, ruleAusgabe]);

        sheet.getRange("A2:A" + lastRow).setNumberFormat("DD.MM.YYYY");
        sheet.getRange("C2:C" + lastRow).setNumberFormat("€#,##0.00;€-#,##0.00");
        sheet.getRange("D2:D" + lastRow).setNumberFormat("€#,##0.00;€-#,##0.00");
        sheet.autoResizeColumns(1, sheet.getLastColumn());

        // Aktualisiert Endsaldo-Zeile oder fügt sie hinzu
        const lastRowText = sheet.getRange(lastRow, 2).getValue().toString().trim().toLowerCase();
        const formattedDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd.MM.yyyy");
        if (lastRowText === "endsaldo") {
            sheet.getRange(lastRow, 1).setValue(formattedDate);
            sheet.getRange(lastRow, 4).setFormula(`=D${lastRow - 1}`);
        } else {
            sheet.appendRow([formattedDate, "Endsaldo", "", sheet.getRange(lastRow, 4).getValue(), "", "", "", "", "", "", "", ""]);
        }
    };

    // Aktualisiert das aktuell aktive Blatt
    const refreshActiveSheet = () => {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getActiveSheet();
        const sheetName = sheet.getName();
        if (sheetName === "Einnahmen" || sheetName === "Ausgaben") {
            refreshDataSheet(sheet);
            SpreadsheetApp.getUi().alert(`Das Blatt "${sheetName}" wurde aktualisiert.`);
        } else if (sheetName === "Bankbewegungen") {
            refreshBankSheet(sheet);
            SpreadsheetApp.getUi().alert(`Das Blatt "${sheetName}" wurde aktualisiert.`);
        } else {
            SpreadsheetApp.getUi().alert("Für dieses Blatt gibt es keine Refresh-Funktion.");
        }
    };

    // Aktualisiert alle relevanten Blätter
    const refreshAllSheets = () => {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const revenueSheet = ss.getSheetByName("Einnahmen");
        const expenseSheet = ss.getSheetByName("Ausgaben");
        const bankSheet = ss.getSheetByName("Bankbewegungen");
        if (revenueSheet) refreshDataSheet(revenueSheet);
        if (expenseSheet) refreshDataSheet(expenseSheet);
        if (bankSheet) refreshBankSheet(bankSheet);
    };

    return { refreshActiveSheet, refreshAllSheets };
})();

// =================== Modul: GuV-Berechnung ===================
const GuVCalculator = (() => {
    // Erstellt ein leeres GuV-Objekt mit Standardwerten
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

    // Verarbeitet eine Zeile der GuV und fügt Werte hinzu
    const processGuVRow = (row, guvData, isIncome) => {
        const paymentDate = Helpers.parseDate(row[12]);
        if (!paymentDate || paymentDate > new Date()) return;
        const month = paymentDate.getMonth() + 1;
        const netto = Helpers.parseCurrency(row[4]);
        const restNetto = Helpers.parseCurrency(row[9]) || 0;
        const gezahltNetto = Math.max(0, netto - restNetto);
        const mwstRate = Helpers.parseMwstRate(row[5]);
        const tax = gezahltNetto * (mwstRate / 100);
        if (isIncome) {
            guvData[month].einnahmen += gezahltNetto;
            guvData[month][`ust_${Math.round(mwstRate)}`] += tax;
        } else {
            guvData[month].ausgaben += gezahltNetto;
            guvData[month][`vst_${Math.round(mwstRate)}`] += tax;
        }
    };

    // Formatiert eine GuV-Zeile für die Ausgabe
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

    // Aggregiert GuV-Daten über einen Zeitraum
    const aggregateGuV = (guvData, start, end) => {
        const sum = createEmptyGuV();
        for (let m = start; m <= end; m++) {
            for (const key in sum) {
                sum[key] += guvData[m][key] || 0;
            }
        }
        return sum;
    };

    // Berechnet die GuV und schreibt die Ergebnisse ins "GuV"-Sheet
    const calculateGuV = () => {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const revenueSheet = ss.getSheetByName("Einnahmen");
        const expenseSheet = ss.getSheetByName("Ausgaben");
        if (!revenueSheet || !expenseSheet) {
            SpreadsheetApp.getUi().alert("Fehlendes Blatt: 'Einnahmen' oder 'Ausgaben'");
            return;
        }
        // Validierung über den Validator (nur Revenue und Ausgaben)
        if (!Validator.validateAllSheets(revenueSheet, expenseSheet)) return;
        // Liest die Daten aus beiden Blättern
        const revenueData = revenueSheet.getDataRange().getValues().slice(1);
        const expenseData = expenseSheet.getDataRange().getValues().slice(1);
        const guvData = {};
        for (let m = 1; m <= 12; m++) {
            guvData[m] = createEmptyGuV();
        }
        // Verarbeitet alle Zeilen zur GuV-Berechnung
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
            "Januar",
            "Februar",
            "März",
            "April",
            "Mai",
            "Juni",
            "Juli",
            "August",
            "September",
            "Oktober",
            "November",
            "Dezember"
        ];
        // Fügt für jeden Monat und Quartal die GuV-Daten ein
        for (let m = 1; m <= 12; m++) {
            guvSheet.appendRow(formatGuVRow(monthNames[m - 1], guvData[m]));
            if (m % 3 === 0)
                guvSheet.appendRow(formatGuVRow(`Quartal ${m / 3}`, aggregateGuV(guvData, m - 2, m)));
        }
        const yearTotal = aggregateGuV(guvData, 1, 12);
        guvSheet.appendRow(formatGuVRow("Gesamtjahr", yearTotal));
        guvSheet.getRange(2, 2, guvSheet.getLastRow() - 1, 1).setNumberFormat("#,##0.00€");
        guvSheet.autoResizeColumns(1, guvSheet.getLastColumn());
        SpreadsheetApp.getUi().alert("GuV wurde aktualisiert!");
    };

    return { calculateGuV };
})();

// =================== Modul: BWACalculator ===================
const BWACalculator = (() => {
    // Ermittelt anhand der Kategorie das passende BWA-Mapping
    const getBwaCategory = (category, isIncome, rowIndex, fehlendeKategorien, type = "operativ") => {
        const mapping =
            type === "bank"
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

    // Berechnet die BWA und schreibt die Ergebnisse ins "BWA"-Sheet
    const calculateBWA = () => {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const revenueSheet = ss.getSheetByName("Einnahmen");
        const expenseSheet = ss.getSheetByName("Ausgaben");
        const bankSheet = ss.getSheetByName("Bankbewegungen");
        const bwaSheet = ss.getSheetByName("BWA") || ss.insertSheet("BWA");

        // Validierung aller beteiligten Sheets (Revenue, Ausgaben, Bank)
        if (!Validator.validateAllSheets(revenueSheet, expenseSheet, bankSheet)) return;

        // Initialisiert Kategorien-Summen
        const categories = {
            einnahmen: { umsatzerloese: 0, provisionserloese: 0, sonstigeErtraege: 0, privateinlage: 0, darlehen: 0, zinsen: 0 },
            ausgaben: { wareneinsatz: 0, betriebskosten: 0, marketing: 0, reisen: 0, personalkosten: 0, sonstigeAufwendungen: 0, eigenbeleg: 0, privatentnahme: 0, darlehen: 0, zinsen: 0 }
        };
        let totalEinnahmen = 0,
            totalAusgaben = 0;
        let offeneForderungen = 0,
            offeneVerbindlichkeiten = 0;
        let totalLiquiditaet = 0;
        const fehlendeKategorien = [];
        // Verarbeitet Einnahmen-Daten
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
        // Verarbeitet Ausgaben-Daten
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
        // Liest den letzten Saldo aus dem Bank-Sheet
        if (bankSheet) {
            const bankData = bankSheet.getDataRange().getValues();
            for (let i = 1; i < bankData.length - 1; i++) {
                const row = bankData[i];
                const saldo = parseFloat(row[3]) || 0;
                totalLiquiditaet = saldo;
            }
        }
        // Schreibt die BWA-Ergebnisse in das BWA-Sheet
        bwaSheet.clearContents();
        bwaSheet.appendRow(["Position", "Betrag (€)"]);
        bwaSheet.appendRow(["--- EINNAHMEN ---", ""]);
        Object.keys(categories.einnahmen).forEach(key =>
            bwaSheet.appendRow([key, categories.einnahmen[key]])
        );
        bwaSheet.appendRow(["Gesamteinnahmen", totalEinnahmen]);
        bwaSheet.appendRow(["Erhaltene Einnahmen", totalEinnahmen - offeneForderungen]);
        bwaSheet.appendRow(["Offene Forderungen", offeneForderungen]);
        bwaSheet.appendRow(["--- AUSGABEN ---", ""]);
        Object.keys(categories.ausgaben).forEach(key =>
            bwaSheet.appendRow([key, -categories.ausgaben[key]])
        );
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
        bwaSheet.getRange(2, 2, bwaSheet.getLastRow() - 1, 1).setNumberFormat("€#,##0.00;€-#,##0.00");
        bwaSheet.autoResizeColumns(1, bwaSheet.getLastColumn());
        // Zeigt Warnungen für nicht zugeordnete Kategorien an
        fehlendeKategorien.length
            ? SpreadsheetApp.getUi().alert("Folgende Kategorien konnten nicht zugeordnet werden:\n" + fehlendeKategorien.join("\n"))
            : SpreadsheetApp.getUi().alert("BWA wurde erfolgreich berechnet und aktualisiert!");
    };

    return { calculateBWA };
})();

// =================== Globale Funktionen ===================
// Erstellt das Menü und setzt Trigger
const onOpen = () => {
    SpreadsheetApp.getUi()
        .createMenu("📂 Buchhaltung")
        .addItem("📥 Dateien importieren", "importDriveFiles")
        .addItem("🔄 Refresh Active Sheet", "refreshSheet")
        .addItem("📊 GuV berechnen", "calculateGuV")
        .addItem("📈 BWA berechnen", "calculateBWA")
        .addToUi();
};

const setupTrigger = () => {
    const triggers = ScriptApp.getProjectTriggers();
    if (!triggers.some(t => t.getHandlerFunction() === "onOpen"))
        SpreadsheetApp.newTrigger("onOpen")
            .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
            .onOpen()
            .create();
};

// Aktualisiert nur das aktuell aktive Blatt
const refreshSheet = () => RefreshModule.refreshActiveSheet();
// Aktualisiert alle Blätter und berechnet die GuV
const calculateGuV = () => {
    RefreshModule.refreshAllSheets();
    GuVCalculator.calculateGuV();
};
// Aktualisiert alle Blätter und berechnet die BWA
const calculateBWA = () => {
    RefreshModule.refreshAllSheets();
    BWACalculator.calculateBWA();
};
// Importiert Dateien und aktualisiert alle Blätter
const importDriveFiles = () => {
    ImportModule.importDriveFiles();
    RefreshModule.refreshAllSheets();
};
