// =================== Zentrale Konfiguration ===================
// Konfiguration der erlaubten Kategorien, Konto- und BWA-Mappings
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
// Gemeinsame Hilfsfunktionen
const Helpers = (() => {

    // Konvertiert String oder Datum in ein Date-Objekt
    const parseDate = value => {
        const d = typeof value === "string" ? new Date(value) : value instanceof Date ? value : null;
        return d && !isNaN(d.getTime()) ? d : null;
    };

    // Wandelt einen String in einen Float um und entfernt dabei unerwünschte Zeichen
    const parseCurrency = value =>
        parseFloat(value.toString().replace(/[^\d,.-]/g, "").replace(",", ".")) || 0;

    // Extrahiert einen Mehrwertsteuersatz aus String oder Zahl
    const parseMwstRate = value => {
        let rate = typeof value === "number" ? (value < 1 ? value * 100 : value)
            : parseFloat(value.toString().replace("%", "").replace(",", "."));
        return isNaN(rate) ? 19 : rate;
    };

    // Sucht in einem Ordner nach einem Unterordner mit dem angegebenen Namen
    const getFolderByName = (parent, name) => {
        const folderIter = parent.getFoldersByName(name);
        return folderIter.hasNext() && folderIter.next();
    };

    return { parseDate, parseCurrency, parseMwstRate, getFolderByName };
})();

// =================== Modul: Validator ===================
// Verantwortlich für die Validierung von Daten und Ausgabe von Warnmeldungen
const Validator = (() => {

    // Prüft, ob ein Wert leer ist
    const isEmpty = v => v == null || v.toString().trim() === "";

    // Setzt die Dropdown-Validierung auf einen bestimmten Range
    // Parameter: sheet, Startzeile, Startspalte, Anzahl der Zeilen, Anzahl der Spalten, Liste der erlaubten Werte
    const validateDropdown = (sheet, row, col, numRows, numCols, list) => {
        sheet.getRange(row, col, numRows, numCols).setDataValidation(
            SpreadsheetApp.newDataValidation().requireValueInList(list, true).build()
        );
    };

    // Validiert eine Zeile aus Revenue oder Ausgaben
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

    // Validiert das Bank-Sheet, führt das Konto-Mapping durch und schreibt die Änderungen zurück
    const validateBanking = bankSheet => {
        const data = bankSheet.getDataRange().getValues();
        const warnings = data.reduce((acc, row, i) => {
            const idx = i + 1;
            // Für Header und Endzeile
            if (i === 1 || i === data.length - 1) {
                isEmpty(row[0]) && acc.push(`Zeile ${idx} (Bank): Buchungsdatum fehlt.`);
                isEmpty(row[1]) && acc.push(`Zeile ${idx} (Bank): Buchungstext fehlt.`);
                (!isEmpty(row[2]) || !isNaN(parseFloat(row[2].toString().trim()))) &&
                acc.push(`Zeile ${idx} (Bank): Betrag darf nicht gesetzt sein.`);
                (isEmpty(row[3]) || isNaN(parseFloat(row[3].toString().trim()))) &&
                acc.push(`Zeile ${idx} (Bank): Saldo fehlt oder ungültig.`);
                !isEmpty(row[4]) && acc.push(`Zeile ${idx} (Bank): Typ darf nicht gesetzt sein.`);
                !isEmpty(row[5]) && acc.push(`Zeile ${idx} (Bank): Kategorie darf nicht gesetzt sein.`);
                !isEmpty(row[6]) && acc.push(`Zeile ${idx} (Bank): Konto (Soll) darf nicht gesetzt sein.`);
                !isEmpty(row[7]) && acc.push(`Zeile ${idx} (Bank): Gegenkonto (Haben) darf nicht gesetzt sein.`);
                // für alle Zeilen außer der ersten und letzten
            } else if (i > 1 && i < data.length - 1) {
                isEmpty(row[0]) && acc.push(`Zeile ${idx} (Bank): Buchungsdatum fehlt.`);
                isEmpty(row[1]) && acc.push(`Zeile ${idx} (Bank): Buchungstext fehlt.`);
                (isEmpty(row[2]) || isNaN(parseFloat(row[2].toString().trim()))) &&
                acc.push(`Zeile ${idx} (Bank): Betrag fehlt oder ungültig.`);
                (isEmpty(row[3]) || isNaN(parseFloat(row[3].toString().trim()))) &&
                acc.push(`Zeile ${idx} (Bank): Saldo fehlt oder ungültig.`);
                isEmpty(row[4]) && acc.push(`Zeile ${idx} (Bank): Typ fehlt.`);
                isEmpty(row[5]) && acc.push(`Zeile ${idx} (Bank): Kategorie fehlt.`);
            }
            // Führe das Konto-Mapping für alle Zeilen außer der ersten und letzten durch
            if (i > 1 && i < data.length - 1) {
                const [ , , , , type, cat ] = row;
                let mapping = type === "Einnahme"
                    ? CategoryConfig.einnahmen.kontoMapping[cat]
                    : type === "Ausgabe"
                        ? CategoryConfig.ausgaben.kontoMapping[cat]
                        : null;
                !mapping && acc.push(`Zeile ${idx} (Bank): Kein Konto-Mapping für Kategorie "${cat || "N/A"}" gefunden – bitte manuell zuordnen!`);
                mapping = mapping || { soll: "Manuell prüfen", gegen: "Manuell prüfen" };
                row[6] = mapping.soll;
                row[7] = mapping.gegen;
            }
            return acc;
        }, []);
        // Schreibe das modifizierte Array (mit Mapping) einmalig zurück ins Sheet
        bankSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
        return warnings;
    };

    // Aggregiert Warnungen aus Revenue, Ausgaben und optional Bank und gibt einen Alert aus, falls nötig
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
        const existingMain = getExistingFiles(mainSheet, 1);
        const existingImport = getExistingFiles(importSheet, 0);
        const newMainRows = [], newImportRows = [], newHistoryRows = [];
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
                    `=E${rowIndex}*F${rowIndex}`, `=E${rowIndex}+G${rowIndex}`,
                    "", `=E${rowIndex}-(I${rowIndex}-G${rowIndex})`,
                    `=IF(A${rowIndex}=""; ""; ROUNDUP(MONTH(A${rowIndex})/3;0))`,
                    `=IF(OR(I${rowIndex}=""; I${rowIndex}=0); "Offen"; IF(I${rowIndex}>=H${rowIndex}; "Bezahlt"; "Teilbezahlt"))`,
                    "", fileName, fileUrl, timestamp
                ]);
                existingMain.add(fileName); wasImported = true;
            }
            if (!existingImport.has(fileName)) {
                newImportRows.push([fileName, fileUrl, fileName]);
                existingImport.add(fileName); wasImported = true;
            }
            wasImported && newHistoryRows.push([timestamp, type, fileName, fileUrl]);
        }
        newImportRows.length && importSheet.getRange(importSheet.getLastRow() + 1, 1, newImportRows.length, newImportRows[0].length).setValues(newImportRows);
        newMainRows.length && mainSheet.getRange(mainSheet.getLastRow() + 1, 1, newMainRows.length, newMainRows[0].length).setValues(newMainRows);
        newHistoryRows.length && historySheet.getRange(historySheet.getLastRow() + 1, 1, newHistoryRows.length, newHistoryRows[0].length).setValues(newHistoryRows);
    };

    const importDriveFiles = () => {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheets = {
            revenueMain: ss.getSheetByName("Einnahmen"),
            expenseMain: ss.getSheetByName("Ausgaben"),
            revenue: ss.getSheetByName("Rechnungen Einnahmen") || ss.insertSheet("Rechnungen Einnahmen"),
            expense: ss.getSheetByName("Rechnungen Ausgaben") || ss.insertSheet("Rechnungen Ausgaben"),
            history: ss.getSheetByName("Änderungshistorie") || ss.insertSheet("Änderungshistorie")
        };
        // Initialisiere Header, falls die Sheets leer sind
        if (sheets.history.getLastRow() === 0)
            sheets.history.appendRow(["Datum", "Rechnungstyp", "Dateiname", "Link zur Datei"]);
        if (sheets.revenue.getLastRow() === 0)
            sheets.revenue.appendRow(["Dateiname", "Link zur Datei", "Rechnungsnummer"]);
        if (sheets.expense.getLastRow() === 0)
            sheets.expense.appendRow(["Dateiname", "Link zur Datei", "Rechnungsnummer"]);
        const file = DriveApp.getFileById(ss.getId());
        const parentFolder = file.getParents()?.hasNext() && file.getParents().next();
        if (!parentFolder) {
            SpreadsheetApp.getUi().alert("Kein übergeordneter Ordner gefunden.");
            return;
        }
        const revenueFolder = Helpers.getFolderByName(parentFolder, "Einnahmen");
        const expenseFolder = Helpers.getFolderByName(parentFolder, "Ausgaben");
        revenueFolder
            ? importFilesFromFolder(revenueFolder, sheets.revenue, sheets.revenueMain, "Einnahme", sheets.history)
            : SpreadsheetApp.getUi().alert("Fehler: 'Einnahmen'-Ordner nicht gefunden.");
        expenseFolder
            ? importFilesFromFolder(expenseFolder, sheets.expense, sheets.expenseMain, "Ausgabe", sheets.history)
            : SpreadsheetApp.getUi().alert("Fehler: 'Ausgaben'-Ordner nicht gefunden.");
    };

    return { importDriveFiles };
})();

// =================== Modul: RefreshModule ===================
// Aktualisiert Datenblätter (Formeln, Formate, Validierungen) und das Bank-Sheet separat
const RefreshModule = (() => {
    // Aktualisiert Einnahmen/Ausgaben-Blätter
    const refreshDataSheet = sheet => {
        const lastRow = sheet.getLastRow();
        if (lastRow < 2) return;
        const numRows = lastRow - 1;
        // Erzeuge Formeln für alle Zeilen
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
        // Setze Formeln in die entsprechenden Spalten
        sheet.getRange(2, 7, numRows, 1).setFormulas(formulas.map(f => [f.mwst]));
        sheet.getRange(2, 8, numRows, 1).setFormulas(formulas.map(f => [f.brutto]));
        sheet.getRange(2, 10, numRows, 1).setFormulas(formulas.map(f => [f.rest]));
        sheet.getRange(2, 11, numRows, 1).setFormulas(formulas.map(f => [f.quartal]));
        sheet.getRange(2, 12, numRows, 1).setFormulas(formulas.map(f => [f.status]));
        // Setze Datums- und Währungsformate
        const dateCols = ["A", "M", "P"], currencyCols = ["E", "G", "H", "I", "J"];
        dateCols.forEach(col => sheet.getRange(`${col}2:${col}${lastRow}`).setNumberFormat("DD.MM.YYYY"));
        currencyCols.forEach(col => sheet.getRange(`${col}2:${col}${lastRow}`).setNumberFormat("€#,##0.00;€-#,##0.00"));
        sheet.getRange(`F2:F${lastRow}`).setNumberFormat("0.00%");
        // Setze 0 in Spalte 9, falls leer
        for (let r = 2; r <= lastRow; r++) {
            const cell = sheet.getRange(r, 9);
            if (cell.getValue() === "" || cell.getValue() === null) cell.setValue(0);
        }
        // Setze Data Validation für Kategorien
        const name = sheet.getName();
        if (name === "Einnahmen")
            Validator.validateDropdown(sheet, 2, 3, lastRow - 1, 1, CategoryConfig.einnahmen.allowed);
        else if (name === "Ausgaben")
            Validator.validateDropdown(sheet, 2, 3, lastRow - 1, 1, CategoryConfig.ausgaben.allowed);
        sheet.autoResizeColumns(1, sheet.getLastColumn());
    };

    // Aktualisiert das Bank-Sheet (inkl. bedingter Formatierung und Endsaldo)
    const refreshBankSheet = sheet => {
        const lastRow = sheet.getLastRow();
        if (lastRow < 3) return;
        const firstDataRow = 3, transRows = lastRow - firstDataRow - 1;
        if (transRows > 0) {
            sheet.getRange(firstDataRow, 4, transRows, 1).setFormulas(
                Array.from({ length: transRows }, (_, i) => [`=D${firstDataRow + i - 1}+C${firstDataRow + i}`])
            );
        }
        // Setze Transaktionstyp basierend auf Betrag
        for (let row = firstDataRow; row <= lastRow; row++) {
            const amount = parseFloat(sheet.getRange(row, 3).getValue()) || 0;
            const typeCell = sheet.getRange(row, 5);
            amount > 0 ? typeCell.setValue("Einnahme") : amount < 0 ? typeCell.setValue("Ausgabe") : typeCell.clearContent();
        }
        // Data Validation für Typ und Kategorie
        Validator.validateDropdown(sheet, firstDataRow, 5, lastRow - firstDataRow + 1, 1, CategoryConfig.bank.typeAllowed);
        Validator.validateDropdown(sheet, firstDataRow, 6, lastRow - firstDataRow + 1, 1, CategoryConfig.bank.allowed);
        // Erzeuge erlaubte Listen für Konto-Mapping
        const allowedKontoSoll = Object.values(CategoryConfig.einnahmen.kontoMapping)
            .concat(Object.values(CategoryConfig.ausgaben.kontoMapping))
            .map(m => m.soll);
        const allowedGegenkonto = Object.values(CategoryConfig.einnahmen.kontoMapping)
            .concat(Object.values(CategoryConfig.ausgaben.kontoMapping))
            .map(m => m.gegen);
        Validator.validateDropdown(sheet, firstDataRow, 7, lastRow - firstDataRow + 1, 1, allowedKontoSoll);
        Validator.validateDropdown(sheet, firstDataRow, 8, lastRow - firstDataRow + 1, 1, allowedGegenkonto);
        // Bedingte Formatierung für Einnahme/Ausgabe
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
        // Setze Zahlenformate für Spalten A, C und D
        sheet.getRange("A2:A" + lastRow).setNumberFormat("DD.MM.YYYY");
        ["C", "D"].forEach(col => sheet.getRange(`${col}2:${col}${lastRow}`).setNumberFormat("€#,##0.00;€-#,##0.00"));
        // Aktualisiere Endsaldo-Zeile oder füge diese hinzu
        const lastRowText = sheet.getRange(lastRow, 2).getValue().toString().trim().toLowerCase();
        const formattedDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd.MM.yyyy");
        lastRowText === "endsaldo"
            ? (sheet.getRange(lastRow, 1).setValue(formattedDate),
                sheet.getRange(lastRow, 4).setFormula(`=D${lastRow - 1}`))
            : sheet.appendRow([formattedDate, "Endsaldo", "", sheet.getRange(lastRow, 4).getValue(), "", "", "", "", "", "", "", ""]);
        sheet.autoResizeColumns(1, sheet.getLastColumn());
    };

    // Aktualisiert das aktuell aktive Blatt basierend auf dessen Namen
    const refreshActiveSheet = () => {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getActiveSheet();
        const name = sheet.getName();
        if (["Einnahmen", "Ausgaben"].includes(name)) {
            refreshDataSheet(sheet);
            SpreadsheetApp.getUi().alert(`Das Blatt "${name}" wurde aktualisiert.`);
        } else if (name === "Bankbewegungen") {
            refreshBankSheet(sheet);
            SpreadsheetApp.getUi().alert(`Das Blatt "${name}" wurde aktualisiert.`);
        } else {
            SpreadsheetApp.getUi().alert("Für dieses Blatt gibt es keine Refresh-Funktion.");
        }
    };

    // Aktualisiert alle relevanten Blätter
    const refreshAllSheets = () => {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        ["Einnahmen", "Ausgaben", "Bankbewegungen"].forEach(name => {
            const sheet = ss.getSheetByName(name);
            if (sheet) name === "Bankbewegungen" ? refreshBankSheet(sheet) : refreshDataSheet(sheet);
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
    // Ermittelt anhand der Kategorie das passende BWA-Mapping;
    // Bei unbekannten Kategorien wird ein Fallback-Wert verwendet und eine Warnung gesammelt.
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

    // Hauptfunktion zur Berechnung der BWA und Ausgabe ins "BWA"-Sheet
    const calculateBWA = () => {
        const ss = SpreadsheetApp.getActiveSpreadsheet(),
            revenueSheet = ss.getSheetByName("Einnahmen"),
            expenseSheet = ss.getSheetByName("Ausgaben"),
            bankSheet = ss.getSheetByName("Bankbewegungen"),
            bwaSheet = ss.getSheetByName("BWA") || ss.insertSheet("BWA");
        // Validierung aller beteiligten Datenblätter
        if (!Validator.validateAllSheets(revenueSheet, expenseSheet, bankSheet)) return;
        bankSheet.autoResizeColumns(1, sheet.getLastColumn());
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
