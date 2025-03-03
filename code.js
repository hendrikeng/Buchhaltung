// ------------------ Modul: Buchhaltung ------------------
const Buchhaltung = (() => {
    // Erzeugt den onOpen-Trigger, falls nicht vorhanden
    const setupTrigger = () => {
        const triggers = ScriptApp.getProjectTriggers();
        if (!triggers.some(t => t.getHandlerFunction() === "onOpen")) {
            ScriptApp.newTrigger("onOpen")
                .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
                .onOpen()
                .create();
            Logger.log("onOpen Trigger wurde erfolgreich erstellt!");
        }
    };

    // Erzeugt das Menü im Spreadsheet
    const onOpen = () => {
        SpreadsheetApp.getUi()
            .createMenu("📂 Buchhaltung")
            .addItem("📥 Rechnungen importieren", "importDriveFiles")
            .addItem("🔄 Aktualisieren (Formeln & Formatierung)", "refreshSheets")
            .addItem("📊 GuV berechnen", "calculateGuV")
            .addItem("📈 BWA berechnen", "calculateBWA")
            .addToUi();
    };

    // Importiert Rechnungen aus den Unterordnern "Einnahmen" und "Ausgaben"
    const importDriveFiles = () => {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheets = {
            revenue: ss.getSheetByName("Rechnungen Einnahmen"),
            expense: ss.getSheetByName("Rechnungen Ausgaben"),
            revenueMain: ss.getSheetByName("Einnahmen"),
            expenseMain: ss.getSheetByName("Ausgaben"),
            history: ss.getSheetByName("Änderungshistorie") || ss.insertSheet("Änderungshistorie")
        };

        // Kopfzeile für Änderungshistorie setzen, falls leer
        if (sheets.history.getLastRow() === 0) {
            sheets.history.appendRow(["Datum", "Rechnungstyp", "Dateiname", "Link zur Datei"]);
        }
        // Überschriften in Rechnungs-Sheets setzen, falls leer
        if (sheets.revenue.getLastRow() === 0) {
            sheets.revenue.appendRow(["Dateiname", "Link zur Datei", "Rechnungsnummer"]);
        }
        if (sheets.expense.getLastRow() === 0) {
            sheets.expense.appendRow(["Dateiname", "Link zur Datei", "Rechnungsnummer"]);
        }

        // Bestimme den übergeordneten Ordner des Sheets
        const file = DriveApp.getFileById(ss.getId());
        const parentFolder = file.getParents()?.hasNext() ? file.getParents().next() : null;
        if (!parentFolder) {
            Logger.log("Kein übergeordneter Ordner gefunden.");
            return;
        }

        // Suche die Unterordner "Einnahmen" und "Ausgaben"
        const revenueFolder = getFolderByName(parentFolder, "Einnahmen");
        const expenseFolder = getFolderByName(parentFolder, "Ausgaben");

        if (revenueFolder) {
            importFiles(revenueFolder, sheets.revenue, sheets.revenueMain, "Einnahme", sheets.history);
            sheets.revenue.autoResizeColumns(1, sheets.revenue.getLastColumn());
            if (sheets.revenueMain.getLastRow() > 1) refreshSheet(sheets.revenueMain);
        } else {
            Logger.log("Fehler: Der 'Einnahmen'-Ordner wurde nicht gefunden.");
        }

        if (expenseFolder) {
            importFiles(expenseFolder, sheets.expense, sheets.expenseMain, "Ausgabe", sheets.history);
            sheets.expense.autoResizeColumns(1, sheets.expense.getLastColumn());
            if (sheets.expenseMain.getLastRow() > 1) refreshSheet(sheets.expenseMain);
        } else {
            Logger.log("Fehler: Der 'Ausgaben'-Ordner wurde nicht gefunden.");
        }
    };

    // Importiert Dateien aus einem Ordner in die angegebenen Sheets und aktualisiert die Historie
    const importFiles = (folder, importSheet, mainSheet, type, historySheet) => {
        const files = folder.getFiles();

        // Liefert vorhandene Dateinamen aus der gewünschten Spalte als Set zurück
        const getExistingFiles = (sheet, colIndex) =>
            new Set(sheet.getDataRange().getValues().slice(1).map(row => row[colIndex]));

        const existingMain = getExistingFiles(mainSheet, 1); // Spalte B
        const existingImport = getExistingFiles(importSheet, 0); // Spalte A

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
                    timestamp,                   // Datum
                    fileName,                    // Dateiname
                    "",                          // Kategorie
                    "",                          // Kunde
                    "",                          // Nettobetrag
                    "",                          // Prozentsatz
                    `=E${rowIndex}*F${rowIndex}`, // MwSt.
                    `=E${rowIndex}+G${rowIndex}`, // Bruttobetrag
                    "",                          // Währung
                    `=E${rowIndex}-(I${rowIndex}-G${rowIndex})`, // Restbetrag
                    `=IF(A${rowIndex}=""; ""; ROUNDUP(MONTH(A${rowIndex})/3;0))`, // Quartal
                    `=IF(OR(I${rowIndex}=""; I${rowIndex}=0); "Offen"; IF(I${rowIndex}>=H${rowIndex}; "Bezahlt"; "Teilbezahlt"))`, // Zahlungsstatus
                    "",                          // Datum (wird formatiert)
                    fileName,                    // Wiederholter Dateiname
                    fileUrl,                     // Link zur Datei
                    timestamp                  // Letzte Aktualisierung
                ]);
                existingMain.add(fileName);
                wasImported = true;
            }

            if (!existingImport.has(fileName)) {
                newImportRows.push([fileName, fileUrl, fileName]);
                existingImport.add(fileName);
                wasImported = true;
            }

            if (wasImported) newHistoryRows.push([timestamp, type, fileName, fileUrl]);
        }

        if (newImportRows.length > 0)
            importSheet.getRange(importSheet.getLastRow() + 1, 1, newImportRows.length, newImportRows[0].length).setValues(newImportRows);
        if (newMainRows.length > 0)
            mainSheet.getRange(mainSheet.getLastRow() + 1, 1, newMainRows.length, newMainRows[0].length).setValues(newMainRows);
        if (newHistoryRows.length > 0)
            historySheet.getRange(historySheet.getLastRow() + 1, 1, newHistoryRows.length, newHistoryRows[0].length).setValues(newHistoryRows);
    };

    // Sucht in einem Ordner nach einem Unterordner mit dem angegebenen Namen
    const getFolderByName = (parent, name) => {
        const folderIter = parent.getFoldersByName(name);
        return folderIter.hasNext() ? folderIter.next() : null;
    };

    // Aktualisiert ein einzelnes Sheet (Formeln und Formatierung)
    const refreshSheet = (sheet) => {
        const lastRow = sheet.getLastRow();
        if (lastRow < 2) return;
        const numRows = lastRow - 1;

        // Erstelle ein Array von Formeln für jede Zeile (Zeile 2 bis lastRow)
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

        // Formatierungen
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
        sheet.autoResizeColumns(1, sheet.getLastColumn());
    };

    // Aktualisiert beide Hauptsheets "Einnahmen" und "Ausgaben"
    const refreshSheets = () => {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const revenueSheet = ss.getSheetByName("Einnahmen");
        const expenseSheet = ss.getSheetByName("Ausgaben");

        if (!revenueSheet || !expenseSheet) {
            SpreadsheetApp.getUi().alert("Eines der Blätter 'Einnahmen' oder 'Ausgaben' wurde nicht gefunden.");
            return;
        }

        refreshSheet(revenueSheet);
        refreshSheet(expenseSheet);
        SpreadsheetApp.getUi().alert("Einnahmen und Ausgaben wurden erfolgreich aktualisiert!");
    };

    return {
        setupTrigger,
        onOpen,
        importDriveFiles,
        refreshSheets,
        refreshSheet
    };
})();

// ------------------ Modul: GuV-Berechnung ------------------
const GuVCalculator = (() => {
    // Parst einen Mehrwertsteuersatz
    const parseMwstRate = value => {
        let rate = parseFloat(value.toString().replace("%", "").replace(",", "."));
        if (rate < 1) rate *= 100;
        return isNaN(rate) ? 19 : rate;
    };

    // Liefert ein leeres GuV-Datenobjekt
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

    // Parst ein Datum
    const parseDate = value => {
        const d = typeof value === "string" ? new Date(value) : (value instanceof Date ? value : null);
        return d && !isNaN(d.getTime()) ? d : null;
    };

    // Parst einen Währungswert
    const parseCurrency = value =>
        parseFloat(value.toString().replace(/[^\d,.-]/g, "").replace(",", ".")) || 0;

    // Verarbeitet eine Zeile (GuV) und aktualisiert das Monatsobjekt
    const processGuVRow = (row, guvData, isIncome) => {
        const paymentDate = parseDate(row[12]);
        if (!paymentDate || paymentDate > new Date()) return;
        const month = paymentDate.getMonth() + 1;
        const netto = parseCurrency(row[4]);
        const restNetto = parseCurrency(row[9]);
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

    // Formatiert eine Zeile für das GuV-Sheet
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

    // Aggregiert die GuV-Daten über den angegebenen Zeitraum
    const aggregateGuV = (guvData, start, end) => {
        const sum = createEmptyGuV();
        for (let m = start; m <= end; m++) {
            for (const key in sum) {
                sum[key] += guvData[m][key] || 0;
            }
        }
        return sum;
    };

    // Liest die Daten aus den Sheets "Einnahmen" und "Ausgaben", verarbeitet sie und schreibt die Ergebnisse ins "GuV"-Sheet
    const calculateGuV = () => {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const revenueSheet = ss.getSheetByName("Einnahmen");
        const expenseSheet = ss.getSheetByName("Ausgaben");

        if (!revenueSheet || !expenseSheet) {
            SpreadsheetApp.getUi().alert("Fehlendes Blatt: 'Einnahmen' oder 'Ausgaben'");
            return;
        }

        const revenueData = revenueSheet.getDataRange().getValues().slice(1);
        const expenseData = expenseSheet.getDataRange().getValues().slice(1);
        const guvData = {};
        // Initialisiere GuV-Daten für jeden Monat (1 bis 12)
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
            if (m % 3 === 0) {
                const quarter = aggregateGuV(guvData, m - 2, m);
                guvSheet.appendRow(formatGuVRow(`Quartal ${m / 3}`, quarter));
            }
        }
        const yearTotal = aggregateGuV(guvData, 1, 12);
        guvSheet.appendRow(formatGuVRow("Gesamtjahr", yearTotal));

        guvSheet.getRange(2, 2, guvSheet.getLastRow() - 1, 9).setNumberFormat("#,##0.00€");
        guvSheet.autoResizeColumns(1, guvSheet.getLastColumn());
        SpreadsheetApp.getUi().alert("GuV wurde aktualisiert!");
    };

    return { calculateGuV };
})();

// ------------------ Globale Funktionen (für Trigger & Menüs) ------------------
const setupTrigger = () => Buchhaltung.setupTrigger();
const onOpen = () => Buchhaltung.onOpen();
const importDriveFiles = () => Buchhaltung.importDriveFiles();
const refreshSheets = () => Buchhaltung.refreshSheets();
const calculateGuV = () => GuVCalculator.calculateGuV();

// Dummy-Funktion für BWA-Berechnung
const calculateBWA = () => {
    SpreadsheetApp.getUi().alert("BWA-Berechnung ist noch nicht implementiert.");
};
