import Helpers from "./helpers.js";

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

export default ImportModule;