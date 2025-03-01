function setupTrigger() {
    const triggers = ScriptApp.getProjectTriggers();
    const triggerExists = triggers.some(trigger => trigger.getHandlerFunction() === "onOpen");

    if (!triggerExists) {
        ScriptApp.newTrigger("onOpen")
            .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
            .onOpen()
            .create();
        Logger.log("onOpen Trigger wurde erfolgreich erstellt!");
    }
}

function onOpen() {
    const ui = SpreadsheetApp.getUi();

    // Submenü für globale Funktionen (Rechnungen importieren)
    const rechnungenMenu = ui.createMenu("🧾 Rechnungen");
    rechnungenMenu.addItem("📥 Rechnungen importieren", "importDriveFiles");

    // Submenü für Funktionen, die nur für "Einnahmen" und "Ausgaben" gelten
    const einnahmenAusgabenMenu = ui.createMenu("💰 Einnahmen/Ausgaben");
    einnahmenAusgabenMenu.addItem("🔀 Tabelle sortieren", "sortCurrentSheetByColumn");
    einnahmenAusgabenMenu.addItem("🔄 Aktualisieren (Formeln & Formatierung)", "updateSheetOnCurrentSheet");

    // Submenü für GUV
    const guvMenu = ui.createMenu("📊 GUV");
    guvMenu.addItem("✅ GUV berechnen", "calculateGUV");

    // Hauptmenü "Buchhaltung" mit den drei Submenüs
    ui.createMenu("📂 Buchhaltung")
        .addSubMenu(rechnungenMenu)
        .addSubMenu(einnahmenAusgabenMenu)
        .addSubMenu(guvMenu)
        .addToUi();
}

/* ----- Globale Funktionen für den Rechnungsimport ----- */
function importDriveFiles() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetEinnahmen = ss.getSheetByName("Rechnungen Einnahmen");
    const sheetAusgaben = ss.getSheetByName("Rechnungen Ausgaben");
    const einnahmenTab = ss.getSheetByName("Einnahmen");
    const ausgabenTab = ss.getSheetByName("Ausgaben");
    let historyTab = ss.getSheetByName("Änderungshistorie");

    // Änderungshistorie erstellen, falls nicht vorhanden
    if (!historyTab) {
        historyTab = ss.insertSheet("Änderungshistorie");
        historyTab.appendRow(["Datum", "Rechnungstyp", "Dateiname", "Link zur Datei"]);
    }

    // Überschriften in den Rechnungs-Sheets setzen, falls leer
    if (sheetEinnahmen.getLastRow() === 0) {
        sheetEinnahmen.appendRow(["Dateiname", "Link zur Datei", "Rechnungsnummer"]);
    }
    if (sheetAusgaben.getLastRow() === 0) {
        sheetAusgaben.appendRow(["Dateiname", "Link zur Datei", "Rechnungsnummer"]);
    }

    // Hauptordner ermitteln (Ordner, in dem sich das aktuelle Sheet befindet)
    const fileId = ss.getId();
    const file = DriveApp.getFileById(fileId);
    const parentIterator = file.getParents();
    if (!parentIterator.hasNext()) {
        Logger.log("Kein übergeordneter Ordner gefunden.");
        return;
    }
    const parentFolder = parentIterator.next();

    // Unterordner "Einnahmen" und "Ausgaben" suchen
    const einnahmenOrdner = getFolderByName(parentFolder, "Einnahmen");
    const ausgabenOrdner = getFolderByName(parentFolder, "Ausgaben");

    if (einnahmenOrdner) {
        importFilesFromFolder(einnahmenOrdner, sheetEinnahmen, einnahmenTab, "Einnahme", historyTab);
        if (einnahmenTab.getLastRow() > 1) {
            // Nach dem Import: nur Formeln und Formatierung aktualisieren (keine Sortierung)
            updateFormulasOnSheet(einnahmenTab);
            applyFormatting(einnahmenTab);
        }
    } else {
        Logger.log("Fehler: Der 'Einnahmen'-Ordner wurde nicht gefunden.");
    }

    if (ausgabenOrdner) {
        importFilesFromFolder(ausgabenOrdner, sheetAusgaben, ausgabenTab, "Ausgabe", historyTab);
        if (ausgabenTab.getLastRow() > 1) {
            updateFormulasOnSheet(ausgabenTab);
            applyFormatting(ausgabenTab);
        }
    } else {
        Logger.log("Fehler: Der 'Ausgaben'-Ordner wurde nicht gefunden.");
    }
}

function importFilesFromFolder(folder, importSheet, mainSheet, type, historyTab) {
    const files = folder.getFiles();

    // Vorhandene Dateinamen in Hauptblatt und Import-Sheet ermitteln:
    const existingMainData = mainSheet.getDataRange().getValues();
    const existingMainFiles = new Set(existingMainData.slice(1).map(row => row[1])); // Spalte 2

    const existingImportData = importSheet.getDataRange().getValues();
    const existingImportFiles = new Set(existingImportData.slice(1).map(row => row[0])); // Spalte 1

    const newImportRows = [];
    const newMainRows = [];
    const newHistoryRows = [];

    // Startindex für neue Zeilen im Hauptblatt:
    const mainStartRow = mainSheet.getLastRow() + 1;

    while (files.hasNext()) {
        const file = files.next();
        const fileNameWithExtension = file.getName();
        const fileName = fileNameWithExtension.replace(/\.[^/.]+$/, ""); // Dateiendung entfernen
        const fileUrl = file.getUrl();
        const timestamp = new Date();
        let wasImported = false;

        // Eintrag im Hauptblatt hinzufügen, wenn noch nicht vorhanden:
        if (!existingMainFiles.has(fileName)) {
            const currentRow = mainStartRow + newMainRows.length;
            newMainRows.push([
                timestamp,      // Datum
                fileName,       // Dateiname
                "", "", "",     // Leere Felder
                `=D${currentRow}*E${currentRow}`, // MwSt.-Betrag
                `=D${currentRow}+F${currentRow}`, // Bruttobetrag
                "",             // Leeres Feld
                `=D${currentRow}-(H${currentRow}-F${currentRow})`, // Restbetrag Netto
                `=IF(A${currentRow}=""; ""; ROUNDUP(MONTH(A${currentRow})/3;0))`, // Quartal
                `=IF(H${currentRow}>=G${currentRow}; "Bezahlt"; "Offen")`, // Zahlungsstatus
                "",             // Leeres Feld
                fileName,       // Dateiname
                fileUrl,        // Link zur Datei
                timestamp       // Letzte Aktualisierung
            ]);
            existingMainFiles.add(fileName);
            wasImported = true;
        }

        // Eintrag im Import-Sheet hinzufügen, wenn noch nicht vorhanden:
        if (!existingImportFiles.has(fileName)) {
            newImportRows.push([fileName, fileUrl, fileName]);
            existingImportFiles.add(fileName);
            wasImported = true;
        }

        // Historieneintrag nur, wenn die Datei tatsächlich neu importiert wurde:
        if (wasImported) {
            newHistoryRows.push([timestamp, type, fileName, fileUrl]);
        }
    }

    if (newImportRows.length > 0) {
        const importStartRow = importSheet.getLastRow() + 1;
        importSheet.getRange(importStartRow, 1, newImportRows.length, newImportRows[0].length).setValues(newImportRows);
    }

    if (newMainRows.length > 0) {
        mainSheet.getRange(mainStartRow, 1, newMainRows.length, newMainRows[0].length).setValues(newMainRows);
    }

    if (newHistoryRows.length > 0) {
        const historyStartRow = historyTab.getLastRow() + 1;
        historyTab.getRange(historyStartRow, 1, newHistoryRows.length, newHistoryRows[0].length).setValues(newHistoryRows);
    }
}

// Hilfsfunktion: Sucht einen Unterordner anhand des Namens
function getFolderByName(parentFolder, folderName) {
    const folderIter = parentFolder.getFoldersByName(folderName);
    return folderIter.hasNext() ? folderIter.next() : null;
}

/* ----- Funktionen ausschließlich für die Tabellen "Einnahmen" und "Ausgaben" ----- */
function sortCurrentSheetByColumn() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    const name = sheet.getName();
    if (name !== "Einnahmen" && name !== "Ausgaben") {
        SpreadsheetApp.getUi().alert("Diese Funktion ist nur für die Blätter 'Einnahmen' und 'Ausgaben' verfügbar.");
        return;
    }
    sortSheetByColumn(sheet, 2);
}

function updateSheetOnCurrentSheet() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    const name = sheet.getName();
    if (name !== "Einnahmen" && name !== "Ausgaben") {
        SpreadsheetApp.getUi().alert("Diese Funktion ist nur für die Blätter 'Einnahmen' und 'Ausgaben' verfügbar.");
        return;
    }
    updateFormulasOnSheet(sheet);
    applyFormatting(sheet);
}

function sortSheetByColumn(sheet, column) {
    const range = sheet.getDataRange();
    const values = range.getValues();
    if (values.length <= 1) {
        Logger.log("Keine Daten zum Sortieren.");
        return;
    }
    const header = values.shift();
    values.sort((a, b) => {
        const valA = a[column - 1];
        const valB = b[column - 1];
        if (typeof valA === "number" && typeof valB === "number") {
            return valA - valB;
        }
        return String(valA).localeCompare(String(valB));
    });
    sheet.clearContents();
    sheet.appendRow(header);
    sheet.getRange(2, 1, values.length, values[0].length).setValues(values);
    updateFormulasOnSheet(sheet);
}

function updateFormulasOnSheet(sheet) {
    const lastRow = sheet.getLastRow();
    const numRows = lastRow - 1;
    if (numRows < 1) return;

    const formulas6 = [];
    const formulas7 = [];
    const formulas9 = [];
    const formulas10 = [];
    const formulas11 = [];

    for (let i = 2; i <= lastRow; i++) {
        formulas6.push([`=D${i}*E${i}`]);
        formulas7.push([`=D${i}+F${i}`]);
        formulas9.push([`=D${i}-(H${i}-F${i})`]);
        formulas10.push([`=IF(A${i}=""; ""; ROUNDUP(MONTH(A${i})/3;0))`]);
        formulas11.push([`=IF(H${i}>=G${i}; "Bezahlt"; "Offen")`]);
    }

    sheet.getRange(2, 6, numRows, 1).setFormulas(formulas6);
    sheet.getRange(2, 7, numRows, 1).setFormulas(formulas7);
    sheet.getRange(2, 9, numRows, 1).setFormulas(formulas9);
    sheet.getRange(2, 10, numRows, 1).setFormulas(formulas10);
    sheet.getRange(2, 11, numRows, 1).setFormulas(formulas11);
}

function applyFormatting(sheet) {
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
        sheet.getRange(`A2:A${lastRow}`).setNumberFormat("DD.MM.YYYY");
        sheet.getRange(`L2:L${lastRow}`).setNumberFormat("DD.MM.YYYY");
        sheet.getRange(`O2:O${lastRow}`).setNumberFormat("DD.MM.YYYY");
        sheet.getRange(`D2:D${lastRow}`).setNumberFormat("€#,##0.00;€-#,##0.00");
        sheet.getRange(`F2:I${lastRow}`).setNumberFormat("€#,##0.00;€-#,##0.00");
        sheet.getRange(`E2:E${lastRow}`).setNumberFormat("0.00%");
    }
}

/* ----- Neue Funktion: GUV-Berechnung ----- */
function calculateGUV() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const einnahmenSheet = ss.getSheetByName("Einnahmen");
    const ausgabenSheet = ss.getSheetByName("Ausgaben");

    if (!einnahmenSheet || !ausgabenSheet) {
        SpreadsheetApp.getUi().alert("Eines der Blätter 'Einnahmen' oder 'Ausgaben' wurde nicht gefunden.");
        return;
    }

    // Daten (ohne Header) einlesen
    const einnahmenData = einnahmenSheet.getDataRange().getValues().slice(1);
    const ausgabenData = ausgabenSheet.getDataRange().getValues().slice(1);

    let fehlendeDaten = [];

    // Monatsweise und Quartalsweise Berechnung
    let guvData = {};
    for (let m = 1; m <= 12; m++) {
        guvData[m] = {
            einnahmen: 0, einnahmenOffen: 0, ausgaben: 0, ausgabenOffen: 0,
            umsatzsteuer: 0, vorsteuer: 0, ustZahlung: 0, ergebnis: 0
        };
    }

    // Einnahmen durchgehen
    einnahmenData.forEach((row, index) => {
        let date = row[0]; // Datum in Spalte A
        if (!(date instanceof Date)) {
            fehlendeDaten.push(`Einnahmen - Zeile ${index + 2}`);
            return;
        }
        let netto = parseFloat(row[3]) || 0; // Netto (€) in Spalte D
        let mwst = parseFloat(row[5]) || 0; // MwSt.-Betrag (€) in Spalte F
        let bezahlt = parseFloat(row[7]) || 0; // Bezahlter Betrag (€) in Spalte H
        let monat = date.getMonth() + 1;

        guvData[monat].einnahmen += bezahlt; // Nur tatsächlich bezahlte Einnahmen
        guvData[monat].umsatzsteuer += bezahlt > 0 ? mwst : 0; // MwSt. nur auf bezahlte Beträge
        guvData[monat].einnahmenOffen += netto - bezahlt; // Offene Forderungen
    });

    // Ausgaben durchgehen
    ausgabenData.forEach((row, index) => {
        let date = row[0]; // Datum in Spalte A
        if (!(date instanceof Date)) {
            fehlendeDaten.push(`Ausgaben - Zeile ${index + 2}`);
            return;
        }
        let netto = parseFloat(row[3]) || 0; // Netto (€) in Spalte D
        let mwst = parseFloat(row[5]) || 0; // MwSt.-Betrag (€) in Spalte F
        let bezahlt = parseFloat(row[7]) || 0; // Bezahlter Betrag (€) in Spalte H
        let monat = date.getMonth() + 1;

        guvData[monat].ausgaben += bezahlt; // Nur tatsächlich bezahlte Ausgaben
        guvData[monat].vorsteuer += bezahlt > 0 ? mwst : 0; // Vorsteuer nur auf bezahlte Beträge
        guvData[monat].ausgabenOffen += netto - bezahlt; // Offene Verbindlichkeiten
    });

    // Falls fehlende Daten gefunden wurden, abbrechen und Benutzer informieren
    if (fehlendeDaten.length > 0) {
        SpreadsheetApp.getUi().alert("Fehler: Es gibt Einträge ohne Datum! Bitte überprüfe folgende Zeilen:\n" + fehlendeDaten.join("\n"));
        return;
    }

    // GUV-Blatt erstellen oder bereinigen
    let guvSheet = ss.getSheetByName("GUV");
    if (!guvSheet) {
        guvSheet = ss.insertSheet("GUV");
    } else {
        guvSheet.clearContents();
    }

    // Kopfzeile schreiben
    guvSheet.appendRow(["Monat", "Einnahmen", "Offene Forderungen", "Ausgaben", "Offene Verbindlichkeiten",
        "Umsatzsteuer", "Vorsteuer", "USt-Zahlung", "Ergebnis"]);

    let gesamtErgebnis = 0;

    // Daten pro Monat eintragen
    for (let m = 1; m <= 12; m++) {
        let einnahmen = guvData[m].einnahmen;
        let einnahmenOffen = guvData[m].einnahmenOffen;
        let ausgaben = guvData[m].ausgaben;
        let ausgabenOffen = guvData[m].ausgabenOffen;
        let umsatzsteuer = guvData[m].umsatzsteuer;
        let vorsteuer = guvData[m].vorsteuer;
        let ustZahlung = umsatzsteuer - vorsteuer;
        let ergebnis = einnahmen - ausgaben;

        gesamtErgebnis += ergebnis;

        guvSheet.appendRow([`Monat ${m}`, einnahmen, einnahmenOffen, ausgaben, ausgabenOffen, umsatzsteuer, vorsteuer, ustZahlung, ergebnis]);
    }

    guvSheet.appendRow(["Gesamtjahr", "", "", "", "", "", "", "", gesamtErgebnis]);

    SpreadsheetApp.getUi().alert("GUV-Berechnung abgeschlossen und aktualisiert.");
}




