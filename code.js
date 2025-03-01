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

    // Submenü für BWA
    const bwaMenu = ui.createMenu("📈 BWA");
    bwaMenu.addItem("📊 BWA berechnen", "calculateBWA");

    // Hauptmenü "Buchhaltung" mit allen Submenüs
    ui.createMenu("📂 Buchhaltung")
        .addSubMenu(rechnungenMenu)
        .addSubMenu(einnahmenAusgabenMenu)
        .addSubMenu(guvMenu)
        .addSubMenu(bwaMenu)  // Hier wird BWA hinzugefügt
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
        // Import in Rechnungen-Eingangsblatt und Hauptübersicht "Einnahmen"
        importFilesFromFolder(einnahmenOrdner, sheetEinnahmen, einnahmenTab, "Einnahme", historyTab);
        // Spaltenbreite in Rechnungen-Eingang anpassen
        sheetEinnahmen.autoResizeColumns(1, sheetEinnahmen.getLastColumn());
        if (einnahmenTab.getLastRow() > 1) {
            updateFormulasOnSheet(einnahmenTab);
            applyFormatting(einnahmenTab);
            einnahmenTab.autoResizeColumns(1, einnahmenTab.getLastColumn());
        }
    } else {
        Logger.log("Fehler: Der 'Einnahmen'-Ordner wurde nicht gefunden.");
    }

    if (ausgabenOrdner) {
        // Import in Rechnungen-Ausgangsblatt und Hauptübersicht "Ausgaben"
        importFilesFromFolder(ausgabenOrdner, sheetAusgaben, ausgabenTab, "Ausgabe", historyTab);
        // Spaltenbreite in Rechnungen-Ausgang anpassen
        sheetAusgaben.autoResizeColumns(1, sheetAusgaben.getLastColumn());
        if (ausgabenTab.getLastRow() > 1) {
            updateFormulasOnSheet(ausgabenTab);
            applyFormatting(ausgabenTab);
            ausgabenTab.autoResizeColumns(1, ausgabenTab.getLastColumn());
        }
    } else {
        Logger.log("Fehler: Der 'Ausgaben'-Ordner wurde nicht gefunden.");
    }
}

function importFilesFromFolder(folder, importSheet, mainSheet, type, historyTab) {
    const files = folder.getFiles();

    // Vorhandene Dateinamen in Hauptblatt und Import-Sheet ermitteln:
    const existingMainData = mainSheet.getDataRange().getValues();
    const existingMainFiles = new Set(existingMainData.slice(1).map(row => row[1])); // Spalte B: Dateiname

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

        // Nur hinzufügen, wenn Datei im Hauptblatt noch nicht vorhanden ist:
        if (!existingMainFiles.has(fileName)) {
            const currentRow = mainStartRow + newMainRows.length;
            newMainRows.push([
                timestamp,               // Spalte A: Datum
                fileName,                // Spalte B: Dateiname
                "",                      // Spalte C: Kategorie (neu)
                "",                      // Spalte D: Kunde
                "",                      // Spalte E: Nettobetrag
                "",                      // Spalte F: Prozentsatz (als Prozent formatiert)
                `=E${currentRow}*F${currentRow}`, // Spalte G: MwSt.-Betrag
                `=E${currentRow}+G${currentRow}`, // Spalte H: Bruttobetrag
                "",                      // Spalte I: (wird als Währung formatiert)
                `=E${currentRow}-(I${currentRow}-G${currentRow})`, // Spalte J: Restbetrag Netto
                `=IF(A${currentRow}=""; ""; ROUNDUP(MONTH(A${currentRow})/3;0))`, // Spalte K: Quartal
                `=IF(I${currentRow}>=H${currentRow}; "Bezahlt"; "Offen")`, // Spalte L: Zahlungsstatus
                "",                      // Spalte M: (wird als Datum formatiert)
                fileName,                // Spalte N: Dateiname (wiederholt)
                fileUrl,                 // Spalte O: Link zur Datei
                timestamp                // Spalte P: Letzte Aktualisierung
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
    sheet.autoResizeColumns(1, sheet.getLastColumn());
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
        formulas6.push([`=E${i}*F${i}`]); // MwSt.-Betrag (Spalte G)
        formulas7.push([`=E${i}+G${i}`]); // Bruttobetrag (Spalte H)
        formulas9.push([`=E${i}-(I${i}-G${i})`]); // Restbetrag Netto (Spalte J)
        formulas10.push([`=IF(A${i}=""; ""; ROUNDUP(MONTH(A${i})/3;0))`]); // Quartal (Spalte K)
        formulas11.push([`=IF(OR(I${i}=""; I${i}=0); "Offen"; IF(I${i}>=H${i}; "Bezahlt"; "Offen"))`]); // Zahlungsstatus (Spalte L)
    }

    sheet.getRange(2, 7, numRows, 1).setFormulas(formulas6);
    sheet.getRange(2, 8, numRows, 1).setFormulas(formulas7);
    sheet.getRange(2, 10, numRows, 1).setFormulas(formulas9);
    sheet.getRange(2, 11, numRows, 1).setFormulas(formulas10);
    sheet.getRange(2, 12, numRows, 1).setFormulas(formulas11);
}

function applyFormatting(sheet) {
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
        // Spalte A (Datum)
        sheet.getRange(`A2:A${lastRow}`).setNumberFormat("DD.MM.YYYY");

        // Spalte M (z.B. weiterer Datumswert)
        sheet.getRange(`M2:M${lastRow}`).setNumberFormat("DD.MM.YYYY");

        // Spalte I als Währung (Accounting)
        sheet.getRange(`I2:I${lastRow}`).setNumberFormat("€#,##0.00;€-#,##0.00");

        // Spalte P (Letzte Aktualisierung, Datum)
        sheet.getRange(`P2:P${lastRow}`).setNumberFormat("DD.MM.YYYY");

        // Spalte E (Nettobetrag) als Währung
        sheet.getRange(`E2:E${lastRow}`).setNumberFormat("€#,##0.00;€-#,##0.00");

        // MwSt, Bruttobetrag und Restbetrag in G, H und J als Währung
        sheet.getRange(`G2:G${lastRow}`).setNumberFormat("€#,##0.00;€-#,##0.00");
        sheet.getRange(`H2:H${lastRow}`).setNumberFormat("€#,##0.00;€-#,##0.00");
        sheet.getRange(`J2:J${lastRow}`).setNumberFormat("€#,##0.00;€-#,##0.00");

        // Spalte F (Prozentsatz)
        sheet.getRange(`F2:F${lastRow}`).setNumberFormat("0.00%");
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

    const einnahmenData = einnahmenSheet.getDataRange().getValues().slice(1);
    const ausgabenData = ausgabenSheet.getDataRange().getValues().slice(1);

    let fehlendeDaten = [];

    let guvData = {};
    for (let m = 1; m <= 12; m++) {
        guvData[m] = {
            einnahmen: 0, einnahmenOffen: 0, ausgaben: 0, ausgabenOffen: 0,
            umsatzsteuer: 0, vorsteuer: 0, ustZahlung: 0, ergebnis: 0
        };
    }

    einnahmenData.forEach((row, index) => {
        let date = row[0];
        if (!(date instanceof Date)) {
            fehlendeDaten.push(`Einnahmen - Zeile ${index + 2}`);
            return;
        }
        let monat = date.getMonth() + 1;
        let netto = parseFloat(row[4]) || 0;
        let mwst = parseFloat(row[6]) || 0;
        let bezahlt = parseFloat(row[8]) || 0;

        guvData[monat].einnahmen += bezahlt;
        guvData[monat].umsatzsteuer += bezahlt > 0 ? mwst : 0;
        guvData[monat].einnahmenOffen += netto - bezahlt;
    });

    ausgabenData.forEach((row, index) => {
        let date = row[0];
        if (!(date instanceof Date)) {
            fehlendeDaten.push(`Ausgaben - Zeile ${index + 2}`);
            return;
        }
        let monat = date.getMonth() + 1;
        let netto = parseFloat(row[4]) || 0;
        let mwst = parseFloat(row[6]) || 0;
        let bezahlt = parseFloat(row[8]) || 0;

        guvData[monat].ausgaben += bezahlt;
        guvData[monat].vorsteuer += bezahlt > 0 ? mwst : 0;
        guvData[monat].ausgabenOffen += netto - bezahlt;
    });

    if (fehlendeDaten.length > 0) {
        SpreadsheetApp.getUi().alert("Fehler: Es gibt Einträge ohne Datum! Bitte überprüfe folgende Zeilen:\n" + fehlendeDaten.join("\n"));
        return;
    }

    let guvSheet = ss.getSheetByName("GUV");
    if (!guvSheet) {
        guvSheet = ss.insertSheet("GUV");
    } else {
        guvSheet.clearContents();
    }

    guvSheet.appendRow(["Zeitraum", "Einnahmen", "Offene Forderungen", "Ausgaben", "Offene Verbindlichkeiten",
        "Umsatzsteuer", "Vorsteuer", "USt-Zahlung", "Ergebnis"]);

    let gesamtErgebnis = 0;
    let quartalsDaten = {1: {}, 2: {}, 3: {}, 4: {}};

    for (let q = 1; q <= 4; q++) {
        quartalsDaten[q] = {
            einnahmen: 0, einnahmenOffen: 0, ausgaben: 0, ausgabenOffen: 0,
            umsatzsteuer: 0, vorsteuer: 0, ustZahlung: 0, ergebnis: 0
        };
    }

    for (let m = 1; m <= 12; m++) {
        let q = Math.ceil(m / 3);
        let data = guvData[m];
        quartalsDaten[q].einnahmen += data.einnahmen;
        quartalsDaten[q].einnahmenOffen += data.einnahmenOffen;
        quartalsDaten[q].ausgaben += data.ausgaben;
        quartalsDaten[q].ausgabenOffen += data.ausgabenOffen;
        quartalsDaten[q].umsatzsteuer += data.umsatzsteuer;
        quartalsDaten[q].vorsteuer += data.vorsteuer;
        quartalsDaten[q].ustZahlung += data.umsatzsteuer - data.vorsteuer;
        quartalsDaten[q].ergebnis += data.einnahmen - data.ausgaben;

        guvSheet.appendRow([`Monat ${m}`, data.einnahmen, data.einnahmenOffen, data.ausgaben, data.ausgabenOffen,
            data.umsatzsteuer, data.vorsteuer, data.umsatzsteuer - data.vorsteuer, data.einnahmen - data.ausgaben]);
    }

    for (let q = 1; q <= 4; q++) {
        let data = quartalsDaten[q];
        guvSheet.appendRow([`Quartal ${q}`, data.einnahmen, data.einnahmenOffen, data.ausgaben, data.ausgabenOffen,
            data.umsatzsteuer, data.vorsteuer, data.umsatzsteuer - data.vorsteuer, data.ergebnis]);
    }

    guvSheet.appendRow(["Gesamtjahr", ...Object.values(quartalsDaten).reduce((acc, q) => acc.map((val, i) => val + Object.values(q)[i]), [0, 0, 0, 0, 0, 0, 0, 0])]);

    let lastRow = guvSheet.getLastRow();
    if (lastRow > 1) {
        guvSheet.getRange(`B2:I${lastRow}`).setNumberFormat("#,##0.00€");
    }

    SpreadsheetApp.getUi().alert("GUV-Berechnung abgeschlossen und aktualisiert.");
}

function calculateBWA() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const einnahmenSheet = ss.getSheetByName("Einnahmen");
    const ausgabenSheet = ss.getSheetByName("Ausgaben");

    if (!einnahmenSheet || !ausgabenSheet) {
        SpreadsheetApp.getUi().alert("Eines der Blätter 'Einnahmen' oder 'Ausgaben' wurde nicht gefunden.");
        return;
    }

    const einnahmenData = einnahmenSheet.getDataRange().getValues().slice(1);
    const ausgabenData = ausgabenSheet.getDataRange().getValues().slice(1);
    let fehlendeDaten = [];

    let bwaData = {};
    for (let m = 1; m <= 12; m++) {
        bwaData[m] = {
            dienstleistungen: 0, produkte: 0, sonstigeEinnahmen: 0, gesamtEinnahmen: 0,
            betriebskosten: 0, marketing: 0, reisen: 0, einkauf: 0, personal: 0, gesamtAusgaben: 0,
            offeneForderungen: 0, offeneVerbindlichkeiten: 0,
            rohertrag: 0, betriebsergebnis: 0, ergebnisVorSteuern: 0, ergebnisNachSteuern: 0,
            liquiditaet: 0
        };
    }

    function getCategoryMapping(category) {
        const categoryMap = {
            "Dienstleistungen": "dienstleistungen",
            "Produkte & Waren": "produkte",
            "Sonstige Einnahmen": "sonstigeEinnahmen",
            "Betriebskosten": "betriebskosten",
            "Marketing & Werbung": "marketing",
            "Reisen & Mobilität": "reisen",
            "Wareneinkauf & Dienstleistungen": "einkauf",
            "Personal & Gehälter": "personal"
        };
        return categoryMap[category] || null;
    }

    einnahmenData.forEach((row, index) => {
        let date = row[0];
        let category = row[2];
        let netto = parseFloat(row[4]) || 0;
        let bezahlt = parseFloat(row[8]) || 0;
        let monat = date instanceof Date ? date.getMonth() + 1 : null;

        if (!monat) {
            fehlendeDaten.push(`Einnahmen - Zeile ${index + 2}`);
            return;
        }
        if (!category) {
            fehlendeDaten.push(`Einnahmen ohne Kategorie - Zeile ${index + 2}`);
            return;
        }

        let mappedCategory = getCategoryMapping(category);
        if (!mappedCategory) {
            fehlendeDaten.push(`Ungültige Kategorie in Einnahmen - Zeile ${index + 2}`);
            return;
        }

        bwaData[monat][mappedCategory] += bezahlt;
        bwaData[monat].gesamtEinnahmen += bezahlt;
        bwaData[monat].offeneForderungen += netto - bezahlt;
    });

    ausgabenData.forEach((row, index) => {
        let date = row[0];
        let category = row[2];
        let netto = parseFloat(row[4]) || 0;
        let bezahlt = parseFloat(row[8]) || 0;
        let monat = date instanceof Date ? date.getMonth() + 1 : null;

        if (!monat) {
            fehlendeDaten.push(`Ausgaben - Zeile ${index + 2}`);
            return;
        }
        if (!category) {
            fehlendeDaten.push(`Ausgaben ohne Kategorie - Zeile ${index + 2}`);
            return;
        }

        let mappedCategory = getCategoryMapping(category);
        if (!mappedCategory) {
            fehlendeDaten.push(`Ungültige Kategorie in Ausgaben - Zeile ${index + 2}`);
            return;
        }

        bwaData[monat][mappedCategory] += bezahlt;
        bwaData[monat].gesamtAusgaben += bezahlt;
        bwaData[monat].offeneVerbindlichkeiten += netto - bezahlt;
    });

    if (fehlendeDaten.length > 0) {
        SpreadsheetApp.getUi().alert("Fehler: Fehlende oder ungültige Daten! Bitte überprüfen:\n" + fehlendeDaten.join("\n"));
        return;
    }

    let bwaSheet = ss.getSheetByName("BWA");
    if (!bwaSheet) {
        bwaSheet = ss.insertSheet("BWA");
    } else {
        bwaSheet.clearContents();
    }

    bwaSheet.appendRow(["Zeitraum", "Dienstleistungen", "Produkte", "Sonstige Einnahmen", "Gesamt-Einnahmen",
        "Betriebskosten", "Marketing", "Reisen", "Einkauf", "Personal", "Gesamt-Ausgaben",
        "Offene Forderungen", "Offene Verbindlichkeiten", "Rohertrag", "Betriebsergebnis",
        "Ergebnis vor Steuern", "Ergebnis nach Steuern", "Liquidität"]);

    for (let q = 1; q <= 4; q++) {
        let quartalDaten = Array(17).fill(0);
        for (let m = (q - 1) * 3 + 1; m <= q * 3; m++) {
            let data = Object.values(bwaData[m]);
            quartalDaten = quartalDaten.map((val, i) => val + data[i]);
        }
        bwaSheet.appendRow([`Quartal ${q}`, ...quartalDaten]);
    }

    for (let m = 1; m <= 12; m++) {
        let data = bwaData[m];
        bwaSheet.appendRow([`Monat ${m}`, ...Object.values(data)]);
    }

    bwaSheet.appendRow(["Gesamtjahr", ...Object.values(bwaData).reduce((acc, q) => acc.map((val, i) => val + Object.values(q)[i]), Array(17).fill(0))]);

    let lastRow = bwaSheet.getLastRow();
    if (lastRow > 1) {
        bwaSheet.getRange(`B2:R${lastRow}`).setNumberFormat("#,##0.00€");
    }

    SpreadsheetApp.getUi().alert("BWA-Berechnung abgeschlossen und aktualisiert.");
}

