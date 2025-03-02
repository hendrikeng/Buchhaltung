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
    ui.createMenu("📂 Buchhaltung")
        .addItem("📥 Rechnungen importieren", "importDriveFiles")
        .addItem("🔄 Aktualisieren (Formeln & Formatierung)", "updateSheetOnCurrentSheet")
        .addItem("📊 GUV berechnen", "calculateGUV")
        .addItem("📈 BWA berechnen", "calculateBWA")
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

/**
 * calculateGUV (Ist-Besteuerung) – überarbeitete Version
 *
 * Zweck:
 *  - Führt eine GuV-Berechnung durch, bei der Einnahmen und Ausgaben nur dann erfasst werden,
 *    wenn sie tatsächlich bezahlt wurden (Ist-Besteuerung).
 *  - Offene Posten (kein Zahlungsdatum) werden nicht in die Monats-/Quartalswerte übernommen,
 *    sondern nur als Gesamtsumme offener Forderungen/Verbindlichkeiten in der Jahresübersicht ausgewiesen.
 *
 * Verbesserungen:
 *  - Dynamische Ermittlung des Mehrwertsteuersatzes (statt pauschal 19% anzunehmen)
 *  - Validierung, dass das Zahlungsdatum nicht in der Zukunft liegt
 *  - Behandlung von Gutschriften/Erstattungen (negative Werte) durch direkte Berechnung des offenen Betrags
 *
 * Benötigte Tabellenblätter:
 *  - "Einnahmen" (mit mindestens diesen Spalten):
 *      - Netto-Betrag (z. B. Spalte E)
 *      - MwSt-Satz (z. B. Spalte G; als String, z. B. "19,00%" oder "7,00%")
 *      - Bezahlter Brutto-Betrag (z. B. Spalte I)
 *      - Zahlungsdatum (z. B. Spalte M)
 *  - "Ausgaben" (mit denselben relevanten Spalten)
 *  - "GUV" (wird automatisch erzeugt oder überschrieben, falls schon vorhanden)
 *
 * Wichtige Schritte:
 *  1) Einlesen aller Daten aus "Einnahmen" und "Ausgaben".
 *  2) Für jede Zeile:
 *     - Wird geprüft, ob ein gültiges Zahlungsdatum vorhanden ist und ob dieses nicht in der Zukunft liegt.
 *     - Der tatsächliche MwSt.-Satz wird dynamisch ermittelt und die Netto- sowie MwSt-Anteile werden berechnet.
 *     - Bei Teilzahlungen wird der offene Anteil (netto - bezahltNetto) berechnet – auch negative Differenzen
 *       (z. B. bei Gutschriften) bleiben erhalten.
 *     - Ohne Zahlungsdatum wird der gesamte Netto-Betrag als offen in der Jahresübersicht erfasst.
 *  3) Aufsummierung der Daten in Monats- und Quartalsübersicht.
 *  4) Ausgabe in einem "GUV"-Sheet inklusive einer Gesamtjahreszeile.
 *  5) Formatierung der Ausgabedaten (z. B. Euro-Formatierung).
 *
 * Voraussetzungen & Hinweise:
 *  - Es muss sichergestellt sein, dass in Spalte "Letzte Zahlung am" (z. B. Spalte M) nur gültige Datumswerte stehen,
 *    wenn ein Teil- oder Vollbetrag bezahlt wurde.
 *  - Ist-Besteuerung bedeutet, dass nur tatsächlich geflossene Zahlungen für Umsatzsteuer, Vorsteuer,
 *    Gewinn und Verlust berücksichtigt werden.
 *  - Offene Forderungen/Ausgaben fließen nicht in die Monats-/Quartalswerte ein, sondern nur in die Jahresübersicht.
 */
function calculateGUV() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const einnahmenSheet = ss.getSheetByName("Einnahmen");
    const ausgabenSheet = ss.getSheetByName("Ausgaben");

    if (!einnahmenSheet || !ausgabenSheet) {
        SpreadsheetApp.getUi().alert("Eines der Blätter 'Einnahmen' oder 'Ausgaben' wurde nicht gefunden.");
        return;
    }

    // Alle Daten einlesen (ohne Kopfzeile)
    const einnahmenData = einnahmenSheet.getDataRange().getValues().slice(1);
    const ausgabenData = ausgabenSheet.getDataRange().getValues().slice(1);

    // ✅ Fix: Richtig initialisierte `guvData` Struktur
    let guvData = {};
    for (let m = 1; m <= 12; m++) {
        guvData[m] = {
            einnahmen: 0, einnahmenOffen: 0,
            ausgaben: 0, ausgabenOffen: 0,
            umsatzsteuer: 0, vorsteuer: 0,
            ustZahlung: 0, ergebnis: 0
        };
    }

    // ✅ Fix: `quartalsDaten` korrekt initialisiert
    let quartalsDaten = {};
    for (let q = 1; q <= 4; q++) {
        quartalsDaten[q] = {
            einnahmen: 0, einnahmenOffen: 0,
            ausgaben: 0, ausgabenOffen: 0,
            umsatzsteuer: 0, vorsteuer: 0,
            ustZahlung: 0, ergebnis: 0
        };
    }

    let offeneEinnahmenGesamt = 0;
    let offeneAusgabenGesamt = 0;
    let fehlendeDaten = [];

    //--------------------------------
    // Einnahmen verarbeiten
    //--------------------------------
    einnahmenData.forEach((row, index) => {
        let zahlungsDatum = parseDate(row[12]);
        let netto = parseCurrency(row[4]);
        let mwstRate = parseMwstRate(row[5]);
        let bezahltBrutto = parseCurrency(row[8]);

        let factor = (mwstRate > 0) ? 1 + (mwstRate / 100) : 1;
        let bezahltNetto = bezahltBrutto / factor;
        let bezahltMwst = bezahltBrutto - bezahltNetto;

        if (bezahltBrutto !== 0 && !zahlungsDatum) {
            fehlendeDaten.push(`Einnahmen - Zeile ${index + 2}: Fehlendes Zahlungsdatum trotz Zahlung`);
            return;
        }

        if (zahlungsDatum && zahlungsDatum > new Date()) {
            fehlendeDaten.push(`Einnahmen - Zeile ${index + 2}: Zahlungsdatum liegt in der Zukunft`);
            return;
        }

        if (zahlungsDatum) {
            let monat = zahlungsDatum.getMonth() + 1;
            let quartal = Math.ceil(monat / 3);

            // ✅ Fix: Stelle sicher, dass `guvData[monat]` existiert
            if (!guvData[monat]) guvData[monat] = {};
            if (!quartalsDaten[quartal]) quartalsDaten[quartal] = {};

            guvData[monat].einnahmen += bezahltNetto;
            guvData[monat].umsatzsteuer += bezahltMwst;
            quartalsDaten[quartal].einnahmen += bezahltNetto;
            quartalsDaten[quartal].umsatzsteuer += bezahltMwst;

            let offenerBetrag = Math.max(0, netto - bezahltNetto);
            if (offenerBetrag > 0) {
                guvData[monat].einnahmenOffen += offenerBetrag;
                quartalsDaten[quartal].einnahmenOffen += offenerBetrag;
            }
        } else {
            offeneEinnahmenGesamt += netto;
        }
    });

    //--------------------------------
    // Ausgaben verarbeiten
    //--------------------------------
    ausgabenData.forEach((row, index) => {
        let zahlungsDatum = parseDate(row[12]);
        let netto = parseCurrency(row[4]);
        let mwstRate = parseMwstRate(row[5]);
        let bezahltBrutto = parseCurrency(row[8]);

        let factor = (mwstRate > 0) ? 1 + (mwstRate / 100) : 1;
        let bezahltNetto = bezahltBrutto / factor;
        let bezahltMwst = bezahltBrutto - bezahltNetto;

        if (zahlungsDatum && zahlungsDatum > new Date()) {
            fehlendeDaten.push(`Ausgaben - Zeile ${index + 2}: Zahlungsdatum liegt in der Zukunft`);
            return;
        }

        if (zahlungsDatum) {
            let monat = zahlungsDatum.getMonth() + 1;
            let quartal = Math.ceil(monat / 3);

            guvData[monat].ausgaben += bezahltNetto;
            guvData[monat].vorsteuer += bezahltMwst;
            quartalsDaten[quartal].ausgaben += bezahltNetto;
            quartalsDaten[quartal].vorsteuer += bezahltMwst;

            // Falls nur Teilbetrag bezahlt wurde, ist der Rest offen
            let offenerBetrag = Math.max(0, netto - bezahltNetto);
            if (offenerBetrag > 0) {
                guvData[monat].ausgabenOffen += offenerBetrag;
                quartalsDaten[quartal].ausgabenOffen += offenerBetrag;
            }
        } else {
            // KEIN Zahlungsdatum => kompletter Betrag offen
            offeneAusgabenGesamt += netto;
        }
    });

    // Falls Fehler (Bezahlung ohne Datum), abbrechen
    if (fehlendeDaten.length > 0) {
        SpreadsheetApp.getUi().alert(
            "Fehlende Zahlungsdaten gefunden:\n" + fehlendeDaten.join("\n")
        );
        return;
    }

    //--------------------------------
    // GUV-Sheet aufbauen
    //--------------------------------
    let guvSheet = ss.getSheetByName("GUV") || ss.insertSheet("GUV");
    guvSheet.clearContents();

    guvSheet.appendRow([
        "Zeitraum",
        "Einnahmen (netto)",
        "Offene Forderungen (netto)",
        "Ausgaben (netto)",
        "Offene Verbindlichkeiten (netto)",
        "Umsatzsteuer (19%)",
        "Vorsteuer (19%)",
        "USt-Zahlung",
        "Ergebnis (Gewinn/Verlust netto)"
    ]);

    // Monate ausgeben
    for (let m = 1; m <= 12; m++) {
        let data = guvData[m];

        // USt-Zahlung = Umsatzsteuer - Vorsteuer
        let ustZahlung = data.umsatzsteuer - data.vorsteuer;
        // Gewinn/Verlust
        let ergebnis = data.einnahmen - data.ausgaben - ustZahlung;

        // Werte in der Struktur aktualisieren (damit sie später summiert werden können)
        data.ustZahlung = ustZahlung;
        data.ergebnis = ergebnis;

        guvSheet.appendRow([
            `Monat ${m}`,
            data.einnahmen,
            data.einnahmenOffen,
            data.ausgaben,
            data.ausgabenOffen,
            data.umsatzsteuer,
            data.vorsteuer,
            ustZahlung,
            ergebnis
        ]);

        // Jedes Quartal nach dem 3., 6., 9. und 12. Monat ausgeben
        if (m % 3 === 0) {
            let q = Math.ceil(m / 3);
            let qData = quartalsDaten[q];

            // USt-Zahlung = Umsatzsteuer - Vorsteuer
            let qUstZahlung = qData.umsatzsteuer - qData.vorsteuer;
            // Gewinn/Verlust
            let qErgebnis = qData.einnahmen - qData.ausgaben - qUstZahlung;

            // Werte in der Struktur aktualisieren
            qData.ustZahlung = qUstZahlung;
            qData.ergebnis = qErgebnis;

            guvSheet.appendRow([
                `Quartal ${q}`,
                qData.einnahmen,
                qData.einnahmenOffen,
                qData.ausgaben,
                qData.ausgabenOffen,
                qData.umsatzsteuer,
                qData.vorsteuer,
                qUstZahlung,
                qErgebnis
            ]);
        }
    }

    //--------------------------------
    // Jahreswerte berechnen und ausgeben
    //--------------------------------
    // Alle Monate summieren
    let jahresSumme = {
        einnahmen: 0,
        einnahmenOffen: 0,
        ausgaben: 0,
        ausgabenOffen: 0,
        umsatzsteuer: 0,
        vorsteuer: 0,
        ustZahlung: 0,
        ergebnis: 0
    };

    for (let m = 1; m <= 12; m++) {
        jahresSumme.einnahmen += guvData[m].einnahmen;
        jahresSumme.einnahmenOffen += guvData[m].einnahmenOffen;
        jahresSumme.ausgaben += guvData[m].ausgaben;
        jahresSumme.ausgabenOffen += guvData[m].ausgabenOffen;
        jahresSumme.umsatzsteuer += guvData[m].umsatzsteuer;
        jahresSumme.vorsteuer += guvData[m].vorsteuer;
        jahresSumme.ustZahlung += guvData[m].ustZahlung;
        jahresSumme.ergebnis += guvData[m].ergebnis;
    }

    // Nun noch die komplett offenen Posten (ohne Zahlungsdatum) hinzufügen
    // Sie tauchen hier als "Offene Forderungen" bzw. "Offene Verbindlichkeiten" im Gesamtjahr auf
    jahresSumme.einnahmenOffen += offeneEinnahmenGesamt;
    jahresSumme.ausgabenOffen += offeneAusgabenGesamt;

    // Ergebnis bleibt davon unberührt, da bei Ist-Besteuerung erst bei Zahlung relevant
    // (nur wenn du willst, kannst du das "theoretische" Ergebnis um die offenen Beträge ergänzen)

    guvSheet.appendRow([
        "Gesamtjahr",
        jahresSumme.einnahmen,
        jahresSumme.einnahmenOffen,
        jahresSumme.ausgaben,
        jahresSumme.ausgabenOffen,
        jahresSumme.umsatzsteuer,
        jahresSumme.vorsteuer,
        jahresSumme.ustZahlung,
        jahresSumme.ergebnis
    ]);

    // Formatierungen
    let lastRow = guvSheet.getLastRow();
    if (lastRow > 1) {
        guvSheet.getRange(`B2:I${lastRow}`).setNumberFormat("#,##0.00€");
        guvSheet.autoResizeColumns(1, guvSheet.getLastColumn());
    }

    SpreadsheetApp.getUi().alert("GUV-Berechnung abgeschlossen und aktualisiert.");
}


// Hilfsfunktionen
function parseCurrency(value) {
    return parseFloat(value.toString().replace(/[^\d,.-]/g, "").replace(",", ".")) || 0;
}

function parseDate(value) {
    return (typeof value === "string") ? new Date(value) : (value instanceof Date ? value : null);
}

function parseMwstRate(value) {
    let rate = (typeof value === "number") ? (value < 1 ? value * 100 : value) :
        parseFloat(value.toString().replace("%", "").replace(",", "."));
    return isNaN(rate) ? 19 : rate;
}


function calculateBWA() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const einnahmenSheet = ss.getSheetByName("Einnahmen");
    const ausgabenSheet = ss.getSheetByName("Ausgaben");
    const bankSheet = ss.getSheetByName("Bankbewegungen");

    if (!einnahmenSheet || !ausgabenSheet || !bankSheet) {
        SpreadsheetApp.getUi().alert("Eines der Blätter 'Einnahmen', 'Ausgaben' oder 'Bankbewegungen' wurde nicht gefunden.");
        return;
    }

    const einnahmenData = einnahmenSheet.getDataRange().getValues().slice(1);
    const ausgabenData = ausgabenSheet.getDataRange().getValues().slice(1);
    const bankData = bankSheet.getDataRange().getValues().slice(1);
    let fehlendeKategorien = [];

    let bwaData = {};
    let lastSaldo = 0;

    for (let m = 1; m <= 12; m++) {
        bwaData[m] = {
            dienstleistungen: 0, produkte: 0, sonstigeEinnahmen: 0, gesamtEinnahmen: 0,
            betriebskosten: 0, marketing: 0, reisen: 0, einkauf: 0, personal: 0, gesamtAusgaben: 0,
            offeneForderungen: 0, offeneVerbindlichkeiten: 0,
            rohertrag: 0, betriebsergebnis: 0, ergebnisVorSteuern: 0, ergebnisNachSteuern: 0,
            liquiditaet: lastSaldo
        };
    }

    function getCategoryMapping(category, zeile, typ) {
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

        if (!category || !categoryMap[category]) {
            let ersatzKategorie = typ === "Einnahme" ? "sonstigeEinnahmen" : "betriebskosten";
            fehlendeKategorien.push(`Zeile ${zeile}: ${category || "KEINE KATEGORIE"} → Ersetzt durch '${ersatzKategorie}'`);
            return ersatzKategorie;
        }

        return categoryMap[category];
    }

    einnahmenData.forEach((row, index) => {
        let date = row[0];
        let category = row[2];
        let netto = parseFloat(row[4]) || 0;
        let bezahlt = parseFloat(row[8]) || 0;
        let monat = date instanceof Date ? date.getMonth() + 1 : null;

        if (!monat) return;

        let mappedCategory = getCategoryMapping(category, index + 2, "Einnahme");
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

        if (!monat) return;

        let mappedCategory = getCategoryMapping(category, index + 2, "Ausgabe");
        bwaData[monat][mappedCategory] += bezahlt;
        bwaData[monat].gesamtAusgaben += bezahlt;
        bwaData[monat].offeneVerbindlichkeiten += netto - bezahlt;
    });

    bankData.forEach(row => {
        let date = row[0];
        let saldo = parseFloat(row[3]) || lastSaldo;
        let monat = date instanceof Date ? date.getMonth() + 1 : null;
        if (monat) {
            bwaData[monat].liquiditaet = saldo;
            lastSaldo = saldo;
        }
    });

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

    for (let m = 1; m <= 12; m++) {
        let data = bwaData[m];
        bwaSheet.appendRow([`Monat ${m}`, ...Object.values(data)]);
        if (m % 3 === 0) {
            let quartalDaten = Object.values(bwaData).slice(m - 2, m + 1).reduce((acc, q) => acc.map((val, i) => val + Object.values(q)[i]), Array(17).fill(0));
            bwaSheet.appendRow([`Quartal ${Math.ceil(m / 3)}`, ...quartalDaten]);
        }
    }

    let jahreswerte = Object.values(bwaData).reduce((acc, q) => acc.map((val, i) => val + Object.values(q)[i]), Array(17).fill(0));
    bwaSheet.appendRow(["Gesamtjahr", ...jahreswerte]);

    let lastRow = bwaSheet.getLastRow();
    if (lastRow > 1) {
        bwaSheet.getRange(`B2:R${lastRow}`).setNumberFormat("#,##0.00€");
        bwaSheet.autoResizeColumns(1, bwaSheet.getLastColumn());
    }

    if (fehlendeKategorien.length > 0) {
        SpreadsheetApp.getUi().alert(
            "Achtung! Einige Kategorien fehlen oder sind unbekannt:\n" +
            fehlendeKategorien.join("\n") +
            "\n→ Standardkategorie 'Sonstige Einnahmen' oder 'Betriebskosten' wurde genutzt."
        );
    }

    SpreadsheetApp.getUi().alert("BWA-Berechnung abgeschlossen und aktualisiert.");
}






