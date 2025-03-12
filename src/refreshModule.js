// file: src/refreshModule.js
import Validator from "./validator.js";
import config from "./config.js";
import Helpers from "./helpers.js";

/**
 * Modul zum Aktualisieren der Tabellenblätter und Neuberechnen von Formeln
 */
const RefreshModule = (() => {
    /**
     * Aktualisiert ein Datenblatt (Einnahmen, Ausgaben, Eigenbelege)
     * @param {Sheet} sheet - Das zu aktualisierende Sheet
     * @returns {boolean} - true bei erfolgreicher Aktualisierung
     */
    const refreshDataSheet = sheet => {
        try {
            const lastRow = sheet.getLastRow();
            if (lastRow < 2) return true; // Keine Daten zum Aktualisieren

            const numRows = lastRow - 1;
            const name = sheet.getName();

            // Passende Spaltenkonfiguration für das entsprechende Sheet auswählen
            let columns;
            if (name === "Einnahmen") {
                columns = config.sheets.einnahmen.columns;
            } else if (name === "Ausgaben") {
                columns = config.sheets.ausgaben.columns;
            } else if (name === "Eigenbelege") {
                columns = config.sheets.eigenbelege.columns;
            } else {
                return false; // Unbekanntes Sheet
            }

            // Spaltenbuchstaben aus den Indizes generieren
            const columnLetters = {};
            Object.entries(columns).forEach(([key, index]) => {
                columnLetters[key] = Helpers.getColumnLetter(index);
            });

            // Formeln für verschiedene Spalten setzen (mit konfigurierten Spaltenbuchstaben)
            const formulas = {};

            // MwSt-Betrag
            formulas[columns.mwstBetrag] = row =>
                `=${columnLetters.nettobetrag}${row}*${columnLetters.mwstSatz}${row}`;

            // Brutto-Betrag
            formulas[columns.bruttoBetrag] = row =>
                `=${columnLetters.nettobetrag}${row}+${columnLetters.mwstBetrag}${row}`;

            // Steuerbemessungsgrundlage - für Teilzahlungen
            formulas[columns.steuerbemessung] = row =>
                `=(${columnLetters.bruttoBetrag}${row}-${columnLetters.bezahlt}${row})/(1+${columnLetters.mwstSatz}${row})`;

            // Quartal
            formulas[columns.quartal] = row =>
                `=IF(${columnLetters.datum}${row}="";"";ROUNDUP(MONTH(${columnLetters.datum}${row})/3;0))`;

            // Zahlungsstatus
            formulas[columns.zahlungsstatus] = row =>
                `=IF(VALUE(${columnLetters.bezahlt}${row})=0;"Offen";IF(VALUE(${columnLetters.bezahlt}${row})>=VALUE(${columnLetters.bruttoBetrag}${row});"Bezahlt";"Teilbezahlt"))`;

            // Formeln für jede Spalte anwenden
            Object.entries(formulas).forEach(([col, formulaFn]) => {
                const formulasArray = Array.from({length: numRows}, (_, i) => [formulaFn(i + 2)]);
                sheet.getRange(2, Number(col), numRows, 1).setFormulas(formulasArray);
            });

            // Bezahlter Betrag - Leerzeichen durch 0 ersetzen für Berechnungen
            const bezahltRange = sheet.getRange(2, columns.bezahlt, numRows, 1);
            const bezahltValues = bezahltRange.getValues().map(([val]) => (val === "" || val === null ? 0 : val));
            bezahltRange.setValues(bezahltValues.map(val => [val]));

            // Dropdown-Validierungen je nach Sheet-Typ setzen
            if (name === "Einnahmen") {
                Validator.validateDropdown(
                    sheet, 2, columns.kategorie, numRows, 1,
                    Object.keys(config.einnahmen.categories)
                );
            } else if (name === "Ausgaben") {
                Validator.validateDropdown(
                    sheet, 2, columns.kategorie, numRows, 1,
                    Object.keys(config.ausgaben.categories)
                );
            } else if (name === "Eigenbelege") {
                Validator.validateDropdown(
                    sheet, 2, columns.kategorie, numRows, 1,
                    config.eigenbelege.category
                );

                // Für Eigenbelege: Status-Dropdown hinzufügen
                Validator.validateDropdown(
                    sheet, 2, columns.status, numRows, 1,
                    config.eigenbelege.status
                );

                // Bedingte Formatierung für Status-Spalte (nur für Eigenbelege)
                Helpers.setConditionalFormattingForColumn(sheet, columnLetters.status, [
                    {value: "Offen", background: "#FFC7CE", fontColor: "#9C0006"},
                    {value: "Erstattet", background: "#FFEB9C", fontColor: "#9C6500"},
                    {value: "Gebucht", background: "#C6EFCE", fontColor: "#006100"}
                ]);
            }

            // Zahlungsart-Dropdown für alle Blätter
            Validator.validateDropdown(
                sheet, 2, columns.zahlungsart, numRows, 1,
                config.common.paymentType
            );

            // Bedingte Formatierung für Zahlungsstatus-Spalte (für alle außer Eigenbelege)
            if (name !== "Eigenbelege") {
                Helpers.setConditionalFormattingForColumn(sheet, columnLetters.zahlungsstatus, [
                    {value: "Offen", background: "#FFC7CE", fontColor: "#9C0006"},
                    {value: "Teilbezahlt", background: "#FFEB9C", fontColor: "#9C6500"},
                    {value: "Bezahlt", background: "#C6EFCE", fontColor: "#006100"}
                ]);
            }

            // Spaltenbreiten automatisch anpassen
            sheet.autoResizeColumns(1, sheet.getLastColumn());

            return true;
        } catch (e) {
            console.error(`Fehler beim Aktualisieren von ${sheet.getName()}:`, e);
            return false;
        }
    };

    /**
     * Aktualisiert das Bankbewegungen-Sheet
     * @param {Sheet} sheet - Das Bankbewegungen-Sheet
     * @returns {boolean} - true bei erfolgreicher Aktualisierung
     */
    const refreshBankSheet = sheet => {
        try {
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            const lastRow = sheet.getLastRow();
            if (lastRow < 3) return true; // Nicht genügend Daten zum Aktualisieren

            const firstDataRow = 3; // Erste Datenzeile (nach Header-Zeile)
            const numDataRows = lastRow - firstDataRow + 1;
            const transRows = lastRow - firstDataRow - 1; // Anzahl der Transaktionszeilen ohne die letzte Zeile

            // Saldo-Formeln setzen (jede Zeile verwendet den Saldo der vorherigen Zeile + aktuellen Betrag)
            if (transRows > 0) {
                sheet.getRange(firstDataRow, 4, transRows, 1).setFormulas(
                    Array.from({length: transRows}, (_, i) =>
                        [`=D${firstDataRow + i - 1}+C${firstDataRow + i}`]
                    )
                );
            }

            // Transaktionstyp basierend auf dem Betrag setzen (Einnahme/Ausgabe)
            const amounts = sheet.getRange(firstDataRow, 3, numDataRows, 1).getValues();
            const typeValues = amounts.map(([val]) => {
                const amt = parseFloat(val) || 0;
                return [amt > 0 ? "Einnahme" : amt < 0 ? "Ausgabe" : ""];
            });
            sheet.getRange(firstDataRow, 5, numDataRows, 1).setValues(typeValues);

            // Dropdown-Validierungen für Typ, Kategorie und Konten
            Validator.validateDropdown(
                sheet, firstDataRow, 5, numDataRows, 1,
                config.bank.type
            );

            Validator.validateDropdown(
                sheet, firstDataRow, 6, numDataRows, 1,
                config.bank.category
            );

            // Konten für Dropdown-Validierung sammeln
            const allowedKontoSoll = Object.values(config.einnahmen.kontoMapping)
                .concat(Object.values(config.ausgaben.kontoMapping))
                .map(m => m.soll);

            const allowedGegenkonto = Object.values(config.einnahmen.kontoMapping)
                .concat(Object.values(config.ausgaben.kontoMapping))
                .map(m => m.gegen);

            // Dropdown-Validierungen für Konten setzen
            Validator.validateDropdown(
                sheet, firstDataRow, 7, numDataRows, 1,
                allowedKontoSoll
            );

            Validator.validateDropdown(
                sheet, firstDataRow, 8, numDataRows, 1,
                allowedGegenkonto
            );

            // Bedingte Formatierung für Transaktionstyp-Spalte
            Helpers.setConditionalFormattingForColumn(sheet, "E", [
                {value: "Einnahme", background: "#C6EFCE", fontColor: "#006100"},
                {value: "Ausgabe", background: "#FFC7CE", fontColor: "#9C0006"}
            ]);

            // REFERENZEN-MATCHING: Suche nach Referenzen in Einnahmen- und Ausgaben-Sheets

            // Daten aus Einnahmen-Sheet
            const einnahmenSheet = ss.getSheetByName("Einnahmen");
            let einnahmenData = [];
            if (einnahmenSheet && einnahmenSheet.getLastRow() > 1) {
                const numEinnahmenRows = einnahmenSheet.getLastRow() - 1;
                // Jetzt auch Beträge (E) und bezahlte Beträge (I) laden
                einnahmenData = einnahmenSheet.getRange(2, 2, numEinnahmenRows, 8).getDisplayValues();
                // B (0), E (3), I (7) Spaltenindizes nach 0-basierter Nummerierung
            }

            // Daten aus Ausgaben-Sheet
            const ausgabenSheet = ss.getSheetByName("Ausgaben");
            let ausgabenData = [];
            if (ausgabenSheet && ausgabenSheet.getLastRow() > 1) {
                const numAusgabenRows = ausgabenSheet.getLastRow() - 1;
                // Jetzt auch Beträge (E) und bezahlte Beträge (I) laden
                ausgabenData = ausgabenSheet.getRange(2, 2, numAusgabenRows, 8).getDisplayValues();
                // B (0), E (3), I (7) Spaltenindizes nach 0-basierter Nummerierung
            }

            // Bankbewegungen Daten für Verarbeitung holen
            const bankData = sheet.getRange(firstDataRow, 1, numDataRows, 12).getDisplayValues();

            // Cache für schnellere Suche
            const einnahmenMap = createReferenceMap(einnahmenData);
            const ausgabenMap = createReferenceMap(ausgabenData);

            // Durchlaufe jede Bankbewegung und suche nach Übereinstimmungen
            for (let i = 0; i < bankData.length; i++) {
                const rowIndex = i + firstDataRow;
                const row = bankData[i];

                // Prüfe, ob es sich um die Endsaldo-Zeile handelt
                const label = row[1] ? row[1].toString().trim().toLowerCase() : "";
                if (rowIndex === lastRow && label === "endsaldo") continue;

                const tranType = row[4]; // Spalte E: Einnahme/Ausgabe
                const refNumber = row[8]; // Spalte I: Referenznummer

                // Nur prüfen, wenn Referenz nicht leer ist
                if (refNumber && refNumber.trim() !== "") {
                    let matchFound = false;
                    let matchInfo = "";

                    const refTrimmed = refNumber.toString().trim();

                    // Betrag für den Vergleich (als absoluter Wert)
                    const betragValue = Math.abs(parseFloat(row[2]) || 0);

                    // In Abhängigkeit vom Typ im entsprechenden Sheet suchen
                    if (tranType === "Einnahme") {
                        // Optimierte Suche in Einnahmen mittels Map
                        const matchResult = findMatch(refTrimmed, einnahmenMap, betragValue);
                        if (matchResult) {
                            matchFound = true;
                            let matchStatus = "";

                            // Je nach Match-Typ unterschiedliche Statusinformationen
                            if (matchResult.matchType) {
                                // Bei "Unsichere Zahlung" auch die Differenz anzeigen
                                if (matchResult.matchType === "Unsichere Zahlung" && matchResult.betragsDifferenz) {
                                    matchStatus = ` (${matchResult.matchType}, Diff: ${matchResult.betragsDifferenz}€)`;
                                } else {
                                    matchStatus = ` (${matchResult.matchType})`;
                                }
                            }

                            // Prüfen, ob Zahldatum in Einnahmen/Ausgaben aktualisiert werden soll
                            if ((matchResult.matchType === "Vollständige Zahlung" ||
                                    matchResult.matchType === "Teilzahlung") &&
                                tranType === "Einnahme") {
                                // Bankbewegungsdatum holen (aus Spalte A)
                                const zahlungsDatum = row[0];
                                if (zahlungsDatum) {
                                    // Zahldatum im Einnahmen-Sheet aktualisieren (nur wenn leer)
                                    const einnahmenRow = matchResult.row;
                                    const zahldatumRange = einnahmenSheet.getRange(einnahmenRow, 14); // Spalte N für Zahldatum
                                    const aktuellDatum = zahldatumRange.getValue();

                                    if (!aktuellDatum || aktuellDatum === "") {
                                        zahldatumRange.setValue(zahlungsDatum);
                                        matchStatus += " ✓ Datum aktualisiert";
                                    }
                                }
                            }

                            matchInfo = `Einnahme #${matchResult.row}${matchStatus}`;
                        }
                    } else if (tranType === "Ausgabe") {
                        // Optimierte Suche in Ausgaben mittels Map
                        const matchResult = findMatch(refTrimmed, ausgabenMap, betragValue);
                        if (matchResult) {
                            matchFound = true;
                            let matchStatus = "";

                            // Je nach Match-Typ unterschiedliche Statusinformationen für Ausgaben
                            if (matchResult.matchType) {
                                if (matchResult.matchType === "Unsichere Zahlung" && matchResult.betragsDifferenz) {
                                    matchStatus = ` (${matchResult.matchType}, Diff: ${matchResult.betragsDifferenz}€)`;
                                } else {
                                    matchStatus = ` (${matchResult.matchType})`;
                                }
                            }

                            // Prüfen, ob Zahldatum in Ausgaben aktualisiert werden soll
                            if ((matchResult.matchType === "Vollständige Zahlung" ||
                                    matchResult.matchType === "Teilzahlung") &&
                                ausgabenSheet) {
                                // Bankbewegungsdatum holen (aus Spalte A)
                                const zahlungsDatum = row[0];
                                if (zahlungsDatum) {
                                    // Zahldatum im Ausgaben-Sheet aktualisieren (nur wenn leer)
                                    const ausgabenRow = matchResult.row;
                                    const zahldatumRange = ausgabenSheet.getRange(ausgabenRow, 14); // Spalte N für Zahldatum
                                    const aktuellDatum = zahldatumRange.getValue();

                                    if (!aktuellDatum || aktuellDatum === "") {
                                        zahldatumRange.setValue(zahlungsDatum);
                                        matchStatus += " ✓ Datum aktualisiert";
                                    }
                                }
                            }

                            matchInfo = `Ausgabe #${matchResult.row}${matchStatus}`;
                        }

                        // FALLS keine Übereinstimmung, auch in Einnahmen suchen (für Gutschriften)
                        if (!matchFound) {
                            // Bei einer Ausgabe, die möglicherweise eine Gutschrift ist,
                            // ignorieren wir den Betrag für den ersten Vergleich, um die entsprechende Einnahme zu finden
                            const gutschriftMatch = findMatch(refTrimmed, einnahmenMap);

                            if (gutschriftMatch) {
                                matchFound = true;
                                let matchStatus = "";

                                // Bei Gutschriften könnte der Betrag abweichen (z.B. Teilgutschrift)
                                // Prüfen, ob die Beträge plausibel sind
                                const gutschriftBetrag = Math.abs(gutschriftMatch.betrag);

                                if (Math.abs(betragValue - gutschriftBetrag) <= 0.01) {
                                    // Beträge stimmen genau überein
                                    matchStatus = " (Vollständige Gutschrift)";
                                } else if (betragValue < gutschriftBetrag) {
                                    // Teilgutschrift (Gutschriftbetrag kleiner als ursprünglicher Rechnungsbetrag)
                                    matchStatus = " (Teilgutschrift)";
                                } else {
                                    // Ungewöhnlicher Fall - Gutschriftbetrag größer als Rechnungsbetrag
                                    matchStatus = " (Ungewöhnliche Gutschrift)";
                                }

                                // Bei Gutschriften auch im Einnahmen-Sheet aktualisieren - hier als negative Zahlung
                                if (einnahmenSheet) {
                                    // Bankbewegungsdatum holen (aus Spalte A)
                                    const gutschriftDatum = row[0];
                                    if (gutschriftDatum) {
                                        // Gutschriftdatum im Einnahmen-Sheet aktualisieren und "G-" vor die Referenz setzen
                                        const einnahmenRow = gutschriftMatch.row;

                                        // Zahldatum aktualisieren (nur wenn leer)
                                        const zahldatumRange = einnahmenSheet.getRange(einnahmenRow, 14); // Spalte N für Zahldatum
                                        const aktuellDatum = zahldatumRange.getValue();

                                        if (!aktuellDatum || aktuellDatum === "") {
                                            zahldatumRange.setValue(gutschriftDatum);
                                            matchStatus += " ✓ Datum aktualisiert";
                                        }

                                        // Optional: Die Referenz mit "G-" kennzeichnen, um Gutschrift zu markieren
                                        const refRange = einnahmenSheet.getRange(einnahmenRow, 2); // Spalte B für Referenz
                                        const currentRef = refRange.getValue();
                                        if (currentRef && !currentRef.toString().startsWith("G-")) {
                                            refRange.setValue("G-" + currentRef);
                                            matchStatus += " ✓ Ref. markiert";
                                        }
                                    }
                                }

                                matchInfo = `Gutschrift zu Einnahme #${gutschriftMatch.row}${matchStatus}`;
                            }
                        }
                    }

                    // Spezialfälle prüfen
                    if (!matchFound) {
                        const lcRef = refTrimmed.toLowerCase();
                        if (lcRef.includes("gesellschaftskonto")) {
                            matchFound = true;
                            matchInfo = tranType === "Einnahme"
                                ? "Gesellschaftskonto (Einnahme)"
                                : "Gesellschaftskonto (Ausgabe)";
                        } else if (lcRef.includes("holding")) {
                            matchFound = true;
                            matchInfo = tranType === "Einnahme"
                                ? "Holding (Einnahme)"
                                : "Holding (Ausgabe)";
                        }
                    }

                    // Ergebnis in Spalte L speichern
                    row[11] = matchFound ? matchInfo : "";
                } else {
                    row[11] = ""; // Leere Spalte L
                }

                // Kontonummern basierend auf Kategorie setzen
                const category = row[5] || "";
                let mapping = null;

                if (tranType === "Einnahme") {
                    mapping = config.einnahmen.kontoMapping[category];
                } else if (tranType === "Ausgabe") {
                    mapping = config.ausgaben.kontoMapping[category];
                }

                if (!mapping) {
                    mapping = {soll: "Manuell prüfen", gegen: "Manuell prüfen"};
                }

                row[6] = mapping.soll;
                row[7] = mapping.gegen;
            }

            // Zuerst nur Spalte L aktualisieren (für bessere Performance und Fehlerbehandlung)
            const matchColumn = bankData.map(row => [row[11]]);
            sheet.getRange(firstDataRow, 12, numDataRows, 1).setValues(matchColumn);

            // Dann die restlichen Daten zurückschreiben
            sheet.getRange(firstDataRow, 1, numDataRows, 11).setValues(
                bankData.map(row => row.slice(0, 11))
            );

            // Verzögerung hinzufügen, um sicherzustellen, dass die Daten verarbeitet wurden
            console.log("Daten wurden zurückgeschrieben, warte 1 Sekunde vor der Formatierung...");
            Utilities.sleep(1000);

            // Log zur Überprüfung
            console.log("Beginne mit Zeilenformatierung für " + bankData.length + " Zeilen");

            // Formatiere die gesamten Zeilen basierend auf dem Match-Typ
            // Wir verarbeiten jede Zeile einzeln, um Probleme zu isolieren
            for (let i = 0; i < bankData.length; i++) {
                try {
                    const rowIndex = firstDataRow + i;
                    const matchInfo = bankData[i][11]; // Spalte L mit Match-Info

                    // Nur formatieren, wenn die Zeile existiert und eine Match-Info hat
                    if (rowIndex <= sheet.getLastRow() && matchInfo && matchInfo.trim() !== "") {
                        console.log(`Versuche Formatierung für Zeile ${rowIndex}, Match: "${matchInfo}"`);

                        // Zuerst den Hintergrund zurücksetzen
                        const rowRange = sheet.getRange(rowIndex, 1, 1, 12);
                        rowRange.setBackground(null);

                        // Kurze Pause, um die Sheets-API nicht zu überlasten
                        if (i % 5 === 0) Utilities.sleep(100);

                        // Dann die neue Farbe anwenden, je nach Match-Typ
                        if (matchInfo.includes("Einnahme")) {
                            if (matchInfo.includes("Vollständige Zahlung")) {
                                console.log(`Setze Grün für Zeile ${rowIndex}`);
                                rowRange.setBackground("#C6EFCE"); // Kräftiges Grün
                            } else if (matchInfo.includes("Teilzahlung")) {
                                console.log(`Setze Orange für Zeile ${rowIndex}`);
                                rowRange.setBackground("#FCE4D6"); // Helles Orange
                            } else {
                                console.log(`Setze Hellgrün für Zeile ${rowIndex}`);
                                rowRange.setBackground("#EAF1DD"); // Helles Grün
                            }
                        } else if (matchInfo.includes("Ausgabe")) {
                            if (matchInfo.includes("Vollständige Zahlung")) {
                                console.log(`Setze Rot für Zeile ${rowIndex}`);
                                rowRange.setBackground("#FFC7CE"); // Helles Rot
                            } else if (matchInfo.includes("Teilzahlung")) {
                                console.log(`Setze Orange für Zeile ${rowIndex}`);
                                rowRange.setBackground("#FCE4D6"); // Helles Orange
                            } else {
                                console.log(`Setze Rosa für Zeile ${rowIndex}`);
                                rowRange.setBackground("#FFCCCC"); // Helles Rosa
                            }
                        } else if (matchInfo.includes("Gutschrift")) {
                            console.log(`Setze Lila für Zeile ${rowIndex}`);
                            rowRange.setBackground("#E6E0FF"); // Helles Lila
                        } else if (matchInfo.includes("Gesellschaftskonto") || matchInfo.includes("Holding")) {
                            console.log(`Setze Gelb für Zeile ${rowIndex}`);
                            rowRange.setBackground("#FFEB9C"); // Helles Gelb
                        }
                    }
                } catch (err) {
                    console.error(`Fehler beim Formatieren von Zeile ${firstDataRow + i}:`, err);
                }
            }

            console.log("Zeilenformatierung abgeschlossen");

            // Bedingte Formatierung für Match-Spalte mit verbesserten Farben
            Helpers.setConditionalFormattingForColumn(sheet, "L", [
                // Grundlegende Match-Typen
                {value: "Einnahme", background: "#C6EFCE", fontColor: "#006100", pattern: "beginsWith"},
                {value: "Ausgabe", background: "#FFC7CE", fontColor: "#9C0006", pattern: "beginsWith"},
                {value: "Gutschrift", background: "#E6E0FF", fontColor: "#5229A3", pattern: "beginsWith"},
                {value: "Gesellschaftskonto", background: "#FFEB9C", fontColor: "#9C6500", pattern: "beginsWith"},
                {value: "Holding", background: "#FFEB9C", fontColor: "#9C6500", pattern: "beginsWith"},

                // Zusätzliche Betragstypen
                {value: "Vollständige Zahlung", background: "#C6EFCE", fontColor: "#006100", pattern: "contains"},
                {value: "Teilzahlung", background: "#FCE4D6", fontColor: "#974706", pattern: "contains"},
                {value: "Unsichere Zahlung", background: "#F8CBAD", fontColor: "#843C0C", pattern: "contains"},
                {value: "Vollständige Gutschrift", background: "#E6E0FF", fontColor: "#5229A3", pattern: "contains"},
                {value: "Teilgutschrift", background: "#E6E0FF", fontColor: "#5229A3", pattern: "contains"},
                {value: "Ungewöhnliche Gutschrift", background: "#FFD966", fontColor: "#7F6000", pattern: "contains"}
            ]);

            // Endsaldo-Zeile aktualisieren
            const lastRowText = sheet.getRange(lastRow, 2).getDisplayValue().toString().trim().toLowerCase();
            const formattedDate = Utilities.formatDate(
                new Date(),
                Session.getScriptTimeZone(),
                "dd.MM.yyyy"
            );

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

            // Spaltenbreiten anpassen
            sheet.autoResizeColumns(1, sheet.getLastColumn());

            // Verzögerung vor dem Aufruf von markPaidInvoices
            console.log("Warte 1 Sekunde vor dem Markieren der bezahlten Rechnungen...");
            Utilities.sleep(1000);

            // Setze farbliche Markierung in den Einnahmen/Ausgaben Sheets basierend auf Zahlungsstatus
            markPaidInvoices(einnahmenSheet, ausgabenSheet);

            return true;
        } catch (e) {
            console.error("Fehler beim Aktualisieren des Bankbewegungen-Sheets:", e);
            return false;
        }
    };

    /**
     * Erstellt eine Map aus Referenznummern für schnellere Suche
     * @param {Array} data - Array mit Referenznummern und Beträgen
     * @returns {Object} - Map mit Referenznummern als Keys
     */
    function createReferenceMap(data) {
        const map = {};
        for (let i = 0; i < data.length; i++) {
            const ref = data[i][0]; // Referenz in Spalte B (Index 0)
            if (ref && ref.trim() !== "") {
                // Entferne "G-" Prefix für den Key, falls vorhanden (für Gutschriften)
                let key = ref.trim();
                const isGutschrift = key.startsWith("G-");
                if (isGutschrift) {
                    key = key.substring(2); // Entferne "G-" Prefix für den Schlüssel
                }

                // Netto-Betrag aus Spalte E (Index 3)
                let betrag = 0;
                if (data[i][3] !== undefined && data[i][3] !== null && data[i][3] !== "") {
                    // Betragsstring säubern
                    const betragStr = data[i][3].toString().replace(/[^0-9.,-]/g, "").replace(",", ".");
                    betrag = parseFloat(betragStr) || 0;

                    // Bei Gutschriften ist der Betrag im Sheet negativ, wir speichern den Absolutwert
                    betrag = Math.abs(betrag);
                }

                // MwSt-Satz aus Spalte F (Index 4)
                let mwstRate = 0;
                if (data[i][4] !== undefined && data[i][4] !== null && data[i][4] !== "") {
                    // MwSt-Satz säubern und parsen
                    const mwstStr = data[i][4].toString().replace(/[^0-9.,-]/g, "").replace(",", ".");
                    mwstRate = parseFloat(mwstStr) || 0;

                    // Wenn der Wert > 1 ist, nehmen wir an, dass es sich um einen Prozentsatz handelt
                    if (mwstRate > 1) {
                        mwstRate = mwstRate;
                    } else {
                        // Ansonsten als Dezimalwert interpretieren und in Prozent umrechnen
                        mwstRate = mwstRate * 100;
                    }
                }

                // Bezahlter Betrag aus Spalte I (Index 7)
                let bezahlt = 0;
                if (data[i][7] !== undefined && data[i][7] !== null && data[i][7] !== "") {
                    // Betragsstring säubern
                    const bezahltStr = data[i][7].toString().replace(/[^0-9.,-]/g, "").replace(",", ".");
                    bezahlt = parseFloat(bezahltStr) || 0;

                    // Bei Gutschriften ist der bezahlte Betrag im Sheet negativ, wir speichern den Absolutwert
                    bezahlt = Math.abs(bezahlt);
                }

                // Speichere auch den Zeilen-Index und die Beträge
                map[key] = {
                    ref: ref.trim(), // Originale Referenz mit G-Prefix, falls vorhanden
                    row: i + 2,
                    betrag: betrag,
                    mwstRate: mwstRate,
                    bezahlt: bezahlt,
                    offen: betrag * (1 + mwstRate/100) - bezahlt,
                    isGutschrift: isGutschrift
                };
            }
        }
        return map;
    }

    /**
     * Markiert bezahlte Einnahmen und Ausgaben farblich basierend auf dem Zahlungsstatus
     * @param {Sheet} einnahmenSheet - Das Einnahmen-Sheet
     * @param {Sheet} ausgabenSheet - Das Ausgaben-Sheet
     */
    function markPaidInvoices(einnahmenSheet, ausgabenSheet) {
        // Sammle zugeordnete Referenzen aus dem Bankbewegungen-Sheet
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const bankSheet = ss.getSheetByName("Bankbewegungen");

        // Map zum Speichern der zugeordneten Referenzen und ihrer Bankbewegungsinformationen
        const bankZuordnungen = {};

        if (bankSheet && bankSheet.getLastRow() > 2) {
            const bankData = bankSheet.getRange(3, 1, bankSheet.getLastRow() - 2, 12).getDisplayValues();

            for (const row of bankData) {
                const matchInfo = row[11]; // Spalte L: Match-Info
                const transTyp = row[4];   // Spalte E: Einnahme/Ausgabe
                const bankDatum = row[0];  // Spalte A: Datum

                if (matchInfo && matchInfo.trim() !== "") {
                    // Spezialfall: Gutschriften zu Einnahmen
                    // Das Format der Match-Info ist "Gutschrift zu Einnahme #12"
                    const gutschriftPattern = matchInfo.match(/Gutschrift zu Einnahme #(\d+)/i);

                    if (gutschriftPattern && gutschriftPattern.length >= 2) {
                        const rowNum = parseInt(gutschriftPattern[1]);

                        // Schlüssel für Einnahme erstellen, nicht für Gutschrift
                        const key = `einnahme#${rowNum}`;

                        // Informationen zur Bankbewegung speichern oder aktualisieren
                        if (!bankZuordnungen[key]) {
                            bankZuordnungen[key] = {
                                typ: "einnahme",
                                row: rowNum,
                                bankDatum: bankDatum,
                                matchInfo: matchInfo,
                                transTyp: "Gutschrift"
                            };
                        } else {
                            // Wenn bereits ein Eintrag existiert, füge diesen als zusätzliche Info hinzu
                            bankZuordnungen[key].additional = bankZuordnungen[key].additional || [];
                            bankZuordnungen[key].additional.push({
                                bankDatum: bankDatum,
                                matchInfo: matchInfo
                            });
                        }
                    }
                    // Standardfall: Normale Einnahmen/Ausgaben
                    else {
                        // Z.B. "Einnahme #5 (Vollständige Zahlung)" oder "Ausgabe #7 (Teilzahlung)"
                        const matchParts = matchInfo.match(/(Einnahme|Ausgabe|Gutschrift)[^#]*#(\d+)/i);

                        if (matchParts && matchParts.length >= 3) {
                            const typ = matchParts[1].toLowerCase(); // "einnahme", "ausgabe", "gutschrift"
                            const rowNum = parseInt(matchParts[2]);  // Zeilennummer im Sheet

                            // Schlüssel für die Map erstellen
                            const key = `${typ}#${rowNum}`;

                            // Informationen zur Bankbewegung speichern oder aktualisieren
                            if (!bankZuordnungen[key]) {
                                bankZuordnungen[key] = {
                                    typ: typ,
                                    row: rowNum,
                                    bankDatum: bankDatum,
                                    matchInfo: matchInfo,
                                    transTyp: transTyp
                                };
                            } else {
                                // Wenn bereits ein Eintrag existiert, füge diesen als zusätzliche Info hinzu
                                bankZuordnungen[key].additional = bankZuordnungen[key].additional || [];
                                bankZuordnungen[key].additional.push({
                                    bankDatum: bankDatum,
                                    matchInfo: matchInfo
                                });
                            }
                        }
                    }
                }
            }
        }

        // Markiere bezahlte Einnahmen
        if (einnahmenSheet && einnahmenSheet.getLastRow() > 1) {
            const numEinnahmenRows = einnahmenSheet.getLastRow() - 1;

            // Hole Werte aus dem Einnahmen-Sheet
            const einnahmenData = einnahmenSheet.getRange(2, 1, numEinnahmenRows, 14).getValues();

            // Für jede Zeile prüfen
            for (let i = 0; i < einnahmenData.length; i++) {
                const row = i + 2; // Aktuelle Zeile im Sheet
                const nettoBetrag = parseFloat(einnahmenData[i][4]) || 0; // Spalte E
                const bezahltBetrag = parseFloat(einnahmenData[i][8]) || 0; // Spalte I
                const zahlungsDatum = einnahmenData[i][13]; // Spalte N
                const referenz = einnahmenData[i][1]; // Spalte B

                // Prüfe, ob es eine Gutschrift ist
                const isGutschrift = referenz && referenz.toString().startsWith("G-");

                // Prüfe, ob diese Einnahme im Banking-Sheet zugeordnet wurde
                const zuordnungsKey = `einnahme#${row}`;
                const hatBankzuordnung = bankZuordnungen[zuordnungsKey] !== undefined;

                // Zeilenbereich für die Formatierung
                const rowRange = einnahmenSheet.getRange(row, 1, 1, 14);

                // Status basierend auf Zahlung setzen
                if (Math.abs(bezahltBetrag) >= Math.abs(nettoBetrag) * 0.999) { // 99.9% bezahlt wegen Rundungsfehlern
                    // Vollständig bezahlt
                    if (zahlungsDatum) {
                        if (isGutschrift) {
                            // Gutschriften in Lila markieren
                            rowRange.setBackground("#E6E0FF"); // Helles Lila
                        } else if (hatBankzuordnung) {
                            // Bezahlte Rechnungen mit Bank-Zuordnung in kräftigerem Grün markieren
                            rowRange.setBackground("#C6EFCE"); // Kräftigeres Grün
                        } else {
                            // Bezahlte Rechnungen ohne Bank-Zuordnung in hellerem Grün markieren
                            rowRange.setBackground("#EAF1DD"); // Helles Grün
                        }
                    } else {
                        // Bezahlt aber kein Datum - in Orange markieren
                        rowRange.setBackground("#FCE4D6"); // Helles Orange
                    }
                } else if (bezahltBetrag > 0) {
                    // Teilzahlung
                    if (hatBankzuordnung) {
                        rowRange.setBackground("#FFC7AA"); // Kräftigeres Orange für Teilzahlungen mit Bank-Zuordnung
                    } else {
                        rowRange.setBackground("#FCE4D6"); // Helles Orange für normale Teilzahlungen
                    }
                } else {
                    // Unbezahlt - keine spezielle Farbe
                    rowRange.setBackground(null);
                }

                // Optional: Wenn eine Bankzuordnung existiert, in Spalte O einen Hinweis setzen
                if (hatBankzuordnung) {
                    const zuordnungsInfo = bankZuordnungen[zuordnungsKey];
                    let infoText = "✓ Bank: " + zuordnungsInfo.bankDatum;

                    // Wenn es mehrere Zuordnungen gibt (z.B. bei aufgeteilten Zahlungen)
                    if (zuordnungsInfo.additional && zuordnungsInfo.additional.length > 0) {
                        infoText += " + " + zuordnungsInfo.additional.length + " weitere";
                    }

                    // Titelzeile für Spalte O setzen, falls noch nicht vorhanden
                    if (einnahmenSheet.getRange(1, 15).getValue() === "") {
                        einnahmenSheet.getRange(1, 15).setValue("Bank-Abgleich");
                    }

                    einnahmenSheet.getRange(row, 15).setValue(infoText);
                } else {
                    // Titelzeile für Spalte O setzen, falls noch nicht vorhanden
                    if (einnahmenSheet.getRange(1, 15).getValue() === "") {
                        einnahmenSheet.getRange(1, 15).setValue("Bank-Abgleich");
                    }

                    // Setze die Zelle leer, falls keine Zuordnung existiert
                    const currentValue = einnahmenSheet.getRange(row, 15).getValue();
                    if (currentValue && currentValue.toString().startsWith("✓ Bank:")) {
                        einnahmenSheet.getRange(row, 15).setValue("");
                    }
                }
            }
        }

        // Markiere bezahlte Ausgaben (ähnliche Logik wie bei Einnahmen)
        if (ausgabenSheet && ausgabenSheet.getLastRow() > 1) {
            const numAusgabenRows = ausgabenSheet.getLastRow() - 1;

            // Hole Werte aus dem Ausgaben-Sheet
            const ausgabenData = ausgabenSheet.getRange(2, 1, numAusgabenRows, 14).getValues();

            // Für jede Zeile prüfen
            for (let i = 0; i < ausgabenData.length; i++) {
                const row = i + 2; // Aktuelle Zeile im Sheet
                const nettoBetrag = parseFloat(ausgabenData[i][4]) || 0; // Spalte E
                const bezahltBetrag = parseFloat(ausgabenData[i][8]) || 0; // Spalte I
                const zahlungsDatum = ausgabenData[i][13]; // Spalte N

                // Prüfe, ob diese Ausgabe im Banking-Sheet zugeordnet wurde
                const zuordnungsKey = `ausgabe#${row}`;
                const hatBankzuordnung = bankZuordnungen[zuordnungsKey] !== undefined;

                // Zeilenbereich für die Formatierung
                const rowRange = ausgabenSheet.getRange(row, 1, 1, 14);

                // Status basierend auf Zahlung setzen
                if (Math.abs(bezahltBetrag) >= Math.abs(nettoBetrag) * 0.999) { // 99.9% bezahlt wegen Rundungsfehlern
                    // Vollständig bezahlt
                    if (zahlungsDatum) {
                        if (hatBankzuordnung) {
                            // Bezahlte Ausgaben mit Bank-Zuordnung in kräftigerem Grün markieren
                            rowRange.setBackground("#C6EFCE"); // Kräftigeres Grün
                        } else {
                            // Bezahlte Ausgaben ohne Bank-Zuordnung in hellerem Grün markieren
                            rowRange.setBackground("#EAF1DD"); // Helles Grün
                        }
                    } else {
                        // Bezahlt aber kein Datum - in Orange markieren
                        rowRange.setBackground("#FCE4D6"); // Helles Orange
                    }
                } else if (bezahltBetrag > 0) {
                    // Teilzahlung
                    if (hatBankzuordnung) {
                        rowRange.setBackground("#FFC7AA"); // Kräftigeres Orange für Teilzahlungen mit Bank-Zuordnung
                    } else {
                        rowRange.setBackground("#FCE4D6"); // Helles Orange für normale Teilzahlungen
                    }
                } else {
                    // Unbezahlt - keine spezielle Farbe
                    rowRange.setBackground(null);
                }

                // Optional: Wenn eine Bankzuordnung existiert, in Spalte O einen Hinweis setzen
                if (hatBankzuordnung) {
                    const zuordnungsInfo = bankZuordnungen[zuordnungsKey];
                    let infoText = "✓ Bank: " + zuordnungsInfo.bankDatum;

                    // Wenn es mehrere Zuordnungen gibt (z.B. bei aufgeteilten Zahlungen)
                    if (zuordnungsInfo.additional && zuordnungsInfo.additional.length > 0) {
                        infoText += " + " + zuordnungsInfo.additional.length + " weitere";
                    }

                    // Titelzeile für Spalte O setzen, falls noch nicht vorhanden
                    if (ausgabenSheet.getRange(1, 15).getValue() === "") {
                        ausgabenSheet.getRange(1, 15).setValue("Bank-Abgleich");
                    }

                    ausgabenSheet.getRange(row, 15).setValue(infoText);
                } else {
                    // Titelzeile für Spalte O setzen, falls noch nicht vorhanden
                    if (ausgabenSheet.getRange(1, 15).getValue() === "") {
                        ausgabenSheet.getRange(1, 15).setValue("Bank-Abgleich");
                    }

                    // Setze die Zelle leer, falls keine Zuordnung existiert
                    const currentValue = ausgabenSheet.getRange(row, 15).getValue();
                    if (currentValue && currentValue.toString().startsWith("✓ Bank:")) {
                        ausgabenSheet.getRange(row, 15).setValue("");
                    }
                }
            }
        }
    }

    /**
     * Findet eine Übereinstimmung in der Referenz-Map
     * @param {string} reference - Zu suchende Referenz
     * @param {Object} refMap - Map mit Referenznummern
     * @param {number} betrag - Betrag der Bankbewegung (absoluter Wert)
     * @returns {Object|null} - Gefundene Übereinstimmung oder null
     */
    function findMatch(reference, refMap, betrag = null) {
        // 1. Exakte Übereinstimmung
        if (refMap[reference]) {
            const match = refMap[reference];

            // Wenn ein Betrag angegeben ist
            if (betrag !== null) {
                const matchNetto = Math.abs(match.betrag);
                const matchMwstRate = parseFloat(match.mwstRate || 0) / 100;

                // Bruttobetrag berechnen (Netto + MwSt)
                const matchBrutto = matchNetto * (1 + matchMwstRate);
                const matchBezahlt = Math.abs(match.bezahlt);

                // Beträge mit größerer Toleranz vergleichen (2 Cent Unterschied erlauben)
                const tolerance = 0.02;

                // Fall 1: Betrag entspricht genau dem Bruttobetrag (Vollständige Zahlung)
                if (Math.abs(betrag - matchBrutto) <= tolerance) {
                    match.matchType = "Vollständige Zahlung";
                    return match;
                }

                // Fall 2: Position ist bereits vollständig bezahlt
                if (Math.abs(matchBezahlt - matchBrutto) <= tolerance && matchBezahlt > 0) {
                    match.matchType = "Vollständige Zahlung";
                    return match;
                }

                // Fall 3: Teilzahlung (Bankbetrag kleiner als Rechnungsbetrag)
                // Nur als Teilzahlung markieren, wenn der Betrag deutlich kleiner ist (> 10% Differenz)
                if (betrag < matchBrutto && (matchBrutto - betrag) > (matchBrutto * 0.1)) {
                    match.matchType = "Teilzahlung";
                    return match;
                }

                // Fall 4: Betrag ist größer als Bruttobetrag, aber vermutlich trotzdem die richtige Zahlung
                // (z.B. wegen Rundungen oder kleinen Gebühren)
                if (betrag > matchBrutto && (betrag - matchBrutto) <= tolerance) {
                    match.matchType = "Vollständige Zahlung";
                    return match;
                }

                // Fall 5: Bei allen anderen Fällen (Beträge weichen stärker ab)
                match.matchType = "Unsichere Zahlung";
                match.betragsDifferenz = Math.abs(betrag - matchBrutto).toFixed(2);
                return match;
            } else {
                // Ohne Betragsvergleich
                match.matchType = "Referenz-Match";
                return match;
            }
        }

        // 2. Teilweise Übereinstimmung (in beide Richtungen)
        for (const key in refMap) {
            if (reference.includes(key) || key.includes(reference)) {
                const match = refMap[key];

                // Nur Betragsvergleich wenn nötig
                if (betrag !== null) {
                    const matchNetto = Math.abs(match.betrag);
                    const matchMwstRate = parseFloat(match.mwstRate || 0) / 100;

                    // Bruttobetrag berechnen (Netto + MwSt)
                    const matchBrutto = matchNetto * (1 + matchMwstRate);
                    const matchBezahlt = Math.abs(match.bezahlt);

                    // Größere Toleranz für Teilweise Übereinstimmungen
                    const tolerance = 0.02;

                    // Beträge stimmen mit Toleranz überein
                    if (Math.abs(betrag - matchBrutto) <= tolerance) {
                        match.matchType = "Vollständige Zahlung";
                        return match;
                    }

                    // Wenn Position bereits vollständig bezahlt ist
                    if (Math.abs(matchBezahlt - matchBrutto) <= tolerance && matchBezahlt > 0) {
                        match.matchType = "Vollständige Zahlung";
                        return match;
                    }

                    // Teilzahlung mit größerer Toleranz für nicht-exakte Referenzen
                    // Nur als Teilzahlung markieren, wenn der Betrag deutlich kleiner ist (> 10% Differenz)
                    if (betrag < matchBrutto && (matchBrutto - betrag) > (matchBrutto * 0.1)) {
                        match.matchType = "Teilzahlung";
                        return match;
                    }

                    // Bei allen anderen Fällen: Unsichere Zahlung
                    match.matchType = "Unsichere Zahlung";
                    match.betragsDifferenz = Math.abs(betrag - matchBrutto).toFixed(2);
                    return match;
                } else {
                    // Ohne Betragsvergleich
                    match.matchType = "Teilw. Match";
                    return match;
                }
            }
        }

        return null;
    }

    /**
     * Aktualisiert das aktive Tabellenblatt
     */
    const refreshActiveSheet = () => {
        try {
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
        } catch (e) {
            console.error("Fehler beim Aktualisieren des aktiven Sheets:", e);
            SpreadsheetApp.getUi().alert("Ein Fehler ist beim Aktualisieren aufgetreten: " + e.toString());
        }
    };

    /**
     * Aktualisiert alle relevanten Tabellenblätter
     */
    const refreshAllSheets = () => {
        try {
            const ss = SpreadsheetApp.getActiveSpreadsheet();

            ["Einnahmen", "Ausgaben", "Eigenbelege", "Bankbewegungen"].forEach(name => {
                const sheet = ss.getSheetByName(name);
                if (sheet) {
                    name === "Bankbewegungen" ? refreshBankSheet(sheet) : refreshDataSheet(sheet);
                }
            });
        } catch (e) {
            console.error("Fehler beim Aktualisieren aller Sheets:", e);
        }
    };

    // Öffentliche API des Moduls
    return {
        refreshActiveSheet,
        refreshAllSheets
    };
})();

export default RefreshModule;