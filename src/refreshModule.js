// file: src/refreshModule.js
import Validator from "./validator.js";
import config from "./config.js";
import Helpers from "./helpers.js";

/**
 * Modul zum Aktualisieren der Tabellenblätter und Neuberechnen von Formeln
 */
const RefreshModule = (() => {
    // Cache für wiederholte Berechnungen
    const _cache = {
        // Sheet-Referenzen Cache, um unnötige getSheetByName-Aufrufe zu vermeiden
        sheets: new Map(),
        // Referenz-Maps für schnellere Suche nach Rechnungsnummern
        references: {
            einnahmen: null,
            ausgaben: null,
            eigenbelege: null
        }
    };

    /**
     * Cache zurücksetzen
     */
    const clearCache = () => {
        _cache.sheets.clear();
        _cache.references.einnahmen = null;
        _cache.references.ausgaben = null;
        _cache.references.eigenbelege = null;
    };

    /**
     * Holt ein Sheet aus dem Cache oder vom Spreadsheet
     * @param {string} name - Name des Sheets
     * @returns {Sheet|null} - Das Sheet oder null, wenn nicht gefunden
     */
    const getSheet = (name) => {
        if (_cache.sheets.has(name)) {
            return _cache.sheets.get(name);
        }

        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
        if (sheet) {
            _cache.sheets.set(name, sheet);
        }
        return sheet;
    };

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
                columns = config.einnahmen.columns;
            } else if (name === "Ausgaben") {
                columns = config.ausgaben.columns;
            } else if (name === "Eigenbelege") {
                columns = config.eigenbelege.columns;
            } else {
                return false; // Unbekanntes Sheet
            }

            // Spaltenbuchstaben aus den Indizes generieren (mit Cache für Effizienz)
            const columnLetters = {};
            Object.entries(columns).forEach(([key, index]) => {
                columnLetters[key] = Helpers.getColumnLetter(index);
            });

            // Batch-Array für Formeln erstellen (effizienter als einzelne Range-Updates)
            const formulasBatch = {};

            // MwSt-Betrag
            formulasBatch[columns.mwstBetrag] = Array.from(
                {length: numRows},
                (_, i) => [`=${columnLetters.nettobetrag}${i + 2}*${columnLetters.mwstSatz}${i + 2}`]
            );

            // Brutto-Betrag
            formulasBatch[columns.bruttoBetrag] = Array.from(
                {length: numRows},
                (_, i) => [`=${columnLetters.nettobetrag}${i + 2}+${columnLetters.mwstBetrag}${i + 2}`]
            );

            // Bezahlter Betrag - für Teilzahlungen
            formulasBatch[columns.restbetragNetto] = Array.from(
                {length: numRows},
                (_, i) => [`=(${columnLetters.bruttoBetrag}${i + 2}-${columnLetters.bezahlt}${i + 2})/(1+${columnLetters.mwstSatz}${i + 2})`]
            );

            // Quartal
            formulasBatch[columns.quartal] = Array.from(
                {length: numRows},
                (_, i) => [`=IF(${columnLetters.datum}${i + 2}="";"";ROUNDUP(MONTH(${columnLetters.datum}${i + 2})/3;0))`]
            );

            if (name !== "Eigenbelege") {
                // Für Einnahmen und Ausgaben: Zahlungsstatus
                formulasBatch[columns.zahlungsstatus] = Array.from(
                    {length: numRows},
                    (_, i) => [`=IF(VALUE(${columnLetters.bezahlt}${i + 2})=0;"Offen";IF(VALUE(${columnLetters.bezahlt}${i + 2})>=VALUE(${columnLetters.bruttoBetrag}${i + 2});"Bezahlt";"Teilbezahlt"))`]
                );
            } else {
                // Für Eigenbelege: Zahlungsstatus
                formulasBatch[columns.zahlungsstatus] = Array.from(
                    {length: numRows},
                    (_, i) => [`=IF(VALUE(${columnLetters.bezahlt}${i + 2})=0;"Offen";IF(VALUE(${columnLetters.bezahlt}${i + 2})>=VALUE(${columnLetters.bruttoBetrag}${i + 2});"Erstattet";"Teilerstattet"))`]
                );
            }

            // Formeln in Batches anwenden (weniger API-Calls)
            Object.entries(formulasBatch).forEach(([col, formulas]) => {
                sheet.getRange(2, Number(col), numRows, 1).setFormulas(formulas);
            });

            // Bezahlter Betrag - Leerzeichen durch 0 ersetzen für Berechnungen
            const bezahltRange = sheet.getRange(2, columns.bezahlt, numRows, 1);
            const bezahltValues = bezahltRange.getValues();
            const updatedBezahltValues = bezahltValues.map(
                ([val]) => [Helpers.isEmpty(val) ? 0 : val]
            );
            bezahltRange.setValues(updatedBezahltValues);

            // Dropdown-Validierungen je nach Sheet-Typ setzen
            setDropdownValidations(sheet, name, numRows, columns);

            // Bedingte Formatierung für Status-Spalte
            if (name !== "Eigenbelege") {
                // Für Einnahmen und Ausgaben: Zahlungsstatus
                Helpers.setConditionalFormattingForColumn(sheet, columnLetters.zahlungsstatus, [
                    {value: "Offen", background: "#FFC7CE", fontColor: "#9C0006"},
                    {value: "Teilbezahlt", background: "#FFEB9C", fontColor: "#9C6500"},
                    {value: "Bezahlt", background: "#C6EFCE", fontColor: "#006100"}
                ]);
            } else {
                // Für Eigenbelege: Status
                Helpers.setConditionalFormattingForColumn(sheet, columnLetters.zahlungsstatus, [
                    {value: "Offen", background: "#FFC7CE", fontColor: "#9C0006"},
                    {value: "Teilerstattet", background: "#FFEB9C", fontColor: "#9C6500"},
                    {value: "Erstattet", background: "#C6EFCE", fontColor: "#006100"}
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
     * Setzt Dropdown-Validierungen für ein Sheet
     * @param {Sheet} sheet - Das Sheet
     * @param {string} sheetName - Name des Sheets
     * @param {number} numRows - Anzahl der Datenzeilen
     * @param {Object} columns - Spaltenkonfiguration
     */
    const setDropdownValidations = (sheet, sheetName, numRows, columns) => {
        if (sheetName === "Einnahmen") {
            Validator.validateDropdown(
                sheet, 2, columns.kategorie, numRows, 1,
                Object.keys(config.einnahmen.categories)
            );
        } else if (sheetName === "Ausgaben") {
            Validator.validateDropdown(
                sheet, 2, columns.kategorie, numRows, 1,
                Object.keys(config.ausgaben.categories)
            );
        } else if (sheetName === "Eigenbelege") {
            // Kategorie-Dropdown für Eigenbelege
            Validator.validateDropdown(
                sheet, 2, columns.kategorie, numRows, 1,
                Object.keys(config.eigenbelege.categories)
            );

            // Dropdown für "Ausgelegt von" hinzufügen (merged aus shareholders und employees)
            const ausleger = [
                ...config.common.shareholders,
                ...config.common.employees
            ];

            Validator.validateDropdown(
                sheet, 2, columns.ausgelegtVon, numRows, 1,
                ausleger
            );
        }

        // Zahlungsart-Dropdown für alle Blätter
        Validator.validateDropdown(
            sheet, 2, columns.zahlungsart, numRows, 1,
            config.common.paymentType
        );
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

            // Bankbewegungen-Konfiguration für Spalten
            const columns = config.bankbewegungen.columns;

            // Spaltenbuchstaben aus den Indizes generieren
            const columnLetters = {};
            Object.entries(columns).forEach(([key, index]) => {
                columnLetters[key] = Helpers.getColumnLetter(index);
            });

            // 1. Saldo-Formeln setzen (jede Zeile verwendet den Saldo der vorherigen Zeile + aktuellen Betrag)
            if (transRows > 0) {
                sheet.getRange(firstDataRow, columns.saldo, transRows, 1).setFormulas(
                    Array.from({length: transRows}, (_, i) =>
                        [`=${columnLetters.saldo}${firstDataRow + i - 1}+${columnLetters.betrag}${firstDataRow + i}`]
                    )
                );
            }

            // 2. Transaktionstyp basierend auf dem Betrag setzen (Einnahme/Ausgabe)
            const amounts = sheet.getRange(firstDataRow, columns.betrag, numDataRows, 1).getValues();
            const typeValues = amounts.map(([val]) => {
                const amt = parseFloat(val) || 0;
                return [amt > 0 ? "Einnahme" : amt < 0 ? "Ausgabe" : ""];
            });
            sheet.getRange(firstDataRow, columns.transaktionstyp, numDataRows, 1).setValues(typeValues);

            // 3. Dropdown-Validierungen für Typ, Kategorie und Konten
            applyBankSheetValidations(sheet, firstDataRow, numDataRows, columns);

            // 4. Bedingte Formatierung für Transaktionstyp-Spalte
            Helpers.setConditionalFormattingForColumn(sheet, columnLetters.transaktionstyp, [
                {value: "Einnahme", background: "#C6EFCE", fontColor: "#006100"},
                {value: "Ausgabe", background: "#FFC7CE", fontColor: "#9C0006"}
            ]);

            // 5. REFERENZEN-MATCHING: Suche nach Referenzen in Einnahmen- und Ausgaben-Sheets
            performBankReferenceMatching(ss, sheet, firstDataRow, numDataRows, lastRow, columns, columnLetters);

            // 6. Endsaldo-Zeile aktualisieren
            updateEndSaldoRow(sheet, lastRow, columns, columnLetters);

            // 7. Spaltenbreiten anpassen
            sheet.autoResizeColumns(1, sheet.getLastColumn());

            return true;
        } catch (e) {
            console.error("Fehler beim Aktualisieren des Bankbewegungen-Sheets:", e);
            return false;
        }
    };

    /**
     * Wendet Validierungen auf das Bankbewegungen-Sheet an
     * @param {Sheet} sheet - Das Sheet
     * @param {number} firstDataRow - Erste Datenzeile
     * @param {number} numDataRows - Anzahl der Datenzeilen
     * @param {Object} columns - Spaltenkonfiguration
     */
    const applyBankSheetValidations = (sheet, firstDataRow, numDataRows, columns) => {
        // Validierung für Transaktionstyp
        Validator.validateDropdown(
            sheet, firstDataRow, columns.transaktionstyp, numDataRows, 1,
            config.bankbewegungen.types
        );

        // Validierung für Kategorie
        Validator.validateDropdown(
            sheet, firstDataRow, columns.kategorie, numDataRows, 1,
            config.bankbewegungen.categories
        );

        // Konten für Dropdown-Validierung sammeln
        const allowedKontoSoll = new Set();
        const allowedGegenkonto = new Set();

        // Konten aus Einnahmen und Ausgaben sammeln
        Object.values(config.einnahmen.kontoMapping).forEach(m => {
            if (m.soll) allowedKontoSoll.add(m.soll);
            if (m.gegen) allowedGegenkonto.add(m.gegen);
        });

        Object.values(config.ausgaben.kontoMapping).forEach(m => {
            if (m.soll) allowedKontoSoll.add(m.soll);
            if (m.gegen) allowedGegenkonto.add(m.gegen);
        });

        // Konten aus Eigenbelegen sammeln
        Object.values(config.eigenbelege.kontoMapping).forEach(m => {
            if (m.soll) allowedKontoSoll.add(m.soll);
            if (m.gegen) allowedGegenkonto.add(m.gegen);
        });

        // Dropdown-Validierungen für Konten setzen
        Validator.validateDropdown(
            sheet, firstDataRow, columns.kontoSoll, numDataRows, 1,
            Array.from(allowedKontoSoll)
        );

        Validator.validateDropdown(
            sheet, firstDataRow, columns.kontoHaben, numDataRows, 1,
            Array.from(allowedGegenkonto)
        );
    };

    /**
     * Aktualisiert die Endsaldo-Zeile im Bankbewegungen-Sheet
     * @param {Sheet} sheet - Das Sheet
     * @param {number} lastRow - Letzte Zeile
     * @param {Object} columns - Spaltenkonfiguration
     * @param {Object} columnLetters - Buchstaben für die Spalten
     */
    const updateEndSaldoRow = (sheet, lastRow, columns, columnLetters) => {
        // Endsaldo-Zeile überprüfen
        const lastRowText = sheet.getRange(lastRow, columns.buchungstext).getDisplayValue().toString().trim().toLowerCase();
        const formattedDate = Utilities.formatDate(
            new Date(),
            Session.getScriptTimeZone(),
            "dd.MM.yyyy"
        );

        if (lastRowText === "endsaldo") {
            // Endsaldo-Zeile aktualisieren
            sheet.getRange(lastRow, columns.datum).setValue(formattedDate);
            sheet.getRange(lastRow, columns.saldo).setFormula(`=${columnLetters.saldo}${lastRow - 1}`);
        } else {
            // Neue Endsaldo-Zeile erstellen
            const newRow = Array(sheet.getLastColumn()).fill("");
            newRow[columns.datum - 1] = formattedDate;
            newRow[columns.buchungstext - 1] = "Endsaldo";
            newRow[columns.saldo - 1] = sheet.getRange(lastRow, columns.saldo).getValue();

            sheet.appendRow(newRow);
        }
    };

    /**
     * Erstellt eine Map aus Referenznummern für schnellere Suche
     * @param {Array} data - Array mit Referenznummern und Beträgen
     * @param {string} sheetType - Optional: Typ des Sheets ("einnahmen", "ausgaben", "eigenbelege")
     * @returns {Object} - Map mit Referenznummern als Keys
     */
    const createReferenceMap = (data, sheetType = null) => {
        // Cache prüfen und nutzen
        if (arguments.length === 1 && typeof data === 'string') {
            if (data === "einnahmen" && _cache.references.einnahmen) {
                return _cache.references.einnahmen;
            }
            if (data === "ausgaben" && _cache.references.ausgaben) {
                return _cache.references.ausgaben;
            }
            if (data === "eigenbelege" && _cache.references.eigenbelege) {
                return _cache.references.eigenbelege;
            }

            // Wenn ein Sheetname übergeben wurde, diesen als sheetType setzen
            sheetType = data;
        }

        const map = {};

        // Keine Daten vorhanden
        if (!data || !data.length) {
            return map;
        }

        // Get column configuration for the sheet type
        const sheetConfig = sheetType ? config[sheetType].columns : null;

        // When working with loaded data range, the first column is always the reference number (rechnungsnummer)
        // Calculate relative positions of other columns based on this
        const rechnungsnummerIdx = 0;

        // Determine indices of required columns within the loaded data range
        let nettoIdx, mwstIdx, bezahltIdx;

        if (sheetConfig) {
            // Calculate relative positions from rechnungsnummer within the loaded range
            nettoIdx = sheetConfig.nettobetrag - sheetConfig.rechnungsnummer;
            mwstIdx = sheetConfig.mwstSatz - sheetConfig.rechnungsnummer;
            bezahltIdx = sheetConfig.bezahlt - sheetConfig.rechnungsnummer;
        } else {
            // Default values if no config provided (for backward compatibility)
            nettoIdx = 3;
            mwstIdx = 4;
            bezahltIdx = 7;
        }

        for (let i = 0; i < data.length; i++) {
            const ref = data[i][rechnungsnummerIdx];
            if (Helpers.isEmpty(ref)) continue;

            // Entferne "G-" Prefix für den Key, falls vorhanden (für Gutschriften)
            let key = ref.trim();
            const isGutschrift = key.startsWith("G-");
            if (isGutschrift) {
                key = key.substring(2);
            }

            const normalizedKey = Helpers.normalizeText(key);

            // Werte extrahieren mit den berechneten Indizes
            let betrag = 0;
            if (!Helpers.isEmpty(data[i][nettoIdx])) {
                betrag = Helpers.parseCurrency(data[i][nettoIdx]);
                betrag = Math.abs(betrag);
            }

            let mwstRate = 0;
            if (!Helpers.isEmpty(data[i][mwstIdx])) {
                mwstRate = Helpers.parseMwstRate(data[i][mwstIdx]);
            }

            let bezahlt = 0;
            if (!Helpers.isEmpty(data[i][bezahltIdx])) {
                bezahlt = Helpers.parseCurrency(data[i][bezahltIdx]);
                bezahlt = Math.abs(bezahlt);
            }

            // Bruttobetrag berechnen
            const brutto = betrag * (1 + mwstRate/100);

            map[key] = {
                ref: ref.trim(),
                normalizedRef: normalizedKey,
                row: i + 2,
                betrag: betrag,
                mwstRate: mwstRate,
                brutto: brutto,
                bezahlt: bezahlt,
                offen: Helpers.round(brutto - bezahlt, 2),
                isGutschrift: isGutschrift
            };

            if (normalizedKey !== key && !map[normalizedKey]) {
                map[normalizedKey] = map[key];
            }
        }

        // Cache result if a sheet type was provided
        if (sheetType) {
            _cache.references[sheetType] = map;
        }

        return map;
    };

    /**
     * Führt das Matching von Bankbewegungen mit Rechnungen durch
     * @param {Spreadsheet} ss - Das Spreadsheet
     * @param {Sheet} sheet - Das Bankbewegungen-Sheet
     * @param {number} firstDataRow - Erste Datenzeile
     * @param {number} numDataRows - Anzahl der Datenzeilen
     * @param {number} lastRow - Letzte Zeile
     * @param {Object} columns - Spaltenkonfiguration
     * @param {Object} columnLetters - Buchstaben für die Spalten
     */
    const performBankReferenceMatching = (ss, sheet, firstDataRow, numDataRows, lastRow, columns, columnLetters) => {
        // Konfigurationen für Spaltenindizes
        const einnahmenCols = config.einnahmen.columns;
        const ausgabenCols = config.ausgaben.columns;
        const eigenbelegeCols = config.eigenbelege.columns;

        // Referenzdaten laden für Einnahmen
        const einnahmenSheet = getSheet("Einnahmen");
        let einnahmenData = [], einnahmenMap = {};

        if (einnahmenSheet && einnahmenSheet.getLastRow() > 1) {
            const numEinnahmenRows = einnahmenSheet.getLastRow() - 1;
            // Die relevanten Spalten laden basierend auf der Konfiguration
            einnahmenData = einnahmenSheet.getRange(2, einnahmenCols.rechnungsnummer, numEinnahmenRows, 8).getDisplayValues();
            einnahmenMap = createReferenceMap(einnahmenData);
        }

        // Referenzdaten laden für Ausgaben
        const ausgabenSheet = getSheet("Ausgaben");
        let ausgabenData = [], ausgabenMap = {};

        if (ausgabenSheet && ausgabenSheet.getLastRow() > 1) {
            const numAusgabenRows = ausgabenSheet.getLastRow() - 1;
            // Die relevanten Spalten laden basierend auf der Konfiguration
            ausgabenData = ausgabenSheet.getRange(2, ausgabenCols.rechnungsnummer, numAusgabenRows, 8).getDisplayValues();
            ausgabenMap = createReferenceMap(ausgabenData);
        }

        // Referenzdaten laden für Eigenbelege
        const eigenbelegeSheet = getSheet("Eigenbelege");
        let eigenbelegeData = [], eigenbelegeMap = {};

        if (eigenbelegeSheet && eigenbelegeSheet.getLastRow() > 1) {
            const numEigenbelegeRows = eigenbelegeSheet.getLastRow() - 1;

            // Calculate the number of columns needed based on config
            const columnsToLoad = eigenbelegeCols.bezahlt - eigenbelegeCols.rechnungsnummer + 1;

            // Load data with the correct column span
            eigenbelegeData = eigenbelegeSheet.getRange(
                2, eigenbelegeCols.rechnungsnummer,
                numEigenbelegeRows,
                columnsToLoad
            ).getDisplayValues();

            // Pass both the data and the sheet type
            eigenbelegeMap = createReferenceMap(eigenbelegeData, "eigenbelege");
        }

        // Bankbewegungen Daten für Verarbeitung holen
        const bankData = sheet.getRange(firstDataRow, 1, numDataRows, columns.matchInfo).getDisplayValues();

        // Ergebnis-Arrays für Batch-Update
        const matchResults = [];
        const kontoSollResults = [];
        const kontoHabenResults = [];

        // Banking-Zuordnungen für spätere Synchronisierung mit Einnahmen/Ausgaben/Eigenbelegen
        const bankZuordnungen = {};

        // Sammeln aller gültigen Konten für die Validierung
        const allowedKontoSoll = new Set();
        const allowedGegenkonto = new Set();

        // Konten aus Einnahmen, Ausgaben und Eigenbelegen sammeln
        Object.values(config.einnahmen.kontoMapping).forEach(m => {
            if (m.soll) allowedKontoSoll.add(m.soll);
            if (m.gegen) allowedGegenkonto.add(m.gegen);
        });

        Object.values(config.ausgaben.kontoMapping).forEach(m => {
            if (m.soll) allowedKontoSoll.add(m.soll);
            if (m.gegen) allowedGegenkonto.add(m.gegen);
        });

        Object.values(config.eigenbelege.kontoMapping).forEach(m => {
            if (m.soll) allowedKontoSoll.add(m.soll);
            if (m.gegen) allowedGegenkonto.add(m.gegen);
        });

        // Fallback-Konto wenn kein Match - das erste Konto aus den erlaubten Konten
        const fallbackKontoSoll = allowedKontoSoll.size > 0 ? Array.from(allowedKontoSoll)[0] : "";
        const fallbackKontoHaben = allowedGegenkonto.size > 0 ? Array.from(allowedGegenkonto)[0] : "";

        // Durchlaufe jede Bankbewegung und suche nach Übereinstimmungen
        for (let i = 0; i < bankData.length; i++) {
            const rowIndex = i + firstDataRow;
            const row = bankData[i];

            // Prüfe, ob es sich um die Endsaldo-Zeile handelt
            const label = row[columns.buchungstext - 1] ? row[columns.buchungstext - 1].toString().trim().toLowerCase() : "";
            if (rowIndex === lastRow && label === "endsaldo") {
                matchResults.push([""]);
                kontoSollResults.push([""]);
                kontoHabenResults.push([""]);
                continue;
            }

            const tranType = row[columns.transaktionstyp - 1]; // Einnahme/Ausgabe
            const refNumber = row[columns.referenz - 1];       // Referenznummer
            let matchInfo = "";

            // Kontonummern basierend auf Kategorie vorbereiten
            let kontoSoll = "", kontoHaben = "";
            const category = row[columns.kategorie - 1] ? row[columns.kategorie - 1].toString().trim() : "";

            // Nur prüfen, wenn Referenz nicht leer ist
            if (!Helpers.isEmpty(refNumber)) {
                let matchFound = false;
                const betragValue = Math.abs(Helpers.parseCurrency(row[columns.betrag - 1]));

                if (tranType === "Einnahme") {
                    // Suche in Einnahmen
                    const matchResult = findMatch(refNumber, einnahmenMap, betragValue);
                    if (matchResult) {
                        matchFound = true;
                        matchInfo = processEinnahmeMatch(matchResult, betragValue, row, columns, einnahmenSheet, einnahmenCols);

                        // Für spätere Markierung merken
                        const key = `einnahme#${matchResult.row}`;
                        bankZuordnungen[key] = {
                            typ: "einnahme",
                            row: matchResult.row,
                            bankDatum: row[columns.datum - 1],
                            matchInfo: matchInfo,
                            transTyp: tranType
                        };
                    }
                } else if (tranType === "Ausgabe") {
                    // Suche in Ausgaben
                    const matchResult = findMatch(refNumber, ausgabenMap, betragValue);
                    if (matchResult) {
                        matchFound = true;
                        matchInfo = processAusgabeMatch(matchResult, betragValue, row, columns, ausgabenSheet, ausgabenCols);

                        // Für spätere Markierung merken
                        const key = `ausgabe#${matchResult.row}`;
                        bankZuordnungen[key] = {
                            typ: "ausgabe",
                            row: matchResult.row,
                            bankDatum: row[columns.datum - 1],
                            matchInfo: matchInfo,
                            transTyp: tranType
                        };
                    }

                    // Wenn keine Übereinstimmung in Ausgaben, dann in Eigenbelegen suchen
                    if (!matchFound) {
                        const eigenbelegMatch = findMatch(refNumber, eigenbelegeMap, betragValue);
                        if (eigenbelegMatch) {
                            matchFound = true;
                            matchInfo = processEigenbelegMatch(eigenbelegMatch, betragValue, row, columns, eigenbelegeSheet, eigenbelegeCols);

                            // Für spätere Markierung merken
                            const key = `eigenbeleg#${eigenbelegMatch.row}`;
                            bankZuordnungen[key] = {
                                typ: "eigenbeleg",
                                row: eigenbelegMatch.row,
                                bankDatum: row[columns.datum - 1],
                                matchInfo: matchInfo,
                                transTyp: tranType
                            };
                        }
                    }

                    // FALLS keine Übereinstimmung, auch in Einnahmen suchen (für Gutschriften)
                    if (!matchFound) {
                        const gutschriftMatch = findMatch(refNumber, einnahmenMap);
                        if (gutschriftMatch) {
                            matchFound = true;
                            matchInfo = processGutschriftMatch(gutschriftMatch, betragValue, row, columns, einnahmenSheet, einnahmenCols);

                            // Für spätere Markierung merken
                            const key = `einnahme#${gutschriftMatch.row}`;
                            bankZuordnungen[key] = {
                                typ: "einnahme",
                                row: gutschriftMatch.row,
                                bankDatum: row[columns.datum - 1],
                                matchInfo: matchInfo,
                                transTyp: "Gutschrift"
                            };
                        }
                    }

                    // FALLS immer noch keine Übereinstimmung, auch in Eigenbelegen suchen (für Erstattungen)
                    if (!matchFound) {
                        const eigenbelegMatch = findMatch(refNumber, eigenbelegeMap, betragValue);
                        if (eigenbelegMatch) {
                            matchFound = true;
                            matchInfo = processEigenbelegMatch(eigenbelegMatch, betragValue, row, columns, eigenbelegeSheet, eigenbelegeCols);

                            // Für spätere Markierung merken
                            const key = `eigenbeleg#${eigenbelegMatch.row}`;
                            bankZuordnungen[key] = {
                                typ: "eigenbeleg",
                                row: eigenbelegMatch.row,
                                bankDatum: row[columns.datum - 1],
                                matchInfo: matchInfo,
                                transTyp: tranType
                            };
                        }
                    }
                }

                // Spezialfälle prüfen (Gesellschaftskonto, Holding)
                if (!matchFound) {
                    const lcRef = refNumber.toString().toLowerCase();
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
            }

            // Kontonummern basierend auf Kategorie setzen
            if (category) {
                // Den richtigen Mapping-Typ basierend auf der Transaktionsart auswählen
                const mappingSource = tranType === "Einnahme" ?
                    config.einnahmen.kontoMapping :
                    config.ausgaben.kontoMapping;

                // Mapping für die Kategorie finden
                if (mappingSource && mappingSource[category]) {
                    const mapping = mappingSource[category];
                    // Nutze die Kontonummern aus dem Mapping
                    kontoSoll = mapping.soll || fallbackKontoSoll;
                    kontoHaben = mapping.gegen || fallbackKontoHaben;
                } else {
                    // Fallback wenn kein Mapping gefunden wurde - erstes zulässiges Konto
                    kontoSoll = fallbackKontoSoll;
                    kontoHaben = fallbackKontoHaben;
                }
            }

            // Ergebnisse für Batch-Update sammeln
            matchResults.push([matchInfo]);
            kontoSollResults.push([kontoSoll]);
            kontoHabenResults.push([kontoHaben]);
        }

        // Batch-Updates ausführen für bessere Performance
        sheet.getRange(firstDataRow, columns.matchInfo, numDataRows, 1).setValues(matchResults);
        sheet.getRange(firstDataRow, columns.kontoSoll, numDataRows, 1).setValues(kontoSollResults);
        sheet.getRange(firstDataRow, columns.kontoHaben, numDataRows, 1).setValues(kontoHabenResults);

        // Formatiere die gesamten Zeilen basierend auf dem Match-Typ
        formatMatchedRows(sheet, firstDataRow, matchResults, columns);

        // Bedingte Formatierung für Match-Spalte mit verbesserten Farben
        setMatchColumnFormatting(sheet, columnLetters.matchInfo);

        // Setze farbliche Markierung in den Einnahmen/Ausgaben/Eigenbelege Sheets basierend auf Zahlungsstatus
        markPaidInvoices(einnahmenSheet, ausgabenSheet, eigenbelegeSheet, bankZuordnungen);
    };

    /**
     * Findet eine Übereinstimmung in der Referenz-Map
     * @param {string} reference - Zu suchende Referenz
     * @param {Object} refMap - Map mit Referenznummern
     * @param {number} betrag - Betrag der Bankbewegung (absoluter Wert)
     * @returns {Object|null} - Gefundene Übereinstimmung oder null
     */
    const findMatch = (reference, refMap, betrag = null) => {
        // Keine Daten
        if (!reference || !refMap) return null;

        // Normalisierte Suche vorbereiten
        const normalizedRef = Helpers.normalizeText(reference);

        // Priorität der Matching-Strategien:
        // 1. Exakte Übereinstimmung
        // 2. Normalisierte Übereinstimmung
        // 3. Enthält-Beziehung (in beide Richtungen)

        // 1. Exakte Übereinstimmung
        if (refMap[reference]) {
            return evaluateMatch(refMap[reference], betrag);
        }

        // 2. Normalisierte Übereinstimmung
        if (normalizedRef && refMap[normalizedRef]) {
            return evaluateMatch(refMap[normalizedRef], betrag);
        }

        // 3. Teilweise Übereinstimmung (mit Cache für Performance)
        const candidateKeys = Object.keys(refMap);

        // Zuerst prüfen wir, ob die Referenz in einem Schlüssel enthalten ist
        for (const key of candidateKeys) {
            if (key.includes(reference) || reference.includes(key)) {
                return evaluateMatch(refMap[key], betrag);
            }
        }

        // Falls keine exakten Treffer, probieren wir es mit normalisierten Werten
        for (const key of candidateKeys) {
            const normalizedKey = Helpers.normalizeText(key);
            if (normalizedKey.includes(normalizedRef) || normalizedRef.includes(normalizedKey)) {
                return evaluateMatch(refMap[key], betrag);
            }
        }

        return null;
    };

    /**
     * Bewertet die Qualität einer Übereinstimmung basierend auf Beträgen
     * @param {Object} match - Die gefundene Übereinstimmung
     * @param {number} betrag - Der Betrag zum Vergleich
     * @returns {Object} - Übereinstimmung mit zusätzlichen Infos
     */
    const evaluateMatch = (match, betrag = null) => {
        if (!match) return null;

        // Behalte die ursprüngliche Übereinstimmung
        const result = { ...match };

        // Wenn kein Betrag zum Vergleich angegeben wurde
        if (betrag === null) {
            result.matchType = "Referenz-Match";
            return result;
        }

        // Hole die Beträge aus der Übereinstimmung
        const matchBrutto = Math.abs(match.brutto);
        const matchBezahlt = Math.abs(match.bezahlt);

        // Toleranzwert für Betragsabweichungen (2 Cent)
        const tolerance = 0.02;

        // Fall 1: Betrag entspricht genau dem Bruttobetrag (Vollständige Zahlung)
        if (Helpers.isApproximatelyEqual(betrag, matchBrutto, tolerance)) {
            result.matchType = "Vollständige Zahlung";
            return result;
        }

        // Fall 2: Position ist bereits vollständig bezahlt
        if (Helpers.isApproximatelyEqual(matchBezahlt, matchBrutto, tolerance) && matchBezahlt > 0) {
            result.matchType = "Vollständige Zahlung";
            return result;
        }

        // Fall 3: Teilzahlung (Bankbetrag kleiner als Rechnungsbetrag)
        // Nur als Teilzahlung markieren, wenn der Betrag deutlich kleiner ist (> 10% Differenz)
        if (betrag < matchBrutto && (matchBrutto - betrag) > (matchBrutto * 0.1)) {
            result.matchType = "Teilzahlung";
            return result;
        }

        // Fall 4: Betrag ist größer als Bruttobetrag, aber vermutlich trotzdem die richtige Zahlung
        // (z.B. wegen Rundungen oder kleinen Gebühren)
        if (betrag > matchBrutto && Helpers.isApproximatelyEqual(betrag, matchBrutto, tolerance)) {
            result.matchType = "Vollständige Zahlung";
            return result;
        }

        // Fall 5: Bei allen anderen Fällen (Beträge weichen stärker ab)
        result.matchType = "Unsichere Zahlung";
        result.betragsDifferenz = Helpers.round(Math.abs(betrag - matchBrutto), 2);
        return result;
    };

    /**
     * Verarbeitet eine Einnahmen-Übereinstimmung
     * @returns {string} Formatierte Match-Information
     */
    const processEinnahmeMatch = (matchResult, betragValue, row, columns, einnahmenSheet, einnahmenCols) => {
        let matchInfo = "";
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
            einnahmenSheet) {
            // Bankbewegungsdatum holen (aus Spalte A)
            const zahlungsDatum = row[columns.datum - 1];
            if (zahlungsDatum) {
                // Zahldatum im Einnahmen-Sheet aktualisieren (nur wenn leer)
                const einnahmenRow = matchResult.row;
                const zahldatumRange = einnahmenSheet.getRange(einnahmenRow, einnahmenCols.zahlungsdatum);
                const aktuellDatum = zahldatumRange.getValue();

                if (Helpers.isEmpty(aktuellDatum)) {
                    zahldatumRange.setValue(zahlungsDatum);
                    matchStatus += " ✓ Datum aktualisiert";
                }
            }
        }

        return `Einnahme #${matchResult.row}${matchStatus}`;
    };

    /**
     * Verarbeitet eine Ausgaben-Übereinstimmung
     * @returns {string} Formatierte Match-Information
     */
    const processAusgabeMatch = (matchResult, betragValue, row, columns, ausgabenSheet, ausgabenCols) => {
        let matchInfo = "";
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
            const zahlungsDatum = row[columns.datum - 1];
            if (zahlungsDatum) {
                // Zahldatum im Ausgaben-Sheet aktualisieren (nur wenn leer)
                const ausgabenRow = matchResult.row;
                const zahldatumRange = ausgabenSheet.getRange(ausgabenRow, ausgabenCols.zahlungsdatum);
                const aktuellDatum = zahldatumRange.getValue();

                if (Helpers.isEmpty(aktuellDatum)) {
                    zahldatumRange.setValue(zahlungsDatum);
                    matchStatus += " ✓ Datum aktualisiert";
                }
            }
        }

        return `Ausgabe #${matchResult.row}${matchStatus}`;
    };

    /**
     * Verarbeitet eine Gutschrift-Übereinstimmung
     * @returns {string} Formatierte Match-Information
     */
    const processGutschriftMatch = (gutschriftMatch, betragValue, row, columns, einnahmenSheet, einnahmenCols) => {
        let matchStatus = "";

        // Bei Gutschriften könnte der Betrag abweichen (z.B. Teilgutschrift)
        // Prüfen, ob die Beträge plausibel sind
        const gutschriftBetrag = Math.abs(gutschriftMatch.brutto);

        if (Helpers.isApproximatelyEqual(betragValue, gutschriftBetrag, 0.01)) {
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
            const gutschriftDatum = row[columns.datum - 1];
            if (gutschriftDatum) {
                // Gutschriftdatum im Einnahmen-Sheet aktualisieren und "G-" vor die Referenz setzen
                const einnahmenRow = gutschriftMatch.row;

                // Zahldatum aktualisieren (nur wenn leer)
                const zahldatumRange = einnahmenSheet.getRange(einnahmenRow, einnahmenCols.zahlungsdatum);
                const aktuellDatum = zahldatumRange.getValue();

                if (Helpers.isEmpty(aktuellDatum)) {
                    zahldatumRange.setValue(gutschriftDatum);
                    matchStatus += " ✓ Datum aktualisiert";
                }

                // Optional: Die Referenz mit "G-" kennzeichnen, um Gutschrift zu markieren
                const refRange = einnahmenSheet.getRange(einnahmenRow, einnahmenCols.rechnungsnummer);
                const currentRef = refRange.getValue();
                if (currentRef && !currentRef.toString().startsWith("G-")) {
                    refRange.setValue("G-" + currentRef);
                    matchStatus += " ✓ Ref. markiert";
                }
            }
        }

        return `Gutschrift zu Einnahme #${gutschriftMatch.row}${matchStatus}`;
    };

    /**
     * Verarbeitet eine Eigenbeleg-Übereinstimmung
     * @returns {string} Formatierte Match-Information
     */
    const processEigenbelegMatch = (eigenbelegMatch, betragValue, row, columns, eigenbelegeSheet, eigenbelegeCols) => {
        let matchInfo = "";
        let matchStatus = "";

        // Je nach Match-Typ unterschiedliche Statusinformationen
        if (eigenbelegMatch.matchType) {
            // Bei "Unsichere Zahlung" auch die Differenz anzeigen
            if (eigenbelegMatch.matchType === "Unsichere Zahlung" && eigenbelegMatch.betragsDifferenz) {
                matchStatus = ` (${eigenbelegMatch.matchType}, Diff: ${eigenbelegMatch.betragsDifferenz}€)`;
            } else {
                matchStatus = ` (${eigenbelegMatch.matchType})`;
            }
        }

        // Prüfen, ob Erstattungsdatum in Eigenbelegen aktualisiert werden soll
        if ((eigenbelegMatch.matchType === "Vollständige Zahlung" ||
                eigenbelegMatch.matchType === "Teilzahlung") &&
            eigenbelegeSheet) {
            // Bankbewegungsdatum holen (aus Spalte A)
            const erstattungsDatum = row[columns.datum - 1];
            if (erstattungsDatum) {
                // Erstattungsdatum im Eigenbelege-Sheet aktualisieren (nur wenn leer)
                const eigenbelegeRow = eigenbelegMatch.row;
                const erstattungsdatumRange = eigenbelegeSheet.getRange(eigenbelegeRow, eigenbelegeCols.zahlungsdatum);
                const aktuellDatum = erstattungsdatumRange.getValue();

                if (Helpers.isEmpty(aktuellDatum)) {
                    erstattungsdatumRange.setValue(erstattungsDatum);
                    matchStatus += " ✓ Datum aktualisiert";
                }
            }
        }

        return `Eigenbeleg #${eigenbelegMatch.row}${matchStatus}`;
    };

    /**
     * Formatiert Zeilen basierend auf dem Match-Typ
     * @param {Sheet} sheet - Das Sheet
     * @param {number} firstDataRow - Erste Datenzeile
     * @param {Array} matchResults - Array mit Match-Infos
     * @param {Object} columns - Spaltenkonfiguration
     */
    const formatMatchedRows = (sheet, firstDataRow, matchResults, columns) => {
        // Performance-optimiertes Batch-Update vorbereiten
        const formatBatches = {
            'Einnahme': { rows: [], color: "#EAF1DD" },  // Helles Grün (Grundfarbe für Einnahmen)
            'Vollständige Zahlung (Einnahme)': { rows: [], color: "#C6EFCE" },  // Kräftiges Grün
            'Teilzahlung (Einnahme)': { rows: [], color: "#FCE4D6" },  // Helles Orange
            'Ausgabe': { rows: [], color: "#FFCCCC" },   // Helles Rosa (Grundfarbe für Ausgaben)
            'Vollständige Zahlung (Ausgabe)': { rows: [], color: "#FFC7CE" },  // Helles Rot
            'Teilzahlung (Ausgabe)': { rows: [], color: "#FCE4D6" },  // Helles Orange
            'Eigenbeleg': { rows: [], color: "#DDEBF7" },  // Helles Blau (Grundfarbe für Eigenbelege)
            'Vollständige Zahlung (Eigenbeleg)': { rows: [], color: "#9BC2E6" },  // Kräftigeres Blau
            'Teilzahlung (Eigenbeleg)': { rows: [], color: "#FCE4D6" },  // Helles Orange
            'Gutschrift': { rows: [], color: "#E6E0FF" },  // Helles Lila
            'Gesellschaftskonto/Holding': { rows: [], color: "#FFEB9C" }  // Helles Gelb
        };

        // Zeilen nach Kategorien gruppieren
        matchResults.forEach((matchInfo, index) => {
            const rowIndex = firstDataRow + index;
            const matchText = (matchInfo && matchInfo[0]) ? matchInfo[0].toString() : "";

            if (!matchText) return; // Überspringe leere Matches

            if (matchText.includes("Einnahme")) {
                if (matchText.includes("Vollständige Zahlung")) {
                    formatBatches['Vollständige Zahlung (Einnahme)'].rows.push(rowIndex);
                } else if (matchText.includes("Teilzahlung")) {
                    formatBatches['Teilzahlung (Einnahme)'].rows.push(rowIndex);
                } else {
                    formatBatches['Einnahme'].rows.push(rowIndex);
                }
            } else if (matchText.includes("Ausgabe")) {
                if (matchText.includes("Vollständige Zahlung")) {
                    formatBatches['Vollständige Zahlung (Ausgabe)'].rows.push(rowIndex);
                } else if (matchText.includes("Teilzahlung")) {
                    formatBatches['Teilzahlung (Ausgabe)'].rows.push(rowIndex);
                } else {
                    formatBatches['Ausgabe'].rows.push(rowIndex);
                }
            } else if (matchText.includes("Eigenbeleg")) {
                if (matchText.includes("Vollständige Zahlung")) {
                    formatBatches['Vollständige Zahlung (Eigenbeleg)'].rows.push(rowIndex);
                } else if (matchText.includes("Teilzahlung")) {
                    formatBatches['Teilzahlung (Eigenbeleg)'].rows.push(rowIndex);
                } else {
                    formatBatches['Eigenbeleg'].rows.push(rowIndex);
                }
            } else if (matchText.includes("Gutschrift")) {
                formatBatches['Gutschrift'].rows.push(rowIndex);
            } else if (matchText.includes("Gesellschaftskonto") || matchText.includes("Holding")) {
                formatBatches['Gesellschaftskonto/Holding'].rows.push(rowIndex);
            }
        });

        // Batch-Formatting anwenden
        Object.values(formatBatches).forEach(batch => {
            if (batch.rows.length > 0) {
                // Gruppen von maximal 20 Zeilen formatieren (API-Limits vermeiden)
                const chunkSize = 20;
                for (let i = 0; i < batch.rows.length; i += chunkSize) {
                    const chunk = batch.rows.slice(i, i + chunkSize);
                    chunk.forEach(rowIndex => {
                        try {
                            sheet.getRange(rowIndex, 1, 1, columns.matchInfo)
                                .setBackground(batch.color);
                        } catch (e) {
                            console.error(`Fehler beim Formatieren von Zeile ${rowIndex}:`, e);
                        }
                    });

                    // Kurze Pause um API-Limits zu vermeiden
                    if (i + chunkSize < batch.rows.length) {
                        Utilities.sleep(50);
                    }
                }
            }
        });
    };

    /**
     * Setzt bedingte Formatierung für die Match-Spalte
     * @param {Sheet} sheet - Das Sheet
     * @param {string} columnLetter - Buchstabe für die Spalte
     */
    const setMatchColumnFormatting = (sheet, columnLetter) => {
        Helpers.setConditionalFormattingForColumn(sheet, columnLetter, [
            // Grundlegende Match-Typen mit beginsWith Pattern
            {value: "Einnahme", background: "#C6EFCE", fontColor: "#006100", pattern: "beginsWith"},
            {value: "Ausgabe", background: "#FFC7CE", fontColor: "#9C0006", pattern: "beginsWith"},
            {value: "Eigenbeleg", background: "#DDEBF7", fontColor: "#2F5597", pattern: "beginsWith"},
            {value: "Gutschrift", background: "#E6E0FF", fontColor: "#5229A3", pattern: "beginsWith"},
            {value: "Gesellschaftskonto", background: "#FFEB9C", fontColor: "#9C6500", pattern: "beginsWith"},
            {value: "Holding", background: "#FFEB9C", fontColor: "#9C6500", pattern: "beginsWith"},

            // Zusätzliche Betragstypen mit contains Pattern
            {value: "Vollständige Zahlung", background: "#C6EFCE", fontColor: "#006100", pattern: "contains"},
            {value: "Teilzahlung", background: "#FCE4D6", fontColor: "#974706", pattern: "contains"},
            {value: "Unsichere Zahlung", background: "#F8CBAD", fontColor: "#843C0C", pattern: "contains"},
            {value: "Vollständige Gutschrift", background: "#E6E0FF", fontColor: "#5229A3", pattern: "contains"},
            {value: "Teilgutschrift", background: "#E6E0FF", fontColor: "#5229A3", pattern: "contains"},
            {value: "Ungewöhnliche Gutschrift", background: "#FFD966", fontColor: "#7F6000", pattern: "contains"}
        ]);
    };

    /**
     * Markiert bezahlte Einnahmen, Ausgaben und Eigenbelege farblich basierend auf dem Zahlungsstatus
     * @param {Sheet} einnahmenSheet - Das Einnahmen-Sheet
     * @param {Sheet} ausgabenSheet - Das Ausgaben-Sheet
     * @param {Sheet} eigenbelegeSheet - Das Eigenbelege-Sheet
     * @param {Object} bankZuordnungen - Zuordnungen aus dem Bankbewegungen-Sheet
     */
    const markPaidInvoices = (einnahmenSheet, ausgabenSheet, eigenbelegeSheet, bankZuordnungen) => {
        // Markiere bezahlte Einnahmen
        if (einnahmenSheet && einnahmenSheet.getLastRow() > 1) {
            markPaidRows(einnahmenSheet, "Einnahmen", bankZuordnungen);
        }

        // Markiere bezahlte Ausgaben
        if (ausgabenSheet && ausgabenSheet.getLastRow() > 1) {
            markPaidRows(ausgabenSheet, "Ausgaben", bankZuordnungen);
        }

        // Markiere bezahlte Eigenbelege
        if (eigenbelegeSheet && eigenbelegeSheet.getLastRow() > 1) {
            markPaidRows(eigenbelegeSheet, "Eigenbelege", bankZuordnungen);
        }
    };

    /**
     * Markiert bezahlte Zeilen in einem Sheet
     * @param {Sheet} sheet - Das zu aktualisierende Sheet
     * @param {string} sheetType - Typ des Sheets ("Einnahmen", "Ausgaben" oder "Eigenbelege")
     * @param {Object} bankZuordnungen - Zuordnungen aus dem Bankbewegungen-Sheet
     */
    const markPaidRows = (sheet, sheetType, bankZuordnungen) => {
        // Konfiguration für das Sheet
        let columns;

        if (sheetType === "Einnahmen") {
            columns = config.einnahmen.columns;
        } else if (sheetType === "Ausgaben") {
            columns = config.ausgaben.columns;
        } else if (sheetType === "Eigenbelege") {
            columns = config.eigenbelege.columns;
        } else {
            return; // Unbekannter Sheet-Typ
        }

        // Hole Werte aus dem Sheet
        const numRows = sheet.getLastRow() - 1;
        const data = sheet.getRange(2, 1, numRows, columns.zahlungsdatum).getValues();

        // Batch-Arrays für die verschiedenen Status
        const fullPaidWithBankRows = [];
        const fullPaidRows = [];
        const partialPaidWithBankRows = [];
        const partialPaidRows = [];
        const gutschriftRows = [];
        const normalRows = [];

        // Bank-Abgleich-Updates sammeln
        const bankabgleichUpdates = [];

        // Datenzeilen durchgehen und in Kategorien einteilen
        for (let i = 0; i < data.length; i++) {
            const row = i + 2; // Aktuelle Zeile im Sheet
            const nettobetrag = Helpers.parseCurrency(data[i][columns.nettobetrag - 1]);
            const bezahltBetrag = Helpers.parseCurrency(data[i][columns.bezahlt - 1]);
            const zahlungsDatum = data[i][columns.zahlungsdatum - 1];
            const referenz = data[i][columns.rechnungsnummer - 1];

            // Prüfe, ob es eine Gutschrift ist
            const isGutschrift = referenz && referenz.toString().startsWith("G-");

            // Prüfe, ob diese Position im Banking-Sheet zugeordnet wurde
            // Bei Eigenbelegen verwenden wir den key "eigenbeleg#row"
            const zuordnungsKey = `${sheetType === "Eigenbelege" ? "eigenbeleg" : sheetType.toLowerCase().slice(0, -1)}#${row}`;
            const hatBankzuordnung = bankZuordnungen[zuordnungsKey] !== undefined;

            // Zahlungsstatus berechnen
            const mwst = Helpers.parseMwstRate(data[i][columns.mwstSatz - 1]) / 100;
            const bruttoBetrag = nettobetrag * (1 + mwst);
            const isPaid = Math.abs(bezahltBetrag) >= Math.abs(bruttoBetrag) * 0.999; // 99.9% bezahlt wegen Rundungsfehlern
            const isPartialPaid = !isPaid && bezahltBetrag > 0;

            // Kategorie bestimmen
            if (isGutschrift) {
                gutschriftRows.push(row);
                // Bank-Abgleich-Info setzen
                if (hatBankzuordnung) {
                    bankabgleichUpdates.push({
                        row,
                        value: getZuordnungsInfo(bankZuordnungen[zuordnungsKey])
                    });
                }
            } else if (isPaid) {
                if (zahlungsDatum) {
                    if (hatBankzuordnung) {
                        fullPaidWithBankRows.push(row);
                        bankabgleichUpdates.push({
                            row,
                            value: getZuordnungsInfo(bankZuordnungen[zuordnungsKey])
                        });
                    } else {
                        fullPaidRows.push(row);
                    }
                } else {
                    // Bezahlt aber kein Datum
                    partialPaidRows.push(row);
                }
            } else if (isPartialPaid) {
                if (hatBankzuordnung) {
                    partialPaidWithBankRows.push(row);
                    bankabgleichUpdates.push({
                        row,
                        value: getZuordnungsInfo(bankZuordnungen[zuordnungsKey])
                    });
                } else {
                    partialPaidRows.push(row);
                }
            } else {
                // Unbezahlt - normale Zeile
                normalRows.push(row);
                // Vorhandene Bank-Abgleich-Info entfernen falls vorhanden
                if (sheet.getRange(row, columns.bankabgleich).getValue().toString().startsWith("✓ Bank:")) {
                    bankabgleichUpdates.push({
                        row,
                        value: ""
                    });
                }
            }
        }

        // Färbe Zeilen im Batch-Modus (für bessere Performance)
        applyColorToRows(sheet, fullPaidWithBankRows, "#C6EFCE"); // Kräftigeres Grün
        applyColorToRows(sheet, fullPaidRows, "#EAF1DD"); // Helles Grün
        applyColorToRows(sheet, partialPaidWithBankRows, "#FFC7AA"); // Kräftigeres Orange
        applyColorToRows(sheet, partialPaidRows, "#FCE4D6"); // Helles Orange
        applyColorToRows(sheet, gutschriftRows, "#E6E0FF"); // Helles Lila
        applyColorToRows(sheet, normalRows, null); // Keine Farbe / Zurücksetzen

        // Bank-Abgleich-Updates in Batches ausführen
        if (bankabgleichUpdates.length > 0) {
            // Gruppiere Updates nach Wert für effizientere Batch-Updates
            const groupedUpdates = {};

            bankabgleichUpdates.forEach(update => {
                const { row, value } = update;
                if (!groupedUpdates[value]) {
                    groupedUpdates[value] = [];
                }
                groupedUpdates[value].push(row);
            });

            // Führe Batch-Updates für jede Gruppe aus
            Object.entries(groupedUpdates).forEach(([value, rows]) => {
                // Verarbeite in Batches von maximal 20 Zeilen
                const batchSize = 20;
                for (let i = 0; i < rows.length; i += batchSize) {
                    const batchRows = rows.slice(i, i + batchSize);
                    batchRows.forEach(row => {
                        sheet.getRange(row, columns.bankabgleich).setValue(value);
                    });

                    // Kurze Pause, um API-Limits zu vermeiden
                    if (i + batchSize < rows.length) {
                        Utilities.sleep(50);
                    }
                }
            });
        }
    };

    /**
     * Wendet eine Farbe auf mehrere Zeilen an
     * @param {Sheet} sheet - Das Sheet
     * @param {Array} rows - Die zu färbenden Zeilennummern
     * @param {string|null} color - Die Hintergrundfarbe oder null zum Zurücksetzen
     */
    const applyColorToRows = (sheet, rows, color) => {
        if (!rows || rows.length === 0) return;

        // Verarbeite in Batches von maximal 20 Zeilen
        const batchSize = 20;
        for (let i = 0; i < rows.length; i += batchSize) {
            const batchRows = rows.slice(i, i + batchSize);
            batchRows.forEach(row => {
                try {
                    const range = sheet.getRange(row, 1, 1, sheet.getLastColumn());
                    if (color) {
                        range.setBackground(color);
                    } else {
                        range.setBackground(null);
                    }
                } catch (e) {
                    console.error(`Fehler beim Formatieren von Zeile ${row}:`, e);
                }
            });

            // Kurze Pause um API-Limits zu vermeiden
            if (i + batchSize < rows.length) {
                Utilities.sleep(50);
            }
        }
    };

    /**
     * Erstellt einen Informationstext für eine Bank-Zuordnung
     * @param {Object} zuordnung - Die Zuordnungsinformation
     * @returns {string} - Formatierter Informationstext
     */
    const getZuordnungsInfo = (zuordnung) => {
        if (!zuordnung) return "";

        let infoText = "✓ Bank: " + zuordnung.bankDatum;

        // Wenn es mehrere Zuordnungen gibt (z.B. bei aufgeteilten Zahlungen)
        if (zuordnung.additional && zuordnung.additional.length > 0) {
            infoText += " + " + zuordnung.additional.length + " weitere";
        }

        return infoText;
    };

    /**
     * Aktualisiert das aktive Tabellenblatt
     */
    const refreshActiveSheet = () => {
        try {
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            const sheet = ss.getActiveSheet();
            const name = sheet.getName();
            const ui = SpreadsheetApp.getUi();

            // Cache zurücksetzen
            clearCache();

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

            // Cache zurücksetzen
            clearCache();

            // Sheets in der richtigen Reihenfolge aktualisieren, um Abhängigkeiten zu berücksichtigen
            const refreshOrder = ["Einnahmen", "Ausgaben", "Eigenbelege", "Bankbewegungen"];

            for (const name of refreshOrder) {
                const sheet = getSheet(name);
                if (sheet) {
                    name === "Bankbewegungen" ? refreshBankSheet(sheet) : refreshDataSheet(sheet);
                    // Kurze Pause einfügen, um API-Limits zu vermeiden
                    Utilities.sleep(100);
                }
            }
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