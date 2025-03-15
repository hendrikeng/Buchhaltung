// modules/refreshModule/matchingHandler.js
import stringUtils from '../../utils/stringUtils.js';
import numberUtils from '../../utils/numberUtils.js';
import globalCache from '../../utils/cacheUtils.js';

/**
 * Führt das Matching von Bankbewegungen mit Rechnungen durch
 * @param {Spreadsheet} ss - Das Spreadsheet
 * @param {Sheet} sheet - Das Bankbewegungen-Sheet
 * @param {number} firstDataRow - Erste Datenzeile
 * @param {number} numDataRows - Anzahl der Datenzeilen
 * @param {number} lastRow - Letzte Zeile
 * @param {Object} columns - Spaltenkonfiguration
 * @param {Object} columnLetters - Buchstaben für die Spalten
 * @param {Object} config - Die Konfiguration
 */
function performBankReferenceMatching(ss, sheet, firstDataRow, numDataRows, lastRow, columns, columnLetters, config) {
    // Konfigurationen für Spaltenindizes
    const einnahmenCols = config.einnahmen.columns;
    const ausgabenCols = config.ausgaben.columns;
    const eigenbelegeCols = config.eigenbelege.columns;
    const gesellschafterCols = config.gesellschafterkonto.columns;
    const holdingCols = config.holdingTransfers.columns;

    // Referenzdaten laden für alle relevanten Sheets
    const sheetData = loadReferenceData(ss, {
        einnahmen: einnahmenCols,
        ausgaben: ausgabenCols,
        eigenbelege: eigenbelegeCols,
        gesellschafterkonto: gesellschafterCols,
        holdingTransfers: holdingCols
    });

    // Bankbewegungen Daten für Verarbeitung holen
    const bankData = sheet.getRange(firstDataRow, 1, numDataRows, columns.matchInfo).getDisplayValues();

    // Ergebnis-Arrays für Batch-Update
    const matchResults = [];
    const kontoSollResults = [];
    const kontoHabenResults = [];

    // Banking-Zuordnungen für spätere Synchronisierung mit anderen Sheets
    const bankZuordnungen = {};

    // Sammeln aller gültigen Konten für die Validierung
    const allowedKontoSoll = new Set();
    const allowedGegenkonto = new Set();

    // Konten aus allen relevanten Mappings sammeln
    [
        config.einnahmen.kontoMapping,
        config.ausgaben.kontoMapping,
        config.eigenbelege.kontoMapping,
        config.gesellschafterkonto.kontoMapping,
        config.holdingTransfers.kontoMapping
    ].forEach(mapping => {
        Object.values(mapping).forEach(m => {
            if (m.soll) allowedKontoSoll.add(m.soll);
            if (m.gegen) allowedGegenkonto.add(m.gegen);
        });
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

        // Matching-Logik für Bankbewegungen
        if (!stringUtils.isEmpty(refNumber)) {
            let matchFound = false;
            const betragValue = Math.abs(numberUtils.parseCurrency(row[columns.betrag - 1]));

            // Je nach Transaktionstyp in unterschiedlichen Sheets suchen
            if (tranType === "Einnahme") {
                // Einnahmen prüfen
                processEinnahmeMatching(
                    refNumber, betragValue, row, columns, sheetData.einnahmen,
                    einnahmenCols, matchResults, kontoSollResults, kontoHabenResults,
                    bankZuordnungen, matchFound, config
                );
            } else if (tranType === "Ausgabe") {
                // Ausgaben, Eigenbelege, Gesellschafterkonto und Holding Transfers prüfen
                processAusgabeMatching(
                    refNumber, betragValue, row, columns, {
                        ausgaben: sheetData.ausgaben,
                        eigenbelege: sheetData.eigenbelege,
                        gesellschafterkonto: sheetData.gesellschafterkonto,
                        holdingTransfers: sheetData.holdingTransfers
                    },
                    {
                        ausgabenCols, eigenbelegeCols, gesellschafterCols, holdingCols
                    },
                    matchResults, kontoSollResults, kontoHabenResults,
                    bankZuordnungen, matchFound, config
                );
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

        // Sicherstellen, dass die Arrays die richtigen Werte haben
        if (matchResults.length <= i) {
            matchResults.push([matchInfo]);
        }

        if (kontoSollResults.length <= i) {
            kontoSollResults.push([kontoSoll]);
        }

        if (kontoHabenResults.length <= i) {
            kontoHabenResults.push([kontoHaben]);
        }
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
    markPaidInvoices(ss, bankZuordnungen, config);
}

/**
 * Lädt Referenzdaten aus allen relevanten Sheets
 * @param {Spreadsheet} ss - Das Spreadsheet
 * @param {Object} columns - Spaltenkonfigurationen für alle Sheets
 * @returns {Object} - Referenzdaten für alle Sheets
 */
function loadReferenceData(ss, columns) {
    const result = {};

    // Sheets und ihre Konfigurationen
    const sheets = {
        einnahmen: { name: "Einnahmen", cols: columns.einnahmen },
        ausgaben: { name: "Ausgaben", cols: columns.ausgaben },
        eigenbelege: { name: "Eigenbelege", cols: columns.eigenbelege },
        gesellschafterkonto: { name: "Gesellschafterkonto", cols: columns.gesellschafterkonto },
        holdingTransfers: { name: "Holding Transfers", cols: columns.holdingTransfers }
    };

    // Für jedes Sheet Referenzdaten laden
    Object.entries(sheets).forEach(([key, config]) => {
        const sheet = ss.getSheetByName(config.name);

        if (sheet && sheet.getLastRow() > 1) {
            result[key] = loadSheetReferenceData(sheet, config.cols, key);
        } else {
            result[key] = {};
        }
    });

    return result;
}

/**
 * Lädt Referenzdaten aus einem einzelnen Sheet
 * @param {Sheet} sheet - Das Sheet
 * @param {Object} columns - Spaltenkonfiguration
 * @param {string} sheetType - Typ des Sheets
 * @returns {Object} - Referenzdaten für das Sheet
 */
function loadSheetReferenceData(sheet, columns, sheetType) {
    // Cache prüfen
    if (globalCache.has('references', sheetType)) {
        return globalCache.get('references', sheetType);
    }

    const map = {};
    const numRows = sheet.getLastRow() - 1;

    if (numRows <= 0) return map;

    // Die benötigten Spalten bestimmen
    const columnsToLoad = Math.max(
        columns.rechnungsnummer,
        columns.nettobetrag,
        columns.mwstSatz || 0,
        columns.bezahlt || 0
    );

    // Die relevanten Spalten laden
    const data = sheet.getRange(2, 1, numRows, columnsToLoad).getValues();

    for (let i = 0; i < data.length; i++) {
        const row = data[i];
        const ref = row[columns.rechnungsnummer - 1];

        if (stringUtils.isEmpty(ref)) continue;

        // Entferne "G-" Prefix für den Key, falls vorhanden (für Gutschriften)
        let key = ref.toString().trim();
        const isGutschrift = key.startsWith("G-");

        if (isGutschrift) {
            key = key.substring(2);
        }

        const normalizedKey = stringUtils.normalizeText(key);

        // Werte extrahieren
        let betrag = 0;
        if (!stringUtils.isEmpty(row[columns.nettobetrag - 1])) {
            betrag = numberUtils.parseCurrency(row[columns.nettobetrag - 1]);
            betrag = Math.abs(betrag);
        }

        let mwstRate = 0;
        if (columns.mwstSatz && !stringUtils.isEmpty(row[columns.mwstSatz - 1])) {
            mwstRate = numberUtils.parseMwstRate(row[columns.mwstSatz - 1]);
        }

        let bezahlt = 0;
        if (columns.bezahlt && !stringUtils.isEmpty(row[columns.bezahlt - 1])) {
            bezahlt = numberUtils.parseCurrency(row[columns.bezahlt - 1]);
            bezahlt = Math.abs(bezahlt);
        }

        // Bruttobetrag berechnen
        const brutto = betrag * (1 + mwstRate/100);

        map[key] = {
            ref: ref.toString().trim(),
            normalizedRef: normalizedKey,
            row: i + 2, // 1-basiert + Überschrift
            betrag: betrag,
            mwstRate: mwstRate,
            brutto: brutto,
            bezahlt: bezahlt,
            offen: numberUtils.round(brutto - bezahlt, 2),
            isGutschrift: isGutschrift
        };

        // Normalisierter Key als Alternative
        if (normalizedKey !== key && !map[normalizedKey]) {
            map[normalizedKey] = map[key];
        }
    }

    // Referenzdaten cachen
    globalCache.set('references', sheetType, map);

    return map;
}

/**
 * Verarbeitet Matching für Einnahmen
 * @param {string} refNumber - Referenznummer
 * @param {number} betragValue - Betrag
 * @param {Array} row - Zeile aus dem Bankbewegungen-Sheet
 * @param {Object} columns - Spaltenkonfiguration
 * @param {Object} einnahmenData - Referenzdaten für Einnahmen
 * @param {Object} einnahmenCols - Spaltenkonfiguration für Einnahmen
 * @param {Array} matchResults - Array für Match-Ergebnisse
 * @param {Array} kontoSollResults - Array für Konto-Soll-Ergebnisse
 * @param {Array} kontoHabenResults - Array für Konto-Haben-Ergebnisse
 * @param {Object} bankZuordnungen - Speicher für Bankzuordnungen
 * @param {boolean} matchFound - Flag für gefundene Übereinstimmung
 * @param {Object} config - Die Konfiguration
 */
function processEinnahmeMatching(refNumber, betragValue, row, columns, einnahmenData, einnahmenCols, matchResults, kontoSollResults, kontoHabenResults, bankZuordnungen, matchFound, config) {
    // Einnahmen-Match prüfen
    const matchResult = findMatch(refNumber, einnahmenData, betragValue);

    if (matchResult) {
        matchFound = true;
        const matchInfo = processEinnahmeMatch(matchResult, betragValue, row, columns, einnahmenCols);
        matchResults.push([matchInfo]);

        // Konten aus dem Mapping setzen
        const category = row[columns.kategorie - 1] ? row[columns.kategorie - 1].toString().trim() : "";
        let kontoSoll = "", kontoHaben = "";

        if (category && config.einnahmen.kontoMapping[category]) {
            const mapping = config.einnahmen.kontoMapping[category];
            kontoSoll = mapping.soll || "";
            kontoHaben = mapping.gegen || "";
        }

        kontoSollResults.push([kontoSoll]);
        kontoHabenResults.push([kontoHaben]);

        // Für spätere Markierung merken
        const key = `einnahme#${matchResult.row}`;
        bankZuordnungen[key] = {
            typ: "einnahme",
            row: matchResult.row,
            bankDatum: row[columns.datum - 1],
            matchInfo: matchInfo,
            transTyp: "Einnahme"
        };
    }
}

/**
 * Verarbeitet Matching für Ausgaben und andere Belege
 * @param {string} refNumber - Referenznummer
 * @param {number} betragValue - Betrag
 * @param {Array} row - Zeile aus dem Bankbewegungen-Sheet
 * @param {Object} columns - Spaltenkonfiguration
 * @param {Object} sheetData - Referenzdaten für verschiedene Sheets
 * @param {Object} sheetCols - Spaltenkonfigurationen für verschiedene Sheets
 * @param {Array} matchResults - Array für Match-Ergebnisse
 * @param {Array} kontoSollResults - Array für Konto-Soll-Ergebnisse
 * @param {Array} kontoHabenResults - Array für Konto-Haben-Ergebnisse
 * @param {Object} bankZuordnungen - Speicher für Bankzuordnungen
 * @param {boolean} matchFound - Flag für gefundene Übereinstimmung
 * @param {Object} config - Die Konfiguration
 */
function processAusgabeMatching(refNumber, betragValue, row, columns, sheetData, sheetCols, matchResults, kontoSollResults, kontoHabenResults, bankZuordnungen, matchFound, config) {
    // Ausgaben-Match prüfen
    const ausgabenMatch = findMatch(refNumber, sheetData.ausgaben, betragValue);
    if (ausgabenMatch) {
        matchFound = true;
        const matchInfo = processAusgabeMatch(ausgabenMatch, betragValue, row, columns, sheetCols.ausgabenCols);
        matchResults.push([matchInfo]);

        // Konten aus dem Mapping setzen
        const category = row[columns.kategorie - 1] ? row[columns.kategorie - 1].toString().trim() : "";
        let kontoSoll = "", kontoHaben = "";

        if (category && config.ausgaben.kontoMapping[category]) {
            const mapping = config.ausgaben.kontoMapping[category];
            kontoSoll = mapping.soll || "";
            kontoHaben = mapping.gegen || "";
        }

        kontoSollResults.push([kontoSoll]);
        kontoHabenResults.push([kontoHaben]);

        // Für spätere Markierung merken
        const key = `ausgabe#${ausgabenMatch.row}`;
        bankZuordnungen[key] = {
            typ: "ausgabe",
            row: ausgabenMatch.row,
            bankDatum: row[columns.datum - 1],
            matchInfo: matchInfo,
            transTyp: "Ausgabe"
        };
        return;
    }

    // Wenn keine Übereinstimmung in Ausgaben, dann in Eigenbelegen suchen
    if (!matchFound) {
        const eigenbelegMatch = findMatch(refNumber, sheetData.eigenbelege, betragValue);
        if (eigenbelegMatch) {
            matchFound = true;
            const matchInfo = processEigenbelegMatch(eigenbelegMatch, betragValue, row, columns, sheetCols.eigenbelegeCols);
            matchResults.push([matchInfo]);

            // Konten aus dem Mapping setzen
            const category = row[columns.kategorie - 1] ? row[columns.kategorie - 1].toString().trim() : "";
            let kontoSoll = "", kontoHaben = "";

            if (category && config.eigenbelege.kontoMapping[category]) {
                const mapping = config.eigenbelege.kontoMapping[category];
                kontoSoll = mapping.soll || "";
                kontoHaben = mapping.gegen || "";
            }

            kontoSollResults.push([kontoSoll]);
            kontoHabenResults.push([kontoHaben]);

            // Für spätere Markierung merken
            const key = `eigenbeleg#${eigenbelegMatch.row}`;
            bankZuordnungen[key] = {
                typ: "eigenbeleg",
                row: eigenbelegMatch.row,
                bankDatum: row[columns.datum - 1],
                matchInfo: matchInfo,
                transTyp: "Ausgabe"
            };
            return;
        }
    }

    // Wenn keine Übereinstimmung, auch in Gesellschafterkonto suchen
    if (!matchFound) {
        const gesellschafterMatch = findMatch(refNumber, sheetData.gesellschafterkonto, betragValue);
        if (gesellschafterMatch) {
            matchFound = true;
            const matchInfo = processGesellschafterMatch(gesellschafterMatch, betragValue, row, columns, sheetCols.gesellschafterCols);
            matchResults.push([matchInfo]);

            // Konten aus dem Mapping setzen
            const category = row[columns.kategorie - 1] ? row[columns.kategorie - 1].toString().trim() : "";
            let kontoSoll = "", kontoHaben = "";

            if (category && config.gesellschafterkonto.kontoMapping[category]) {
                const mapping = config.gesellschafterkonto.kontoMapping[category];
                kontoSoll = mapping.soll || "";
                kontoHaben = mapping.gegen || "";
            }

            kontoSollResults.push([kontoSoll]);
            kontoHabenResults.push([kontoHaben]);

            // Für spätere Markierung merken
            const key = `gesellschafterkonto#${gesellschafterMatch.row}`;
            bankZuordnungen[key] = {
                typ: "gesellschafterkonto",
                row: gesellschafterMatch.row,
                bankDatum: row[columns.datum - 1],
                matchInfo: matchInfo,
                transTyp: "Ausgabe"
            };
            return;
        }
    }

    // Wenn keine Übereinstimmung, auch in Holding Transfers suchen
    if (!matchFound) {
        const holdingMatch = findMatch(refNumber, sheetData.holdingTransfers, betragValue);
        if (holdingMatch) {
            matchFound = true;
            const matchInfo = processHoldingMatch(holdingMatch, betragValue, row, columns, sheetCols.holdingCols);
            matchResults.push([matchInfo]);

            // Konten aus dem Mapping setzen
            const category = row[columns.kategorie - 1] ? row[columns.kategorie - 1].toString().trim() : "";
            let kontoSoll = "", kontoHaben = "";

            if (category && config.holdingTransfers.kontoMapping[category]) {
                const mapping = config.holdingTransfers.kontoMapping[category];
                kontoSoll = mapping.soll || "";
                kontoHaben = mapping.gegen || "";
            }

            kontoSollResults.push([kontoSoll]);
            kontoHabenResults.push([kontoHaben]);

            // Für spätere Markierung merken
            const key = `holdingtransfer#${holdingMatch.row}`;
            bankZuordnungen[key] = {
                typ: "holdingtransfer",
                row: holdingMatch.row,
                bankDatum: row[columns.datum - 1],
                matchInfo: matchInfo,
                transTyp: "Ausgabe"
            };
            return;
        }
    }

    // FALLS keine Übereinstimmung, auch in Einnahmen suchen (für Gutschriften)
    if (!matchFound) {
        const gutschriftMatch = findMatch(refNumber, sheetData.einnahmen);
        if (gutschriftMatch) {
            matchFound = true;
            const matchInfo = processGutschriftMatch(gutschriftMatch, betragValue, row, columns, sheetCols.einnahmenCols);
            matchResults.push([matchInfo]);

            // Konten aus dem Mapping setzen
            const category = row[columns.kategorie - 1] ? row[columns.kategorie - 1].toString().trim() : "";
            let kontoSoll = "", kontoHaben = "";

            if (category && config.einnahmen.kontoMapping[category]) {
                const mapping = config.einnahmen.kontoMapping[category];
                kontoSoll = mapping.soll || "";
                kontoHaben = mapping.gegen || "";
            }

            kontoSollResults.push([kontoSoll]);
            kontoHabenResults.push([kontoHaben]);

            // Für spätere Markierung merken
            const key = `einnahme#${gutschriftMatch.row}`;
            bankZuordnungen[key] = {
                typ: "einnahme",
                row: gutschriftMatch.row,
                bankDatum: row[columns.datum - 1],
                matchInfo: matchInfo,
                transTyp: "Gutschrift"
            };
            return;
        }
    }

    // Wenn keine Übereinstimmung gefunden wurde, leere Ergebnisse hinzufügen
    if (!matchFound) {
        matchResults.push([""]);
        kontoSollResults.push([""]);
        kontoHabenResults.push([""]);
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
    // Keine Daten
    if (!reference || !refMap) return null;

    // Normalisierte Suche vorbereiten
    const normalizedRef = stringUtils.normalizeText(reference);

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

    // 3. Teilweise Übereinstimmung
    const candidateKeys = Object.keys(refMap);

    // Zuerst prüfen wir, ob die Referenz in einem Schlüssel enthalten ist
    for (const key of candidateKeys) {
        if (key.includes(reference) || reference.includes(key)) {
            return evaluateMatch(refMap[key], betrag);
        }
    }

    // Falls keine exakten Treffer, probieren wir es mit normalisierten Werten
    for (const key of candidateKeys) {
        const normalizedKey = stringUtils.normalizeText(key);
        if (normalizedKey.includes(normalizedRef) || normalizedRef.includes(normalizedKey)) {
            return evaluateMatch(refMap[key], betrag);
        }
    }

    return null;
}

/**
 * Bewertet die Qualität einer Übereinstimmung basierend auf Beträgen
 * @param {Object} match - Die gefundene Übereinstimmung
 * @param {number} betrag - Der Betrag zum Vergleich
 * @returns {Object} - Übereinstimmung mit zusätzlichen Infos
 */
function evaluateMatch(match, betrag = null) {
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
    if (numberUtils.isApproximatelyEqual(betrag, matchBrutto, tolerance)) {
        result.matchType = "Vollständige Zahlung";
        return result;
    }

    // Fall 2: Position ist bereits vollständig bezahlt
    if (numberUtils.isApproximatelyEqual(matchBezahlt, matchBrutto, tolerance) && matchBezahlt > 0) {
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
    if (betrag > matchBrutto && numberUtils.isApproximatelyEqual(betrag, matchBrutto, tolerance)) {
        result.matchType = "Vollständige Zahlung";
        return result;
    }

    // Fall 5: Bei allen anderen Fällen (Beträge weichen stärker ab)
    result.matchType = "Unsichere Zahlung";
    result.betragsDifferenz = numberUtils.round(Math.abs(betrag - matchBrutto), 2);
    return result;
}

/**
 * Verarbeitet eine Einnahmen-Übereinstimmung
 * @param {Object} matchResult - Das Match-Ergebnis
 * @param {number} betragValue - Der Betrag
 * @param {Array} row - Die Bankbewegungszeile
 * @param {Object} columns - Die Spaltenkonfiguration
 * @param {Object} einnahmenCols - Die Spaltenkonfiguration für Einnahmen
 * @returns {string} Formatierte Match-Information
 */
function processEinnahmeMatch(matchResult, betragValue, row, columns, einnahmenCols) {
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

    return `Einnahme #${matchResult.row}${matchStatus}`;
}

/**
 * Verarbeitet eine Ausgaben-Übereinstimmung
 * @param {Object} matchResult - Das Match-Ergebnis
 * @param {number} betragValue - Der Betrag
 * @param {Array} row - Die Bankbewegungszeile
 * @param {Object} columns - Die Spaltenkonfiguration
 * @param {Object} ausgabenCols - Die Spaltenkonfiguration für Ausgaben
 * @returns {string} Formatierte Match-Information
 */
function processAusgabeMatch(matchResult, betragValue, row, columns, ausgabenCols) {
    let matchStatus = "";

    // Je nach Match-Typ unterschiedliche Statusinformationen für Ausgaben
    if (matchResult.matchType) {
        if (matchResult.matchType === "Unsichere Zahlung" && matchResult.betragsDifferenz) {
            matchStatus = ` (${matchResult.matchType}, Diff: ${matchResult.betragsDifferenz}€)`;
        } else {
            matchStatus = ` (${matchResult.matchType})`;
        }
    }

    return `Ausgabe #${matchResult.row}${matchStatus}`;
}

/**
 * Verarbeitet eine Eigenbeleg-Übereinstimmung
 * @param {Object} eigenbelegMatch - Das Match-Ergebnis
 * @param {number} betragValue - Der Betrag
 * @param {Array} row - Die Bankbewegungszeile
 * @param {Object} columns - Die Spaltenkonfiguration
 * @param {Object} eigenbelegeCols - Die Spaltenkonfiguration für Eigenbelege
 * @returns {string} Formatierte Match-Information
 */
function processEigenbelegMatch(eigenbelegMatch, betragValue, row, columns, eigenbelegeCols) {
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

    return `Eigenbeleg #${eigenbelegMatch.row}${matchStatus}`;
}

/**
 * Verarbeitet eine Gesellschafterkonto-Übereinstimmung
 * @param {Object} gesellschafterMatch - Das Match-Ergebnis
 * @param {number} betragValue - Der Betrag
 * @param {Array} row - Die Bankbewegungszeile
 * @param {Object} columns - Die Spaltenkonfiguration
 * @param {Object} gesellschafterCols - Die Spaltenkonfiguration für Gesellschafterkonto
 * @returns {string} Formatierte Match-Information
 */
function processGesellschafterMatch(gesellschafterMatch, betragValue, row, columns, gesellschafterCols) {
    let matchStatus = "";

    // Je nach Match-Typ unterschiedliche Statusinformationen
    if (gesellschafterMatch.matchType) {
        // Bei "Unsichere Zahlung" auch die Differenz anzeigen
        if (gesellschafterMatch.matchType === "Unsichere Zahlung" && gesellschafterMatch.betragsDifferenz) {
            matchStatus = ` (${gesellschafterMatch.matchType}, Diff: ${gesellschafterMatch.betragsDifferenz}€)`;
        } else {
            matchStatus = ` (${gesellschafterMatch.matchType})`;
        }
    }

    return `Gesellschafterkonto #${gesellschafterMatch.row}${matchStatus}`;
}

/**
 * Verarbeitet eine Holding-Transfers-Übereinstimmung
 * @param {Object} holdingMatch - Das Match-Ergebnis
 * @param {number} betragValue - Der Betrag
 * @param {Array} row - Die Bankbewegungszeile
 * @param {Object} columns - Die Spaltenkonfiguration
 * @param {Object} holdingCols - Die Spaltenkonfiguration für Holding Transfers
 * @returns {string} Formatierte Match-Information
 */
function processHoldingMatch(holdingMatch, betragValue, row, columns, holdingCols) {
    let matchStatus = "";

    // Je nach Match-Typ unterschiedliche Statusinformationen
    if (holdingMatch.matchType) {
        // Bei "Unsichere Zahlung" auch die Differenz anzeigen
        if (holdingMatch.matchType === "Unsichere Zahlung" && holdingMatch.betragsDifferenz) {
            matchStatus = ` (${holdingMatch.matchType}, Diff: ${holdingMatch.betragsDifferenz}€)`;
        } else {
            matchStatus = ` (${holdingMatch.matchType})`;
        }
    }

    return `Holding Transfer #${holdingMatch.row}${matchStatus}`;
}

/**
 * Verarbeitet eine Gutschrift-Übereinstimmung
 * @param {Object} gutschriftMatch - Das Match-Ergebnis
 * @param {number} betragValue - Der Betrag
 * @param {Array} row - Die Bankbewegungszeile
 * @param {Object} columns - Die Spaltenkonfiguration
 * @param {Object} einnahmenCols - Die Spaltenkonfiguration für Einnahmen
 * @returns {string} Formatierte Match-Information
 */
function processGutschriftMatch(gutschriftMatch, betragValue, row, columns, einnahmenCols) {
    let matchStatus = "";

    // Bei Gutschriften könnte der Betrag abweichen (z.B. Teilgutschrift)
    // Prüfen, ob die Beträge plausibel sind
    const gutschriftBetrag = Math.abs(gutschriftMatch.brutto);

    if (numberUtils.isApproximatelyEqual(betragValue, gutschriftBetrag, 0.01)) {
        // Beträge stimmen genau überein
        matchStatus = " (Vollständige Gutschrift)";
    } else if (betragValue < gutschriftBetrag) {
        // Teilgutschrift (Gutschriftbetrag kleiner als ursprünglicher Rechnungsbetrag)
        matchStatus = " (Teilgutschrift)";
    } else {
        // Ungewöhnlicher Fall - Gutschriftbetrag größer als Rechnungsbetrag
        matchStatus = " (Ungewöhnliche Gutschrift)";
    }

    return `Gutschrift zu Einnahme #${gutschriftMatch.row}${matchStatus}`;
}

/**
 * Formatiert Zeilen basierend auf dem Match-Typ
 * @param {Sheet} sheet - Das Sheet
 * @param {number} firstDataRow - Erste Datenzeile
 * @param {Array} matchResults - Array mit Match-Infos
 * @param {Object} columns - Spaltenkonfiguration
 */
function formatMatchedRows(sheet, firstDataRow, matchResults, columns) {
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
        'Gesellschafterkonto': { rows: [], color: "#E2EFDA" },  // Helles Grün für Gesellschafterkonto
        'Holding Transfer': { rows: [], color: "#FFF2CC" },  // Helles Gelb für Holding Transfers
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
        } else if (matchText.includes("Gesellschafterkonto")) {
            formatBatches['Gesellschafterkonto'].rows.push(rowIndex);
        } else if (matchText.includes("Holding Transfer")) {
            formatBatches['Holding Transfer'].rows.push(rowIndex);
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
}

/**
 * Setzt bedingte Formatierung für die Match-Spalte
 * @param {Sheet} sheet - Das Sheet
 * @param {string} columnLetter - Buchstabe für die Spalte
 */
function setMatchColumnFormatting(sheet, columnLetter) {
    const conditions = [
        // Grundlegende Match-Typen mit beginsWith Pattern
        {value: "Einnahme", background: "#C6EFCE", fontColor: "#006100", pattern: "beginsWith"},
        {value: "Ausgabe", background: "#FFC7CE", fontColor: "#9C0006", pattern: "beginsWith"},
        {value: "Eigenbeleg", background: "#DDEBF7", fontColor: "#2F5597", pattern: "beginsWith"},
        {value: "Gesellschafterkonto", background: "#E2EFDA", fontColor: "#375623", pattern: "beginsWith"},
        {value: "Holding Transfer", background: "#FFF2CC", fontColor: "#7F6000", pattern: "beginsWith"},
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
    ];

    // Bedingte Formatierung für die Match-Spalte setzen
    sheetUtils.setConditionalFormattingForColumn(sheet, columnLetter, conditions);
}

/**
 * Markiert bezahlte Einnahmen, Ausgaben und Eigenbelege farblich basierend auf dem Zahlungsstatus
 * @param {Spreadsheet} ss - Das Spreadsheet
 * @param {Object} bankZuordnungen - Zuordnungen aus dem Bankbewegungen-Sheet
 * @param {Object} config - Die Konfiguration
 */
function markPaidInvoices(ss, bankZuordnungen, config) {
    // Alle relevanten Sheets abrufen
    const sheets = {
        einnahmen: ss.getSheetByName("Einnahmen"),
        ausgaben: ss.getSheetByName("Ausgaben"),
        eigenbelege: ss.getSheetByName("Eigenbelege"),
        gesellschafterkonto: ss.getSheetByName("Gesellschafterkonto"),
        holdingTransfers: ss.getSheetByName("Holding Transfers")
    };

    // Für jedes Sheet Zahlungsdaten aktualisieren
    Object.entries(sheets).forEach(([type, sheet]) => {
        if (sheet && sheet.getLastRow() > 1) {
            markPaidRows(sheet, type, bankZuordnungen, config);
        }
    });
}

/**
 * Markiert bezahlte Zeilen in einem Sheet
 * @param {Sheet} sheet - Das zu aktualisierende Sheet
 * @param {string} sheetType - Typ des Sheets ("einnahmen", "ausgaben", "eigenbelege", etc.)
 * @param {Object} bankZuordnungen - Zuordnungen aus dem Bankbewegungen-Sheet
 * @param {Object} config - Die Konfiguration
 */
function markPaidRows(sheet, sheetType, bankZuordnungen, config) {
    // Konfiguration für das Sheet
    const columns = config[sheetType].columns;

    // Hole Werte aus dem Sheet
    const numRows = sheet.getLastRow() - 1;
    if (numRows <= 0) return;

    // Bestimme die maximale Anzahl an Spalten, die benötigt werden
    const maxCol = Math.max(
        columns.zahlungsdatum || 0,
        columns.bankabgleich || 0
    );

    if (maxCol === 0) return; // Keine relevanten Spalten verfügbar

    const data = sheet.getRange(2, 1, numRows, maxCol).getValues();

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
        const nettobetrag = numberUtils.parseCurrency(data[i][columns.nettobetrag - 1]);
        const bezahltBetrag = numberUtils.parseCurrency(data[i][columns.bezahlt - 1]);
        const zahlungsDatum = data[i][columns.zahlungsdatum - 1];
        const referenz = data[i][columns.rechnungsnummer - 1];

        // Prüfe, ob es eine Gutschrift ist
        const isGutschrift = referenz && referenz.toString().startsWith("G-");

        // Prüfe, ob diese Position im Banking-Sheet zugeordnet wurde
        // Bei Eigenbelegen verwenden wir den key "eigenbeleg#row"
        const prefix = sheetType === "eigenbelege" ? "eigenbeleg" :
            sheetType === "gesellschafterkonto" ? "gesellschafterkonto" :
                sheetType === "holdingTransfers" ? "holdingtransfer" :
                    sheetType;

        const zuordnungsKey = `${prefix}#${row}`;
        const hatBankzuordnung = bankZuordnungen[zuordnungsKey] !== undefined;

        // Zahlungsstatus berechnen
        const mwst = columns.mwstSatz ? numberUtils.parseMwstRate(data[i][columns.mwstSatz - 1], config.tax.defaultMwst) / 100 : 0;
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
            if (columns.bankabgleich && data[i][columns.bankabgleich - 1]?.toString().startsWith("✓ Bank:")) {
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
    if (bankabgleichUpdates.length > 0 && columns.bankabgleich) {
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
}

/**
 * Wendet eine Farbe auf mehrere Zeilen an
 * @param {Sheet} sheet - Das Sheet
 * @param {Array} rows - Die zu färbenden Zeilennummern
 * @param {string|null} color - Die Hintergrundfarbe oder null zum Zurücksetzen
 */
function applyColorToRows(sheet, rows, color) {
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
}

/**
 * Erstellt einen Informationstext für eine Bank-Zuordnung
 * @param {Object} zuordnung - Die Zuordnungsinformation
 * @returns {string} - Formatierter Informationstext
 */
function getZuordnungsInfo(zuordnung) {
    if (!zuordnung) return "";

    let infoText = "✓ Bank: " + zuordnung.bankDatum;

    // Wenn es mehrere Zuordnungen gibt (z.B. bei aufgeteilten Zahlungen)
    if (zuordnung.additional && zuordnung.additional.length > 0) {
        infoText += " + " + zuordnung.additional.length + " weitere";
    }

    return infoText;
}

export default {
    performBankReferenceMatching,
    processEinnahmeMatching,
    processAusgabeMatching,
    findMatch,
    evaluateMatch,
    formatMatchedRows,
    setMatchColumnFormatting,
    markPaidInvoices
};