// src/modules/bankReconciliationModule/matchingHandler.js
import stringUtils from '../../utils/stringUtils.js';
import numberUtils from '../../utils/numberUtils.js';
import globalCache from '../../utils/cacheUtils.js';
import dateUtils from '../../utils/dateUtils.js';

/**
 * Führt das Matching von Bankbewegungen mit Rechnungen durch
 * @param {Spreadsheet} ss - Das Spreadsheet
 * @param {Sheet} sheet - Das Bankbewegungen-Sheet
 * @param {number} firstDataRow - Erste Datenzeile
 * @param {number} numDataRows - Anzahl der Datenzeilen
 * @param {number} lastRow - Letzte Zeile
 * @param {Object} columns - Spaltenkonfiguration
 * @param {Object} config - Die Konfiguration
 * @returns {Object} Ergebnis des Matchings mit matchInfo und bankZuordnungen
 */
function performBankReferenceMatching(ss, sheet, firstDataRow, numDataRows, lastRow, columns, config) {
    // Optimierung: Laden und Verarbeiten der Daten in einer Operation
    console.log('Starting bank reference matching...');

    // 1. Referenzdaten in einem Batch laden - dies vermeidet wiederholte Sheet-Zugriffe
    const sheetData = loadReferenceData(ss, config);

    // 2. Bankdaten in einem Batch laden
    const bankData = sheet.getRange(firstDataRow, 1, numDataRows, sheet.getLastColumn()).getValues();

    // 3. Matching vorbereiten
    const results = {
        matchInfo: [],
        kontoSollResults: [],
        kontoHabenResults: [],
        categoryResults: [],
        bankZuordnungen: {},
    };

    // 4. Kontoinformationen sammeln (einmalig)
    const accountInfo = collectAccountInfo(config);

    // 5. Matching in einem Durchgang durchführen
    processBankEntries(
        bankData,
        firstDataRow,
        lastRow,
        columns,
        sheetData,
        results,
        accountInfo,
        config,
    );

    // 6. Batch-Updates durchführen
    performBatchUpdates(
        sheet,
        firstDataRow,
        numDataRows,
        columns,
        results.matchInfo,
        results.kontoSollResults,
        results.kontoHabenResults,
        results.categoryResults,
    );

    return {
        matchInfo: results.matchInfo,
        bankZuordnungen: results.bankZuordnungen,
    };
}

/**
 * Lädt Referenzdaten aus allen relevanten Sheets in einem Batch
 * @param {Spreadsheet} ss - Das Spreadsheet
 * @param {Object} config - Die Konfiguration
 * @returns {Object} Die geladenen Referenzdaten
 */
function loadReferenceData(ss, config) {
    const result = {};

    // Optimierung: Sheets und Konfigurationen in einem Array definieren
    const sheetConfig = [
        { key: 'einnahmen', name: 'Einnahmen', cols: config.einnahmen.columns },
        { key: 'ausgaben', name: 'Ausgaben', cols: config.ausgaben.columns },
        { key: 'eigenbelege', name: 'Eigenbelege', cols: config.eigenbelege.columns },
        { key: 'gesellschafterkonto', name: 'Gesellschafterkonto', cols: config.gesellschafterkonto.columns },
        { key: 'holdingTransfers', name: 'Holding Transfers', cols: config.holdingTransfers.columns },
    ];

    // Für jedes Sheet Referenzdaten laden - parallelisiert
    sheetConfig.forEach(cfg => {
        // Prüfe zuerst den Cache
        if (globalCache.has('references', cfg.key)) {
            result[cfg.key] = globalCache.get('references', cfg.key);
            return;
        }

        const sheet = ss.getSheetByName(cfg.name);
        if (sheet && sheet.getLastRow() > 1) {
            result[cfg.key] = loadSheetReferenceData(sheet, cfg.cols, cfg.key);
        } else {
            result[cfg.key] = {};
        }
    });

    return result;
}

/**
 * Lädt Referenzdaten aus einem einzelnen Sheet mit optimierter Verarbeitung
 * @param {Sheet} sheet - Das Sheet
 * @param {Object} columns - Die Spaltenkonfiguration
 * @param {string} sheetType - Der Sheet-Typ
 * @returns {Object} Die geladenen Referenzdaten
 */
function loadSheetReferenceData(sheet, columns, sheetType) {
    // Cache prüfen
    if (globalCache.has('references', sheetType)) {
        return globalCache.get('references', sheetType);
    }

    const map = {};
    const numRows = sheet.getLastRow() - 1;

    if (numRows <= 0) return map;

    // Optimierung: Nur benötigte Spalten laden
    const requiredCols = [
        columns.rechnungsnummer,
        columns.nettobetrag,
        columns.mwstSatz,
        columns.bezahlt,
        columns.kategorie,
        columns.buchungskonto,
    ].filter(col => col !== undefined);

    const maxCol = Math.max(...requiredCols);

    // Daten in einem Batch laden
    const data = sheet.getRange(2, 1, numRows, maxCol).getValues();

    // Effizienter Aufbau der Referenzdaten
    for (let i = 0; i < data.length; i++) {
        const row = data[i];
        const ref = row[columns.rechnungsnummer - 1];

        if (stringUtils.isEmpty(ref)) continue;

        // Key immer ohne G-Prefix
        const key = String(ref).trim();

        // Werte vorab extrahieren und berechnen
        const betrag = numberUtils.parseCurrency(row[columns.nettobetrag - 1] || 0);
        const isGutschrift = betrag < 0;
        const normalizedKey = stringUtils.normalizeText(key);
        const mwstRate = columns.mwstSatz ?
            numberUtils.parseMwstRate(row[columns.mwstSatz - 1] || 0) : 0;
        const bezahlt = columns.bezahlt ?
            Math.abs(numberUtils.parseCurrency(row[columns.bezahlt - 1] || 0)) : 0;
        const kategorie = columns.kategorie && row[columns.kategorie - 1] ?
            row[columns.kategorie - 1].toString().trim() : '';
        const buchungskonto = columns.buchungskonto && row[columns.buchungskonto - 1] ?
            row[columns.buchungskonto - 1].toString().trim() : '';

        // Bruttobetrag einmal berechnen
        const bruttoAbs = Math.abs(betrag) * (1 + mwstRate/100);

        // Optimiertes Datenformat mit allen benötigten Informationen
        map[key] = {
            ref: key,
            normalizedRef: normalizedKey,
            row: i + 2, // 1-basiert + Überschrift
            betrag: betrag,
            mwstRate: mwstRate,
            brutto: bruttoAbs * (isGutschrift ? -1 : 1), // Preserve sign for display
            bezahlt: bezahlt,
            offen: numberUtils.round(bruttoAbs - bezahlt, 2) * (isGutschrift ? -1 : 1), // Preserve sign
            isGutschrift: isGutschrift,
            kategorie: kategorie,
            buchungskonto: buchungskonto,
            originalData: {
                kategorie: kategorie,
                buchungskonto: buchungskonto,
                betrag: betrag, // Keep the original value with sign
                rowData: row, // Komplette Zeile speichern
                sheetType: sheetType, // Typ des Sheets speichern
            },
        };

        // Normalisierter Key als Alternative für besseres Matching
        if (normalizedKey !== key && !map[normalizedKey]) {
            map[normalizedKey] = map[key];
        }
    }

    // Referenzdaten cachen
    globalCache.set('references', sheetType, map);
    return map;
}

/**
 * Sammelt Kontoinformationen aus der Konfiguration (optimiert)
 * @param {Object} config - Die Konfiguration
 * @returns {Object} Gesammelte Kontoinformationen
 */
function collectAccountInfo(config) {
    const allowedKontoSoll = new Set();
    const allowedGegenkonto = new Set();

    // Alle relevanten Mappings in einem Array
    const mappings = [
        config.einnahmen.kontoMapping,
        config.ausgaben.kontoMapping,
        config.eigenbelege.kontoMapping,
        config.gesellschafterkonto.kontoMapping,
        config.holdingTransfers.kontoMapping,
    ];

    // Effizienter Durchlauf durch alle Mappings in einer Schleife
    mappings.forEach(mapping => {
        Object.values(mapping).forEach(m => {
            if (m.soll) allowedKontoSoll.add(m.soll);
            if (m.gegen) allowedGegenkonto.add(m.gegen);
        });
    });

    // Fallback-Konten bestimmen
    const fallbackKontoSoll = allowedKontoSoll.size > 0 ? Array.from(allowedKontoSoll)[0] : '';
    const fallbackKontoHaben = allowedGegenkonto.size > 0 ? Array.from(allowedGegenkonto)[0] : '';

    return {
        allowedKontoSoll,
        allowedGegenkonto,
        fallbackKontoSoll,
        fallbackKontoHaben,
    };
}

/**
 * Verarbeitet Bankeinträge und sucht nach Übereinstimmungen
 * @param {Array} bankData - Die Bankdaten
 * @param {number} firstDataRow - Erste Datenzeile
 * @param {number} lastRow - Letzte Zeile
 * @param {Object} columns - Spaltenkonfiguration
 * @param {Object} sheetData - Referenzdaten
 * @param {Object} results - Ergebnisobjekt
 * @param {Object} accountInfo - Kontoinformationen
 * @param {Object} config - Die Konfiguration
 */
function processBankEntries(bankData, firstDataRow, lastRow, columns, sheetData,
    results, accountInfo, config) {

    // Destrukturiere die accountInfo für bessere Lesbarkeit
    const { fallbackKontoSoll, fallbackKontoHaben } = accountInfo;

    // Durchlaufe jede Bankbewegung einmal
    for (let i = 0; i < bankData.length; i++) {
        const rowIndex = i + firstDataRow;
        const row = bankData[i];

        // Prüfe, ob es sich um die Endsaldo-Zeile handelt
        if (isEndSaldoRow(rowIndex, lastRow, row, columns)) {
            addEmptyResults(results);
            continue;
        }

        // Extrahiere alle benötigten Informationen aus der Zeile
        const tranType = row[columns.transaktionstyp - 1]; // Einnahme/Ausgabe
        const refNumber = row[columns.referenz - 1];       // Referenznummer
        const betragValue = Math.abs(numberUtils.parseCurrency(row[columns.betrag - 1]));
        const betragOriginal = numberUtils.parseCurrency(row[columns.betrag - 1]); // Original with sign
        const category = row[columns.kategorie - 1] ? row[columns.kategorie - 1].toString().trim() : '';

        // Flag für gefundene Übereinstimmung
        let matchFound = false;

        // Nur Matching durchführen, wenn eine Referenznummer vorhanden ist
        if (!stringUtils.isEmpty(refNumber)) {
            // Für Ausgaben mit negativem Betrag zuerst auf Gutschriften prüfen
            if (tranType === 'Ausgabe' && betragOriginal < 0) {
                matchFound = processDocumentTypeMatching(
                    sheetData.einnahmen, refNumber, betragValue, row,
                    columns, 'gutschrift', config.einnahmen.columns, config.einnahmen.kontoMapping,
                    results, category, config, firstDataRow, true);

                // Bei Gutschrift-Treffer den nächsten Eintrag verarbeiten
                if (matchFound) continue;
            }

            // Standard-Logik für Einnahmen
            if (tranType === 'Einnahme' && !matchFound) {
                matchFound = processDocumentTypeMatching(
                    sheetData.einnahmen, refNumber, betragValue, row,
                    columns, 'einnahme', config.einnahmen.columns, config.einnahmen.kontoMapping,
                    results, category, config, firstDataRow);
            }

            // Standard-Logik für Ausgaben - Optimiert mit Array von Dokumenttypen
            if (tranType === 'Ausgabe' && !matchFound) {
                const docTypes = [
                    { data: sheetData.ausgaben, type: 'ausgabe', cols: config.ausgaben.columns, mapping: config.ausgaben.kontoMapping },
                    { data: sheetData.eigenbelege, type: 'eigenbeleg', cols: config.eigenbelege.columns, mapping: config.eigenbelege.kontoMapping },
                    { data: sheetData.gesellschafterkonto, type: 'gesellschafterkonto', cols: config.gesellschafterkonto.columns, mapping: config.gesellschafterkonto.kontoMapping },
                    { data: sheetData.holdingTransfers, type: 'holdingtransfer', cols: config.holdingTransfers.columns, mapping: config.holdingTransfers.kontoMapping },
                ];

                // Optimiert: Verwende some() für frühzeitigen Abbruch bei Treffer
                matchFound = docTypes.some(docType =>
                    processDocumentTypeMatching(
                        docType.data, refNumber, betragValue, row,
                        columns, docType.type, docType.cols, docType.mapping,
                        results, category, config, firstDataRow),
                );
            }
        }

        // Wenn keine Übereinstimmung gefunden, default Konten basierend auf Kategorie setzen
        if (!matchFound) {
            setDefaultAccounts(
                category, tranType, config, fallbackKontoSoll, fallbackKontoHaben,
                results,
            );
        }
    }
}

/**
 * Prüft, ob es sich um die Endsaldo-Zeile handelt
 * @param {number} rowIndex - Zeilenindex
 * @param {number} lastRow - Letzte Zeile
 * @param {Array} row - Zeilendaten
 * @param {Object} columns - Spaltenkonfiguration
 * @returns {boolean} true wenn es die Endsaldo-Zeile ist
 */
function isEndSaldoRow(rowIndex, lastRow, row, columns) {
    const label = row[columns.buchungstext - 1] ? row[columns.buchungstext - 1].toString().trim().toLowerCase() : '';
    return rowIndex === lastRow && label === 'endsaldo';
}

/**
 * Fügt leere Ergebnisse zu den Ergebnis-Arrays hinzu
 * @param {Object} results - Das Ergebnisobjekt
 */
function addEmptyResults(results) {
    results.matchInfo.push(['']);
    results.kontoSollResults.push(['']);
    results.kontoHabenResults.push(['']);
    results.categoryResults.push(['']);
}

/**
 * Optimierte Version des Document-Type-Matchings
 * @param {Object} docData - Referenzdaten
 * @param {string} refNumber - Referenznummer
 * @param {number} betragValue - Betrag
 * @param {Array} row - Zeilendaten
 * @param {Object} columns - Spaltenkonfiguration
 * @param {string} docType - Dokumenttyp
 * @param {Object} docCols - Dokumentspaltenkonfiguration
 * @param {Object} kontoMapping - Kontomapping
 * @param {Object} results - Ergebnisobjekt
 * @param {string} category - Kategorie
 * @param {Object} config - Konfiguration
 * @param {number} firstDataRow - Erste Datenzeile
 * @param {boolean} isGutschrift - Ist eine Gutschrift
 * @returns {boolean} true wenn eine Übereinstimmung gefunden wurde
 */
function processDocumentTypeMatching(docData, refNumber, betragValue, row, columns, docType, docCols, kontoMapping,
    results, category, config, firstDataRow, isGutschrift = false) {

    // Nur Gutschrift-Verarbeitung für Einnahmen erlauben
    if (isGutschrift && docType !== 'einnahme' && docType !== 'gutschrift') {
        return false;
    }

    // Übereinstimmung suchen
    const matchResult = findMatch(refNumber, docData, isGutschrift ? null : betragValue);
    if (!matchResult) return false;

    // Für Gutschrift-Matching prüfen, ob der Betrag negativ ist
    if (isGutschrift && (!matchResult.originalData || matchResult.originalData.betrag >= 0)) {
        return false; // Keine tatsächliche Gutschrift, diese Übereinstimmung überspringen
    }

    // Referenz-Übereinstimmung validieren
    if (matchResult.ref && !isGoodReferenceMatch(refNumber, matchResult.ref)) {
        return false;
    }

    // Match-Info erstellen
    const matchInfoText = createMatchInfo(matchResult, docType, betragValue);

    // Kategorie und Konto aus dem gefundenen Dokument verwenden
    let sourceCategory = '';
    let sourceKonto = '';

    if (matchResult.originalData) {
        sourceCategory = matchResult.originalData.kategorie || '';
        sourceKonto = matchResult.originalData.buchungskonto || '';
    }

    // Konten aus dem Mapping setzen
    let kontoSoll = '', kontoHaben = '';

    if (sourceCategory && kontoMapping[sourceCategory]) {
        const mapping = kontoMapping[sourceCategory];
        kontoSoll = mapping.soll || '';
        kontoHaben = mapping.gegen || '';
    } else if (category && kontoMapping[category]) {
        // Fallback auf Bank-Kategorie
        const mapping = kontoMapping[category];
        kontoSoll = mapping.soll || '';
        kontoHaben = mapping.gegen || '';
    }

    // Ergebnisse zu den Arrays hinzufügen
    results.matchInfo.push([matchInfoText]);
    results.kontoSollResults.push([kontoSoll]);
    results.kontoHabenResults.push([kontoHaben]);
    // Behalte die Bank-Kategorie, die Kategorie aus dem Dokument wird im Approval-Prozess vorgeschlagen
    results.categoryResults.push([category]);

    // Optimierte Schlüsselgenerierung für Bankzuordnungen
    const key = `${docType}#${matchResult.row}`;

    // Verbesserte Gutschrift-Erkennung
    const isActualGutschrift = docType === 'gutschrift' ||
        (matchResult.originalData && matchResult.originalData.betrag < 0);

    // Information für spätere Markierung speichern
    results.bankZuordnungen[key] = {
        typ: isActualGutschrift ? 'gutschrift' : (docType === 'gutschrift' ? 'einnahme' : docType),
        row: matchResult.row,
        bankRow: firstDataRow + results.matchInfo.length - 1, // Speichern der aktuellen Bankzeile
        bankDatum: row[columns.datum - 1],
        matchInfo: matchInfoText,
        transTyp: isActualGutschrift ? 'Gutschrift' : row[columns.transaktionstyp - 1],
        category: category,
        sourceCategory: sourceCategory, // Add original document category separately
        kontoSoll: kontoSoll,
        kontoHaben: kontoHaben,
        isGutschrift: isActualGutschrift,
        originalRef: matchResult.ref,
    };

    return true;
}

/**
 * Sucht eine Übereinstimmung in der Referenz-Map mit optimierten Suchstrategien
 * @param {string} reference - Die Referenz
 * @param {Object} refMap - Die Referenzdaten
 * @param {number} betrag - Der Betrag
 * @returns {Object|null} Die gefundene Übereinstimmung oder null
 */
function findMatch(reference, refMap, betrag = null) {
    // Keine Daten oder keine Referenz
    if (!reference || !refMap) return null;

    // Ensure reference is a string
    const refString = String(reference).trim();
    if (!refString) return null;

    // Normalisierte Suche vorbereiten
    const normalizedRef = stringUtils.normalizeText(refString);
    if (!normalizedRef) return null;

    // Matching-Strategien in effizienter Reihenfolge
    let match = null;
    let matchedKey = null;

    // 1. Exakte Übereinstimmung (am häufigsten und schnellsten)
    if (refMap[refString]) {
        match = evaluateMatch(refMap[refString], betrag);
        if (match) matchedKey = refString;
    }

    // 2. Nur wenn keine exakte Übereinstimmung gefunden wurde: Normalisierte Übereinstimmung
    if (!match && normalizedRef && refMap[normalizedRef]) {
        match = evaluateMatch(refMap[normalizedRef], betrag);
        if (match) matchedKey = normalizedRef;
    }

    // 3. Nur wenn immer noch kein Match: Teilweise Übereinstimmung mit Optimierung
    if (!match) {
        // Optimierung: Zuerst prüfen wir Schlüssel, die die Referenz enthalten
        // Dies ist eine schnellere Strategie als alle Schlüssel zu prüfen
        const candidateKeys = Object.keys(refMap).filter(key =>
            typeof key === 'string' &&
            (key.includes(refString) || refString.includes(key)),
        );

        // Wenn wir potenzielle Kandidaten haben, prüfen wir diese
        for (const key of candidateKeys) {
            if (isGoodReferenceMatch(key, refString)) {
                match = evaluateMatch(refMap[key], betrag);
                if (match) {
                    matchedKey = key;
                    break;
                }
            }
        }
    }

    // Originaldata hinzufügen, wenn ein Match gefunden wurde
    if (match && matchedKey) {
        match.originalData = refMap[matchedKey].originalData;
        match.ref = refMap[matchedKey].ref; // Original-Referenz für Validierung
    }

    return match;
}

/**
 * Bewertet die Qualität einer Übereinstimmung basierend auf Beträgen
 * @param {Object} match - Die gefundene Übereinstimmung
 * @param {number} betrag - Der zu vergleichende Betrag
 * @returns {Object|null} Die bewertete Übereinstimmung oder null
 */
function evaluateMatch(match, betrag = null) {
    if (!match) return null;

    // Ergebnis mit ursprünglichen Daten initialisieren
    const result = { ...match };

    // Wenn kein Betrag zum Vergleich angegeben wurde
    if (betrag === null) {
        result.matchType = 'Referenz-Match';
        return result;
    }

    // Absolute Beträge für Vergleich verwenden
    const matchBrutto = Math.abs(match.brutto);
    const matchBezahlt = Math.abs(match.bezahlt);

    // Toleranzwert für Betragsabweichungen (2 Cent)
    const tolerance = 0.02;

    // Optimierter Vergleichsalgorithmus mit frühzeitigem Abbruch

    // Fall 1: Betrag entspricht genau dem Bruttobetrag (Vollständige Zahlung)
    if (numberUtils.isApproximatelyEqual(betrag, matchBrutto, tolerance)) {
        result.matchType = 'Vollständige Zahlung';
        return result;
    }

    // Fall 2: Position ist bereits vollständig bezahlt
    if (numberUtils.isApproximatelyEqual(matchBezahlt, matchBrutto, tolerance) && matchBezahlt > 0) {
        result.matchType = 'Vollständige Zahlung';
        return result;
    }

    // Fall 3: Teilzahlung - nur wenn signifikant unter dem Bruttobetrag
    if (betrag < matchBrutto && (matchBrutto - betrag) > (matchBrutto * 0.1)) {
        result.matchType = 'Teilzahlung';
        return result;
    }

    // Fall 4: Betrag größer, aber nahe dem Bruttobetrag
    if (betrag > matchBrutto && numberUtils.isApproximatelyEqual(betrag, matchBrutto, tolerance)) {
        result.matchType = 'Vollständige Zahlung';
        return result;
    }

    // Fall 5: Alle anderen Fälle
    result.matchType = 'Unsichere Zahlung';
    result.betragsDifferenz = numberUtils.round(Math.abs(betrag - matchBrutto), 2);
    return result;
}

/**
 * Erstellt einen Informationstext für eine Match-Übereinstimmung
 * @param {Object} matchResult - Das Matching-Ergebnis
 * @param {string} docType - Der Dokumenttyp
 * @param {number} betragValue - Der Betrag
 * @returns {string} Der Informationstext
 */
function createMatchInfo(matchResult, docType, betragValue) {
    let matchStatus = '';

    // Match-Status mit Differenzanzeige
    if (matchResult.matchType) {
        if (matchResult.matchType === 'Unsichere Zahlung' && matchResult.betragsDifferenz) {
            matchStatus = ` (${matchResult.matchType}, Diff: ${matchResult.betragsDifferenz}€)`;
        } else {
            matchStatus = ` (${matchResult.matchType})`;
        }
    }

    // Verbesserte Gutschrift-Erkennung
    const isActualGutschrift = docType === 'gutschrift' ||
        (matchResult.originalData && matchResult.originalData.betrag < 0);

    // Spezielle Behandlung für Gutschriften
    if (isActualGutschrift) {
        const gutschriftBetrag = Math.abs(matchResult.originalData ?
            matchResult.originalData.betrag : matchResult.brutto);

        // Status für Gutschriften basierend auf Betrag
        if (numberUtils.isApproximatelyEqual(betragValue, gutschriftBetrag, 0.01)) {
            matchStatus = ' (Vollständige Gutschrift)';
        } else if (betragValue < gutschriftBetrag) {
            matchStatus = ' (Teilgutschrift)';
        } else {
            matchStatus = ' (Ungewöhnliche Gutschrift)';
        }

        return `Gutschrift zu Einnahme #${matchResult.row}${matchStatus}`;
    }

    // Konstanten für Dokumenttypen verbessern die Lesbarkeit und Performance
    const docTypeNames = {
        'einnahme': 'Einnahme',
        'ausgabe': 'Ausgabe',
        'eigenbeleg': 'Eigenbeleg',
        'gesellschafterkonto': 'Gesellschafterkonto',
        'holdingtransfer': 'Holding Transfer',
    };

    return `${docTypeNames[docType] || docType} #${matchResult.row}${matchStatus}`;
}

/**
 * Setzt Standard-Konten basierend auf der Kategorie und Transaktionstyp
 * @param {string} category - Die Kategorie
 * @param {string} tranType - Der Transaktionstyp
 * @param {Object} config - Die Konfiguration
 * @param {string} fallbackKontoSoll - Das Fallback-Konto Soll
 * @param {string} fallbackKontoHaben - Das Fallback-Konto Haben
 * @param {Object} results - Das Ergebnisobjekt
 */
function setDefaultAccounts(category, tranType, config, fallbackKontoSoll, fallbackKontoHaben, results) {
    let kontoSoll = fallbackKontoSoll;
    let kontoHaben = fallbackKontoHaben;

    if (category) {
        // Mapping-Quelle basierend auf Transaktionstyp auswählen
        const mappingSource = tranType === 'Einnahme' ?
            config.einnahmen.kontoMapping :
            config.ausgaben.kontoMapping;

        // Mapping für die Kategorie finden und anwenden
        if (mappingSource && mappingSource[category]) {
            const mapping = mappingSource[category];
            kontoSoll = mapping.soll || fallbackKontoSoll;
            kontoHaben = mapping.gegen || fallbackKontoHaben;
        }
    }

    // Ergebnisse hinzufügen
    results.matchInfo.push(['']);
    results.kontoSollResults.push([kontoSoll]);
    results.kontoHabenResults.push([kontoHaben]);
    results.categoryResults.push([category]); // Ursprüngliche Kategorie beibehalten
}

/**
 * Führt Batch-Updates für alle betroffenen Spalten aus
 * @param {Sheet} sheet - Das Sheet
 * @param {number} firstDataRow - Erste Datenzeile
 * @param {number} numDataRows - Anzahl der Datenzeilen
 * @param {Object} columns - Spaltenkonfiguration
 * @param {Array} matchInfo - Match-Informationen
 * @param {Array} kontoSollResults - Konto-Soll-Ergebnisse
 * @param {Array} kontoHabenResults - Konto-Haben-Ergebnisse
 * @param {Array} categoryResults - Kategorie-Ergebnisse
 */
function performBatchUpdates(sheet, firstDataRow, numDataRows, columns,
    matchInfo, kontoSollResults, kontoHabenResults, categoryResults) {

    // Batch-Updates für alle Spalten in separaten API-Calls
    if (columns.matchInfo) {
        sheet.getRange(firstDataRow, columns.matchInfo, numDataRows, 1).setValues(matchInfo);
    }

    if (columns.kontoSoll) {
        sheet.getRange(firstDataRow, columns.kontoSoll, numDataRows, 1).setValues(kontoSollResults);
    }

    if (columns.kontoHaben) {
        sheet.getRange(firstDataRow, columns.kontoHaben, numDataRows, 1).setValues(kontoHabenResults);
    }

    // Kategorien aktualisieren, wenn Kategorie-Spalte existiert
    if (columns.kategorie && categoryResults.length === numDataRows) {
        sheet.getRange(firstDataRow, columns.kategorie, numDataRows, 1).setValues(categoryResults);
    }
}

/**
 * Prüft, ob zwei Referenznummern gut genug übereinstimmen
 * @param {string} ref1 - Erste Referenz
 * @param {string} ref2 - Zweite Referenz
 * @returns {boolean} true wenn die Referenzen übereinstimmen
 */
function isGoodReferenceMatch(ref1, ref2) {
    // Leere Werte behandeln
    if (!ref1 || !ref2) return false;

    // Frühe Rückkehr bei exakter Übereinstimmung
    if (ref1 === ref2) return true;

    // Schnelle Prüfung, ob eine Referenz die andere enthält
    if (ref1.includes(ref2) || ref2.includes(ref1)) {
        // Bei kurzen Referenzen (weniger als 5 Zeichen) strengere Prüfung
        const shorter = ref1.length <= ref2.length ? ref1 : ref2;
        if (shorter.length < 5) {
            return ref1.startsWith(ref2) || ref2.startsWith(ref1);
        }
        return true;
    }

    // Keine gute Übereinstimmung
    return false;
}

export default {
    performBankReferenceMatching,
    isGoodReferenceMatch,
};