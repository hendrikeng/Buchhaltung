// modules/refreshModule/matchingUtils.js

import stringUtils from '../../utils/stringUtils.js';
import numberUtils from '../../utils/numberUtils.js';
import globalCache from '../../utils/cacheUtils.js';

/**
 * Lädt Referenzdaten aus allen relevanten Sheets
 */
function loadReferenceData(ss, config) {
    const result = {};

    // Sheets und ihre Konfigurationen
    const sheets = {
        einnahmen: { name: "Einnahmen", cols: config.einnahmen.columns },
        ausgaben: { name: "Ausgaben", cols: config.ausgaben.columns },
        eigenbelege: { name: "Eigenbelege", cols: config.eigenbelege.columns },
        gesellschafterkonto: { name: "Gesellschafterkonto", cols: config.gesellschafterkonto.columns },
        holdingTransfers: { name: "Holding Transfers", cols: config.holdingTransfers.columns }
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
 * Findet eine Übereinstimmung in der Referenz-Map
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

    // Fall 3: Teilzahlung
    if (betrag < matchBrutto && (matchBrutto - betrag) > (matchBrutto * 0.1)) {
        result.matchType = "Teilzahlung";
        return result;
    }

    // Fall 4: Betrag größer, aber nahe dem Bruttobetrag
    if (betrag > matchBrutto && numberUtils.isApproximatelyEqual(betrag, matchBrutto, tolerance)) {
        result.matchType = "Vollständige Zahlung";
        return result;
    }

    // Fall 5: Bei allen anderen Fällen
    result.matchType = "Unsichere Zahlung";
    result.betragsDifferenz = numberUtils.round(Math.abs(betrag - matchBrutto), 2);
    return result;
}

export default {
    loadReferenceData,
    loadSheetReferenceData,
    findMatch,
    evaluateMatch
};