// modules/bwaModule/calculator.js
import dateUtils from '../../utils/dateUtils.js';
import numberUtils from '../../utils/numberUtils.js';
import stringUtils from '../../utils/stringUtils.js';

// Cache für häufig verwendete Kategoriezuordnungen
const categoryMappingCache = new Map();

/**
 * Verarbeitet Einnahmen und ordnet sie den BWA-Kategorien zu
 * Optimierte Version mit Cache und effizienteren Lookups
 * @param {Array} row - Zeile aus dem Einnahmen-Sheet
 * @param {Object} bwaData - BWA-Datenstruktur
 * @param {Object} config - Die Konfiguration
 */
function processRevenue(row, bwaData, config) {
    try {
        const columns = config.einnahmen.columns;

        // Monat bestimmen
        const m = dateUtils.getMonthFromRow(row, 'einnahmen', config);
        if (!m) return;

        // Betrag extrahieren
        const amount = numberUtils.parseCurrency(row[columns.nettobetrag - 1]);
        if (amount === 0) return;

        // Kategorie extrahieren
        const category = row[columns.kategorie - 1]?.toString().trim() || '';

        // Nicht-betriebliche Buchungen ignorieren
        const nonBusinessCategories = new Set(['Gewinnvortrag', 'Verlustvortrag', 'Gewinnvortrag/Verlustvortrag']);
        if (nonBusinessCategories.has(category)) return;

        // Cache-Schlüssel für Kategorie-Zuordnung
        const cacheKey = `rev_cat_${category}`;

        // Prüfen, ob Zuordnung im Cache
        if (categoryMappingCache.has(cacheKey)) {
            const target = categoryMappingCache.get(cacheKey);
            bwaData[m][target] += amount;
            return;
        }

        // Direkte Zuordnung basierend auf der Kategorie
        const directMappings = {
            'Sonstige betriebliche Erträge': 'sonstigeErtraege',
            'Erträge aus Vermietung/Verpachtung': 'vermietung',
            'Erträge aus Zuschüssen': 'zuschuesse',
            'Erträge aus Währungsgewinnen': 'waehrungsgewinne',
            'Erträge aus Anlagenabgängen': 'anlagenabgaenge',
        };

        if (directMappings[category]) {
            categoryMappingCache.set(cacheKey, directMappings[category]);
            bwaData[m][directMappings[category]] += amount;
            return;
        }

        // BWA-Mapping aus Konfiguration verwenden
        const mapping = config.einnahmen.bwaMapping[category];
        if (mapping === 'umsatzerloese' || mapping === 'provisionserloese') {
            categoryMappingCache.set(cacheKey, mapping);
            bwaData[m][mapping] += amount;
            return;
        }

        // Kategorisierung nach Steuersatz
        if (numberUtils.parseMwstRate(row[columns.mwstSatz - 1], config.tax.defaultMwst) === 0) {
            categoryMappingCache.set(cacheKey, 'steuerfreieInlandEinnahmen');
            bwaData[m].steuerfreieInlandEinnahmen += amount;
        } else {
            categoryMappingCache.set(cacheKey, 'umsatzerloese');
            bwaData[m].umsatzerloese += amount;
        }
    } catch (e) {
        console.error('Fehler bei der Verarbeitung einer Einnahme:', e);
    }
}

/**
 * Verarbeitet Ausgaben und ordnet sie den BWA-Kategorien zu
 * Optimierte Version mit Cache und effizienteren Lookups
 * @param {Array} row - Zeile aus dem Ausgaben-Sheet
 * @param {Object} bwaData - BWA-Datenstruktur
 * @param {Object} config - Die Konfiguration
 */
function processExpense(row, bwaData, config) {
    try {
        const columns = config.ausgaben.columns;

        // Monat bestimmen
        const m = dateUtils.getMonthFromRow(row, 'ausgaben', config);
        if (!m) return;

        // Betrag extrahieren
        const amount = numberUtils.parseCurrency(row[columns.nettobetrag - 1]);
        if (amount === 0) return;

        // Kategorie extrahieren
        const category = row[columns.kategorie - 1]?.toString().trim() || '';

        // Nicht-betriebliche Positionen ignorieren (mit Set für schnelleren Lookup)
        const nonBusinessCategories = new Set([
            'Privatentnahme', 'Privateinlage', 'Holding Transfers',
            'Gewinnvortrag', 'Verlustvortrag', 'Gewinnvortrag/Verlustvortrag',
        ]);

        if (nonBusinessCategories.has(category)) {
            return;
        }

        // Cache-Schlüssel für Kategorie-Zuordnung
        const cacheKey = `exp_cat_${category}`;

        // Prüfen, ob Zuordnung im Cache
        if (categoryMappingCache.has(cacheKey)) {
            const target = categoryMappingCache.get(cacheKey);
            bwaData[m][target] += amount;
            return;
        }

        // Direkte Zuordnung basierend auf der Kategorie
        const directMappings = {
            'Bruttolöhne & Gehälter': 'bruttoLoehne',
            'Soziale Abgaben & Arbeitgeberanteile': 'sozialeAbgaben',
            'Sonstige Personalkosten': 'sonstigePersonalkosten',
            'Gewerbesteuerrückstellungen': 'gewerbesteuerRueckstellungen',
            'Telefon & Internet': 'telefonInternet',
            'Bürokosten': 'buerokosten',
            'Fortbildungskosten': 'fortbildungskosten',
        };

        if (directMappings[category]) {
            categoryMappingCache.set(cacheKey, directMappings[category]);
            bwaData[m][directMappings[category]] += amount;
            return;
        }

        // BWA-Mapping aus Konfiguration verwenden
        const mapping = config.ausgaben.bwaMapping[category];
        if (!mapping) {
            categoryMappingCache.set(cacheKey, 'sonstigeAufwendungen');
            bwaData[m].sonstigeAufwendungen += amount;
            return;
        }

        // Optimierung: Set statt switch für schnelleren Lookup
        categoryMappingCache.set(cacheKey, mapping);
        bwaData[m][mapping] += amount;

    } catch (e) {
        console.error('Fehler bei der Verarbeitung einer Ausgabe:', e);
    }
}

/**
 * Verarbeitet Eigenbelege und ordnet sie den BWA-Kategorien zu
 * Optimierte Version mit Cache und effizienteren Lookups
 * @param {Array} row - Zeile aus dem Eigenbelege-Sheet
 * @param {Object} bwaData - BWA-Datenstruktur
 * @param {Object} config - Die Konfiguration
 */
function processEigen(row, bwaData, config) {
    try {
        const columns = config.eigenbelege.columns;

        // Monat bestimmen
        const m = dateUtils.getMonthFromRow(row, 'eigenbelege', config);
        if (!m) return;

        // Betrag extrahieren
        const amount = numberUtils.parseCurrency(row[columns.nettobetrag - 1]);
        if (amount === 0) return;

        // Kategorie extrahieren
        const category = row[columns.kategorie - 1]?.toString().trim() || '';

        // Cache-Schlüssel für Kategorie-Zuordnung
        const cacheKey = `eig_cat_${category}`;

        // Prüfen, ob Kategorie-Konfiguration im Cache
        let eigenCfg, taxType;

        if (categoryMappingCache.has(cacheKey)) {
            const cached = categoryMappingCache.get(cacheKey);
            taxType = cached.taxType;
        } else {
            eigenCfg = config.eigenbelege.categories[category] ?? {};
            taxType = eigenCfg.taxType ?? 'steuerpflichtig';
            categoryMappingCache.set(cacheKey, { taxType });
        }

        // Basierend auf Steuertyp zuordnen
        if (taxType === 'steuerfrei') {
            bwaData[m].eigenbelegeSteuerfrei += amount;
        } else {
            bwaData[m].eigenbelegeSteuerpflichtig += amount;
        }

        // BWA-Mapping verwenden wenn vorhanden (mit Cache)
        const mappingCacheKey = `eig_map_${category}`;
        let mapping;

        if (categoryMappingCache.has(mappingCacheKey)) {
            mapping = categoryMappingCache.get(mappingCacheKey);
        } else {
            mapping = config.eigenbelege.bwaMapping[category] || 'sonstigeAufwendungen';
            categoryMappingCache.set(mappingCacheKey, mapping);
        }

        bwaData[m][mapping] += amount;

    } catch (e) {
        console.error('Fehler bei der Verarbeitung eines Eigenbelegs:', e);
    }
}

/**
 * Verarbeitet Gesellschafterkonto und ordnet Positionen den BWA-Kategorien zu
 * Optimierte Version mit Cache und effizienteren Zeitstempeln
 * @param {Array} row - Zeile aus dem Gesellschafterkonto-Sheet
 * @param {Object} bwaData - BWA-Datenstruktur
 * @param {Object} config - Die Konfiguration
 */
function processGesellschafter(row, bwaData, config) {
    try {
        const columns = config.gesellschafterkonto.columns;

        // Datum aus der Zeile extrahieren und Monat ermitteln
        const zeitstempelDatum = row[columns.zeitstempel - 1];
        if (!zeitstempelDatum) return;

        // Optimierung: Datum und Jahr gecacht prüfen
        const dateCacheKey = `gesellschafter_date_${zeitstempelDatum}`;

        let d, m;
        if (categoryMappingCache.has(dateCacheKey)) {
            const cached = categoryMappingCache.get(dateCacheKey);
            d = cached.date;
            m = cached.month;
        } else {
            d = dateUtils.parseDate(zeitstempelDatum);
            if (!d) return;

            // Prüfen, ob das Jahr mit dem Konfigurationsjahr übereinstimmt
            const targetYear = config?.tax?.year || new Date().getFullYear();
            if (d.getFullYear() !== targetYear) return;

            m = d.getMonth() + 1; // Monat (1-12)
            categoryMappingCache.set(dateCacheKey, { date: d, month: m });
        }

        // Betrag extrahieren
        const amount = numberUtils.parseCurrency(row[columns.betrag - 1]);
        if (amount === 0) return;

        // Kategorie extrahieren
        const category = row[columns.kategorie - 1]?.toString().trim() || '';

        // BWA-Mapping verwenden (mit Cache)
        const mappingCacheKey = `gesell_map_${category}`;
        let mapping;

        if (categoryMappingCache.has(mappingCacheKey)) {
            mapping = categoryMappingCache.get(mappingCacheKey);
        } else {
            mapping = config.gesellschafterkonto.bwaMapping[category];
            categoryMappingCache.set(mappingCacheKey, mapping || null);
        }

        if (mapping) {
            bwaData[m][mapping] += amount;
        }
    } catch (e) {
        console.error('Fehler bei der Verarbeitung einer Gesellschafterkonto-Position:', e);
    }
}

/**
 * Verarbeitet Holding Transfers und ordnet Positionen den BWA-Kategorien zu
 * Optimierte Version mit Cache und effizienteren Lookups
 * @param {Array} row - Zeile aus dem Holding Transfers-Sheet
 * @param {Object} bwaData - BWA-Datenstruktur
 * @param {Object} config - Die Konfiguration
 */
function processHolding(row, bwaData, config) {
    try {
        const columns = config.holdingTransfers.columns;

        // Datum aus der Zeile extrahieren und Monat ermitteln
        const zeitstempelDatum = row[columns.zeitstempel - 1];
        if (!zeitstempelDatum) return;

        // Optimierung: Datum und Jahr gecacht prüfen
        const dateCacheKey = `holding_date_${zeitstempelDatum}`;

        let d, m;
        if (categoryMappingCache.has(dateCacheKey)) {
            const cached = categoryMappingCache.get(dateCacheKey);
            d = cached.date;
            m = cached.month;
        } else {
            d = dateUtils.parseDate(zeitstempelDatum);
            if (!d) return;

            // Prüfen, ob das Jahr mit dem Konfigurationsjahr übereinstimmt
            const targetYear = config?.tax?.year || new Date().getFullYear();
            if (d.getFullYear() !== targetYear) return;

            m = d.getMonth() + 1; // Monat (1-12)
            categoryMappingCache.set(dateCacheKey, { date: d, month: m });
        }

        // Betrag extrahieren
        const amount = numberUtils.parseCurrency(row[columns.betrag - 1]);
        if (amount === 0) return;

        // Kategorie extrahieren
        const category = row[columns.art - 1]?.toString().trim() || '';

        // BWA-Mapping verwenden (mit Cache)
        const mappingCacheKey = `holding_map_${category}`;
        let mapping;

        if (categoryMappingCache.has(mappingCacheKey)) {
            mapping = categoryMappingCache.get(mappingCacheKey);
        } else {
            mapping = config.holdingTransfers.bwaMapping[category];
            categoryMappingCache.set(mappingCacheKey, mapping || null);
        }

        if (mapping) {
            bwaData[m][mapping] += amount;
        }
    } catch (e) {
        console.error('Fehler bei der Verarbeitung eines Holding Transfers:', e);
    }
}

/**
 * Berechnet Gruppensummen und abgeleitete Werte für alle Monate
 * Optimierte Version mit effizienter Berechnungslogik
 * @param {Object} bwaData - BWA-Datenstruktur mit Rohdaten
 * @param {Object} config - Die Konfiguration
 */
function calculateAggregates(bwaData, config) {
    for (let m = 1; m <= 12; m++) {
        const d = bwaData[m];

        // Optimierung: Berechne Summen in einem Durchgang und verwende
        // Zwischenspeicher für mehrfach verwendete Summen

        // Erlöse
        d.gesamtErloese = numberUtils.round(
            d.umsatzerloese + d.provisionserloese + d.steuerfreieInlandEinnahmen +
            d.steuerfreieAuslandEinnahmen + d.sonstigeErtraege + d.vermietung +
            d.zuschuesse + d.waehrungsgewinne + d.anlagenabgaenge,
            2,
        );

        // Materialkosten
        d.gesamtWareneinsatz = numberUtils.round(
            d.wareneinsatz + d.fremdleistungen + d.rohHilfsBetriebsstoffe,
            2,
        );

        // Betriebsausgaben
        d.gesamtBetriebsausgaben = numberUtils.round(
            d.bruttoLoehne + d.sozialeAbgaben + d.sonstigePersonalkosten +
            d.werbungMarketing + d.reisekosten + d.versicherungen + d.telefonInternet +
            d.buerokosten + d.fortbildungskosten + d.kfzKosten + d.sonstigeAufwendungen,
            2,
        );

        // Abschreibungen & Zinsen
        d.gesamtAbschreibungenZinsen = numberUtils.round(
            d.abschreibungenMaschinen + d.abschreibungenBueromaterial +
            d.abschreibungenImmateriell + d.zinsenBank + d.zinsenGesellschafter +
            d.leasingkosten,
            2,
        );

        // Besondere Posten
        d.gesamtBesonderePosten = numberUtils.round(
            d.eigenkapitalveraenderungen + d.gesellschafterdarlehen + d.ausschuettungen,
            2,
        );

        // Rückstellungen
        d.gesamtRueckstellungenTransfers = numberUtils.round(
            d.steuerrueckstellungen + d.rueckstellungenSonstige,
            2,
        );

        // EBIT (Zwischenspeicher für Steuern)
        d.ebit = numberUtils.round(
            d.gesamtErloese - (d.gesamtWareneinsatz + d.gesamtBetriebsausgaben +
                d.gesamtAbschreibungenZinsen + d.gesamtBesonderePosten),
            2,
        );

        // Steuern berechnen (mit optimierten Zugriff auf Config)
        const taxConfig = config.tax.isHolding ? config.tax.holding : config.tax.operative;

        // Für Holdings gelten spezielle Steuersätze wegen Beteiligungsprivileg
        const steuerfaktor = config.tax.isHolding
            ? taxConfig.gewinnUebertragSteuerpflichtig / 100
            : 1;

        // Steuerberechnung (angepasst für Klarheit und Optimierung)
        d.gewerbesteuer = numberUtils.round(d.ebit * (taxConfig.gewerbesteuer / 10000) * steuerfaktor, 2);
        d.koerperschaftsteuer = numberUtils.round(d.ebit * (taxConfig.koerperschaftsteuer / 100) * steuerfaktor, 2);
        d.solidaritaetszuschlag = numberUtils.round(d.koerperschaftsteuer * (taxConfig.solidaritaetszuschlag / 100), 2);

        // Gesamte Steuerlast
        d.steuerlast = numberUtils.round(
            d.koerperschaftsteuer + d.solidaritaetszuschlag + d.gewerbesteuer,
            2,
        );

        // Gewinn nach Steuern
        d.gewinnNachSteuern = numberUtils.round(d.ebit - d.steuerlast, 2);
    }
}

export default {
    processRevenue,
    processExpense,
    processEigen,
    processGesellschafter,
    processHolding,
    calculateAggregates,
};