// modules/bwaModule/calculator.js
import dateUtils from '../../utils/dateUtils.js';
import numberUtils from '../../utils/numberUtils.js';
import stringUtils from '../../utils/stringUtils.js';

// Cache for frequently used category mappings
const categoryMappingCache = new Map();

/**
 * Calculates group totals and derived values for all months
 * Optimized version with efficient calculation logic
 * @param {Object} bwaData - BWA data structure with raw data
 * @param {Object} config - Configuration
 */
function calculateAggregates(bwaData, config) {
    for (let m = 1; m <= 12; m++) {
        const d = bwaData[m];

        // Calculate sums in one pass using intermediate variables for frequently used sums

        // Revenues
        d.gesamtErloese = numberUtils.round(
            d.umsatzerloese + d.provisionserloese + d.steuerfreieInlandEinnahmen +
            d.steuerfreieAuslandEinnahmen + d.sonstigeErtraege + d.vermietung +
            d.zuschuesse + d.waehrungsgewinne + d.anlagenabgaenge +
            d.steuerlicheKorrekturen,
            2,
        );

        // Material costs
        d.gesamtWareneinsatz = numberUtils.round(
            d.wareneinsatz + d.fremdleistungen + d.rohHilfsBetriebsstoffe,
            2,
        );

        // Operating expenses
        d.gesamtBetriebsausgaben = numberUtils.round(
            d.bruttoLoehne + d.sozialeAbgaben + d.sonstigePersonalkosten +
            d.werbungMarketing + d.reisekosten + d.versicherungen + d.telefonInternet +
            d.buerokosten + d.fortbildungskosten + d.kfzKosten + d.mieteNebenkosten +
            d.sonstigeAufwendungen,
            2,
        );

        // Depreciation & interest
        d.gesamtAbschreibungenZinsen = numberUtils.round(
            d.abschreibungenMaschinen + d.abschreibungenBueromaterial +
            d.abschreibungenImmateriell + d.zinsenBank + d.zinsenGesellschafter +
            d.leasingkosten,
            2,
        );

        // EBIT (Earnings Before Interest and Taxes)
        d.ebit = numberUtils.round(
            d.gesamtErloese - (d.gesamtWareneinsatz + d.gesamtBetriebsausgaben +
                d.gesamtAbschreibungenZinsen),
            2,
        );

        // Calculate profit after taxes (may be from other sources or tax estimation)
        // This is a simplified approach - in reality tax calculation would be more complex
        const taxRate = 0.3; // Approximate combined tax rate
        d.gewinnNachSteuern = numberUtils.round(d.ebit * (1 - taxRate), 2);
    }
}

/**
 * Processes revenue and maps to BWA categories
 * Optimized for cash-based accounting and foreign country handling
 * @param {Array} row - Row from revenue sheet
 * @param {Object} bwaData - BWA data structure
 * @param {Object} config - Configuration
 */
function processRevenue(row, bwaData, config) {
    try {
        const columns = config.einnahmen.columns;

        // Check for valid payment date (cash-based accounting)
        const paymentDate = dateUtils.parseDate(row[columns.zahlungsdatum - 1]);
        if (!paymentDate || paymentDate > new Date()) return;

        // Check if payment date is in target year
        const targetYear = config?.tax?.year || new Date().getFullYear();
        if (paymentDate.getFullYear() !== targetYear) return;

        // Get payment month for cash-based accounting
        const paymentMonth = paymentDate.getMonth() + 1; // 1-12

        // Calculate correct amount based on payment proportion
        const nettobetrag = numberUtils.parseCurrency(row[columns.nettobetrag - 1]);
        const bruttobetrag = numberUtils.parseCurrency(row[columns.bruttoBetrag - 1]);
        const bezahlt = numberUtils.parseCurrency(row[columns.bezahlt - 1]);

        // Skip if nothing paid (for positive amounts) or nothing refunded (for negative amounts)
        if (Math.abs(bezahlt) <= 0) return;

        // Calculate paid proportion correctly for both positive and negative amounts
        const bezahltAnteil = Math.abs(bezahlt) / (Math.abs(bruttobetrag) || 1);
        const amount = nettobetrag * Math.min(bezahltAnteil, 1); // Preserves sign for credits

        if (Math.abs(amount) === 0) return;

        // Extract category
        const category = row[columns.kategorie - 1]?.toString().trim() || '';

        // Check foreign status
        const auslandWert = columns.ausland && row[columns.ausland - 1] ?
            row[columns.ausland - 1].toString().trim() : 'Inland';
        const isEuAusland = auslandWert === 'EU-Ausland';
        const isNichtEuAusland = auslandWert === 'Nicht-EU-Ausland';
        const isAusland = isEuAusland || isNichtEuAusland;

        // Skip non-business categories
        const nonBusinessCategories = new Set(['Gewinnvortrag', 'Verlustvortrag']);
        if (nonBusinessCategories.has(category)) return;

        // Special handling for foreign income
        if (isEuAusland) {
            bwaData[paymentMonth].innergemeinschaftlicheLieferungen += amount;
            return;
        } else if (isNichtEuAusland) {
            bwaData[paymentMonth].steuerfreieAuslandEinnahmen += amount;
            return;
        }

        // Cache key for category mapping
        const cacheKey = `rev_cat_${category}_${auslandWert}`;

        // Check if mapping is cached
        if (categoryMappingCache.has(cacheKey)) {
            const target = categoryMappingCache.get(cacheKey);
            bwaData[paymentMonth][target] += amount;
            return;
        }

        // Direct mappings for common categories
        const directMappings = {
            'Erlöse aus Lieferungen und Leistungen': 'umsatzerloese',
            'Provisionserlöse': 'provisionserloese',
            'Sonstige betriebliche Erträge': 'sonstigeErtraege',
            'Erträge aus Vermietung/Verpachtung': 'vermietung',
            'Erträge aus Zuschüssen': 'zuschuesse',
            'Erträge aus Währungsgewinnen': 'waehrungsgewinne',
            'Erträge aus Anlagenabgängen': 'anlagenabgaenge',
            'Gutschriften (Warenrückgabe)': 'umsatzerloese',
            'Darlehen': 'sonstigeErtraege',
            'Zinsen': 'sonstigeErtraege',
            'Umsatzsteuererstattungen': 'steuerlicheKorrekturen',
        };

        // Check for mapping in config
        const categoryConfig = config.einnahmen.categories[category];
        let mapping = categoryConfig?.bwaMapping;

        // Use direct mapping if available
        if (!mapping && directMappings[category]) {
            mapping = directMappings[category];
        }

        // Apply mapping if found
        if (mapping) {
            categoryMappingCache.set(cacheKey, mapping);
            bwaData[paymentMonth][mapping] += amount;
            return;
        }

        // Fallback categorization based on VAT rate
        if (numberUtils.parseMwstRate(row[columns.mwstSatz - 1], config.tax.defaultMwst) === 0) {
            categoryMappingCache.set(cacheKey, 'steuerfreieInlandEinnahmen');
            bwaData[paymentMonth].steuerfreieInlandEinnahmen += amount;
        } else {
            categoryMappingCache.set(cacheKey, 'umsatzerloese');
            bwaData[paymentMonth].umsatzerloese += amount;
        }
    } catch (e) {
        console.error('Error processing revenue:', e);
    }
}

/**
 * Processes expenses and maps to BWA categories
 * Optimized for cash-based accounting and foreign country handling
 * @param {Array} row - Row from expense sheet
 * @param {Object} bwaData - BWA data structure
 * @param {Object} config - Configuration
 */
function processExpense(row, bwaData, config) {
    try {
        const columns = config.ausgaben.columns;

        // Check for valid payment date (cash-based accounting)
        const paymentDate = dateUtils.parseDate(row[columns.zahlungsdatum - 1]);
        if (!paymentDate || paymentDate > new Date()) return;

        // Check if payment date is in target year
        const targetYear = config?.tax?.year || new Date().getFullYear();
        if (paymentDate.getFullYear() !== targetYear) return;

        // Get payment month for cash-based accounting
        const paymentMonth = paymentDate.getMonth() + 1; // 1-12

        // Calculate correct amount based on payment proportion
        const nettobetrag = numberUtils.parseCurrency(row[columns.nettobetrag - 1]);
        const bruttobetrag = numberUtils.parseCurrency(row[columns.bruttoBetrag - 1]);
        const bezahlt = numberUtils.parseCurrency(row[columns.bezahlt - 1]);

        // Skip if nothing paid
        if (Math.abs(bezahlt) <= 0) return;

        // Calculate paid proportion correctly for both positive and negative amounts
        const bezahltAnteil = Math.abs(bezahlt) / (Math.abs(bruttobetrag) || 1);
        const amount = nettobetrag * Math.min(bezahltAnteil, 1); // Preserves sign for credits

        if (Math.abs(amount) === 0) return;

        // Extract category
        const category = row[columns.kategorie - 1]?.toString().trim() || '';

        // Check foreign status
        const auslandWert = columns.ausland && row[columns.ausland - 1] ?
            row[columns.ausland - 1].toString().trim() : 'Inland';
        const isAusland = auslandWert === 'EU-Ausland' || auslandWert === 'Nicht-EU-Ausland';

        // Skip non-business categories
        const nonBusinessCategories = new Set([
            'Privatentnahme', 'Privateinlage', 'Holding Transfers',
            'Gewinnvortrag', 'Verlustvortrag',
        ]);
        if (nonBusinessCategories.has(category)) return;

        // Track non-deductible VAT for entertainment expenses
        if (category === 'Bewirtung') {
            const mwstRate = numberUtils.parseMwstRate(row[columns.mwstSatz - 1], config.tax.defaultMwst);
            const mwstAmount = amount * (mwstRate / 100);
            // Only 30% of VAT is non-deductible for entertainment
            const nonDeductible = mwstAmount * 0.3;
            bwaData[paymentMonth].nichtAbzugsfaehigeVSt += nonDeductible;
        }

        // Cache key for category mapping
        const cacheKey = `exp_cat_${category}_${auslandWert}`;

        // Check if mapping is cached
        if (categoryMappingCache.has(cacheKey)) {
            const target = categoryMappingCache.get(cacheKey);
            bwaData[paymentMonth][target] += amount;
            return;
        }

        // Direct mappings for common categories
        const directMappings = {
            'Wareneinsatz': 'wareneinsatz',
            'Bezogene Leistungen': 'fremdleistungen',
            'Roh-, Hilfs- & Betriebsstoffe': 'rohHilfsBetriebsstoffe',
            'Bruttolöhne & Gehälter': 'bruttoLoehne',
            'Soziale Abgaben & Arbeitgeberanteile': 'sozialeAbgaben',
            'Sonstige Personalkosten': 'sonstigePersonalkosten',
            'Marketing & Werbung': 'werbungMarketing',
            'Werbung & Marketing': 'werbungMarketing',
            'Reisekosten': 'reisekosten',
            'Versicherungen': 'versicherungen',
            'Telefon & Internet': 'telefonInternet',
            'Bürokosten': 'buerokosten',
            'Fortbildungskosten': 'fortbildungskosten',
            'Kfz-Kosten': 'kfzKosten',
            'Abschreibungen Maschinen': 'abschreibungenMaschinen',
            'Abschreibungen Büroausstattung': 'abschreibungenBueromaterial',
            'Abschreibungen immaterielle Wirtschaftsgüter': 'abschreibungenImmateriell',
            'Zinsen auf Bankdarlehen': 'zinsenBank',
            'Zinsen auf Gesellschafterdarlehen': 'zinsenGesellschafter',
            'Leasingkosten': 'leasingkosten',
            'Sonstige betriebliche Aufwendungen': 'sonstigeAufwendungen',
            'Provisionszahlungen': 'fremdleistungen',
            'IT-Kosten': 'sonstigeAufwendungen',
            'Miete': 'mieteNebenkosten',
            'Nebenkosten': 'mieteNebenkosten',
            'Betriebskosten': 'sonstigeAufwendungen',
            'Porto': 'buerokosten',
            'Bewirtung': 'sonstigeAufwendungen',
            'Bankgebühren': 'sonstigeAufwendungen',
        };

        // Use direct mapping if available
        if (directMappings[category]) {
            categoryMappingCache.set(cacheKey, directMappings[category]);
            bwaData[paymentMonth][directMappings[category]] += amount;
            return;
        }

        // Check for mapping in config
        const categoryConfig = config.ausgaben.categories[category];
        const mapping = categoryConfig?.bwaMapping;

        if (mapping) {
            categoryMappingCache.set(cacheKey, mapping);
            bwaData[paymentMonth][mapping] += amount;
            return;
        }

        // Fallback: categorize as other expenses
        categoryMappingCache.set(cacheKey, 'sonstigeAufwendungen');
        bwaData[paymentMonth].sonstigeAufwendungen += amount;
    } catch (e) {
        console.error('Error processing expense:', e);
    }
}

/**
 * Processes own receipts and maps to BWA categories
 * Optimized for cash-based accounting
 * @param {Array} row - Row from own receipts sheet
 * @param {Object} bwaData - BWA data structure
 * @param {Object} config - Configuration
 */
function processEigen(row, bwaData, config) {
    try {
        const columns = config.eigenbelege.columns;

        // Check for valid payment date (cash-based accounting)
        const paymentDate = dateUtils.parseDate(row[columns.zahlungsdatum - 1]);
        if (!paymentDate || paymentDate > new Date()) return;

        // Check if payment date is in target year
        const targetYear = config?.tax?.year || new Date().getFullYear();
        if (paymentDate.getFullYear() !== targetYear) return;

        // Get payment month for cash-based accounting
        const paymentMonth = paymentDate.getMonth() + 1; // 1-12

        // Calculate correct amount based on payment proportion
        const nettobetrag = numberUtils.parseCurrency(row[columns.nettobetrag - 1]);
        const bruttobetrag = numberUtils.parseCurrency(row[columns.bruttoBetrag - 1]);
        const bezahlt = numberUtils.parseCurrency(row[columns.bezahlt - 1]);

        // Skip if nothing paid
        if (Math.abs(bezahlt) <= 0) return;

        // Calculate paid proportion correctly for both positive and negative amounts
        const bezahltAnteil = Math.abs(bezahlt) / (Math.abs(bruttobetrag) || 1);
        const amount = nettobetrag * Math.min(bezahltAnteil, 1); // Preserves sign for credits

        if (Math.abs(amount) === 0) return;

        // Extract category
        const category = row[columns.kategorie - 1]?.toString().trim() || '';

        // Internal tracking of own receipts type
        if (category === 'Bewirtung') {
            // Track non-deductible VAT for entertainment expenses
            const mwstRate = numberUtils.parseMwstRate(row[columns.mwstSatz - 1], config.tax.defaultMwst);
            const mwstAmount = amount * (mwstRate / 100);
            // Only 30% of VAT is non-deductible for entertainment
            const nonDeductible = mwstAmount * 0.3;
            bwaData[paymentMonth].nichtAbzugsfaehigeVSt += nonDeductible;
        }

        // Use BWA mapping with cache
        const mappingCacheKey = `eig_map_${category}`;
        let mapping;

        if (categoryMappingCache.has(mappingCacheKey)) {
            mapping = categoryMappingCache.get(mappingCacheKey);
        } else {
            const categoryConfig = config.eigenbelege.categories[category];
            mapping = categoryConfig?.bwaMapping || 'sonstigeAufwendungen';
            categoryMappingCache.set(mappingCacheKey, mapping);
        }

        // Map to appropriate BWA category
        bwaData[paymentMonth][mapping] += amount;
    } catch (e) {
        console.error('Error processing own receipt:', e);
    }
}

/**
 * Processes shareholder account positions
 * (Only process items relevant for BWA, not balance sheet items)
 * @param {Array} row - Row from shareholder account sheet
 * @param {Object} bwaData - BWA data structure
 * @param {Object} config - Configuration
 */
function processGesellschafter(row, bwaData, config) {
    try {
        const columns = config.gesellschafterkonto.columns;

        // Check for valid payment date (cash-based accounting)
        const paymentDate = dateUtils.parseDate(row[columns.zahlungsdatum - 1]);
        if (!paymentDate || paymentDate > new Date()) return;

        // Check if payment date is in target year
        const targetYear = config?.tax?.year || new Date().getFullYear();
        if (paymentDate.getFullYear() !== targetYear) return;

        // Get payment month for cash-based accounting
        const paymentMonth = paymentDate.getMonth() + 1; // 1-12

        // Extract amount and category
        const amount = numberUtils.parseCurrency(row[columns.betrag - 1]);
        if (Math.abs(amount) === 0) return;

        const category = row[columns.kategorie - 1]?.toString().trim() || '';

        // Skip balance sheet items that don't belong in BWA
        const skipCategories = new Set([
            'Gesellschafterdarlehen',
            'Ausschüttungen',
            'Kapitalrückführung',
            'Privatentnahme',
            'Privateinlage',
        ]);

        if (skipCategories.has(category)) {
            return;
        }

        // Process only if it's an actual operating expense or income
        // For example, interest payments to shareholders would still be included
        if (category === 'Zinsen auf Gesellschafterdarlehen') {
            bwaData[paymentMonth].zinsenGesellschafter += Math.abs(amount);
        } else if (category === 'Sonstige betriebliche Aufwendungen') {
            bwaData[paymentMonth].sonstigeAufwendungen += Math.abs(amount);
        }
    } catch (e) {
        console.error('Error processing shareholder position:', e);
    }
}

/**
 * Processes holding transfers
 * (Only process items relevant for BWA, not balance sheet items)
 * @param {Array} row - Row from holding transfers sheet
 * @param {Object} bwaData - BWA data structure
 * @param {Object} config - Configuration
 */
function processHolding(row, bwaData, config) {
    // Most holding transfers are balance sheet items and don't belong in BWA
    // This is kept as a placeholder for scenarios where operating expenses
    // might be recorded in the holding transfers sheet
    try {
        const columns = config.holdingTransfers.columns;

        // Check for valid payment date (cash-based accounting)
        const paymentDate = dateUtils.parseDate(row[columns.zahlungsdatum - 1]);
        if (!paymentDate || paymentDate > new Date()) return;

        // Check if payment date is in target year
        const targetYear = config?.tax?.year || new Date().getFullYear();
        if (paymentDate.getFullYear() !== targetYear) return;

        // Extract category
        const category = row[columns.kategorie - 1]?.toString().trim() || '';

        // Generally, holding transfers don't appear in BWA as they're balance sheet items
        // This is a stub for potential exceptions
    } catch (e) {
        console.error('Error processing holding transfer:', e);
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