// modules/bwaModule/calculator.js
import dateUtils from '../../utils/dateUtils.js';
import numberUtils from '../../utils/numberUtils.js';
import stringUtils from '../../utils/stringUtils.js';

// Cache for frequently used category mappings
const categoryMappingCache = new Map();

/**
 * Processes revenue and maps to BWA categories
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

        // Skip if nothing paid or refunded
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

        // Track VAT
        if (!isAusland) {
            const mwstRate = numberUtils.parseMwstRate(row[columns.mwstSatz - 1], config.tax.defaultMwst);
            const mwstAmount = amount * (mwstRate / 100);
            bwaData[paymentMonth].umsatzsteuer += mwstAmount;
        }

        // Skip non-business categories
        const nonBusinessCategories = new Set(['Gewinnvortrag', 'Verlustvortrag']);
        if (nonBusinessCategories.has(category)) return;

        // Special handling for foreign income
        if (isAusland) {
            if (isEuAusland) {
                bwaData[paymentMonth].innergemeinschaftlicheLieferungen += amount;
            } else {
                bwaData[paymentMonth].steuerfreieAuslandErloese += amount;
            }
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

        // DATEV-compliant mappings
        const datevMappings = {
            'Erlöse aus Lieferungen und Leistungen': 'erloeseLieferungenLeistungen',
            'Provisionserlöse': 'provisionserloese',
            'Steuerfreie Inland-Einnahmen': 'steuerfreieInlandErloese',
            'Sonstige betriebliche Erträge': 'sonstigeBetrieblicheErtraege',
            'Erträge aus Vermietung/Verpachtung': 'ertraegeVermietungVerpachtung',
            'Erträge aus Zuschüssen': 'ertraegeZuschuesse',
            'Erträge aus Währungsgewinnen': 'ertraegeKursgewinne',
            'Erträge aus Anlagenabgängen': 'ertraegeAnlagenabgaenge',
            'Gutschriften (Warenrückgabe)': 'erloeseLieferungenLeistungen',
        };

        // Check for mapping in config
        const categoryConfig = config.einnahmen.categories[category];
        let mapping = null;

        // Use DATEV mapping if available
        if (datevMappings[category]) {
            mapping = datevMappings[category];
        }
        // If not found in DATEV mappings, try to get from category config
        else if (categoryConfig?.bwaMapping) {
            // Map legacy bwaMapping to new DATEV fields
            const legacyToDatev = {
                'umsatzerloese': 'erloeseLieferungenLeistungen',
                'provisionserloese': 'provisionserloese',
                'steuerfreieInlandEinnahmen': 'steuerfreieInlandErloese',
                'steuerfreieAuslandEinnahmen': 'steuerfreieAuslandErloese',
                'sonstigeErtraege': 'sonstigeBetrieblicheErtraege',
                'vermietung': 'ertraegeVermietungVerpachtung',
                'zuschuesse': 'ertraegeZuschuesse',
                'waehrungsgewinne': 'ertraegeKursgewinne',
                'anlagenabgaenge': 'ertraegeAnlagenabgaenge',
            };

            mapping = legacyToDatev[categoryConfig.bwaMapping] || 'sonstigeBetrieblicheErtraege';
        }

        // Apply mapping if found
        if (mapping) {
            categoryMappingCache.set(cacheKey, mapping);
            bwaData[paymentMonth][mapping] += amount;
            return;
        }

        // Fallback categorization
        if (numberUtils.parseMwstRate(row[columns.mwstSatz - 1], config.tax.defaultMwst) === 0) {
            categoryMappingCache.set(cacheKey, 'steuerfreieInlandErloese');
            bwaData[paymentMonth].steuerfreieInlandErloese += amount;
        } else {
            categoryMappingCache.set(cacheKey, 'erloeseLieferungenLeistungen');
            bwaData[paymentMonth].erloeseLieferungenLeistungen += amount;
        }
    } catch (e) {
        console.error('Error processing revenue:', e);
    }
}

/**
 * Processes expenses and maps to BWA categories
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
        const isEuAusland = auslandWert === 'EU-Ausland';
        const isNichtEuAusland = auslandWert === 'Nicht-EU-Ausland';
        const isAusland = isEuAusland || isNichtEuAusland;

        // Track VAT
        if (!isAusland) {
            const mwstRate = numberUtils.parseMwstRate(row[columns.mwstSatz - 1], config.tax.defaultMwst);
            const mwstAmount = amount * (mwstRate / 100);
            bwaData[paymentMonth].vorsteuer += mwstAmount;

            // Track non-deductible VAT (for entertainment expenses)
            if (category === 'Bewirtung') {
                // Only 70% of VAT is deductible for entertainment
                const nonDeductible = mwstAmount * 0.3;
                bwaData[paymentMonth].nichtAbzugsfaehigeVSt += nonDeductible;
                bwaData[paymentMonth].vorsteuer -= nonDeductible; // Reduce deductible portion
            }
        }

        // Skip non-business categories
        const nonBusinessCategories = new Set([
            'Privatentnahme', 'Privateinlage', 'Holding Transfers',
            'Gewinnvortrag', 'Verlustvortrag',
        ]);
        if (nonBusinessCategories.has(category)) return;

        // Cache key for category mapping
        const cacheKey = `exp_cat_${category}_${auslandWert}`;

        // Check if mapping is cached
        if (categoryMappingCache.has(cacheKey)) {
            const target = categoryMappingCache.get(cacheKey);
            bwaData[paymentMonth][target] += amount;
            return;
        }

        // DATEV-compliant mappings
        const datevMappings = {
            // Material
            'Wareneinsatz': 'wareneinsatz',
            'Bezogene Leistungen': 'bezogeneLeistungen',
            'Roh-, Hilfs- & Betriebsstoffe': 'rohHilfsBetriebsstoffe',

            // Personal
            'Bruttolöhne & Gehälter': 'loehneGehaelter',
            'Soziale Abgaben & Arbeitgeberanteile': 'sozialeAbgaben',
            'Sonstige Personalkosten': 'sonstigePersonalkosten',

            // Sonstige Aufwendungen
            'Provisionszahlungen': 'provisionszahlungenDritte',
            'IT-Kosten': 'itKosten',
            'Miete': 'mieteLeasing',
            'Nebenkosten': 'mieteLeasing',
            'Marketing & Werbung': 'werbungMarketing',
            'Reisekosten': 'reisekosten',
            'Versicherungen': 'versicherungen',
            'Telefon & Internet': 'telefonInternet',
            'Bürokosten': 'buerokosten',
            'Fortbildungskosten': 'fortbildungskosten',
            'Kfz-Kosten': 'kfzKosten',
            'Beiträge & Abgaben': 'beitraegeAbgaben',
            'Bewirtung': 'bewirtungskosten',
            'Sonstige betriebliche Aufwendungen': 'sonstigeBetrieblicheAufwendungen',

            // Abschreibungen & Zinsen
            'Abschreibungen Maschinen': 'abschreibungenSachanlagen',
            'Abschreibungen Büroausstattung': 'abschreibungenSachanlagen',
            'Abschreibungen immaterielle Wirtschaftsgüter': 'abschreibungenImmaterielleVG',
            'Zinsen auf Bankdarlehen': 'zinsenBankdarlehen',
            'Zinsen auf Gesellschafterdarlehen': 'zinsenGesellschafterdarlehen',
            'Leasingkosten': 'leasingzinsen',
        };

        // Use DATEV mapping if available
        if (datevMappings[category]) {
            categoryMappingCache.set(cacheKey, datevMappings[category]);
            bwaData[paymentMonth][datevMappings[category]] += amount;
            return;
        }

        // Check for mapping in config and map to new DATEV structure
        const categoryConfig = config.ausgaben.categories[category];
        let mapping = null;

        if (categoryConfig?.bwaMapping) {
            // Map legacy bwaMapping to new DATEV fields
            const legacyToDatev = {
                'wareneinsatz': 'wareneinsatz',
                'fremdleistungen': 'bezogeneLeistungen',
                'rohHilfsBetriebsstoffe': 'rohHilfsBetriebsstoffe',
                'bruttoLoehne': 'loehneGehaelter',
                'sozialeAbgaben': 'sozialeAbgaben',
                'sonstigePersonalkosten': 'sonstigePersonalkosten',
                'provisionszahlungen': 'provisionszahlungenDritte',
                'itKosten': 'itKosten',
                'werbungMarketing': 'werbungMarketing',
                'reisekosten': 'reisekosten',
                'versicherungen': 'versicherungen',
                'telefonInternet': 'telefonInternet',
                'buerokosten': 'buerokosten',
                'fortbildungskosten': 'fortbildungskosten',
                'kfzKosten': 'kfzKosten',
                'mieteNebenkosten': 'mieteLeasing',
                'sonstigeAufwendungen': 'sonstigeBetrieblicheAufwendungen',
                'abschreibungenMaschinen': 'abschreibungenSachanlagen',
                'abschreibungenBueromaterial': 'abschreibungenSachanlagen',
                'abschreibungenImmateriell': 'abschreibungenImmaterielleVG',
                'zinsenBank': 'zinsenBankdarlehen',
                'zinsenGesellschafter': 'zinsenGesellschafterdarlehen',
                'leasingkosten': 'leasingzinsen',
                'beitraegeAbgaben': 'beitraegeAbgaben',
                'bewirtungskosten': 'bewirtungskosten',
            };

            mapping = legacyToDatev[categoryConfig.bwaMapping];
        }

        if (mapping) {
            categoryMappingCache.set(cacheKey, mapping);
            bwaData[paymentMonth][mapping] += amount;
            return;
        }

        // Fallback: categorize as other expenses
        categoryMappingCache.set(cacheKey, 'sonstigeBetrieblicheAufwendungen');
        bwaData[paymentMonth].sonstigeBetrieblicheAufwendungen += amount;
    } catch (e) {
        console.error('Error processing expense:', e);
    }
}

/**
 * Processes own receipts and maps to BWA categories
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

        // Check foreign status
        const auslandWert = columns.ausland && row[columns.ausland - 1] ?
            row[columns.ausland - 1].toString().trim() : 'Inland';

        // DATEV-compliant mappings
        const datevMappings = {
            'Kleidung': 'sonstigeBetrieblicheAufwendungen',
            'Trinkgeld': 'sonstigeBetrieblicheAufwendungen',
            'Private Vorauslage': 'sonstigeBetrieblicheAufwendungen',
            'Bürokosten': 'buerokosten',
            'Reisekosten': 'reisekosten',
            'Bewirtung': 'bewirtungskosten',
            'Sonstiges': 'sonstigeBetrieblicheAufwendungen',
        };

        let mapping;
        const mappingCacheKey = `eig_map_${category}_${auslandWert}`;

        if (categoryMappingCache.has(mappingCacheKey)) {
            mapping = categoryMappingCache.get(mappingCacheKey);
        } else if (datevMappings[category]) {
            mapping = datevMappings[category];
        } else {
            const categoryConfig = config.eigenbelege.categories[category];

            if (categoryConfig?.bwaMapping) {
                // Map legacy bwaMapping to new DATEV fields
                const legacyToDatev = {
                    'buerokosten': 'buerokosten',
                    'reisekosten': 'reisekosten',
                    'sonstigeAufwendungen': 'sonstigeBetrieblicheAufwendungen',
                    'bewirtungskosten': 'bewirtungskosten',
                };

                mapping = legacyToDatev[categoryConfig.bwaMapping] || 'sonstigeBetrieblicheAufwendungen';
            } else {
                mapping = 'sonstigeBetrieblicheAufwendungen';
            }

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

        // Shareholder transactions typically appear in balance sheet, not BWA
    } catch (e) {
        console.error('Error processing shareholder position:', e);
    }
}

/**
 * Processes holding transfers
 * @param {Array} row - Row from holding transfers sheet
 * @param {Object} bwaData - BWA data structure
 * @param {Object} config - Configuration
 */
function processHolding(row, bwaData, config) {
    try {
        const columns = config.holdingTransfers.columns;

        // Check for valid payment date (cash-based accounting)
        const paymentDate = dateUtils.parseDate(row[columns.zahlungsdatum - 1]);
        if (!paymentDate || paymentDate > new Date()) return;

        // Check if payment date is in target year
        const targetYear = config?.tax?.year || new Date().getFullYear();
        if (paymentDate.getFullYear() !== targetYear) return;

        // Holding transfers typically appear in balance sheet, not BWA
    } catch (e) {
        console.error('Error processing holding transfer:', e);
    }
}

/**
 * Calculates group totals and derived values for all months
 * @param {Object} bwaData - BWA data structure with raw data
 * @param {Object} config - Configuration
 */
function calculateAggregates(bwaData, config) {
    for (let m = 1; m <= 12; m++) {
        const d = bwaData[m];

        // 1. Betriebserlöse (Einnahmen)
        d.betriebserloese_gesamt = numberUtils.round(
            d.erloeseLieferungenLeistungen +
            d.provisionserloese +
            d.steuerfreieInlandErloese +
            d.steuerfreieAuslandErloese +
            d.innergemeinschaftlicheLieferungen +
            d.sonstigeBetrieblicheErtraege +
            d.ertraegeVermietungVerpachtung +
            d.ertraegeZuschuesse +
            d.ertraegeKursgewinne +
            d.ertraegeAnlagenabgaenge,
            2,
        );

        // 2. Materialaufwand & Wareneinsatz
        d.materialaufwand_gesamt = numberUtils.round(
            d.wareneinsatz +
            d.bezogeneLeistungen +
            d.rohHilfsBetriebsstoffe,
            2,
        );

        // 3. Personalaufwand
        d.personalaufwand_gesamt = numberUtils.round(
            d.loehneGehaelter +
            d.sozialeAbgaben +
            d.sonstigePersonalkosten,
            2,
        );

        // 4. Sonstige betriebliche Aufwendungen
        d.sonstigeAufwendungen_gesamt = numberUtils.round(
            d.provisionszahlungenDritte +
            d.itKosten +
            d.mieteLeasing +
            d.werbungMarketing +
            d.reisekosten +
            d.versicherungen +
            d.telefonInternet +
            d.buerokosten +
            d.fortbildungskosten +
            d.kfzKosten +
            d.beitraegeAbgaben +
            d.bewirtungskosten +
            d.sonstigeBetrieblicheAufwendungen,
            2,
        );

        // 5. Abschreibungen und Zinsen
        d.abschreibungenZinsen_gesamt = numberUtils.round(
            d.abschreibungenSachanlagen +
            d.abschreibungenImmaterielleVG +
            d.zinsenBankdarlehen +
            d.zinsenGesellschafterdarlehen +
            d.leasingzinsen,
            2,
        );

        // 6. Betriebsergebnis (EBIT)
        d.ebit = numberUtils.round(
            d.betriebserloese_gesamt -
            (d.materialaufwand_gesamt +
                d.personalaufwand_gesamt +
                d.sonstigeAufwendungen_gesamt +
                d.abschreibungenZinsen_gesamt),
            2,
        );

        // Calculate taxes
        const taxConfig = config.tax.isHolding ? config.tax.holding : config.tax.operative;
        const steuerfaktor = config.tax.isHolding ? taxConfig.gewinnUebertragSteuerpflichtig / 100 : 1;

        // Tax calculations
        d.gewerbesteuer = Math.max(0, numberUtils.round(d.ebit * (taxConfig.gewerbesteuer / 10000) * steuerfaktor, 2));
        d.koerperschaftsteuer = Math.max(0, numberUtils.round(d.ebit * (taxConfig.koerperschaftsteuer / 100) * steuerfaktor, 2));
        d.solidaritaetszuschlag = Math.max(0, numberUtils.round(d.koerperschaftsteuer * (taxConfig.solidaritaetszuschlag / 100), 2));

        // Total tax burden
        d.steuerlast = numberUtils.round(
            d.koerperschaftsteuer + d.solidaritaetszuschlag + d.gewerbesteuer,
            2,
        );

        // 7. Jahresergebnis (profit after taxes)
        d.jahresergebnis = numberUtils.round(d.ebit - d.steuerlast, 2);
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