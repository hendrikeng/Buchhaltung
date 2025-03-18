// src/modules/refreshModule/refreshUtils.js
import dateUtils from '../../utils/dateUtils.js';

/**
 * Helper function to determine if two reference numbers match well enough
 * @param {string} ref1 - First reference
 * @param {string} ref2 - Second reference
 * @returns {boolean} - Whether they're a good match
 */
function isGoodReferenceMatch(ref1, ref2) {
    // Handle empty values
    if (!ref1 || !ref2) return false;

    // Quick check for exact match
    if (ref1 === ref2) return true;

    // See if one contains the other
    if (ref1.includes(ref2) || ref2.includes(ref1)) {
        // For short references (less than 5 chars), be more strict
        const shorter = ref1.length <= ref2.length ? ref1 : ref2;
        if (shorter.length < 5) {
            return ref1.startsWith(ref2) || ref2.startsWith(ref1);
        }
        return true;
    }

    // If neither contains the other, they're not a good match
    return false;
}

/**
 * Helper function to format dates consistently
 * @param {Date|string} date - The date to format
 * @returns {string} - Formatted date
 */
function formatDate(date) {
    try {
        // Convert string dates to Date objects
        const dateObj = typeof date === 'string' ?
            dateUtils.parseDate(date) :
            date instanceof Date ? date : new Date(date);

        // Important: Use the date directly without timezone conversion
        return Utilities.formatDate(
            dateObj,
            Session.getScriptTimeZone(),
            'dd.MM.yyyy',
        );
    } catch (e) {
        console.error('Error formatting date:', e);
        return String(date);
    }
}

/**
 * Ermittelt das Präfix für einen Dokumenttyp
 * @param {string} sheetType - Sheet type
 * @returns {string} - Document type prefix
 */
function getDocumentTypePrefix(sheetType) {
    const prefixMap = {
        'einnahmen': 'einnahme',
        'ausgaben': 'ausgabe',
        'eigenbelege': 'eigenbeleg',
        'gesellschafterkonto': 'gesellschafterkonto',
        'holdingTransfers': 'holdingtransfer',
        'gutschrift': 'gutschrift',
    };
    return prefixMap[sheetType] || sheetType;
}

/**
 * Erstellt einen Informationstext für eine Bank-Zuordnung
 * @param {Object} zuordnung - The bank assignment
 * @returns {string} - Formatted info text
 */
function getZuordnungsInfo(zuordnung) {
    if (!zuordnung) return '';

    let infoText = '✓ Bank: ';

    // Format the date to Berlin timezone
    try {
        const date = zuordnung.bankDatum;
        if (date) {
            // Format as DD.MM.YYYY
            const formattedDate = Utilities.formatDate(
                new Date(date),
                Session.getScriptTimeZone(),
                'dd.MM.yyyy',
            );
            infoText += formattedDate;
        } else {
            infoText += 'Datum unbekannt';
        }
    } catch (e) {
        infoText += String(zuordnung.bankDatum);
    }

    // Additional info if available
    if (zuordnung.additional && zuordnung.additional.length > 0) {
        infoText += ' + ' + zuordnung.additional.length + ' weitere';
    }

    return infoText;
}

export default {
    isGoodReferenceMatch,
    formatDate,
    getDocumentTypePrefix,
    getZuordnungsInfo,
};