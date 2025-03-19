// modules/validatorModule/index.js
import documentValidator from './documentValidator.js';
import bankValidator from './bankValidator.js';
import cellValidator from './cellValidator.js';

/**
 * Hauptexport der Validator-Funktionalität mit optimierter Struktur
 * Für eine bessere Modulorganisation werden alle Validator-Funktionen exportiert
 */
export default {
    // Document validation functions
    validateDocumentRow: documentValidator.validateDocumentRow,
    validateSheet: documentValidator.validateSheet,
    validateAllSheets: documentValidator.validateAllSheets,

    // Bank validation functions
    validateBanking: bankValidator.validateBanking,

    // Cell validation functions
    validateDropdown: cellValidator.validateDropdown,
    validateCellValue: cellValidator.validateCellValue,
    validateSheetWithRules: cellValidator.validateSheetWithRules,
};