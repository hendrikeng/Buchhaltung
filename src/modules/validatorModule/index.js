// modules/validatorModule/index.js
import documentValidator from './documentValidator.js';
import bankValidator from './bankValidator.js';
import cellValidator from './cellValidator.js';

/**
 * Hauptexport der Validator-Funktionalität
 */
export default {
    ...documentValidator,
    ...bankValidator,
    ...cellValidator
};