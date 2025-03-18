// utils/cacheUtils.js
/**
 * Funktionen für Caching zur Verbesserung der Performance
 */

/**
 * Cache-Objekt für verschiedene Datentypen
 */
const globalCache = {
    // Verschiedene Cache-Maps für unterschiedliche Datentypen
    _store: {
        dates: new Map(),
        currency: new Map(),
        mwstRates: new Map(),
        columnLetters: new Map(),
        sheets: new Map(),
        references: {
            einnahmen: null,
            ausgaben: null,
            eigenbelege: null,
            gesellschafterkonto: null,
            holdingTransfers: null,
        },
        computed: new Map(),
    },

    /**
     * Löscht den gesamten Cache oder einen bestimmten Teilbereich
     * @param {string} type - Optional: Der zu löschende Cache-Bereich
     */
    clear(type = null) {
        if (type === null) {
            // Optimized: Dynamically clear all maps
            Object.entries(this._store).forEach(([key, value]) => {
                if (value instanceof Map) {
                    value.clear();
                } else if (key === 'references') {
                    Object.keys(this._store.references).forEach(refKey => {
                        this._store.references[refKey] = null;
                    });
                }
            });
        } else if (type === 'references') {
            Object.keys(this._store.references).forEach(refKey => {
                this._store.references[refKey] = null;
            });
        } else if (this._store[type]) {
            if (this._store[type] instanceof Map) {
                this._store[type].clear();
            }
        }
    },

    /**
     * Speichert einen Wert im Cache
     * @param {string} type - Cache-Typ (dates, currency, mwstRates, columnLetters, computed)
     * @param {string} key - Schlüssel für den gecacheten Wert
     * @param {*} value - Zu cachender Wert
     */
    set(type, key, value) {
        if (this._store[type] instanceof Map) {
            this._store[type].set(key, value);
        } else if (type === 'references' && key in this._store.references) {
            this._store.references[key] = value;
        }
    },

    /**
     * Prüft ob ein Wert im Cache vorhanden ist
     * @param {string} type - Cache-Typ
     * @param {string} key - Schlüssel des zu prüfenden Werts
     * @returns {boolean} true wenn der Wert im Cache vorhanden ist
     */
    has(type, key) {
        if (this._store[type] instanceof Map) {
            return this._store[type].has(key);
        } else if (type === 'references') {
            return this._store.references[key] !== null;
        }
        return false;
    },

    /**
     * Ruft einen Wert aus dem Cache ab
     * @param {string} type - Cache-Typ
     * @param {string} key - Schlüssel des abzurufenden Werts
     * @returns {*} Der gecachete Wert oder undefined wenn nicht vorhanden
     */
    get(type, key) {
        if (this._store[type] instanceof Map) {
            return this._store[type].get(key);
        } else if (type === 'references') {
            return this._store.references[key];
        }
        return undefined;
    },

    /**
     * Ruft einen Wert aus dem Cache ab oder berechnet ihn, wenn nicht vorhanden
     * @param {string} type - Cache-Typ
     * @param {string} key - Schlüssel des abzurufenden Werts
     * @param {Function} computeFn - Funktion zur Berechnung des Werts, wenn nicht im Cache
     * @returns {*} Der gecachete oder berechnete Wert
     */
    getOrCompute(type, key, computeFn) {
        if (this.has(type, key)) {
            return this.get(type, key);
        }

        // Added error handling for compute function
        try {
            const value = computeFn();
            this.set(type, key, value);
            return value;
        } catch (error) {
            console.error(`Error computing cache value for ${type}.${key}:`, error);
            return undefined;
        }
    },
};

export default globalCache;