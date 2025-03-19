// utils/cacheUtils.js
/**
 * Funktionen für Caching zur Verbesserung der Performance
 * Optimierte Version mit effizienteren Strukturen und verbesserter Verwaltung
 */

/**
 * Cache-Objekt für verschiedene Datentypen
 */
const globalCache = {
    // Verschiedene Cache-Maps für unterschiedliche Datentypen
    _store: {
        // Datums- und Zeit-Operationen
        dates: new Map(),

        // Währungs- und Zahlenoperationen
        currency: new Map(),
        mwstRates: new Map(),

        // Formatierungs-Caches
        columnLetters: new Map(),

        // Sheet-Daten
        sheets: new Map(),

        // Import- und Validierungsdaten
        validation: new Map(),

        // Referenzdaten für Module (BWA, Bilanz, UStVA)
        references: {
            einnahmen: null,
            ausgaben: null,
            eigenbelege: null,
            gesellschafterkonto: null,
            holdingTransfers: null,
        },

        // Berechnete Daten für Module
        computed: new Map(),
    },

    /**
     * Löscht den gesamten Cache oder einen bestimmten Teilbereich
     * Optimierte Version mit effizienterem Löschen
     * @param {string} type - Optional: Der zu löschende Cache-Bereich
     */
    clear(type = null) {
        if (type === null) {
            // Optimiert: Dynamisches Löschen aller Maps
            Object.entries(this._store).forEach(([key, value]) => {
                if (value instanceof Map) {
                    value.clear();
                } else if (key === 'references') {
                    Object.keys(this._store.references).forEach(refKey => {
                        this._store.references[refKey] = null;
                    });
                }
            });

            console.info('Entire cache cleared');
        } else if (type === 'references') {
            Object.keys(this._store.references).forEach(refKey => {
                this._store.references[refKey] = null;
            });
            console.info('References cache cleared');
        } else if (this._store[type]) {
            if (this._store[type] instanceof Map) {
                this._store[type].clear();
                console.info(`Cache for ${type} cleared`);
            }
        }
    },

    /**
     * Speichert einen Wert im Cache
     * @param {string} type - Cache-Typ (dates, currency, mwstRates, columnLetters, computed, etc.)
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

    /**
     * Gibt die Größe eines Cache-Typs zurück
     * @param {string} type - Cache-Typ
     * @returns {number} - Anzahl der Elemente im Cache oder -1 wenn nicht verfügbar
     */
    size(type) {
        if (this._store[type] instanceof Map) {
            return this._store[type].size;
        } else if (type === 'references') {
            return Object.values(this._store.references).filter(v => v !== null).length;
        }
        return -1;
    },

    /**
     * Gibt einen Statusbericht über den Cache zurück
     * @returns {Object} - Statusbericht
     */
    status() {
        const status = {};

        Object.entries(this._store).forEach(([type, store]) => {
            if (store instanceof Map) {
                status[type] = store.size;
            } else if (type === 'references') {
                status[type] = Object.values(store).filter(v => v !== null).length;
            }
        });

        return status;
    },
};

export default globalCache;