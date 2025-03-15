// config/index.js
import common from './common.js';
import tax from './tax.js';
import sheets from './sheets.js';
import accounts from './accounts.js';

/**
 * Konfiguration für die Buchhaltungsanwendung
 * Unterstützt die Buchhaltung für Holding und operative GmbH nach SKR04
 */
const config = {
    // Allgemeine Einstellungen
    common,

    // Steuerliche Einstellungen
    tax,

    // Sheet-Konfigurationen
    ...sheets,

    // Kontenplan
    kontenplan: accounts,

    /**
     * Initialisierungsfunktion für abgeleitete Daten
     * Wird automatisch beim Import aufgerufen
     */
    initialize() {
        // Bankkategorien dynamisch aus den Einnahmen- und Ausgaben-Kategorien befüllen
        this.bankbewegungen.categories = [
            ...Object.keys(this.einnahmen.categories),
            ...Object.keys(this.ausgaben.categories),
            ...Object.keys(this.eigenbelege.categories),
            ...Object.keys(this.gesellschafterkonto.categories),
            ...Object.keys(this.holdingTransfers.categories),
        ];

        // Duplikate aus den Kategorien entfernen
        this.bankbewegungen.categories = [...new Set(this.bankbewegungen.categories)];

        return this;
    },
};

// Initialisierung ausführen und exportieren
export default config.initialize();