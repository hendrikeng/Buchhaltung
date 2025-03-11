/**
 * accountingModule.js - Modul für allgemeine Buchhaltungsfunktionen
 *
 * Stellt Funktionen bereit, die übergreifend für die Buchhaltung benötigt werden:
 * - SKR04-Konten und Zuordnungen
 * - Quartalsberechnungen
 * - Steuerliche Berechnungen
 */

const AccountingModule = (function() {
    // Private Variablen und Funktionen

    /**
     * SKR04-Kontenrahmen mit den wichtigsten Konten für das System
     * Die vollständige Liste würde den Rahmen sprengen
     */
    const SKR04_ACCOUNTS = {
        // Anlagevermögen
        "Sachanlagen": {
            kontoNr: "0300",
            steuerlich: "keine USt"
        },
        "Büroeinrichtung": {
            kontoNr: "0410",
            steuerlich: "Vorsteuer"
        },
        "GWG": {
            kontoNr: "0490",
            steuerlich: "Vorsteuer"
        },

        // Finanzkonten
        "Bank": {
            kontoNr: "1200",
            steuerlich: "keine USt"
        },
        "Kasse": {
            kontoNr: "1000",
            steuerlich: "keine USt"
        },

        // Forderungen und Verbindlichkeiten
        "Forderungen aus L+L": {
            kontoNr: "1400",
            steuerlich: "keine USt"
        },
        "Verbindlichkeiten aus L+L": {
            kontoNr: "1600",
            steuerlich: "keine USt"
        },

        // Umsatzsteuer
        "Umsatzsteuer 19%": {
            kontoNr: "1776",
            steuerlich: "USt 19%"
        },
        "Umsatzsteuer 7%": {
            kontoNr: "1771",
            steuerlich: "USt 7%"
        },
        "Vorsteuer 19%": {
            kontoNr: "1576",
            steuerlich: "VSt 19%"
        },
        "Vorsteuer 7%": {
            kontoNr: "1571",
            steuerlich: "VSt 7%"
        },

        // Erlöskonten
        "Erlöse aus Lieferungen und Leistungen": {
            kontoNr: "8400",
            steuerlich: "USt-pflichtig"
        },
        "Erlöse aus Vermietung und Verpachtung": {
            kontoNr: "8300",
            steuerlich: "USt-pflichtig"
        },
        "Erlöse steuerfrei §4 Nr. 8-28 UStG": {
            kontoNr: "8120",
            steuerlich: "steuerfrei Inland"
        },
        "Erlöse innergemeinschaftliche Lieferungen §4 Nr. 1b UStG": {
            kontoNr: "8125",
            steuerlich: "steuerfrei EU"
        },
        "Erlöse Ausfuhrlieferungen §4 Nr. 1a UStG": {
            kontoNr: "8120",
            steuerlich: "steuerfrei Export"
        },
        "Sonstige betriebliche Erträge": {
            kontoNr: "8100",
            steuerlich: "USt-pflichtig"
        },
        "Zinserträge": {
            kontoNr: "8150",
            steuerlich: "steuerfrei"
        },

        // Aufwandskonten
        "Wareneinkauf 19% VSt": {
            kontoNr: "3300",
            steuerlich: "VSt 19%"
        },
        "Wareneinkauf 7% VSt": {
            kontoNr: "3400",
            steuerlich: "VSt 7%"
        },
        "Fremdleistungen": {
            kontoNr: "3500",
            steuerlich: "VSt 19%"
        },
        "Personalkosten": {
            kontoNr: "4000",
            steuerlich: "keine USt"
        },
        "Raumkosten": {
            kontoNr: "4210",
            steuerlich: "VSt 19%"
        },
        "Fahrzeugkosten": {
            kontoNr: "4530",
            steuerlich: "VSt 19%"
        },
        "Werbe- und Reisekosten": {
            kontoNr: "4600",
            steuerlich: "VSt 19%"
        },
        "Bewirtungskosten": {
            kontoNr: "4650",
            steuerlich: "VSt 19% (70%)"
        },
        "Bürobedarf": {
            kontoNr: "6800",
            steuerlich: "VSt 19%"
        },
        "Versicherungen/Beiträge": {
            kontoNr: "4380",
            steuerlich: "keine USt"
        },
        "Beratungskosten": {
            kontoNr: "6830",
            steuerlich: "VSt 19%"
        }
    };

    /**
     * Mapping von Kategorien zu Buchungssätzen
     * Definiert automatische Buchungssätze für Einnahmen und Ausgaben
     */
    const ACCOUNT_MAPPINGS = {
        // Einnahmen
        einnahmen: {
            "Erlöse aus Lieferungen und Leistungen": {
                soll: "1200", // Bank
                haben: "8400"  // Erlöse aus L+L
            },
            "Erlöse aus Vermietung und Verpachtung": {
                soll: "1200", // Bank
                haben: "8300"  // Erlöse aus Vermietung
            },
            "Erlöse steuerfrei §4 Nr. 8-28 UStG": {
                soll: "1200", // Bank
                haben: "8120"  // Steuerfreie Umsätze
            },
            "Erlöse innergemeinschaftliche Lieferungen §4 Nr. 1b UStG": {
                soll: "1200", // Bank
                haben: "8125"  // Innergemeinschaftliche Lieferungen
            },
            "Zinserträge": {
                soll: "1200", // Bank
                haben: "8150"  // Zinserträge
            }
            // Weitere Einnahmenkategorien...
        },

        // Ausgaben
        ausgaben: {
            "Wareneinkauf 19% VSt": {
                soll: "3300", // Wareneinkauf
                haben: "1200"  // Bank
            },
            "Wareneinkauf 7% VSt": {
                soll: "3400", // Wareneinkauf 7%
                haben: "1200"  // Bank
            },
            "Fremdleistungen": {
                soll: "3500", // Fremdleistungen
                haben: "1200"  // Bank
            },
            "Raumkosten": {
                soll: "4210", // Raumkosten
                haben: "1200"  // Bank
            },
            "Bürobedarf": {
                soll: "6800", // Bürobedarf
                haben: "1200"  // Bank
            },
            "Beratungskosten": {
                soll: "6830", // Beratungskosten
                haben: "1200"  // Bank
            }
            // Weitere Ausgabenkategorien...
        }
    };

    /**
     * Ermittelt das richtige Quartal für ein Datum
     * @param {Date|string} datum - Datum als Date-Objekt oder String
     * @returns {number} Quartalszahl (1-4)
     */
    function getQuartal(datum) {
        const date = (datum instanceof Date) ? datum : new Date(datum);
        const month = date.getMonth(); // 0-basiert (0 = Januar)

        if (month <= 2) return 1;      // Jan-März
        else if (month <= 5) return 2; // Apr-Jun
        else if (month <= 8) return 3; // Jul-Sep
        else return 4;                 // Okt-Dez
    }

    // Öffentliche API
    return {
        /**
         * Liefert die passenden Soll- und Haben-Konten basierend auf SKR04 für eine Kategorie
         * @param {string} kategorie - Kategoriename
         * @param {boolean} isEinnahme - Handelt es sich um eine Einnahme?
         * @returns {Object|null} Objekt mit soll und haben Kontonummern oder null
         */
        getSKR04Accounts: function(kategorie, isEinnahme) {
            const type = isEinnahme ? 'einnahmen' : 'ausgaben';

            // In den Mappings suchen
            if (ACCOUNT_MAPPINGS[type] && ACCOUNT_MAPPINGS[type][kategorie]) {
                return ACCOUNT_MAPPINGS[type][kategorie];
            }

            // Fallback: Manuell zu prüfende Konten
            return {
                soll: "Manuell prüfen",
                haben: "Manuell prüfen"
            };
        },

        /**
         * Berechnet das Quartal für ein Datum
         * @param {Date|string} datum - Datum
         * @returns {number} Quartalszahl (1-4)
         */
        getQuartal: function(datum) {
            return getQuartal(datum);
        },

        /**
         * Berechnet die Mehrwertsteuer basierend auf dem Nettobetrag und Steuersatz
         * @param {number} nettoBetrag - Nettobetrag in Euro
         * @param {number} steuersatz - Steuersatz als Dezimalzahl (z.B. 0.19 für 19%)
         * @returns {number} Mehrwertsteuerbetrag
         */
        calculateMwSt: function(nettoBetrag, steuersatz) {
            return Math.round(nettoBetrag * steuersatz * 100) / 100;
        },

        /**
         * Berechnet Netto- und MwSt-Betrag aus einem Bruttobetrag
         * @param {number} bruttoBetrag - Bruttobetrag in Euro
         * @param {number} steuersatz - Steuersatz als Dezimalzahl (z.B. 0.19 für 19%)
         * @returns {Object} Objekt mit netto und mwst Werten
         */
        calculateNettoFromBrutto: function(bruttoBetrag, steuersatz) {
            const netto = Math.round(bruttoBetrag / (1 + steuersatz) * 100) / 100;
            const mwst = Math.round((bruttoBetrag - netto) * 100) / 100;

            return {
                netto: netto,
                mwst: mwst
            };
        },

        /**
         * Prüft, ob eine Kategorie steuerpflichtig ist
         * @param {string} kategorie - Kategoriename
         * @returns {boolean} True, wenn die Kategorie steuerpflichtig ist
         */
        isSteuerpflichtig: function(kategorie) {
            // Steuerfreie Kategorien enthalten typischerweise bestimmte Schlüsselwörter
            const steuerfrei = kategorie.match(/steuerfrei|§4|innergemeinschaftlich|Ausfuhr/i);
            return !steuerfrei;
        },

        /**
         * Liefert eine Liste aller verfügbaren SKR04-Konten
         * @returns {Object} SKR04-Konten mit Details
         */
        getAllSKR04Accounts: function() {
            return {...SKR04_ACCOUNTS};
        },

        /**
         * Gibt die Details eines spezifischen SKR04-Kontos zurück
         * @param {string} kontoName - Name des Kontos
         * @returns {Object|null} Kontoinformationen oder null wenn nicht gefunden
         */
        getSKR04AccountDetails: function(kontoName) {
            return SKR04_ACCOUNTS[kontoName] || null;
        }
    };
})();