// src/modules/taxModule/optimierungsCalculator.js
import numberUtils from '../../utils/numberUtils.js';
import dividendenCalculator from './dividendenCalculator.js';
import koerperschaftsteuerCalculator from './koerperschaftsteuerCalculator.js';
import gewerbesteuerCalculator from './gewerbesteuerCalculator.js';

/**
 * Modul zur Berechnung von Steueroptimierungsszenarien
 */
const optimierungsCalculator = {
    /**
     * Berechnet verschiedene Steueroptimierungsszenarien
     * @param {Object} bilanzData - Die Bilanzdaten
     * @param {Object} config - Die Konfiguration
     * @returns {Object} - Berechnete Optimierungsszenarien
     */
    calculateOptimierungsszenarien(bilanzData, config) {
        // Gewinn aus Bilanz übernehmen
        const gewinnVorSteuern = bilanzData?.passiva?.jahresueberschuss || 0;

        // Keine Optimierungsszenarien bei Verlust
        if (gewinnVorSteuern <= 0) {
            return {
                thesaurierungVsAusschuettung: null,
                holdingVsDirekt: null,
                simulation: null,
            };
        }

        // 1. Szenario: Thesaurierung vs. Ausschüttung
        const thesaurierungVsAusschuettung = this._compareThesaurierungVsAusschuettung(
            gewinnVorSteuern, config,
        );

        // 2. Szenario: Gewinn in Holding ziehen vs. direkt ausschütten
        const holdingVsDirekt = this._compareHoldingVsDirekt(
            gewinnVorSteuern, config,
        );

        // 3. Simulationsrechner für Netto-Ergebnis
        const simulation = this._simulateNettoErgebnis(
            gewinnVorSteuern, config,
        );

        return {
            thesaurierungVsAusschuettung,
            holdingVsDirekt,
            simulation,
        };
    },

    /**
     * Vergleicht Thesaurierung mit sofortiger Ausschüttung
     * @param {number} gewinn - Gewinn vor Steuern
     * @param {Object} config - Die Konfiguration
     * @returns {Object} - Vergleichsdaten
     */
    _compareThesaurierungVsAusschuettung(gewinn, config) {
        // 1. Unternehmenssteuern berechnen (KSt + Soli + GewSt)
        const kstResult = koerperschaftsteuerCalculator.calculateKoerperschaftsteuer(
            { passiva: { jahresueberschuss: gewinn } }, config,
        );

        const gewStResult = config.tax.isHolding ?
            { gewerbesteuer: 0 } :
            gewerbesteuerCalculator.calculateGewerbesteuer(
                { passiva: { jahresueberschuss: gewinn } }, config,
            );

        const unternehmenssteuern = kstResult.koerperschaftsteuer +
            kstResult.solidaritaetszuschlag +
            gewStResult.gewerbesteuer;

        // Gewinn nach Unternehmenssteuern
        const gewinnNachUnternehmenssteuern = gewinn - unternehmenssteuern;

        // 2. Szenario Thesaurierung: Gewinn verbleibt in der GmbH
        const thesaurierung = {
            gewinnVorSteuern: gewinn,
            unternehmenssteuern: unternehmenssteuern,
            gewinnNachSteuern: gewinnNachUnternehmenssteuern,
            effektiverSteuersatz: (unternehmenssteuern / gewinn) * 100,
        };

        // 3. Szenario Ausschüttung: Gewinn wird ausgeschüttet
        // Typischer Einkommensteuersatz eines Gesellschafters (42%)
        const typischerEStSatz = 42;

        // Ausschüttungsszenarien mit verschiedenen ESt-Sätzen
        const ausschuettungsSzenarien = [25, 35, 42, 45].map(estSatz => {
            const abgeltungResult = dividendenCalculator.calculateAbgeltungsteuer(
                gewinnNachUnternehmenssteuern, config,
            );

            const teVergleich = dividendenCalculator.compareAbgeltungWithEinkommensteuer(
                gewinnNachUnternehmenssteuern, estSatz, config,
            );

            // Günstigere Besteuerungsmethode verwenden
            const ausschuettungssteuer = Math.min(
                abgeltungResult.gesamtsteuer,
                teVergleich.teileinkuenfteverfahren.steuer,
            );

            // Gesamte Steuerbelastung
            const gesamtsteuer = unternehmenssteuern + ausschuettungssteuer;

            // Netto-Ergebnis für den Gesellschafter
            const nettoErgebnis = gewinn - gesamtsteuer;

            // Effektiver Steuersatz
            const effektiverSteuersatz = (gesamtsteuer / gewinn) * 100;

            return {
                estSatz,
                unternehmenssteuern,
                ausschuettungssteuer,
                gesamtsteuer,
                nettoErgebnis,
                effektiverSteuersatz,
            };
        });

        return {
            thesaurierung,
            ausschuettungsSzenarien,
            empfehlung: this._getThesaurierungsEmpfehlung(
                thesaurierung,
                ausschuettungsSzenarien[2], // 42% ESt-Satz als Referenz
            ),
        };
    },

    /**
     * Vergleicht direkte Ausschüttung mit Ausschüttung über Holding
     * @param {number} gewinn - Gewinn vor Steuern
     * @param {Object} config - Die Konfiguration
     * @returns {Object} - Vergleichsdaten
     */
    _compareHoldingVsDirekt(gewinn, config) {
        // Temporäre Konfiguration für operative GmbH
        const operativeConfig = JSON.parse(JSON.stringify(config));
        operativeConfig.tax.isHolding = false;

        // Temporäre Konfiguration für Holding GmbH
        const holdingConfig = JSON.parse(JSON.stringify(config));
        holdingConfig.tax.isHolding = true;

        // 1. Szenario: Direkte Ausschüttung aus operativer GmbH

        // Unternehmenssteuern der operativen GmbH
        const kstResultOperativ = koerperschaftsteuerCalculator.calculateKoerperschaftsteuer(
            { passiva: { jahresueberschuss: gewinn } }, operativeConfig,
        );

        const gewStResultOperativ = gewerbesteuerCalculator.calculateGewerbesteuer(
            { passiva: { jahresueberschuss: gewinn } }, operativeConfig,
        );

        const unternehmenssteuernOperativ = kstResultOperativ.koerperschaftsteuer +
            kstResultOperativ.solidaritaetszuschlag +
            gewStResultOperativ.gewerbesteuer;

        // Gewinn nach Unternehmenssteuern
        const gewinnNachUnternehmenssteuernOperativ = gewinn - unternehmenssteuernOperativ;

        // Abgeltungsteuer auf Ausschüttung
        const abgeltungResultDirekt = dividendenCalculator.calculateAbgeltungsteuer(
            gewinnNachUnternehmenssteuernOperativ, operativeConfig,
        );

        // Gesamte Steuerbelastung bei direkter Ausschüttung
        const gesamtsteuerDirekt = unternehmenssteuernOperativ + abgeltungResultDirekt.gesamtsteuer;

        // Netto-Ergebnis für den Gesellschafter bei direkter Ausschüttung
        const nettoErgebnisDirekt = gewinn - gesamtsteuerDirekt;

        // 2. Szenario: Ausschüttung über Holding

        // Unternehmenssteuern der operativen GmbH bleiben gleich

        // Beteiligungserträge der Holding (95% steuerfrei)
        const beteiligungsertraege = gewinnNachUnternehmenssteuernOperativ;

        // Körperschaftsteuer auf 5% der Beteiligungserträge
        const steuerpflichtigerAnteilHolding = beteiligungsertraege *
            (holdingConfig.tax.holding.gewinnUebertragSteuerpflichtig / 100);

        const kstHolding = steuerpflichtigerAnteilHolding *
            (holdingConfig.tax.holding.koerperschaftsteuer / 100);

        const soliHolding = kstHolding *
            (holdingConfig.tax.holding.solidaritaetszuschlag / 100);

        // Steuern der Holding
        const steuerHolding = kstHolding + soliHolding;

        // Ausschüttbarer Betrag der Holding
        const ausschuettbarHolding = beteiligungsertraege - steuerHolding;

        // Abgeltungsteuer auf Ausschüttung der Holding
        const abgeltungResultHolding = dividendenCalculator.calculateAbgeltungsteuer(
            ausschuettbarHolding, holdingConfig,
        );

        // Gesamte Steuerbelastung bei Ausschüttung über Holding
        const gesamtsteuerHolding = unternehmenssteuernOperativ +
            steuerHolding +
            abgeltungResultHolding.gesamtsteuer;

        // Netto-Ergebnis für den Gesellschafter bei Ausschüttung über Holding
        const nettoErgebnisHolding = gewinn - gesamtsteuerHolding;

        // Steuerersparnis durch Holding-Struktur
        const steuerersparnis = gesamtsteuerDirekt - gesamtsteuerHolding;

        return {
            direkteAusschuettung: {
                unternehmenssteuern: unternehmenssteuernOperativ,
                ausschuettungssteuer: abgeltungResultDirekt.gesamtsteuer,
                gesamtsteuer: gesamtsteuerDirekt,
                nettoErgebnis: nettoErgebnisDirekt,
                effektiverSteuersatz: (gesamtsteuerDirekt / gewinn) * 100,
            },
            ausschuettungUeberHolding: {
                unternehmenssteuernOperativ: unternehmenssteuernOperativ,
                steuerHolding: steuerHolding,
                ausschuettungssteuer: abgeltungResultHolding.gesamtsteuer,
                gesamtsteuer: gesamtsteuerHolding,
                nettoErgebnis: nettoErgebnisHolding,
                effektiverSteuersatz: (gesamtsteuerHolding / gewinn) * 100,
            },
            steuerersparnis,
            empfehlung: this._getHoldingEmpfehlung(steuerersparnis, gewinn),
        };
    },

    /**
     * Simuliert das Netto-Ergebnis nach Steuern in verschiedenen Szenarien
     * @param {number} gewinn - Gewinn vor Steuern
     * @param {Object} config - Die Konfiguration
     * @returns {Object} - Simulationsergebnisse
     */
    _simulateNettoErgebnis(gewinn, config) {
        // Verschiedene Gewinnhöhen simulieren
        const gewinnSzenarien = [
            50000, 100000, 250000, 500000, 1000000,
        ];

        // Ergebnisse für jedes Szenario berechnen
        const ergebnisse = gewinnSzenarien.map(szenarioGewinn => {
            // Berechnung für operative GmbH mit voller Ausschüttung
            const operativeConfig = JSON.parse(JSON.stringify(config));
            operativeConfig.tax.isHolding = false;

            const operativeResult = this._compareThesaurierungVsAusschuettung(
                szenarioGewinn, operativeConfig,
            );

            // Berechnung für Holding-Struktur
            const holdingResult = this._compareHoldingVsDirekt(
                szenarioGewinn, config,
            );

            return {
                gewinn: szenarioGewinn,
                operativeGmbH: {
                    thesaurierung: {
                        steuern: operativeResult.thesaurierung.unternehmenssteuern,
                        netto: operativeResult.thesaurierung.gewinnNachSteuern,
                        effektiverSteuersatz: operativeResult.thesaurierung.effektiverSteuersatz,
                    },
                    ausschuettung: {
                        steuern: operativeResult.ausschuettungsSzenarien[2].gesamtsteuer,
                        netto: operativeResult.ausschuettungsSzenarien[2].nettoErgebnis,
                        effektiverSteuersatz: operativeResult.ausschuettungsSzenarien[2].effektiverSteuersatz,
                    },
                },
                holdingStruktur: {
                    steuern: holdingResult.ausschuettungUeberHolding.gesamtsteuer,
                    netto: holdingResult.ausschuettungUeberHolding.nettoErgebnis,
                    effektiverSteuersatz: holdingResult.ausschuettungUeberHolding.effektiverSteuersatz,
                },
            };
        });

        return {
            szenarien: ergebnisse,
            empfehlung: this._getGesamtEmpfehlung(ergebnisse),
        };
    },

    /**
     * Ermittelt eine Empfehlung für Thesaurierung vs. Ausschüttung
     * @param {Object} thesaurierung - Thesaurierungsszenario
     * @param {Object} ausschuettung - Ausschüttungsszenario
     * @returns {string} - Empfehlung
     */
    _getThesaurierungsEmpfehlung(thesaurierung, ausschuettung) {
        const steuerdifferenz = ausschuettung.gesamtsteuer - thesaurierung.unternehmenssteuern;

        if (steuerdifferenz > 10000) {
            return 'Gewinn in der GmbH belassen (thesaurieren) und nur nach Bedarf ausschütten.';
        } else if (steuerdifferenz > 5000) {
            return 'Tendenziell thesaurieren, aber teilweise Ausschüttung nach persönlichem Bedarf.';
        } else {
            return 'Ausschüttung je nach persönlichem Bedarf, da Steuervorteil der Thesaurierung gering.';
        }
    },

    /**
     * Ermittelt eine Empfehlung für Holding-Struktur vs. direkte Ausschüttung
     * @param {number} steuerersparnis - Steuerersparnis durch Holding
     * @param {number} gewinn - Ursprünglicher Gewinn
     * @returns {string} - Empfehlung
     */
    _getHoldingEmpfehlung(steuerersparnis, gewinn) {
        const relativeErsparnis = (steuerersparnis / gewinn) * 100;

        if (steuerersparnis < 0) {
            return 'Direkte Ausschüttung ohne Holding-Struktur ist günstiger.';
        } else if (relativeErsparnis < 1) {
            return 'Steuerersparnis durch Holding-Struktur ist minimal und rechtfertigt eventuell nicht den administrativen Aufwand.';
        } else if (relativeErsparnis < 3) {
            return 'Moderate Steuerersparnis durch Holding-Struktur. Bei vorhandenem Beteiligungsbesitz sinnvoll.';
        } else {
            return 'Deutliche Steuerersparnis durch Holding-Struktur. Empfehlenswert.';
        }
    },

    /**
     * Ermittelt eine Gesamtempfehlung basierend auf allen Szenarien
     * @param {Array} szenarien - Alle berechneten Szenarien
     * @returns {string} - Gesamtempfehlung
     */
    _getGesamtEmpfehlung(szenarien) {
        // Durchschnittliche Steuerersparnis durch Holding-Struktur berechnen
        let durchschnittlicheErsparnis = 0;
        const anzahlSzenarien = szenarien.length;

        szenarien.forEach(szenario => {
            const direkteSteuern = szenario.operativeGmbH.ausschuettung.steuern;
            const holdingSteuern = szenario.holdingStruktur.steuern;
            durchschnittlicheErsparnis += direkteSteuern - holdingSteuern;
        });

        durchschnittlicheErsparnis /= anzahlSzenarien;

        // Durchschnittlicher Steuersatzvorteil für Thesaurierung
        let durchschnittlicherThesaurierungsVorteil = 0;

        szenarien.forEach(szenario => {
            const ausschuettungsSatz = szenario.operativeGmbH.ausschuettung.effektiverSteuersatz;
            const thesaurierungsSatz = szenario.operativeGmbH.thesaurierung.effektiverSteuersatz;
            durchschnittlicherThesaurierungsVorteil += ausschuettungsSatz - thesaurierungsSatz;
        });

        durchschnittlicherThesaurierungsVorteil /= anzahlSzenarien;

        // Empfehlung basierend auf den Durchschnittswerten
        let empfehlung = '';

        if (durchschnittlicheErsparnis > 20000) {
            empfehlung = 'Die Holding-Struktur bietet erhebliche steuerliche Vorteile';

            if (durchschnittlicherThesaurierungsVorteil > 10) {
                empfehlung += ' in Kombination mit einer Thesaurierungsstrategie.';
            } else {
                empfehlung += ', unabhängig von der Ausschüttungspolitik.';
            }
        } else if (durchschnittlicheErsparnis > 5000) {
            empfehlung = 'Eine Holding-Struktur bietet moderate steuerliche Vorteile';

            if (durchschnittlicherThesaurierungsVorteil > 10) {
                empfehlung += ', besonders in Kombination mit einer Thesaurierungsstrategie.';
            } else {
                empfehlung += ', aber die Kosten der Strukturierung sollten berücksichtigt werden.';
            }
        } else {
            if (durchschnittlicherThesaurierungsVorteil > 10) {
                empfehlung = 'Die Thesaurierung von Gewinnen bietet deutliche steuerliche Vorteile, ' +
                    'während die Einrichtung einer Holding-Struktur nur geringe Steuervorteile bringt.';
            } else {
                empfehlung = 'Weder die Holding-Struktur noch eine strikte Thesaurierungspolitik ' +
                    'bringen erhebliche steuerliche Vorteile. Die Entscheidung sollte primär ' +
                    'nach betriebswirtschaftlichen und nicht nach steuerlichen Gesichtspunkten erfolgen.';
            }
        }

        return empfehlung;
    },
};

export default optimierungsCalculator;