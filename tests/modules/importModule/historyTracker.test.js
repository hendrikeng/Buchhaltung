// tests/modules/importModule/historyTracker.test.js
import { jest, describe, beforeEach, it, expect } from '@jest/globals';
import historyTracker from '../../../src/modules/importModule/historyTracker.js';
import sheetUtils from '../../../src/utils/sheetUtils.js';

// Mocks für die abhängigen Module
jest.mock('../../../src/utils/sheetUtils.js', () => ({
    getOrCreateSheet: jest.fn()
}));

describe('historyTracker', () => {
    beforeEach(() => {
        jest.clearAllMocks();
    });

    describe('getOrCreateHistorySheet', () => {
        it('should get existing history sheet', () => {
            // Einrichten der Mocks
            const mockHistorySheet = {
                getLastRow: jest.fn(() => 5)
            };

            sheetUtils.getOrCreateSheet.mockReturnValue(mockHistorySheet);

            const mockSpreadsheet = {};
            const mockConfig = {
                aenderungshistorie: {
                    columns: {
                        datum: 1,
                        typ: 2,
                        dateiname: 3,
                        dateilink: 4
                    }
                }
            };

            // Test ausführen
            const result = historyTracker.getOrCreateHistorySheet(mockSpreadsheet, mockConfig);

            // Überprüfungen
            expect(sheetUtils.getOrCreateSheet).toHaveBeenCalledWith(mockSpreadsheet, 'Änderungshistorie');
            expect(result).toBe(mockHistorySheet);
            expect(mockHistorySheet.getLastRow).toHaveBeenCalled();
        });

        it('should create history sheet with header if it does not exist', () => {
            // Einrichten der Mocks
            const mockHistorySheet = {
                getLastRow: jest.fn(() => 0),
                appendRow: jest.fn(),
                getRange: jest.fn(() => ({
                    setFontWeight: jest.fn(() => ({}))
                }))
            };

            sheetUtils.getOrCreateSheet.mockReturnValue(mockHistorySheet);

            const mockSpreadsheet = {};
            const mockConfig = {
                aenderungshistorie: {
                    columns: {
                        datum: 1,
                        typ: 2,
                        dateiname: 3,
                        dateilink: 4
                    }
                }
            };

            // Test ausführen
            const result = historyTracker.getOrCreateHistorySheet(mockSpreadsheet, mockConfig);

            // Überprüfungen
            expect(mockHistorySheet.appendRow).toHaveBeenCalled();
            expect(mockHistorySheet.getRange).toHaveBeenCalledWith(1, 1, 1, 4);
            expect(result).toBe(mockHistorySheet);
        });
    });

    describe('collectExistingFiles', () => {
        it('should collect file names from history sheet', () => {
            // Einrichten der Mocks
            const mockHistorySheet = {
                getDataRange: jest.fn(() => ({
                    getValues: jest.fn(() => [
                        ['Datum', 'Typ', 'Dateiname', 'Link'],
                        ['2021-01-15', 'Einnahme', 'Invoice-2021-01-15', 'https://example.com/file1'],
                        ['2021-02-20', 'Ausgabe', 'Receipt-2021-02-20', 'https://example.com/file2']
                    ])
                }))
            };

            // Test ausführen
            const result = historyTracker.collectExistingFiles(mockHistorySheet);

            // Überprüfungen
            expect(result.size).toBe(2);
            expect(result.has('Invoice-2021-01-15')).toBe(true);
            expect(result.has('Receipt-2021-02-20')).toBe(true);
        });
    });
});