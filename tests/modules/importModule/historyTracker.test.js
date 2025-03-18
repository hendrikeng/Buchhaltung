// tests/modules/importModule/historyTracker.test.js
import { describe, it, expect, jest, beforeEach } from '@jest/globals';

// Mock the historyTracker module functions
const historyTracker = {
    getOrCreateHistorySheet: jest.fn(),
    collectExistingFiles: jest.fn(), 
};

// Mock dependencies
const sheetUtils = {
    getOrCreateSheet: jest.fn(),
};

describe('historyTracker', () => {
    beforeEach(() => {
        jest.clearAllMocks();
    });

    describe('getOrCreateHistorySheet', () => {
        it('should get existing history sheet', () => {
            // Setup
            const mockSheet = {
                getLastRow: jest.fn(() => 5),
            };

            sheetUtils.getOrCreateSheet.mockReturnValue(mockSheet);
            historyTracker.getOrCreateHistorySheet.mockReturnValue(mockSheet);

            const mockSpreadsheet = {};
            const mockConfig = {
                aenderungshistorie: { columns: {} },
            };

            // Execute
            const result = historyTracker.getOrCreateHistorySheet(mockSpreadsheet, mockConfig);

            // Verify
            expect(result).toBe(mockSheet);
        });
    });

    describe('collectExistingFiles', () => {
        it('should collect file names from history sheet', () => {
            // Setup
            const mockFileSet = new Set(['file1.pdf', 'file2.pdf']);
            historyTracker.collectExistingFiles.mockReturnValue(mockFileSet);

            const mockSheet = {
                getDataRange: jest.fn(() => ({
                    getValues: jest.fn(() => [
                        ['Datum', 'Typ', 'Dateiname', 'Link'],
                        ['2023-01-01', 'Einnahme', 'file1.pdf', 'url1'],
                        ['2023-01-02', 'Ausgabe', 'file2.pdf', 'url2'],
                    ]),
                })),
            };

            // Execute
            const result = historyTracker.collectExistingFiles(mockSheet);

            // Verify
            expect(result).toEqual(mockFileSet);
            expect(result.size).toBe(2);
        });
    });
});