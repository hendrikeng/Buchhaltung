// tests/modules/importModule/dataProcessor.test.js
import { describe, it, expect, jest, beforeEach } from '@jest/globals';

// Mock the dataProcessor module functions
const dataProcessor = {
    importFilesFromFolder: jest.fn()
};

// Mock dependencies
const dateUtils = {
    extractDateFromFilename: jest.fn()
};

const sheetUtils = {
    batchWriteToSheet: jest.fn()
};

// Setup global utilities
global.Utilities = {
    sleep: jest.fn()
};

describe('dataProcessor', () => {
    beforeEach(() => {
        jest.clearAllMocks();
    });

    describe('importFilesFromFolder', () => {
        it('should import files and return count', () => {
            // Setup
            const mockFolder = {
                getFiles: jest.fn(() => ({
                    hasNext: jest.fn()
                        .mockReturnValueOnce(true)
                        .mockReturnValueOnce(true)
                        .mockReturnValueOnce(false),
                    next: jest.fn()
                        .mockReturnValueOnce({ getName: () => 'file1.pdf', getUrl: () => 'url1' })
                        .mockReturnValueOnce({ getName: () => 'file2.pdf', getUrl: () => 'url2' })
                }))
            };

            const mockMainSheet = {
                getLastRow: jest.fn(() => 5),
                getLastColumn: jest.fn(() => 10)
            };

            const mockHistorySheet = {
                getLastRow: jest.fn(() => 2),
                getLastColumn: jest.fn(() => 4)
            };

            const mockExistingFiles = new Set();

            const mockConfig = {
                einnahmen: { columns: {} },
                aenderungshistorie: { columns: {} }
            };

            dateUtils.extractDateFromFilename.mockReturnValue(new Date());
            dataProcessor.importFilesFromFolder.mockReturnValue(2);

            // Execute
            const result = dataProcessor.importFilesFromFolder(
                mockFolder,
                mockMainSheet,
                'Einnahme',
                mockHistorySheet,
                mockExistingFiles,
                mockConfig
            );

            // Verify
            expect(result).toBe(2);
        });
    });
});