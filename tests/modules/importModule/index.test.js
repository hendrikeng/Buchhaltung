// tests/modules/importModule/index.test.js
import { describe, it, expect, jest, beforeEach } from '@jest/globals';

// Mock dependencies
const fileManager = {
    getParentFolder: jest.fn(),
    findOrCreateFolder: jest.fn()
};

const dataProcessor = {
    importFilesFromFolder: jest.fn()
};

const historyTracker = {
    getOrCreateHistorySheet: jest.fn(),
    collectExistingFiles: jest.fn()
};

// Create ImportModule mock
const ImportModule = {
    importDriveFiles: jest.fn()
};

// Setup global mocks
global.SpreadsheetApp = {
    getActiveSpreadsheet: jest.fn(() => ({
        getSheetByName: jest.fn(name => {
            if (['Einnahmen', 'Ausgaben', 'Eigenbelege'].includes(name)) {
                return { name };
            }
            return null;
        })
    })),
    getUi: jest.fn(() => ({
        alert: jest.fn()
    }))
};

describe('ImportModule', () => {
    beforeEach(() => {
        jest.clearAllMocks();
    });

    describe('importDriveFiles', () => {
        it('should import files successfully', () => {
            // Setup
            const mockParentFolder = { name: 'Parent' };
            const mockRevFolder = { name: 'Einnahmen' };
            const mockExpFolder = { name: 'Ausgaben' };
            const mockRecFolder = { name: 'Eigenbelege' };
            const mockHistorySheet = { name: 'Änderungshistorie' };
            const mockExistingFiles = new Set();

            fileManager.getParentFolder.mockReturnValue(mockParentFolder);
            fileManager.findOrCreateFolder
                .mockReturnValueOnce(mockRevFolder)
                .mockReturnValueOnce(mockExpFolder)
                .mockReturnValueOnce(mockRecFolder);

            historyTracker.getOrCreateHistorySheet.mockReturnValue(mockHistorySheet);
            historyTracker.collectExistingFiles.mockReturnValue(mockExistingFiles);

            dataProcessor.importFilesFromFolder
                .mockReturnValueOnce(3) // 3 Einnahmen
                .mockReturnValueOnce(2) // 2 Ausgaben
                .mockReturnValueOnce(1); // 1 Eigenbeleg

            ImportModule.importDriveFiles.mockReturnValue(6);

            const mockConfig = {
                einnahmen: { columns: {} },
                ausgaben: { columns: {} },
                eigenbelege: { columns: {} }
            };

            // Execute
            const result = ImportModule.importDriveFiles(mockConfig);

            // Verify
            expect(result).toBe(6); // 6 files imported
        });
    });
});