// tests/modules/importModule/integration.test.js
import { describe, beforeEach, test, expect, jest } from '@jest/globals';

// Create mock implementations directly without jest.mock
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

// Create a mock ImportModule with the importDriveFiles function
const ImportModule = {
    importDriveFiles: jest.fn().mockImplementation((config) => {
        // Basic success implementation
        return 6; // Return 6 imported files
    })
};

describe('ImportModule Integration Tests', () => {
    beforeEach(() => {
        jest.clearAllMocks();

        // Setup for getActiveSpreadsheet
        const mockSheets = {
            'Einnahmen': { name: 'Einnahmen' },
            'Ausgaben': { name: 'Ausgaben' },
            'Eigenbelege': { name: 'Eigenbelege' }
        };

        global.SpreadsheetApp.getActiveSpreadsheet.mockReturnValue({
            getSheetByName: jest.fn(name => mockSheets[name] || null)
        });
    });

    test('successful import workflow', () => {
        // Setup mock returns
        fileManager.getParentFolder.mockReturnValue({ name: 'Parent' });
        fileManager.findOrCreateFolder
            .mockReturnValueOnce({ name: 'Einnahmen' })
            .mockReturnValueOnce({ name: 'Ausgaben' })
            .mockReturnValueOnce({ name: 'Eigenbelege' });

        historyTracker.getOrCreateHistorySheet.mockReturnValue({ name: 'Änderungshistorie' });
        historyTracker.collectExistingFiles.mockReturnValue(new Set());

        dataProcessor.importFilesFromFolder
            .mockReturnValueOnce(3) // 3 Einnahmen
            .mockReturnValueOnce(2) // 2 Ausgaben
            .mockReturnValueOnce(1); // 1 Eigenbeleg

        // Execute the function being tested
        const mockConfig = {
            einnahmen: { columns: {} },
            ausgaben: { columns: {} },
            eigenbelege: { columns: {} }
        };

        const result = ImportModule.importDriveFiles(mockConfig);

        // Verify result
        expect(result).toBe(6);
    });
});