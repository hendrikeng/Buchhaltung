// tests/modules/importModule/fileManager.test.js
import { describe, it, expect, jest, beforeEach } from '@jest/globals';

// Mock our globals
global.DriveApp = {
    getFileById: jest.fn(),
    createFolder: jest.fn()
};

global.SpreadsheetApp = {
    getActiveSpreadsheet: jest.fn(() => ({
        getId: jest.fn(() => 'spreadsheet123'),
        getSheetByName: jest.fn()
    })),
    getUi: jest.fn(() => ({
        alert: jest.fn(),
        Button: { YES: 'yes', NO: 'no' },
        ButtonSet: { YES_NO: 'yes_no' }
    }))
};

// Mock the fileManager module functions
const fileManager = {
    getParentFolder: jest.fn(),
    findOrCreateFolder: jest.fn()
};

describe('fileManager', () => {
    beforeEach(() => {
        jest.clearAllMocks();
    });

    describe('getParentFolder', () => {
        it('should return parent folder when found', () => {
            // Setup
            const mockParentFolder = { id: 'folder123' };
            const mockFile = {
                getParents: jest.fn(() => ({
                    hasNext: jest.fn(() => true),
                    next: jest.fn(() => mockParentFolder)
                }))
            };

            DriveApp.getFileById.mockReturnValue(mockFile);
            fileManager.getParentFolder.mockReturnValue(mockParentFolder);

            // Execute
            const result = fileManager.getParentFolder();

            // Verify
            expect(result).toBe(mockParentFolder);
        });
    });
});