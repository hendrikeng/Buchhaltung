// tests/modules/importModule/fileManager.test.js
import { jest, describe, beforeEach, it, expect } from '@jest/globals';
import fileManager from '../../../src/modules/importModule/fileManager.js';

// Mocks für Google Apps Script API
const mockFolders = new Map();
const mockFiles = new Map();

// Hilfsfunktionen für Mock-Objekte
function createMockFolder(name, id = `folder-${Math.random().toString(36).substr(2, 9)}`) {
    const subfolders = new Map();
    const files = new Map();

    return {
        getId: jest.fn(() => id),
        getName: jest.fn(() => name),
        getParents: jest.fn(() => ({
            hasNext: jest.fn(() => mockFolders.size > 0),
            next: jest.fn(() => Array.from(mockFolders.values())[0])
        })),
        getFoldersByName: jest.fn((folderName) => {
            const matches = Array.from(subfolders.values()).filter(f => f.getName() === folderName);
            let index = 0;
            return {
                hasNext: jest.fn(() => index < matches.length),
                next: jest.fn(() => matches[index++])
            };
        }),
        createFolder: jest.fn((name) => {
            const folder = createMockFolder(name);
            subfolders.set(folder.getId(), folder);
            return folder;
        }),
        getFiles: jest.fn(() => {
            let fileIndex = 0;
            const fileArray = Array.from(files.values());
            return {
                hasNext: jest.fn(() => fileIndex < fileArray.length),
                next: jest.fn(() => fileArray[fileIndex++])
            };
        })
    };
}

// Hilfsfunktion zum Zurücksetzen aller Mocks
function resetMocks() {
    mockFolders.clear();
    mockFiles.clear();
    jest.clearAllMocks();
}

describe('fileManager', () => {
    beforeEach(() => {
        resetMocks();
    });

    describe('getParentFolder', () => {
        it('should return the parent folder of the active spreadsheet', () => {
            // Einrichten der Mocks
            const parentFolder = createMockFolder('Parent');
            mockFolders.set(parentFolder.getId(), parentFolder);

            const ssFile = {
                getId: jest.fn(() => 'spreadsheet123'),
                getParents: jest.fn(() => ({
                    hasNext: jest.fn(() => true),
                    next: jest.fn(() => parentFolder)
                }))
            };

            mockFiles.set('spreadsheet123', ssFile);

            // Override für diese Test-Funktion
            DriveApp.getFileById = jest.fn(() => ssFile);

            // Test
            const result = fileManager.getParentFolder();

            // Überprüfungen
            expect(DriveApp.getFileById).toHaveBeenCalledWith('spreadsheet123');
            expect(result).toBe(parentFolder);
        });

        it('should return null if no parent folder is found', () => {
            // Einrichten der Mocks
            const ssFile = {
                getId: jest.fn(() => 'spreadsheet123'),
                getParents: jest.fn(() => ({
                    hasNext: jest.fn(() => false),
                    next: jest.fn(() => null)
                }))
            };

            mockFiles.set('spreadsheet123', ssFile);

            // Override für diese Test-Funktion
            DriveApp.getFileById = jest.fn(() => ssFile);

            // Test
            const result = fileManager.getParentFolder();

            // Überprüfungen
            expect(result).toBeNull();
        });
    });

    describe('findOrCreateFolder', () => {
        it('should return an existing folder if it exists', () => {
            // Einrichten der Mocks
            const parentFolder = createMockFolder('Parent');
            const existingFolder = createMockFolder('Einnahmen');

            mockFolders.set(parentFolder.getId(), parentFolder);
            mockFolders.set(existingFolder.getId(), existingFolder);

            parentFolder.getFoldersByName = jest.fn(() => ({
                hasNext: jest.fn(() => true),
                next: jest.fn(() => existingFolder)
            }));

            const ui = SpreadsheetApp.getUi();

            // Test
            const result = fileManager.findOrCreateFolder(parentFolder, 'Einnahmen', ui);

            // Überprüfungen
            expect(result).toBe(existingFolder);
            expect(parentFolder.getFoldersByName).toHaveBeenCalledWith('Einnahmen');
            expect(parentFolder.createFolder).not.toHaveBeenCalled();
        });

        it('should create a new folder if it does not exist and user confirms', () => {
            // Einrichten der Mocks
            const parentFolder = createMockFolder('Parent');
            mockFolders.set(parentFolder.getId(), parentFolder);

            parentFolder.getFoldersByName = jest.fn(() => ({
                hasNext: jest.fn(() => false),
                next: jest.fn(() => null)
            }));

            const ui = {
                alert: jest.fn(() => 'yes'),
                Button: { YES: 'yes', NO: 'no' },
                ButtonSet: { YES_NO: 'yes_no' }
            };

            const newFolder = createMockFolder('Einnahmen');
            parentFolder.createFolder = jest.fn(() => newFolder);

            // Test
            const result = fileManager.findOrCreateFolder(parentFolder, 'Einnahmen', ui);

            // Überprüfungen
            expect(ui.alert).toHaveBeenCalled();
            expect(parentFolder.createFolder).toHaveBeenCalledWith('Einnahmen');
            expect(result).toBe(newFolder);
        });
    });
});