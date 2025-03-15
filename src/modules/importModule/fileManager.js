// modules/importModule/fileManager.js
import sheetUtils from '../../utils/sheetUtils.js';

/**
 * Ruft den übergeordneten Ordner des aktuellen Spreadsheets ab
 * @returns {Folder|null} - Der übergeordnete Ordner oder null bei Fehler
 */
function getParentFolder() {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const file = DriveApp.getFileById(ss.getId());
        const parents = file.getParents();
        return parents.hasNext() ? parents.next() : null;
    } catch (e) {
        console.error("Fehler beim Abrufen des übergeordneten Ordners:", e);
        return null;
    }
}

/**
 * Findet oder erstellt einen Ordner mit dem angegebenen Namen
 * @param {Folder} parentFolder - Der übergeordnete Ordner
 * @param {string} folderName - Der zu findende oder erstellende Ordnername
 * @param {UI} ui - UI-Objekt für Dialoge
 * @returns {Folder|null} - Der gefundene oder erstellte Ordner oder null bei Fehler
 */
function findOrCreateFolder(parentFolder, folderName, ui) {
    if (!parentFolder) return null;

    try {
        let folder = sheetUtils.getFolderByName(parentFolder, folderName);

        if (!folder) {
            const createFolder = ui.alert(
                `Der Ordner '${folderName}' existiert nicht. Soll er erstellt werden?`,
                ui.ButtonSet.YES_NO
            );

            if (createFolder === ui.Button.YES) {
                folder = parentFolder.createFolder(folderName);
            }
        }

        return folder;
    } catch (e) {
        console.error(`Fehler beim Finden/Erstellen des Ordners ${folderName}:`, e);
        return null;
    }
}

export default {
    getParentFolder,
    findOrCreateFolder
};