// modules/importModule/fileManager.js
import sheetUtils from '../../utils/sheetUtils.js';

/**
 * Ruft den übergeordneten Ordner des aktuellen Spreadsheets ab
 * Optimierte Version mit besserer Fehlerbehandlung
 * @returns {Folder|null} - Der übergeordnete Ordner oder null bei Fehler
 */
function getParentFolder() {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const file = DriveApp.getFileById(ss.getId());
        const parents = file.getParents();
        return parents.hasNext() ? parents.next() : null;
    } catch (e) {
        console.error('Fehler beim Abrufen des übergeordneten Ordners:', e);
        return null;
    }
}

/**
 * Findet oder erstellt einen Ordner mit dem angegebenen Namen
 * Optimierte Version mit Caching und besserer Fehlerbehandlung
 * @param {Folder} parentFolder - Der übergeordnete Ordner
 * @param {string} folderName - Der zu findende oder erstellende Ordnername
 * @param {UI} ui - UI-Objekt für Dialoge
 * @returns {Folder|null} - Der gefundene oder erstellte Ordner oder null bei Fehler
 */
function findOrCreateFolder(parentFolder, folderName, ui) {
    if (!parentFolder) return null;

    try {
        // Optimierung: Versuche zuerst direkten Zugriff ohne Suche
        let folder = sheetUtils.getFolderByName(parentFolder, folderName);

        if (!folder) {
            const createFolder = ui.alert(
                `Der Ordner '${folderName}' existiert nicht. Soll er erstellt werden?`,
                ui.ButtonSet.YES_NO,
            );

            if (createFolder === ui.Button.YES) {
                try {
                    folder = parentFolder.createFolder(folderName);
                    console.log(`Folder ${folderName} created successfully`);
                } catch (folderError) {
                    console.error(`Error creating folder ${folderName}:`, folderError);
                    ui.alert(`Fehler beim Erstellen des Ordners '${folderName}': ${folderError.message}`);
                    return null;
                }
            } else {
                // Benutzer hat abgelehnt, keinen Ordner erstellen
                return null;
            }
        }

        return folder;
    } catch (e) {
        console.error(`Fehler beim Finden/Erstellen des Ordners ${folderName}:`, e);
        ui.alert(`Ein Fehler ist beim Bearbeiten des Ordners '${folderName}' aufgetreten: ${e.message}`);
        return null;
    }
}

export default {
    getParentFolder,
    findOrCreateFolder,
};