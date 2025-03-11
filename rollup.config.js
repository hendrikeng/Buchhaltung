import resolve from '@rollup/plugin-node-resolve';
import fs from 'fs';
import path from 'path';

// Einfaches Plugin, um @include-Anweisungen zu verarbeiten
function includePlugin() {
    return {
        name: 'include-plugin',
        transform(code, id) {
            // @include "filename.js" Muster suchen
            const includePattern = /\/\/\s*@include\s+"([^"]+)"/g;

            if (!includePattern.test(code)) {
                return null; // Code unverändert lassen, wenn keine @include-Anweisungen vorhanden sind
            }

            // @include-Anweisungen durch den tatsächlichen Dateiinhalt ersetzen
            let result = code.replace(includePattern, (match, filename) => {
                try {
                    const includePath = path.resolve(path.dirname(id), filename);
                    const fileContent = fs.readFileSync(includePath, 'utf8');
                    return fileContent;
                } catch (error) {
                    console.error(`Fehler beim Einbinden von ${filename}:`, error);
                    return `// Fehler beim Einbinden von ${filename}: ${error.message}`;
                }
            });

            return {
                code: result,
                map: { mappings: '' }
            };
        }
    };
}

export default {
    input: 'src/code.js',
    output: {
        file: 'dist/code.js',
        format: 'cjs',
        banner: '/* Bundled Code for Google Apps Script */'
    },
    treeshake: false, // Tree Shaking deaktivieren, damit alle Funktionen erhalten bleiben
    plugins: [
        includePlugin(),
        resolve()
    ]
};