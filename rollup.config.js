// file: rollup.config.js
import resolve from '@rollup/plugin-node-resolve';
import terser from '@rollup/plugin-terser';

export default {
    input: 'src/code.js',       // Dein Haupt-Einstiegspunkt
    output: {
        file: 'dist/code.js',     // Gebündelte Datei für Apps Script
        format: 'cjs',            // Verwende CommonJS-Format
        banner: '/* Bundled Code for Google Apps Script */',
    },
    treeshake: false,
    plugins: [
        resolve(),                // Auflösen von Modul-Importen
        // terser({                  // Minimieren des Codes
        //     format: {
        //         comments: false,  // Entfernt alle Kommentare
        //     },
        //     compress: {
        //         drop_console: false, // Behält console.log für Debugging
        //         drop_debugger: true  // Entfernt debugger-Anweisungen
        //     },
        //     mangle: {
        //         reserved: ['global'] // Schützt 'global' vor Umbenennung (wichtig für Apps Script)
        //     }
        // })
    ]
};