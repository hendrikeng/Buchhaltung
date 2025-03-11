import resolve from '@rollup/plugin-node-resolve';

export default {
    input: 'src/code.js',       // Dein Haupt-Einstiegspunkt
    output: {
        file: 'dist/code.js',     // Gebündelte Datei für Apps Script
        format: 'cjs',            // Verwende CommonJS-Format statt IIFE
        banner: '/* Bundled Code for Google Apps Script */',
    },
    treeshake: false,           // Deaktiviert Tree Shaking, falls du sicherstellen willst, dass nichts entfernt wird
    plugins: [resolve()]
};
