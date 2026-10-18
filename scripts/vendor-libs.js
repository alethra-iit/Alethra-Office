// Menyalin library pihak ketiga dari node_modules ke www/lib supaya aplikasi 100% offline.
// Jalankan: npm install && npm run vendor
const fs = require('fs'), path = require('path');
const nm = p => path.join(__dirname, '..', 'node_modules', p);
const out = path.join(__dirname, '..', 'www', 'lib');
fs.mkdirSync(out, { recursive: true });
const files = {
  'mammoth.browser.min.js':        'mammoth/mammoth.browser.min.js',
  'docx.iife.js':                  'docx/dist/index.iife.js',
  'exceljs.min.js':                'exceljs/dist/exceljs.min.js',
  'pdf.min.js':                    'pdfjs-dist/build/pdf.min.js',
  'pdf.worker.min.js':             'pdfjs-dist/build/pdf.worker.min.js',
  'pdf-lib.min.js':                'pdf-lib/dist/pdf-lib.min.js',
  'pdf-lib-plus-encrypt.min.js':   'pdf-lib-plus-encrypt/dist/pdf-lib-plus-encrypt.min.js',
  'xlsx.full.min.js':              'xlsx/dist/xlsx.full.min.js',
};
for (const [dst, src] of Object.entries(files)) {
  fs.copyFileSync(nm(src), path.join(out, dst));
  console.log('lib/' + dst);
}
