import { mkdir,copyFile } from 'node:fs/promises';
await mkdir('public/vendor',{recursive:true});
await copyFile('node_modules/exceljs/dist/exceljs.min.js','public/vendor/exceljs.min.js');
console.log('BSI assets ready. No bank data included.');
