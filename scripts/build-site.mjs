import { mkdir, copyFile } from 'node:fs/promises';

// Publish only the form, keeping Apps Script source and internal files private.
const output = new URL('../public/', import.meta.url);
await mkdir(output, { recursive: true });
await copyFile(new URL('../Asas_Sales_Form.html', import.meta.url), new URL('index.html', output));
console.log('Static site ready: public/index.html');
