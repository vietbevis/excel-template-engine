import { readFile, writeFile } from 'node:fs/promises';
import { ExcelTemplateEngine } from '../src/index.js';

const data = JSON.parse(await readFile('examples/data/production-showcase-data.json', 'utf8')) as Record<string, unknown>;

const engine = new ExcelTemplateEngine();
engine
  .registerHelper('currency', ([value]) => `${Number(value).toLocaleString('en-US')} VND`)
  .registerHelper('sum', ([values]) => Array.isArray(values)
    ? values.reduce((total, value) => total + Number(value), 0)
    : 0);

const result = await engine.render('examples/templates/production-showcase-template.xlsx', data, {
  limits: {
    maxTemplateBytes: 5 * 1024 * 1024,
    maxWorksheets: 10,
    maxRows: 1000,
    maxColumns: 100,
    maxOperations: 10000,
  },
  assetResolver: {
    maxBytes: 2 * 1024 * 1024,
  },
});

await writeFile('examples/output/production-showcase-output.xlsx', result.output);
console.log(JSON.stringify({
  output: 'examples/output/production-showcase-output.xlsx',
  warnings: result.warnings,
}, null, 2));
