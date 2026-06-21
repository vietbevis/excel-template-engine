import { mkdir, writeFile } from 'node:fs/promises';
import { ExcelTemplateEngine } from '../src/index.js';

const engine = new ExcelTemplateEngine();

engine.registerHelper('sum', ([values]) => {
  if (!Array.isArray(values)) {
    return 0;
  }

  return values.reduce((total, value) => total + Number(value), 0);
});

const data = {
  teacher: {
    name: 'Nguyễn Văn A',
    department: 'Khoa Công nghệ thông tin',
  },
  hasSignature: true,
  scores: [8, 9, 10],
  students: [
    { name: 'Trần Thị B', code: 'SV001', scores: [8, 9] },
    { name: 'Lê Văn C', code: 'SV002', scores: [7, 8] },
  ],
  subjects: [
    { name: 'Toán' },
    { name: 'Lý' },
    { name: 'Hóa' },
  ],
};

const result = await engine.render('examples/templates/class-report.xlsx', data, {
  preserveFormulas: true,
  recalculateFormulas: true,
});

await mkdir('examples/output', { recursive: true });
await writeFile('examples/output/class-report.xlsx', result.output);

if (result.warnings.length > 0) {
  console.warn(result.warnings.join('\n'));
}
