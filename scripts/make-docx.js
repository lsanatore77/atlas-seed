// Builds one Word document per controlled document from the JSON that build_document_files.py writes.
// Usage: node scripts/make-docx.js <spec.json> <output folder>
// docx-js keeps the files small (about 10 KB), so link-document-files.ps1 can send each in one Apex run.
const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, ShadingType,
  AlignmentType, Footer, PageNumber, LevelFormat, BorderStyle,
} = require('docx');

const [, , specPath, outDir] = process.argv;
const specs = JSON.parse(fs.readFileSync(specPath, 'utf8'));
const NAVY = '1F3A5F', GREY = '5A6470', LINE = 'C9D1DB', FONT = 'Arial';
const CONTENT_W = 10080; // 8.5in page minus 0.75in margins each side, in DXA

const t = (text, o = {}) => new TextRun({ text, font: FONT, size: (o.size || 10) * 2, bold: !!o.bold, color: o.color });
const p = (runs, o = {}) => new Paragraph({
  children: Array.isArray(runs) ? runs : [runs],
  spacing: { before: o.before || 0, after: o.after === undefined ? 80 : o.after },
  alignment: o.align, keepNext: !!o.keepNext, numbering: o.numbering,
});
const border = { style: BorderStyle.SINGLE, size: 4, color: LINE };
const borders = { top: border, bottom: border, left: border, right: border };

function table(rows, widths) {
  return new Table({
    width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    columnWidths: widths,
    rows: rows.map((row, i) => new TableRow({
      tableHeader: i === 0,
      children: row.map((val, j) => new TableCell({
        width: { size: widths[j], type: WidthType.DXA },
        borders,
        shading: i === 0 ? { type: ShadingType.CLEAR, color: 'auto', fill: NAVY } : undefined,
        margins: { top: 40, bottom: 40, left: 80, right: 80 },
        children: [p(t(String(val), { size: 9, bold: i === 0, color: i === 0 ? 'FFFFFF' : undefined }), { after: 0 })],
      })),
    })),
  });
}

function heading(text) {
  return p(t(text, { size: 11, bold: true, color: NAVY }), { before: 160, after: 60, keepNext: true });
}

function build(d) {
  const children = [];
  // Header band: company name and record type
  const half = CONTENT_W / 2;
  children.push(new Table({
    width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: [half, half],
    rows: [new TableRow({ children: [
      ['ATLAS MANUFACTURING', AlignmentType.LEFT], [d.recordType, AlignmentType.RIGHT],
    ].map(([text, align]) => new TableCell({
      width: { size: half, type: WidthType.DXA },
      shading: { type: ShadingType.CLEAR, color: 'auto', fill: NAVY },
      borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } },
      margins: { top: 80, bottom: 80, left: 120, right: 120 },
      children: [p(t(text, { size: 9, bold: true, color: 'FFFFFF' }), { after: 0, align })],
    })) })],
  }));
  children.push(p(t(`${d.number}  ${d.title}`, { size: 15, bold: true, color: NAVY }), { before: 160, after: 40 }));
  children.push(p(t(`Revision ${d.revision}  ·  Status: ${d.status}  ·  Effective: ${d.effective || 'not yet released'}`,
    { size: 8.5, color: GREY }), { after: 120 }));
  if (d.status === 'Draft') {
    children.push(p(t('DRAFT – not released for use', { size: 11, bold: true, color: 'C0392B' }), { after: 120 }));
  }
  for (const s of d.sections) {
    children.push(heading(s.heading));
    if (s.paragraphs) s.paragraphs.forEach((x) => children.push(p(t(x))));
    if (s.steps) s.steps.forEach((x) => children.push(p(t(x), { after: 60, numbering: { reference: 'steps-' + s.id, level: 0 } })));
    if (s.table) { children.push(table(s.table, [3000, 3000, 4080])); children.push(p(t(''), { after: 40 })); }
  }
  children.push(heading('Revision history'));
  children.push(table([['Rev', 'Date', 'Change'],
    d.effective ? [d.revision, d.effective, 'Current release.'] : [d.revision, '', 'New document, in draft.']], [900, 1600, 7580]));

  // One numbering reference per steps section, so each list starts at 1
  const config = d.sections.filter((s) => s.steps).map((s) => ({
    reference: 'steps-' + s.id,
    levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 360, hanging: 360 } }, run: { font: FONT, size: 20 } } }],
  }));

  return new Document({
    creator: 'Atlas Manufacturing (test data)', title: `${d.number} ${d.title}`,
    styles: { default: { document: { run: { font: FONT, size: 20 } } } },
    numbering: { config },
    sections: [{
      properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 864, bottom: 1008, left: 1080, right: 1080 } } },
      footers: { default: new Footer({ children: [p([
        t('Atlas Manufacturing is a fictional company (test data). Uncontrolled when printed: check CQ for the current revision.   Page ', { size: 7.5, color: GREY }),
        new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 15, color: GREY }),
      ], { after: 0 })] }) },
      children,
    }],
  });
}

(async () => {
  for (const d of specs) {
    const buf = await Packer.toBuffer(build(d));
    fs.writeFileSync(path.join(outDir, d.file), buf);
    console.log(`${d.number}: ${buf.length} bytes`);
  }
  console.log(`${specs.length} files in ${outDir}`);
})();
