const path = require('path');
const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  ImageRun, Header, Footer, AlignmentType, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, PageBreak, HeadingLevel,
  TabStopType, TabStopPosition,
} = require('docx');

const BORDER_NONE = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const BORDER_THIN = { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' };
const BORDER_MED  = { style: BorderStyle.SINGLE, size: 4, color: '1a1a1a' };
const NO_BORDERS  = { top: BORDER_NONE, bottom: BORDER_NONE, left: BORDER_NONE, right: BORDER_NONE };

function cell(children, opts = {}) {
  return new TableCell({
    borders: opts.borders || NO_BORDERS,
    width: opts.width ? { size: opts.width, type: WidthType.DXA } : undefined,
    shading: opts.shade ? { fill: opts.shade, type: ShadingType.CLEAR } : undefined,
    verticalAlign: opts.valign || VerticalAlign.TOP,
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    columnSpan: opts.span,
    children,
  });
}

function para(text, opts = {}) {
  return new Paragraph({
    alignment: opts.align || AlignmentType.LEFT,
    spacing: { before: opts.spaceBefore || 0, after: opts.spaceAfter || 0 },
    children: [new TextRun({
      text,
      bold: opts.bold,
      size: opts.size || 20,
      font: 'Arial',
      color: opts.color || '000000',
    })],
  });
}

function labelValue(label, value) {
  return new Paragraph({
    spacing: { before: 40, after: 40 },
    children: [
      new TextRun({ text: `${label}: `, bold: true, size: 20, font: 'Arial' }),
      new TextRun({ text: value || '', size: 20, font: 'Arial' }),
    ],
  });
}

module.exports = async function handler(req, res) {
  if (req.method === 'OPTIONS') {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
    return res.status(200).end();
  }
  if (req.method !== 'POST') return res.status(405).end();
  res.setHeader('Access-Control-Allow-Origin', '*');

  const { hvItems = [], lvItems = [], meta = {} } = req.body;
  const today = new Date().toLocaleDateString('en-CA', { year: 'numeric', month: 'long', day: 'numeric' });
  const allItems = [...hvItems, ...lvItems];

  // Logo
  const logoPath = path.join(__dirname, 'logo.png');
  const logoData = fs.readFileSync(logoPath);

  // ── PAGE 1: Quotation letter ──────────────────────────────────────────────

  const headerChildren = [
    new Paragraph({
      alignment: AlignmentType.LEFT,
      children: [new ImageRun({ type: 'png', data: logoData, transformation: { width: 200, height: 61 }, altText: { title: 'Logo', description: 'AC Tesla logo', name: 'logo' } })],
      spacing: { after: 0 },
    }),
  ];

  const footerChildren = [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      border: { top: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' } },
      spacing: { before: 80 },
      children: [
        new TextRun({ text: '3348 Harvester Rd., Unit #2  Burlington, Ontario  L7N 3M8  Tel. (905) 637-0637  Fax. (905) 637-0655', size: 16, font: 'Arial', color: '666666' }),
      ],
    }),
  ];

  // Meta info rows
  const infoRows = [
    ['TO:', meta.to || ''],
    ['ATTENTION:', meta.attention || ''],
    ['ACTQ #:', meta.actqNo || ''],
    ['DATE:', today],
    ['RE:', meta.re || 'Electrical Bill of Materials'],
  ].map(([label, value]) =>
    new TableRow({ children: [
      cell([new Paragraph({ children: [new TextRun({ text: label, bold: true, size: 20, font: 'Arial' })] })], { width: 1440 }),
      cell([new Paragraph({ children: [new TextRun({ text: value, size: 20, font: 'Arial' })] })], { width: 7920 }),
    ]}),
  );

  // BOM items table rows for page 1 (summary)
  const itemHeaderRow = new TableRow({
    children: [
      cell([para('ITEM', { bold: true, size: 18 })], { width: 720, shade: 'E8E8E8', borders: { top: BORDER_MED, bottom: BORDER_MED, left: BORDER_MED, right: BORDER_THIN } }),
      cell([para('DESCRIPTION', { bold: true, size: 18 })], { width: 7200, shade: 'E8E8E8', borders: { top: BORDER_MED, bottom: BORDER_MED, left: BORDER_THIN, right: BORDER_THIN } }),
      cell([para('PRICE', { bold: true, size: 18, align: AlignmentType.RIGHT })], { width: 1440, shade: 'E8E8E8', borders: { top: BORDER_MED, bottom: BORDER_MED, left: BORDER_THIN, right: BORDER_MED } }),
    ],
  });

  const itemRows = allItems.map((item, i) =>
    new TableRow({ children: [
      cell([para(String(i + 1), { size: 18 })], { width: 720, borders: { top: BORDER_NONE, bottom: BORDER_THIN, left: BORDER_MED, right: BORDER_THIN } }),
      cell([new Paragraph({ spacing: { before: 40, after: 40 }, children: [
        new TextRun({ text: item.description || '', size: 18, font: 'Arial' }),
        item.tag_ref ? new TextRun({ text: `  [${item.tag_ref}]`, size: 16, font: 'Arial', color: '888888' }) : new TextRun(''),
      ]})], { width: 7200, borders: { top: BORDER_NONE, bottom: BORDER_THIN, left: BORDER_THIN, right: BORDER_THIN } }),
      cell([para('', { align: AlignmentType.RIGHT, size: 18 })], { width: 1440, borders: { top: BORDER_NONE, bottom: BORDER_THIN, left: BORDER_THIN, right: BORDER_MED } }),
    ]}),
  );

  const totalRow = new TableRow({ children: [
    cell([para('', { size: 18 })], { width: 720, borders: { top: BORDER_MED, bottom: BORDER_MED, left: BORDER_MED, right: BORDER_THIN } }),
    cell([new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: 'TOTAL (EXCLUDING HST)', bold: true, size: 20, font: 'Arial' })] })], { width: 7200, borders: { top: BORDER_MED, bottom: BORDER_MED, left: BORDER_THIN, right: BORDER_THIN } }),
    cell([para('$', { bold: true, size: 20, align: AlignmentType.RIGHT })], { width: 1440, borders: { top: BORDER_MED, bottom: BORDER_MED, left: BORDER_THIN, right: BORDER_MED } }),
  ]});

  // ── PAGE 2: Full BOM Table ─────────────────────────────────────────────────

  function bomSection(title, items, shade) {
    if (!items.length) return [];
    const cols = [500, 900, 2500, 500, 500, 1200, 1200, 1200, 1160];
    const headers = ['#', 'Tag/Ref', 'Description', 'Qty', 'Unit', 'Category', 'Manufacturer', 'Part/Model', 'Voltage/Rating'];

    const hdrRow = new TableRow({ children: headers.map((h, i) =>
      cell([para(h, { bold: true, size: 16 })], {
        width: cols[i],
        shade,
        borders: { top: BORDER_MED, bottom: BORDER_MED, left: i === 0 ? BORDER_MED : BORDER_THIN, right: i === cols.length - 1 ? BORDER_MED : BORDER_THIN },
      }),
    )});

    const dataRows = items.map((item, idx) =>
      new TableRow({ children: [
        cell([para(String(idx + 1), { size: 16 })], { width: cols[0], borders: { top: BORDER_NONE, bottom: BORDER_THIN, left: BORDER_MED, right: BORDER_THIN } }),
        cell([para(item.tag_ref || '—', { size: 16 })], { width: cols[1], borders: { top: BORDER_NONE, bottom: BORDER_THIN, left: BORDER_THIN, right: BORDER_THIN } }),
        cell([para(item.description || '—', { size: 16 })], { width: cols[2], borders: { top: BORDER_NONE, bottom: BORDER_THIN, left: BORDER_THIN, right: BORDER_THIN } }),
        cell([para(String(item.quantity ?? '—'), { size: 16 })], { width: cols[3], borders: { top: BORDER_NONE, bottom: BORDER_THIN, left: BORDER_THIN, right: BORDER_THIN } }),
        cell([para(item.unit || 'EA', { size: 16 })], { width: cols[4], borders: { top: BORDER_NONE, bottom: BORDER_THIN, left: BORDER_THIN, right: BORDER_THIN } }),
        cell([para(item.category || '—', { size: 16 })], { width: cols[5], borders: { top: BORDER_NONE, bottom: BORDER_THIN, left: BORDER_THIN, right: BORDER_THIN } }),
        cell([para(item.manufacturer || '—', { size: 16 })], { width: cols[6], borders: { top: BORDER_NONE, bottom: BORDER_THIN, left: BORDER_THIN, right: BORDER_THIN } }),
        cell([para(item.part_model || '—', { size: 16 })], { width: cols[7], borders: { top: BORDER_NONE, bottom: BORDER_THIN, left: BORDER_THIN, right: BORDER_THIN } }),
        cell([para(item.voltage_rating || '—', { size: 16 })], { width: cols[8], borders: { top: BORDER_NONE, bottom: BORDER_THIN, left: BORDER_THIN, right: BORDER_MED } }),
      ]}),
    );

    return [
      para(title, { bold: true, size: 22, spaceBefore: 160, spaceAfter: 80 }),
      new Table({ width: { size: 9660, type: WidthType.DXA }, columnWidths: cols, rows: [hdrRow, ...dataRows] }),
    ];
  }

  const doc = new Document({
    sections: [
      // ── Section 1: Quotation letter ──
      {
        properties: {
          page: {
            size: { width: 12240, height: 15840 },
            margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 },
          },
        },
        headers: { default: new Header({ children: headerChildren }) },
        footers: { default: new Footer({ children: footerChildren }) },
        children: [
          para('QUOTATION', { bold: true, size: 36, align: AlignmentType.CENTER, spaceAfter: 240 }),

          new Table({
            width: { size: 9360, type: WidthType.DXA },
            columnWidths: [1440, 7920],
            rows: infoRows,
          }),

          para('', { spaceAfter: 160 }),

          new Table({
            width: { size: 9360, type: WidthType.DXA },
            columnWidths: [720, 7200, 1440],
            rows: [itemHeaderRow, ...itemRows, totalRow],
          }),

          para('', { spaceBefore: 240 }),
          para('Should you have any questions please contact us at your convenience.', { spaceAfter: 80, size: 20 }),
          para('Best regards,', { spaceAfter: 160, size: 20 }),
          para('Jeff Li', { bold: true, size: 20 }),
          para('AC TESLA Inc.', { size: 20 }),

          // Page break before BOM
          new Paragraph({ children: [new PageBreak()] }),
        ],
      },
      // ── Section 2: Full BOM ──
      {
        properties: {
          page: {
            size: { width: 15840, height: 12240 },
            margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 },
          },
        },
        headers: { default: new Header({ children: headerChildren }) },
        footers: { default: new Footer({ children: footerChildren }) },
        children: [
          para('BILL OF MATERIALS', { bold: true, size: 32, align: AlignmentType.CENTER, spaceAfter: 80 }),
          para(`Project: ${meta.re || ''}`, { size: 18, align: AlignmentType.CENTER, spaceAfter: 160, color: '666666' }),
          ...bomSection('High Voltage Equipment', hvItems, 'FAECE7'),
          ...bomSection('Low Voltage Equipment', lvItems, 'E6F1FB'),
        ],
      },
    ],
  });

  const buffer = await Packer.toBuffer(doc);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
  res.setHeader('Content-Disposition', 'attachment; filename="quotation.docx"');
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.send(buffer);
};
