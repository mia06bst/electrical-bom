const path = require('path');
const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  ImageRun, Header, Footer, AlignmentType, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, PageBreak, HeadingLevel,
  TabStopType,
} = require('docx');

const FONT = 'Times New Roman';
const NONE = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const NO_BORDERS = { top: NONE, bottom: NONE, left: NONE, right: NONE };
const THIN = { style: BorderStyle.SINGLE, size: 4, color: '000000' };
const DOUBLE_BLUE = { style: BorderStyle.DOUBLE, size: 4, color: '0099CC' };

function p(children, opts = {}) {
  return new Paragraph({
    alignment: opts.align || AlignmentType.LEFT,
    spacing: { before: opts.before || 0, after: opts.after || 80 },
    border: opts.border,
    children: Array.isArray(children) ? children : [children],
  });
}

function run(text, opts = {}) {
  return new TextRun({ text, font: FONT, size: opts.size || 20, bold: opts.bold, underline: opts.underline, color: opts.color });
}

function cell(children, opts = {}) {
  return new TableCell({
    borders: opts.borders || NO_BORDERS,
    width: opts.width ? { size: opts.width, type: WidthType.DXA } : undefined,
    shading: opts.shade ? { fill: opts.shade, type: ShadingType.CLEAR } : undefined,
    verticalAlign: opts.valign || VerticalAlign.TOP,
    margins: { top: opts.mt || 60, bottom: opts.mb || 60, left: opts.ml || 100, right: opts.mr || 100 },
    columnSpan: opts.span,
    rowSpan: opts.rowSpan,
    children,
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
  const today = new Date().toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' });
  const allItems = [...hvItems, ...lvItems];

  const logoData = fs.readFileSync(path.join(__dirname, 'logo.png'));

  // ── HEADER: logo left, company info right, blue double bottom border ──
  const headerTable = new Table({
    width: { size: 12240, type: WidthType.DXA },
    columnWidths: [5000, 7240],
    borders: { top: NONE, bottom: DOUBLE_BLUE, left: NONE, right: NONE, insideH: NONE, insideV: NONE },
    rows: [new TableRow({ children: [
      cell([new Paragraph({ children: [new ImageRun({ type: 'png', data: logoData, transformation: { width: 180, height: 55 }, altText: { title: 'Logo', description: 'AC Tesla', name: 'logo' } })] })], { width: 5000, mb: 120 }),
      cell([
        p(run('AC TESLA INC.', { bold: true, size: 20 }), { align: AlignmentType.RIGHT, after: 0 }),
        p(run('3348 Harvester Rd., Unit #2', { size: 18 }), { align: AlignmentType.RIGHT, after: 0 }),
        p(run('Burlington, Ontario  L7N 3M8', { size: 18 }), { align: AlignmentType.RIGHT, after: 0 }),
        p(run('Tel. (905) 637-0637   Fax. (905) 637-0655', { size: 18 }), { align: AlignmentType.RIGHT, after: 0 }),
      ], { width: 7240, mb: 120 }),
    ]})],
  });

  // ── FOOTER ──
  const footer = new Footer({ children: [
    p(run('3348 Harvester Rd., Unit #2  Burlington, Ontario  L7N 3M8  Tel. (905) 637-0637  Fax. (905) 637-0655', { size: 16 }), { align: AlignmentType.CENTER }),
  ]});

  // ── PAGE 1: Quotation letter ──

  // Info block: TO / ATTENTION / ACTQ # / DATE / RE
  const infoTable = new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [1800, 7560],
    borders: { top: NONE, bottom: NONE, left: NONE, right: NONE, insideH: NONE, insideV: NONE },
    rows: [
      ['TO:', meta.to || ''],
      ['ATTENTION:', meta.attention || ''],
      ['ACTQ #:', meta.actqNo || ''],
      ['DATE:', today],
      ['RE:', meta.re || ''],
    ].map(([label, value]) => new TableRow({ children: [
      cell([p(run(label, { bold: true, size: 20 }), { after: 40 })], { width: 1800 }),
      cell([p(run(value, { size: 20 }), { after: 40 })], { width: 7560 }),
    ]})),
  });

  // Items table
  const itemBorderCell = (c, isFirst, isLast) => ({
    top: THIN, bottom: THIN,
    left: isFirst ? THIN : NONE,
    right: isLast ? THIN : NONE,
  });

  const itemHeaderRow = new TableRow({ children: [
    cell([p(run('ITEM', { bold: true, size: 20 }), { align: AlignmentType.CENTER })], { width: 720, shade: 'D9D9D9', borders: { top: THIN, bottom: THIN, left: THIN, right: NONE } }),
    cell([p(run('DESCRIPTION', { bold: true, size: 20 }), { align: AlignmentType.CENTER })], { width: 7200, shade: 'D9D9D9', borders: { top: THIN, bottom: THIN, left: NONE, right: NONE } }),
    cell([p(run('PRICE', { bold: true, size: 20 }), { align: AlignmentType.CENTER })], { width: 1440, shade: 'D9D9D9', borders: { top: THIN, bottom: THIN, left: NONE, right: THIN } }),
  ]});

  // Build description text from BOM items
  const descLines = allItems.map((it, i) => {
    const qty = it.quantity ? `${it.quantity} x ` : '';
    const tag = it.tag_ref ? ` (${it.tag_ref})` : '';
    return `${qty}${it.description}${tag}${it.voltage_rating ? ' – ' + it.voltage_rating : ''}`;
  }).join('\n');

  const descriptionPara = new Paragraph({
    alignment: AlignmentType.BOTH,
    spacing: { before: 60, after: 60 },
    children: [run(
      `Provide labour and material as per the electrical bill of materials for ${meta.re || 'the above-mentioned project'}. ` +
      `All detailed scope of work and equipment list specified on the next attached pages will be taken into consideration in the quotation.`,
      { size: 18 }
    )],
  });

  const notePara = new Paragraph({
    alignment: AlignmentType.BOTH,
    spacing: { before: 60, after: 40 },
    children: [run('Note:', { bold: true, size: 18 }), run(' Price is subject to change if new additional equipment is found to have an impact on the existing system or work is rescheduled to be done during afterhours, weekend or holidays.', { size: 18 })],
  });

  const itemRow = new TableRow({ children: [
    cell([p(run('1', { size: 20 }), { align: AlignmentType.CENTER })], { width: 720, borders: { top: NONE, bottom: NONE, left: THIN, right: NONE }, valign: VerticalAlign.TOP }),
    cell([descriptionPara, notePara], { width: 7200, borders: { top: NONE, bottom: NONE, left: NONE, right: NONE } }),
    cell([p(run('', { size: 20 }))], { width: 1440, borders: { top: NONE, bottom: NONE, left: NONE, right: THIN } }),
  ]});

  const totalRow = new TableRow({ children: [
    cell([p(run('', { size: 20 }))], { width: 720, borders: { top: THIN, bottom: THIN, left: THIN, right: NONE } }),
    cell([p([run('TOTAL (EXCLUDING HST)  ', { bold: true, size: 20 }), run('$', { bold: true, size: 20 })], { align: AlignmentType.RIGHT })], { width: 8640, borders: { top: THIN, bottom: THIN, left: NONE, right: THIN }, span: 2 }),
  ]});

  const itemsTable = new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [720, 7200, 1440],
    borders: { top: NONE, bottom: NONE, left: NONE, right: NONE, insideH: NONE, insideV: NONE },
    rows: [itemHeaderRow, itemRow, totalRow],
  });

  // ── SCOPE OF WORK section (generated from BOM) ──
  const scopeItems = [];

  // HV equipment list
  if (hvItems.length) {
    scopeItems.push(p(run('HV Equipment List:', { bold: true, size: 18 }), { after: 40 }));
    hvItems.forEach(it => {
      const tag = it.tag_ref ? `${it.tag_ref} – ` : '';
      scopeItems.push(new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { before: 0, after: 40 },
        indent: { left: 720 },
        children: [run(`${it.quantity || 1} x ${tag}${it.description}${it.voltage_rating ? ' (' + it.voltage_rating + ')' : ''}`, { size: 18 })],
      }));
    });
  }

  // LV equipment list
  if (lvItems.length) {
    scopeItems.push(p(run('LV Equipment List:', { bold: true, size: 18 }), { before: 80, after: 40 }));
    lvItems.forEach(it => {
      const tag = it.tag_ref ? `${it.tag_ref} – ` : '';
      scopeItems.push(new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { before: 0, after: 40 },
        indent: { left: 720 },
        children: [run(`${it.quantity || 1} x ${tag}${it.description}${it.voltage_rating ? ' (' + it.voltage_rating + ')' : ''}`, { size: 18 })],
      }));
    });
  }

  // ── TERMS AND CONDITIONS ──
  const terms = [
    'All prices are net and firm for thirty (30) days, subject to credit approval, with all applicable taxes extra.',
    'All traveling time, mileage, and expenses are included in the quoted price.',
    'Invoicing will be progressive on a monthly basis as charges accumulate.',
    'The quoted price is based on work being performed on regular time rates.',
    'All work will be performed by AC TESLA Inc. personnel and/or AC TESLA Inc. authorized subcontractors.',
    'It is an understanding that all switching and isolation will be done by others.',
    'AC TESLA Inc. will not be held responsible for delays caused by events beyond their reasonable control, including without limitation, strikes, lockouts and labor disputes. If a delay event happens, the time for performance will be extended by a period equal to the duration of the delay.',
    'The cost of any inspection authority approvals required for this project is not included in the quoted price.',
    'Any costs incurred to satisfy local union requirements will be charged as an extra to the quoted price.',
    'This proposal is based on the design submitted on the specification requirements. If any additions or deletions to these specifications are made, extra charges to the quoted price will apply.',
    'Prior to the start of any work, the customer will be required to provide free and easy access to the proposed work areas. This includes the removal of all materials stored around proposed work areas.',
  ];

  const doc = new Document({
    styles: {
      default: { document: { run: { font: FONT, size: 20 } } },
      paragraphStyles: [
        { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
          run: { size: 28, bold: false, font: FONT },
          paragraph: { alignment: AlignmentType.CENTER, spacing: { before: 120, after: 120 }, outlineLevel: 0 } },
      ],
    },
    sections: [{
      properties: {
        page: {
          size: { width: 12240, height: 15840 },
          margin: { top: 720, right: 1080, bottom: 720, left: 1080 },
        },
      },
      headers: { default: new Header({ children: [headerTable] }) },
      footers: { default: footer },
      children: [
        // QUOTATION title
        new Paragraph({ heading: HeadingLevel.HEADING_1, alignment: AlignmentType.CENTER, spacing: { before: 240, after: 240 }, children: [run('QUOTATION', { bold: false, size: 28 })] }),

        // Info block
        infoTable,

        p(run(''), { after: 120 }),

        // Items table
        itemsTable,

        p(run(''), { after: 120 }),

        // Closing
        p(run('Should you have any questions please contact us at your convenience.', { size: 20 }), { after: 80 }),
        p(run(''), { after: 80 }),
        p(run('Best regards,', { size: 20 }), { after: 40 }),
        p(run(''), { after: 80 }),
        p(run('Jeff Li', { bold: true, size: 20 }), { after: 0 }),

        // Page break → Scope of Work
        new Paragraph({ children: [new PageBreak()] }),

        // SCOPE OF WORK heading
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 40 }, children: [run('SCOPE OF WORK & EQUIPMENT LIST', { bold: true, size: 28 })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 160 }, children: [run(`(${meta.re || ''})`, { size: 20 })] }),

        // Equipment lists from BOM
        ...scopeItems,

        // End of section
        p(run(''), { after: 80 }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 160, after: 0 }, border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: '000000' } }, children: [run('')] }),
        p(run('End of section', { size: 18 }), { align: AlignmentType.LEFT, after: 0 }),

        // Page break → Terms
        new Paragraph({ children: [new PageBreak()] }),

        // TERMS AND CONDITIONS
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 160 }, children: [run('TERMS AND CONDITIONS', { bold: true, size: 20, underline: {} })] }),

        ...terms.map(t => new Paragraph({
          alignment: AlignmentType.BOTH,
          spacing: { before: 0, after: 120 },
          children: [run(t, { size: 20 })],
        })),
      ],
    }],
  });

  const buffer = await Packer.toBuffer(doc);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
  res.setHeader('Content-Disposition', 'attachment; filename="quotation.docx"');
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.send(buffer);
};
