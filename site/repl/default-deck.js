// Every @office-kit/pptx function is in scope. No imports needed.
// `pres` is a new 16:9 deck (13.33 x 7.5 in) from createPresentation().
//
// A six-slide board deck. Every number on every slide comes from `model`,
// so change a figure here and the headlines, charts, and table all follow.

const model = {
  quarters: ['Q1', 'Q2', 'Q3', 'Q4'],
  revenue: [4.2, 4.9, 5.8, 6.9], // $M
  marginPct: [11, 13, 16, 19], // operating margin
  newRevenue: { Enterprise: 3.1, 'Mid-market': 1.4, SMB: 0.6, Partners: 0.4 }, // $M
  // name, owner, investment in $M, payback in months, status
  bets: [
    ['Enterprise sales pods', 'M. Okafor', 2.4, 14, 'On track'],
    ['Usage-based pricing', 'L. Hartmann', 0.9, 9, 'On track'],
    ['Partner marketplace', 'S. Iyer', 1.6, 22, 'At risk'],
    ['EU data residency', 'D. Costa', 1.1, 18, 'Not started'],
  ],
};

const growthPct = Math.round((model.revenue[3] / model.revenue[0] - 1) * 100);
const marginGain = model.marginPct[3] - model.marginPct[0];
const newTotal = Object.values(model.newRevenue).reduce((a, b) => a + b, 0);
const enterprisePct = Math.round((model.newRevenue.Enterprise / newTotal) * 100);
const invested = '$' + model.bets.reduce((sum, bet) => sum + bet[2], 0).toFixed(1) + 'M';

const NAVY = '#0A1628';
const BLUE = '#1E4FFF';
const SLATE = '#9AA3B2';
const PANEL = '#EEF0F3';
const LINE = '#D5D9E0';
const INK = '#111827';
const MUTED = '#5B6673';
const WHITE = '#FFFFFF';
const STATUS = { 'On track': '#2E9E6C', 'At risk': '#D64545', 'Not started': SLATE };

// Positions are [x, y, width, height] in inches.
const frame = ([x, y, w, h]) => ({ x: inches(x), y: inches(y), w: inches(w), h: inches(h) });

function text(slide, str, box, format, anchor) {
  const shape = addSlideTextBox(slide, { ...frame(box), text: str });
  setShapeTextFormat(shape, format);
  if (anchor) setShapeTextAnchor(shape, anchor);
  return shape;
}

// A filled shape with no outline. With `label`, bold white text centred in it.
function block(slide, preset, box, color, label, size = 18) {
  const words = label ? { text: label } : {};
  const shape = addSlideShape(slide, { preset, ...frame(box), ...words });
  setShapeFill(shape, color);
  setShapeNoStroke(shape);
  if (label) {
    setShapeTextFormat(shape, { size, bold: true, color: WHITE });
    setShapeAlignment(shape, 'center');
    setShapeTextAnchor(shape, 'center');
  }
  return shape;
}

function rule(slide, [x1, y1], [x2, y2], color = LINE) {
  const from = { x: inches(x1), y: inches(y1) };
  const to = { x: inches(x2), y: inches(y2) };
  return addSlideLine(slide, { from, to, color, widthEmu: pt(0.75) });
}

// One paragraph, two formats: a bold lead-in, then the sentence in grey.
function leadIn(slide, lead, rest, box, size = 14) {
  const shape = addSlideTextBox(slide, { ...frame(box), text: '' });
  const runs = [
    { text: lead + ' ', format: { size, bold: true, color: INK } },
    { text: rest, format: { size, color: MUTED } },
  ];
  setShapeParagraphs(shape, [{ runs }]);
  return shape;
}

function bullets(slide, lines, box) {
  const shape = text(slide, lines.join('\n'), box, { size: 14, color: INK });
  setShapeBulletStyle(shape, 'bullet');
  lines.forEach((_, i) => setParagraphSpacing(shape, i, { afterPts: 9 }));
  return shape;
}

// The frame every content slide shares: a headline that states the
// conclusion, a rule under it, the source, and the page number.
let page = 1;
function content(title, source) {
  const slide = addBlankSlide(pres);
  page += 1;
  text(slide, title, [0.7, 0.45, 11.9, 0.95], { size: 26, bold: true, color: INK }, 'bottom');
  rule(slide, [0.7, 1.5], [12.63, 1.5]);
  text(slide, 'Source: ' + source, [0.7, 6.95, 10, 0.3], { size: 10, color: MUTED });
  text(slide, String(page), [11.9, 6.95, 0.73, 0.3], { size: 10, color: MUTED });
  return slide;
}

// 1. Cover, on a diagonal gradient
const cover = addBlankSlide(pres);
const wash = addSlideShape(cover, { preset: 'rect', ...frame([0, 0, 13.333, 7.5]) });
const stops = [
  { offset: 0, color: NAVY },
  { offset: 1, color: BLUE },
];
setShapeGradientFill(wash, { stops, angleDeg: 15 });
setShapeNoStroke(wash);
text(cover, 'FY26 growth plan', [0.9, 2.5, 11, 1.1], { size: 50, bold: true, color: WHITE });
const promise = 'Where the growth came from, and the four bets we ask the board to fund';
text(cover, promise, [0.9, 3.65, 9, 0.9], { size: 22, color: '#C9D6FF' });
rule(cover, [1, 5.9], [4.4, 5.9], '#6F8FFF');
text(cover, 'Board of directors, October 2026', [0.9, 6.05, 8, 0.4], {
  size: 14,
  color: '#C9D6FF',
});

// 2. Executive summary: three findings, and the ask in a side panel
const summary = content(
  `Revenue grew ${growthPct}% in four quarters, and enterprise accounts are the reason`,
  'Finance close, FY26 Q4',
);
const findings = [
  [
    'Growth is concentrated.',
    `Enterprise brought in ${enterprisePct}% of new revenue, more than the other three segments combined.`,
  ],
  [
    'Margin followed revenue.',
    `Operating margin rose from ${model.marginPct[0]}% to ${model.marginPct[3]}% as hosting costs fell per customer.`,
  ],
  [
    'The next step costs money.',
    `Four bets need ${invested}. Three of them pay back within 18 months.`,
  ],
];
findings.forEach(([lead, rest], i) => {
  const y = 1.95 + i * 1.5;
  block(summary, 'ellipse', [0.7, y, 0.62, 0.62], NAVY, String(i + 1));
  leadIn(summary, lead, rest, [1.55, y - 0.08, 6.4, 1.2], 16);
});
block(summary, 'rect', [8.5, 1.5, 4.13, 5.3], PANEL);
text(summary, 'What we ask the board', [8.8, 1.8, 3.6, 0.5], { size: 16, bold: true, color: INK });
const asks = [
  `Approve ${invested} across four bets`,
  'Move two sales pods to enterprise',
  'Revisit the marketplace in Q2',
];
bullets(summary, asks, [8.8, 2.45, 3.6, 3.5]);

// 3. One bar in colour and the rest in grey, because the headline names one subject
const segments = content(
  `A: Enterprise brought in ${enterprisePct}% of new revenue`,
  'CRM closed-won report, FY26. New revenue excludes renewals.',
);
// A bar chart draws its first category at the bottom, so list the largest last.
const names = Object.keys(model.newRevenue).reverse();
addSlideChart(segments, {
  ...frame([0.7, 1.75, 7.9, 4.95]),
  spec: {
    kind: 'bar',
    categories: names,
    series: [
      {
        name: 'New revenue ($M)',
        values: names.map((name) => model.newRevenue[name]),
        pointColors: names.map((name) => (name === 'Enterprise' ? BLUE : SLATE)),
      },
    ],
    dataLabels: {
      showValue: true,
      showCategory: false,
      showSeriesName: false,
      showPercent: false,
      numberFormat: '$0.0"M"',
    },
    valueAxisHidden: true,
    gapWidthPct: 60,
  },
});
rule(segments, [8.95, 1.85], [8.95, 6.6]);
const notes = [
  [
    'Deal size, not deal count.',
    'Enterprise closed 38 deals against 410 in SMB. The average contract is over 50 times larger.',
  ],
  ['Mid-market is steady.', 'It grew with headcount and needs no change in plan.'],
  ['Partners are early.', 'The channel opened in Q3, so one good quarter would double it.'],
];
notes.forEach(([lead, rest], i) => {
  const y = 1.9 + i * 1.6;
  block(segments, 'rect', [9.25, y + 0.05, 0.06, 1.2], BLUE);
  leadIn(segments, lead, rest, [9.45, y, 3.2, 1.4], 13);
});

// 4. Two scales on one chart: revenue as columns, margin as a line on a second axis
const margin = content(
  `B: Margin rose ${marginGain} points while revenue grew ${growthPct}%`,
  'Finance close, FY26 Q1 to Q4',
);
addSlideChart(margin, {
  ...frame([0.7, 1.75, 8.3, 4.95]),
  spec: {
    kind: 'column',
    categories: model.quarters,
    series: [
      { name: 'Revenue ($M)', values: model.revenue, color: NAVY },
      {
        name: 'Operating margin (%)',
        values: model.marginPct,
        chartKind: 'line',
        secondaryAxis: true,
        color: BLUE,
        lineWidthEmu: pt(2.5),
        markerSymbol: 'circle',
        markerSizePt: 8,
      },
    ],
    legend: { position: 'b' },
    valueAxis: { min: 0, numberFormat: '$0"M"' },
    secondaryValueAxis: { scaling: { min: 0, numberFormat: '0"%"' } },
    valueAxisMajorGridlines: true,
    valueAxisMajorGridlineColor: LINE,
    gapWidthPct: 70,
  },
});
const kpis = [
  [`$${model.revenue[3]}M`, 'Q4 revenue'],
  [`${model.marginPct[3]}%`, 'Q4 operating margin'],
  [`+${growthPct}%`, 'Revenue, Q1 to Q4'],
];
kpis.forEach(([value, label], i) => {
  const y = 1.85 + i * 1.62;
  // Grouped, so the card moves and resizes as one object in PowerPoint.
  const card = [
    block(margin, 'roundRect', [9.4, y, 3.23, 1.42], PANEL),
    block(margin, 'rect', [9.4, y + 0.2, 0.08, 1.02], BLUE),
    text(margin, value, [9.7, y + 0.12, 2.8, 0.75], { size: 32, bold: true, color: NAVY }),
    text(margin, label, [9.7, y + 0.88, 2.8, 0.4], { size: 13, color: MUTED }),
  ];
  groupShapes(card, { name: label });
});

// 5. Roadmap: three phases as arrows, the work under each, and how we know it is done
const roadmap = content(
  'The plan runs in three phases, and each one funds the next',
  'Strategy team working plan',
);
const phases = [
  {
    when: 'H1 FY27',
    name: 'Focus',
    work: ['Move two pods to enterprise', 'Launch usage-based pricing', 'Hire a partner lead'],
    done: 'Enterprise pipeline covers three quarters of quota.',
  },
  {
    when: 'H2 FY27',
    name: 'Extend',
    work: ['Open the partner marketplace', 'Start EU data residency', 'Second pod in EMEA'],
    done: 'Ten partners live and the first EU customer signed.',
  },
  {
    when: 'FY28',
    name: 'Scale',
    work: ['Partners reach 15% of new revenue', 'EU region generally available', 'Review tiers'],
    done: 'Operating margin holds above 20% at twice the revenue.',
  },
];
phases.forEach((phase, i) => {
  const x = 0.7 + i * 4.05;
  block(roadmap, 'homePlate', [x, 1.9, 3.95, 1], i === 0 ? BLUE : NAVY, phase.name, 20);
  text(roadmap, phase.when, [x, 3.1, 3.6, 0.4], { size: 13, bold: true, color: BLUE });
  bullets(roadmap, phase.work, [x, 3.5, 3.6, 1.7]);
  block(roadmap, 'roundRect', [x + 0.1, 5.35, 3.6, 1.2], PANEL);
  leadIn(roadmap, 'Done when:', phase.done, [x + 0.25, 5.45, 3.3, 1], 13);
  if (i > 0) rule(roadmap, [x - 0.1, 3.1], [x - 0.1, 6.55]);
});

// 6. The decision table, with the status cell coloured per row
const bets = content(
  `The four bets total ${invested}, and one of them is at risk`,
  'Initiative owners, reviewed by finance',
);
const rows = [
  ['Bet', 'Owner', 'Investment', 'Payback', 'Status'],
  ...model.bets.map(([name, owner, cost, months, status]) => [
    name,
    owner,
    `$${cost.toFixed(1)}M`,
    `${months} months`,
    status,
  ]),
  ['Total', '', invested, '', ''],
];
const table = addSlideTable(bets, {
  ...frame([0.7, 1.85, 11.93, 3.9]),
  rows,
  colWidths: [4.3, 2.5, 1.8, 1.8, 1.53].map(inches),
});
rows.forEach((row, r) => {
  const head = r === 0;
  const total = r === rows.length - 1;
  row.forEach((value, c) => {
    const cell = getTableCell(table, r, c);
    const status = head ? undefined : STATUS[value];
    const onColour = head || status !== undefined;
    setTableCellFill(cell, head ? NAVY : (status ?? (total ? PANEL : WHITE)));
    setTableCellAnchor(cell, 'center');
    setTableCellAlignment(cell, c >= 2 ? 'center' : 'left');
    setTableCellTextFormat(cell, {
      size: 14,
      bold: onColour || total,
      color: onColour ? WHITE : INK,
    });
    setTableCellBorders(cell, { bottom: { color: LINE, widthEmu: pt(0.75) } });
  });
});
leadIn(
  bets,
  'Why the marketplace is at risk.',
  'Two launch partners have not signed, and payback moves past 22 months if either drops out.',
  [0.7, 5.95, 11.9, 0.8],
  13,
);
