/** @jsxImportSource @office-kit/pptx-dsl */
import { pt } from '@office-kit/pptx';
import type { Color, RunSpec } from '@office-kit/pptx';
import { Presentation, Slide, Text, Shape, Line, Group, Chart, Table } from '@office-kit/pptx-dsl';
import type { Child, ShapeProps } from '@office-kit/pptx-dsl';

// A six-slide board deck. Every number on every slide comes from `model`,
// so change a figure here and the headlines, charts, and table all follow.

type Status = 'On track' | 'At risk' | 'Not started';
type Bet = [name: string, owner: string, investment: number, paybackMonths: number, status: Status];

const model = {
  quarters: ['Q1', 'Q2', 'Q3', 'Q4'],
  revenue: [4.2, 4.9, 5.8, 6.9], // $M
  marginPct: [11, 13, 16, 19], // operating margin
  newRevenue: { Enterprise: 3.1, 'Mid-market': 1.4, SMB: 0.6, Partners: 0.4 }, // $M
  bets: [
    ['Enterprise sales pods', 'M. Okafor', 2.4, 14, 'On track'],
    ['Usage-based pricing', 'L. Hartmann', 0.9, 9, 'On track'],
    ['Partner marketplace', 'S. Iyer', 1.6, 22, 'At risk'],
    ['EU data residency', 'D. Costa', 1.1, 18, 'Not started'],
  ] satisfies Bet[],
};

const growthPct = Math.round((model.revenue[3]! / model.revenue[0]! - 1) * 100);
const marginGain = model.marginPct[3]! - model.marginPct[0]!;
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
const STATUS: Record<string, Color> = {
  'On track': '#2E9E6C',
  'At risk': '#D64545',
  'Not started': SLATE,
};

// Positions are [x, y, width, height] in inches.
type Box = [x: number, y: number, width: number, height: number];
const at = ([x, y, width, height]: Box) => ({ x, y, width, height });

// A filled shape with no outline. With `label`, bold white text centred in it.
function Block(props: {
  preset: ShapeProps['preset'];
  box: Box;
  color: Color;
  label?: string;
  size?: number;
}) {
  const words =
    props.label === undefined
      ? {}
      : ({
          text: props.label,
          format: { size: props.size ?? 18, bold: true, color: WHITE },
          align: 'center',
          anchor: 'center',
        } as const);
  return (
    <Shape preset={props.preset} {...at(props.box)} fill={props.color} stroke={false} {...words} />
  );
}

function Rule(props: { from: [number, number]; to: [number, number]; color?: Color }) {
  const [x1, y1] = props.from;
  const [x2, y2] = props.to;
  return <Line x1={x1} y1={y1} x2={x2} y2={y2} color={props.color ?? LINE} width={0.75} />;
}

// One paragraph, two formats: a bold lead-in, then the sentence in grey.
function LeadIn(props: { lead: string; rest: string; box: Box; size?: number }) {
  const size = props.size ?? 14;
  // Annotated so the palette's literal types survive into `format.color`; a
  // bare array literal widens them back to `string`, which `Color` rejects.
  const runs: RunSpec[] = [
    { text: props.lead + ' ', format: { size, bold: true, color: INK } },
    { text: props.rest, format: { size, color: MUTED } },
  ];
  return <Text {...at(props.box)} paragraphs={[{ runs }]} />;
}

function Bullets(props: { lines: string[]; box: Box }) {
  return (
    <Text {...at(props.box)} size={14} color={INK} bullets="bullet" paragraphSpacing={{ after: 9 }}>
      {props.lines.join('\n')}
    </Text>
  );
}

// The frame every content slide shares: a headline that states the
// conclusion, a rule under it, the source, and the page number.
function Content(props: { page: number; title: string; source: string; children: Child }) {
  return (
    <Slide>
      <Text {...at([0.7, 0.45, 11.9, 0.95])} size={26} bold color={INK} anchor="bottom">
        {props.title}
      </Text>
      <Rule from={[0.7, 1.5]} to={[12.63, 1.5]} />
      <Text {...at([0.7, 6.95, 10, 0.3])} size={10} color={MUTED}>
        {'Source: ' + props.source}
      </Text>
      <Text {...at([11.9, 6.95, 0.73, 0.3])} size={10} color={MUTED}>
        {props.page}
      </Text>
      {props.children}
    </Slide>
  );
}

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
] as const;
const asks = [
  `Approve ${invested} across four bets`,
  'Move two sales pods to enterprise',
  'Revisit the marketplace in Q2',
];

// A bar chart draws its first category at the bottom, so list the largest last.
const segments = Object.entries(model.newRevenue).reverse();
const segmentNotes = [
  [
    'Deal size, not deal count.',
    'Enterprise closed 38 deals against 410 in SMB. The average contract is over 50 times larger.',
  ],
  ['Mid-market is steady.', 'It grew with headcount and needs no change in plan.'],
  ['Partners are early.', 'The channel opened in Q3, so one good quarter would double it.'],
] as const;

const kpis = [
  [`$${model.revenue[3]}M`, 'Q4 revenue'],
  [`${model.marginPct[3]}%`, 'Q4 operating margin'],
  [`+${growthPct}%`, 'Revenue, Q1 to Q4'],
] as const;

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

const betRows = [
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

export default (
  <Presentation>
    {/* 1. Cover, on a diagonal gradient */}
    <Slide>
      <Shape
        preset="rect"
        {...at([0, 0, 13.333, 7.5])}
        fill={{
          stops: [
            { offset: 0, color: NAVY },
            { offset: 1, color: BLUE },
          ],
          angleDeg: 15,
        }}
        stroke={false}
      />
      <Text {...at([0.9, 2.5, 11, 1.1])} size={50} bold color={WHITE}>
        FY26 growth plan
      </Text>
      <Text {...at([0.9, 3.65, 9, 0.9])} size={22} color="#C9D6FF">
        Where the growth came from, and the four bets we ask the board to fund
      </Text>
      <Rule from={[1, 5.9]} to={[4.4, 5.9]} color="#6F8FFF" />
      <Text {...at([0.9, 6.05, 8, 0.4])} size={14} color="#C9D6FF">
        Board of directors, October 2026
      </Text>
    </Slide>

    {/* 2. Executive summary: three findings, and the ask in a side panel */}
    <Content
      page={2}
      title={`Revenue grew ${growthPct}% in four quarters, and enterprise accounts are the reason`}
      source="Finance close, FY26 Q4"
    >
      {findings.map(([lead, rest], i) => {
        const y = 1.95 + i * 1.5;
        return (
          <>
            <Block preset="ellipse" box={[0.7, y, 0.62, 0.62]} color={NAVY} label={String(i + 1)} />
            <LeadIn lead={lead} rest={rest} box={[1.55, y - 0.08, 6.4, 1.2]} size={16} />
          </>
        );
      })}
      <Block preset="rect" box={[8.5, 1.5, 4.13, 5.3]} color={PANEL} />
      <Text {...at([8.8, 1.8, 3.6, 0.5])} size={16} bold color={INK}>
        What we ask the board
      </Text>
      <Bullets lines={asks} box={[8.8, 2.45, 3.6, 3.5]} />
    </Content>

    {/* 3. One bar in colour and the rest in grey, because the headline names one subject */}
    <Content
      page={3}
      title={`A: Enterprise brought in ${enterprisePct}% of new revenue`}
      source="CRM closed-won report, FY26. New revenue excludes renewals."
    >
      <Chart
        {...at([0.7, 1.75, 7.9, 4.95])}
        spec={{
          kind: 'bar',
          categories: segments.map(([name]) => name),
          series: [
            {
              name: 'New revenue ($M)',
              values: segments.map(([, value]) => value),
              pointColors: segments.map(([name]) => (name === 'Enterprise' ? BLUE : SLATE)),
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
        }}
      />
      <Rule from={[8.95, 1.85]} to={[8.95, 6.6]} />
      {segmentNotes.map(([lead, rest], i) => {
        const y = 1.9 + i * 1.6;
        return (
          <>
            <Block preset="rect" box={[9.25, y + 0.05, 0.06, 1.2]} color={BLUE} />
            <LeadIn lead={lead} rest={rest} box={[9.45, y, 3.2, 1.4]} size={13} />
          </>
        );
      })}
    </Content>

    {/* 4. Two scales on one chart: revenue as columns, margin as a line on a second axis */}
    <Content
      page={4}
      title={`B: Margin rose ${marginGain} points while revenue grew ${growthPct}%`}
      source="Finance close, FY26 Q1 to Q4"
    >
      <Chart
        {...at([0.7, 1.75, 8.3, 4.95])}
        spec={{
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
        }}
      />
      {kpis.map(([value, label], i) => {
        const y = 1.85 + i * 1.62;
        // Grouped, so the card moves and resizes as one object in PowerPoint.
        return (
          <Group name={label}>
            <Block preset="roundRect" box={[9.4, y, 3.23, 1.42]} color={PANEL} />
            <Block preset="rect" box={[9.4, y + 0.2, 0.08, 1.02]} color={BLUE} />
            <Text {...at([9.7, y + 0.12, 2.8, 0.75])} size={32} bold color={NAVY}>
              {value}
            </Text>
            <Text {...at([9.7, y + 0.88, 2.8, 0.4])} size={13} color={MUTED}>
              {label}
            </Text>
          </Group>
        );
      })}
    </Content>

    {/* 5. Roadmap: three phases as arrows, the work under each, and how we know it is done */}
    <Content
      page={5}
      title="The plan runs in three phases, and each one funds the next"
      source="Strategy team working plan"
    >
      {phases.map((phase, i) => {
        const x = 0.7 + i * 4.05;
        return (
          <>
            <Block
              preset="homePlate"
              box={[x, 1.9, 3.95, 1]}
              color={i === 0 ? BLUE : NAVY}
              label={phase.name}
              size={20}
            />
            <Text {...at([x, 3.1, 3.6, 0.4])} size={13} bold color={BLUE}>
              {phase.when}
            </Text>
            <Bullets lines={phase.work} box={[x, 3.5, 3.6, 1.7]} />
            <Block preset="roundRect" box={[x + 0.1, 5.35, 3.6, 1.2]} color={PANEL} />
            <LeadIn lead="Done when:" rest={phase.done} box={[x + 0.25, 5.45, 3.3, 1]} size={13} />
            {i > 0 && <Rule from={[x - 0.1, 3.1]} to={[x - 0.1, 6.55]} />}
          </>
        );
      })}
    </Content>

    {/* 6. The decision table, with the status cell coloured per row */}
    <Content
      page={6}
      title={`The four bets total ${invested}, and one of them is at risk`}
      source="Initiative owners, reviewed by finance"
    >
      <Table
        {...at([0.7, 1.85, 11.93, 3.9])}
        rows={betRows}
        columnWidths={[4.3, 2.5, 1.8, 1.8, 1.53]}
        cellStyle={{
          fill: WHITE,
          format: { size: 14, bold: false, color: INK },
          anchor: 'center',
          borders: { bottom: { color: LINE, width: 0.75 } },
        }}
        headerStyle={{ fill: NAVY, format: { bold: true, color: WHITE } }}
        styleCell={({ row, column, value }) => {
          const align = column >= 2 ? 'center' : 'left';
          const status = row > 0 ? STATUS[value] : undefined;
          if (status) return { align, fill: status, format: { bold: true, color: WHITE } };
          if (row === betRows.length - 1) return { align, fill: PANEL, format: { bold: true } };
          return { align };
        }}
      />
      <LeadIn
        lead="Why the marketplace is at risk."
        rest="Two launch partners have not signed, and payback moves past 22 months if either drops out."
        box={[0.7, 5.95, 11.9, 0.8]}
        size={13}
      />
    </Content>
  </Presentation>
);
