/** @jsxImportSource @office-kit/pptx-dsl */
import { Presentation, Slide, Text, Shape, Chart, Table } from '@office-kit/pptx-dsl';

const accent = '#E5481F';
const kpis = [
  ['$300k', 'Q4 revenue'],
  ['47%', 'Gross margin'],
  ['+25%', 'Quarter on quarter'],
];

export default (
  <Presentation>
    <Slide background="#15171C">
      <Shape
        preset="rect"
        x={0.9}
        y={2.55}
        width={0.14}
        height={1.75}
        fill={accent}
        stroke={false}
      />
      <Text x={1.25} y={2.4} width={10} height={1.1} size={48} bold color="#FFFFFF">
        Q3 business review
      </Text>
      <Text x={1.25} y={3.55} width={10} height={0.6} size={22} color="#B4B9C4">
        Revenue, margin, and what we do next
      </Text>
      <Text x={1.25} y={6.4} width={10} height={0.4} size={14} color="#8A909C">
        Finance team, October 2026
      </Text>
    </Slide>
    <Slide>
      <Text x={0.9} y={0.55} width={11.5} height={0.8} size={30} bold>
        Revenue grew 2.5x in four quarters
      </Text>
      {kpis.map(([value, label], i) => (
        <>
          <Shape
            preset="roundRect"
            x={0.9}
            y={1.7 + i * 1.75}
            width={3.4}
            height={1.5}
            fill="#F3F4F7"
            stroke={false}
          />
          <Text x={1.15} y={1.85 + i * 1.75} width={3} height={0.75} size={34} bold color={accent}>
            {value}
          </Text>
          <Text x={1.15} y={2.6 + i * 1.75} width={3} height={0.4} size={14} color="#5B616E">
            {label}
          </Text>
        </>
      ))}
      <Chart
        x={4.7}
        y={1.6}
        width={7.8}
        height={5.3}
        spec={{
          kind: 'column',
          categories: ['Q1', 'Q2', 'Q3', 'Q4'],
          series: [
            { name: 'Revenue', values: [120, 180, 240, 300], color: accent },
            { name: 'Cost', values: [80, 90, 130, 160], color: '#C9CDD6' },
          ],
          legend: { position: 'b' },
          valueAxisMajorGridlines: true,
          gapWidthPct: 80,
        }}
      />
    </Slide>
    <Slide>
      <Text x={0.9} y={0.55} width={11.5} height={0.8} size={30} bold>
        What we do next
      </Text>
      <Table
        x={0.9}
        y={1.7}
        width={11.5}
        height={3.2}
        rows={[
          ['Owner', 'Action', 'Due'],
          ['Aiko', 'Renegotiate the hosting contract', 'Nov 15'],
          ['Ben', 'Ship annual billing', 'Dec 1'],
          ['Chloe', 'Hire two support engineers', 'Jan 10'],
        ]}
        columnWidths={[2.2, 7, 2.3]}
        cellStyle={{ format: { size: 16 }, anchor: 'center' }}
        headerStyle={{ fill: '#15171C', format: { size: 16, bold: true, color: '#FFFFFF' } }}
        stripeFill="#F3F4F7"
      />
    </Slide>
  </Presentation>
);
