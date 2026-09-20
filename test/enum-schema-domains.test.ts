import { existsSync, readFileSync } from 'node:fs';
import { expect, it } from 'vitest';
import * as common from '../src/internal/enum-values.ts';
import * as chart from '../src/internal/chartml/enum-validation.ts';

const schemaUrl = (file: string): URL =>
  new URL(
    `../references/ecma-376-5th/ECMA-376/OfficeOpenXML-XMLSchema-Transitional/${file}.xsd`,
    import.meta.url,
  );

const domains = [
  ['SHAPE_PRESETS', common.SHAPE_PRESETS, 'dml-main', ['ST_ShapeType']],
  ['LINE_DASHES', common.LINE_DASHES, 'dml-main', ['ST_PresetLineDashVal']],
  ['TEXT_DIRECTIONS', common.TEXT_DIRECTIONS, 'dml-main', ['ST_TextVerticalType']],
  ['AUTO_NUMBER_SCHEMES', common.AUTO_NUMBER_SCHEMES, 'dml-main', ['ST_TextAutonumberScheme']],
  ['UNDERLINES', common.UNDERLINES, 'dml-main', ['ST_TextUnderlineType']],
  ['STRIKES', common.STRIKES, 'dml-main', ['ST_TextStrikeType']],
  ['SLIDE_SIZE_TYPES', common.SLIDE_SIZE_TYPES, 'pml', ['ST_SlideSizeType']],
  ['LABEL_POSITIONS', chart.LABEL_POSITIONS, 'dml-chart', ['ST_DLblPos']],
  ['DISPLAY_UNITS', chart.DISPLAY_UNITS, 'dml-chart', ['ST_BuiltInUnit']],
  ['LABEL_ALIGNMENTS', chart.LABEL_ALIGNMENTS, 'dml-chart', ['ST_LblAlgn']],
  ['DISPLAY_BLANKS', chart.DISPLAY_BLANKS, 'dml-chart', ['ST_DispBlanksAs']],
  ['SCATTER_STYLES', chart.SCATTER_STYLES, 'dml-chart', ['ST_ScatterStyle']],
  ['RADAR_STYLES', chart.RADAR_STYLES, 'dml-chart', ['ST_RadarStyle']],
  ['BAR_3D_SHAPES', chart.BAR_3D_SHAPES, 'dml-chart', ['ST_Shape']],
  ['OF_PIE_TYPES', chart.OF_PIE_TYPES, 'dml-chart', ['ST_OfPieType']],
  ['SPLIT_TYPES', chart.SPLIT_TYPES, 'dml-chart', ['ST_SplitType']],
  ['LAYOUT_TARGETS', chart.LAYOUT_TARGETS, 'dml-chart', ['ST_LayoutTarget']],
  ['AXIS_CROSSES', chart.AXIS_CROSSES, 'dml-chart', ['ST_Crosses']],
  ['LEGEND_POSITIONS', chart.LEGEND_POSITIONS, 'dml-chart', ['ST_LegendPos']],
  ['MARKER_SYMBOLS', chart.MARKER_SYMBOLS, 'dml-chart', ['ST_MarkerStyle']],
  ['TRENDLINE_TYPES', chart.TRENDLINE_TYPES, 'dml-chart', ['ST_TrendlineType']],
  ['ERROR_BAR_TYPES', chart.ERROR_BAR_TYPES, 'dml-chart', ['ST_ErrBarType']],
  ['ERROR_VALUE_TYPES', chart.ERROR_VALUE_TYPES, 'dml-chart', ['ST_ErrValType']],
  ['TICK_MARKS', chart.TICK_MARKS, 'dml-chart', ['ST_TickMark']],
  ['TICK_LABEL_POSITIONS', chart.TICK_LABEL_POSITIONS, 'dml-chart', ['ST_TickLblPos']],
  ['ORIENTATIONS', chart.ORIENTATIONS, 'dml-chart', ['ST_Orientation']],
  ['CROSS_BETWEEN', chart.CROSS_BETWEEN, 'dml-chart', ['ST_CrossBetween']],
  ['TIME_UNITS', chart.TIME_UNITS, 'dml-chart', ['ST_TimeUnit']],
  ['GROUPINGS', chart.GROUPINGS, 'dml-chart', ['ST_Grouping', 'ST_BarGrouping']],
] as const;

for (const [name, values, file, types] of domains) {
  it.skipIf(!existsSync(schemaUrl(file)))(`${name} matches the XSD enum domain exactly`, () => {
    const schema = readFileSync(schemaUrl(file), 'utf8');
    const expected = types.flatMap((type) => {
      const body = schema.match(
        new RegExp(`<xsd:simpleType name="${type}">([\\s\\S]*?)</xsd:simpleType>`),
      )?.[1];
      expect(body, type).toBeDefined();
      const tokens = [...body!.matchAll(/<xsd:enumeration value="([^"]+)"/g)].map(
        (match) => match[1]!,
      );
      expect(tokens.length).toBeGreaterThan(0);
      return tokens;
    });
    // Grouping accepts the union of the bar and line/area schema domains.
    expect([...values].sort()).toEqual([...new Set(expected)].sort());
  });
}
