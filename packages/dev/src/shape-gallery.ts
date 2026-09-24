// Mac PowerPoint gallery order. Presets are shared by insertion validation and UI.
const category = (label: string, entries: string) => ({
  label,
  shapes: entries
    .split('\n')
    .filter(Boolean)
    .map((line) => {
      const [preset, name] = line.split('|');
      return { preset: preset!, label: name! };
    }),
});
export const shapeCategories = [
  category(
    'Rectangles',
    `rect|Rectangle
roundRect|Rounded Rectangle
snip1Rect|Snip Single Corner Rectangle
snip2SameRect|Snip Same Side Corner Rectangle
snip2DiagRect|Snip Diagonal Corner Rectangle
snipRoundRect|Snip and Round Single Corner Rectangle
round1Rect|Round Single Corner Rectangle
round2SameRect|Round Same Side Corner Rectangle
round2DiagRect|Round Diagonal Corner Rectangle`,
  ),
  category(
    'Basic Shapes',
    `ellipse|Oval
triangle|Isosceles Triangle
rtTriangle|Right Triangle
parallelogram|Parallelogram
trapezoid|Trapezoid
diamond|Diamond
pentagon|Regular Pentagon
hexagon|Hexagon
heptagon|Heptagon
octagon|Octagon
decagon|Decagon
dodecagon|Dodecagon
pie|Pie
chord|Chord
teardrop|Teardrop
frame|Frame
halfFrame|Half Frame
corner|L-Shape
diagStripe|Diagonal Stripe
plus|Cross
plaque|Plaque
can|Can
cube|Cube
bevel|Bevel
donut|Donut
noSmoking|"No" Symbol
blockArc|Block Arc
smileyFace|Smiley Face
heart|Heart
lightningBolt|Lightning Bolt
sun|Sun
moon|Moon
cloud|Cloud
arc|Arc
bracketPair|Double Bracket
bracePair|Double Brace
leftBracket|Left Bracket
rightBracket|Right Bracket
leftBrace|Left Brace
rightBrace|Right Brace`,
  ),
  category(
    'Block Arrows',
    `rightArrow|Right Arrow
leftArrow|Left Arrow
upArrow|Up Arrow
downArrow|Down Arrow
leftRightArrow|Left-Right Arrow
upDownArrow|Up-Down Arrow
quadArrow|Quad Arrow
leftRightUpArrow|Left-Right-Up Arrow
bentArrow|Bent Arrow
uturnArrow|U-Turn Arrow
bentUpArrow|Bent-Up Arrow
curvedRightArrow|Curved Right Arrow
curvedLeftArrow|Curved Left Arrow
curvedUpArrow|Curved Up Arrow
curvedDownArrow|Curved Down Arrow
stripedRightArrow|Striped Right Arrow
notchedRightArrow|Notched Right Arrow
homePlate|Pentagon
chevron|Chevron
rightArrowCallout|Right Arrow Callout
downArrowCallout|Down Arrow Callout
leftArrowCallout|Left Arrow Callout
upArrowCallout|Up Arrow Callout
leftRightArrowCallout|Left-Right Arrow Callout
quadArrowCallout|Quad Arrow Callout
circularArrow|Circular Arrow`,
  ),
  category(
    'Equation Shapes',
    `mathPlus|Plus
mathMinus|Minus
mathMultiply|Multiply
mathDivide|Division
mathEqual|Equal
mathNotEqual|Not Equal`,
  ),
  category(
    'Flowchart',
    `flowChartProcess|Process
flowChartAlternateProcess|Alternate Process
flowChartDecision|Decision
flowChartInputOutput|Data
flowChartPredefinedProcess|Predefined Process
flowChartInternalStorage|Internal Storage
flowChartDocument|Document
flowChartMultidocument|Multidocument
flowChartTerminator|Terminator
flowChartPreparation|Preparation
flowChartManualInput|Manual Input
flowChartManualOperation|Manual Operation
flowChartConnector|Connector
flowChartPunchedCard|Card
flowChartPunchedTape|Punched Tape
flowChartSummingJunction|Summing Junction
flowChartOr|Or
flowChartCollate|Collate
flowChartSort|Sort
flowChartExtract|Extract
flowChartMerge|Merge
flowChartOnlineStorage|Stored Data
flowChartDelay|Delay
flowChartMagneticTape|Sequential Access Storage
flowChartMagneticDisk|Magnetic Disk
flowChartMagneticDrum|Direct Access Storage
flowChartDisplay|Display`,
  ),
  category(
    'Stars and Banners',
    `irregularSeal1|Explosion 1
irregularSeal2|Explosion 2
star4|4-Point Star
star5|5-Point Star
star6|6-Point Star
star7|7-Point Star
star8|8-Point Star
star10|10-Point Star
star12|12-Point Star
star16|16-Point Star
star24|24-Point Star
star32|32-Point Star
ribbon2|Up Ribbon
ribbon|Down Ribbon
ellipseRibbon2|Curved Up Ribbon
ellipseRibbon|Curved Down Ribbon
verticalScroll|Vertical Scroll
horizontalScroll|Horizontal Scroll
wave|Wave
doubleWave|Double Wave`,
  ),
  category(
    'Callouts',
    `wedgeRectCallout|Rectangular Callout
wedgeRoundRectCallout|Rounded Rectangular Callout
wedgeEllipseCallout|Oval Callout
cloudCallout|Cloud Callout
borderCallout1|Line Callout 1
borderCallout2|Line Callout 2
borderCallout3|Line Callout 3
accentCallout1|Line Callout 1 (Accent Bar)
callout1|Line Callout 1 (No Border)
callout2|Line Callout 2 (No Border)
callout3|Line Callout 3 (No Border)
accentBorderCallout1|Line Callout 1 (Border and Accent Bar)`,
  ),
  category(
    'Action Buttons',
    `actionButtonBackPrevious|Action Button: Back or Previous
actionButtonForwardNext|Action Button: Forward or Next
actionButtonBeginning|Action Button: Beginning
actionButtonEnd|Action Button: End
actionButtonHome|Action Button: Home
actionButtonInformation|Action Button: Information
actionButtonReturn|Action Button: Return
actionButtonMovie|Action Button: Movie
actionButtonDocument|Action Button: Document
actionButtonSound|Action Button: Sound
actionButtonHelp|Action Button: Help
actionButtonBlank|Action Button: Custom`,
  ),
];
export function defaultButtonAction(preset: string) {
  switch (preset) {
    case 'actionButtonBackPrevious':
      return 'prevSlide';
    case 'actionButtonForwardNext':
      return 'nextSlide';
    case 'actionButtonBeginning':
    case 'actionButtonHome':
      return 'firstSlide';
    case 'actionButtonEnd':
      return 'lastSlide';
    case 'actionButtonReturn':
      return 'lastSlideViewed';
    default:
      return null;
  }
}
export const shapePresets = new Set(shapeCategories.flatMap((c) => c.shapes.map((s) => s.preset)));

export const lineTools = [
  { preset: 'line', label: 'Line', geometry: 'line', arrows: 0 },
  { preset: 'lineArrow', label: 'Arrow', geometry: 'line', arrows: 1 },
  { preset: 'lineDoubleArrow', label: 'Double Arrow', geometry: 'line', arrows: 2 },
  { preset: 'elbow', label: 'Elbow Connector', geometry: 'bentConnector3', arrows: 0 },
  { preset: 'elbowArrow', label: 'Elbow Arrow Connector', geometry: 'bentConnector3', arrows: 1 },
  {
    preset: 'elbowDoubleArrow',
    label: 'Elbow Double-Arrow Connector',
    geometry: 'bentConnector3',
    arrows: 2,
  },
  { preset: 'curved', label: 'Curved Connector', geometry: 'curvedConnector3', arrows: 0 },
  {
    preset: 'curvedArrow',
    label: 'Curved Arrow Connector',
    geometry: 'curvedConnector3',
    arrows: 1,
  },
  {
    preset: 'curvedDoubleArrow',
    label: 'Curved Double-Arrow Connector',
    geometry: 'curvedConnector3',
    arrows: 2,
  },
] as const;
export const lineTool = (preset: string | null | undefined) =>
  lineTools.find((tool) => tool.preset === preset);
