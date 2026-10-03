// Built-in DrawingML table styles used when PowerPoint stores only a style GUID.

import { parseXml, type XmlElement } from '../xml/index.ts';

const BUILTIN_MEDIUM_STYLE_2_ACCENT_1_ID = '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}';

// Native definition captured from PowerPoint's tableStyles.xml fixture. Keep
// parsing lazy: most presentations use an embedded style or no table style.
const BUILTIN_MEDIUM_STYLE_2_ACCENT_1_XML =
  '<a:tblStyle xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" styleId="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}" styleName="Medium Style 2 - Accent 1">' +
  '<a:wholeTbl><a:tcTxStyle><a:fontRef idx="minor"><a:prstClr val="black"/></a:fontRef><a:schemeClr val="dk1"/></a:tcTxStyle><a:tcStyle><a:tcBdr>' +
  '<a:left><a:ln w="12700" cmpd="sng"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></a:left>' +
  '<a:right><a:ln w="12700" cmpd="sng"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></a:right>' +
  '<a:top><a:ln w="12700" cmpd="sng"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></a:top>' +
  '<a:bottom><a:ln w="12700" cmpd="sng"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></a:bottom>' +
  '<a:insideH><a:ln w="12700" cmpd="sng"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></a:insideH>' +
  '<a:insideV><a:ln w="12700" cmpd="sng"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></a:insideV>' +
  '</a:tcBdr><a:fill><a:solidFill><a:schemeClr val="accent1"><a:tint val="20000"/></a:schemeClr></a:solidFill></a:fill></a:tcStyle></a:wholeTbl>' +
  '<a:band1H><a:tcStyle><a:tcBdr/><a:fill><a:solidFill><a:schemeClr val="accent1"><a:tint val="40000"/></a:schemeClr></a:solidFill></a:fill></a:tcStyle></a:band1H>' +
  '<a:band2H><a:tcStyle><a:tcBdr/></a:tcStyle></a:band2H>' +
  '<a:band1V><a:tcStyle><a:tcBdr/><a:fill><a:solidFill><a:schemeClr val="accent1"><a:tint val="40000"/></a:schemeClr></a:solidFill></a:fill></a:tcStyle></a:band1V>' +
  '<a:band2V><a:tcStyle><a:tcBdr/></a:tcStyle></a:band2V>' +
  '<a:lastCol><a:tcTxStyle b="on"><a:fontRef idx="minor"><a:prstClr val="black"/></a:fontRef><a:schemeClr val="lt1"/></a:tcTxStyle><a:tcStyle><a:tcBdr/><a:fill><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:fill></a:tcStyle></a:lastCol>' +
  '<a:firstCol><a:tcTxStyle b="on"><a:fontRef idx="minor"><a:prstClr val="black"/></a:fontRef><a:schemeClr val="lt1"/></a:tcTxStyle><a:tcStyle><a:tcBdr/><a:fill><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:fill></a:tcStyle></a:firstCol>' +
  '<a:lastRow><a:tcTxStyle b="on"><a:fontRef idx="minor"><a:prstClr val="black"/></a:fontRef><a:schemeClr val="lt1"/></a:tcTxStyle><a:tcStyle><a:tcBdr><a:top><a:ln w="38100" cmpd="sng"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></a:top></a:tcBdr><a:fill><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:fill></a:tcStyle></a:lastRow>' +
  '<a:firstRow><a:tcTxStyle b="on"><a:fontRef idx="minor"><a:prstClr val="black"/></a:fontRef><a:schemeClr val="lt1"/></a:tcTxStyle><a:tcStyle><a:tcBdr><a:bottom><a:ln w="38100" cmpd="sng"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></a:bottom></a:tcBdr><a:fill><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:fill></a:tcStyle></a:firstRow>' +
  '</a:tblStyle>';

// ISO/IEC 29500 tableStyle example, reproduced in Microsoft's SDK remarks:
// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.tablestyle
const BUILTIN_TABLE_GRID_ID = '{5940675A-B579-460E-94D1-54222C63F5DA}';
const BUILTIN_TABLE_GRID_XML =
  `<a:tblStyle xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" styleId="${BUILTIN_TABLE_GRID_ID}" styleName="No Style, Table Grid">` +
  '<a:wholeTbl><a:tcTxStyle><a:fontRef idx="minor"><a:scrgbClr r="0" g="0" b="0"/></a:fontRef><a:schemeClr val="tx1"/></a:tcTxStyle><a:tcStyle><a:tcBdr>' +
  ['left', 'right', 'top', 'bottom', 'insideH', 'insideV']
    .map(
      (side) =>
        `<a:${side}><a:ln w="12700" cmpd="sng"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:ln></a:${side}>`,
    )
    .join('') +
  '</a:tcBdr><a:fill><a:noFill/></a:fill></a:tcStyle></a:wholeTbl></a:tblStyle>';

const builtinDefinitions = new Map([
  [BUILTIN_MEDIUM_STYLE_2_ACCENT_1_ID, BUILTIN_MEDIUM_STYLE_2_ACCENT_1_XML],
  [BUILTIN_TABLE_GRID_ID, BUILTIN_TABLE_GRID_XML],
]);
const builtinStyles = new Map<string, XmlElement>();

export const getBuiltinTableStyle = (styleId: string): XmlElement | null => {
  const id = styleId.trim().toUpperCase();
  const cached = builtinStyles.get(id);
  if (cached) return cached;
  const definition = builtinDefinitions.get(id);
  if (!definition) return null;
  const style = parseXml(definition).root;
  builtinStyles.set(id, style);
  return style;
};
