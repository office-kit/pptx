const WORD_BYTES = 4;
const SHORT_BYTES = 2;
const RECORD_HEADER_BYTES = 8;
const HEADER_MIN_BYTES = 88;
const VALUE_RECORD_BYTES = 12;
const PAIR_RECORD_BYTES = 16;
const BOUNDS_RECORD_BYTES = 24;
const BRUSH_RECORD_BYTES = 24;
const EOF_MIN_BYTES = 20;
const HEADER_BOUNDS_OFFSET = 8;
const HEADER_FILE_SIZE_OFFSET = 48;
const HEADER_RECORD_COUNT_OFFSET = 52;
const VALUE_OFFSET = RECORD_HEADER_BYTES;
const SECOND_VALUE_OFFSET = VALUE_OFFSET + WORD_BYTES;
const BRUSH_COLOR_OFFSET = 16;
const RGB_CHANNEL_COUNT = 3;
const HEX_RADIX = 16;
const EMFPLUS_COMMENT = 0x2b464d45;
const SIGNATURE_OFFSET = 40;
const SIGNATURE = 0x464d4520;
const POINT_COUNT_OFFSET = 24;
const POINTS_OFFSET = 28;
const BEZIER_POINT_COUNT = 3;
const STOCK_OBJECT_FLAG = 0x80000000;
const REGION_COPY = 5;
const MAP_TEXT = 1;
const MAP_ANISOTROPIC = 8;
const BRUSH_SOLID = 0;
const BRUSH_NULL = 1;
const FILL_ALTERNATE = 1;
const FILL_WINDING = 2;
const STOCK_BRUSHES = ['#ffffff', '#c0c0c0', '#808080', '#404040', '#000000', 'none'];

const RECORD = {
  HEADER: 1,
  POLYBEZIERTO: 5,
  POLYLINETO: 6,
  SETWINDOWEXTEX: 9,
  SETWINDOWORGEX: 10,
  SETVIEWPORTEXTEX: 11,
  SETVIEWPORTORGEX: 12,
  SETBRUSHORGEX: 13,
  EOF: 14,
  SETMAPMODE: 17,
  SETBKMODE: 18,
  SETPOLYFILLMODE: 19,
  SETSTRETCHBLTMODE: 21,
  SETTEXTALIGN: 22,
  MOVETOEX: 27,
  SELECTOBJECT: 37,
  CREATEBRUSHINDIRECT: 39,
  DELETEOBJECT: 40,
  BEGINPATH: 59,
  ENDPATH: 60,
  CLOSEFIGURE: 61,
  FILLPATH: 62,
  SELECTCLIPPATH: 67,
  COMMENT: 70,
  EXTSELECTCLIPRGN: 75,
  POLYBEZIERTO16: 88,
  POLYLINETO16: 89,
} as const;

// Unsupported drawing commands reject the whole image to avoid silently missing artwork.
export const renderEmfToSvg = (bytes: Uint8Array): string | null => {
  if (bytes.byteLength < HEADER_MIN_BYTES) return null;
  const file = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (
    file.getUint32(0, true) !== RECORD.HEADER ||
    file.getUint32(SIGNATURE_OFFSET, true) !== SIGNATURE
  )
    return null;
  const left = file.getInt32(HEADER_BOUNDS_OFFSET, true);
  const top = file.getInt32(HEADER_BOUNDS_OFFSET + WORD_BYTES, true);
  const width = file.getInt32(HEADER_BOUNDS_OFFSET + WORD_BYTES * 2, true) - left;
  const height = file.getInt32(HEADER_BOUNDS_OFFSET + WORD_BYTES * 3, true) - top;
  if (
    width <= 0 ||
    height <= 0 ||
    file.getUint32(HEADER_FILE_SIZE_OFFSET, true) !== bytes.byteLength
  )
    return null;

  let windowX = 0;
  let windowY = 0;
  let windowWidth = 1;
  let windowHeight = 1;
  let viewportX = 0;
  let viewportY = 0;
  let viewportWidth = 1;
  let viewportHeight = 1;
  let mapMode: number = MAP_TEXT;
  let fillRule = 'evenodd';
  let brush = '#ffffff';
  const brushes = new Map<number, string>();
  let path: string[] = [];
  let recordingPath = false;
  let clip = '';
  let clipCount = 0;
  const definitions: string[] = [];
  const drawings: string[] = [];
  let recordCount = 0;
  const point = (x: number, y: number): string => {
    if (mapMode === MAP_TEXT) return `${x} ${y}`;
    const mappedX = ((x - windowX) * viewportWidth) / windowWidth + viewportX;
    const mappedY = ((y - windowY) * viewportHeight) / windowHeight + viewportY;
    return `${mappedX} ${mappedY}`;
  };

  for (let offset = 0; offset < bytes.byteLength; ) {
    if (bytes.byteLength - offset < RECORD_HEADER_BYTES) return null;
    const type = file.getUint32(offset, true);
    const size = file.getUint32(offset + WORD_BYTES, true);
    if (size < RECORD_HEADER_BYTES || size % WORD_BYTES !== 0 || size > bytes.byteLength - offset)
      return null;
    const record = new DataView(bytes.buffer, bytes.byteOffset + offset, size);
    const unsigned = (at: number) => record.getUint32(at, true);
    const signed = (at: number) => record.getInt32(at, true);
    recordCount++;
    switch (type) {
      case RECORD.HEADER:
        if (offset !== 0 || size < HEADER_MIN_BYTES) return null;
        break;
      case RECORD.EOF:
        if (
          size < EOF_MIN_BYTES ||
          offset + size !== bytes.byteLength ||
          recordingPath ||
          recordCount !== file.getUint32(HEADER_RECORD_COUNT_OFFSET, true)
        )
          return null;
        return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="${left} ${top} ${width} ${height}" width="${width}" height="${height}"><defs>${definitions.join('')}</defs>${drawings.join('')}</svg>`;
      case RECORD.SETMAPMODE:
        if (size < VALUE_RECORD_BYTES) return null;
        mapMode = unsigned(VALUE_OFFSET);
        if (mapMode !== MAP_TEXT && mapMode !== MAP_ANISOTROPIC) return null;
        break;
      case RECORD.SETWINDOWORGEX:
      case RECORD.SETVIEWPORTORGEX:
      case RECORD.SETWINDOWEXTEX:
      case RECORD.SETVIEWPORTEXTEX: {
        if (size < PAIR_RECORD_BYTES) return null;
        const x = signed(VALUE_OFFSET);
        const y = signed(SECOND_VALUE_OFFSET);
        if (type === RECORD.SETWINDOWORGEX) {
          windowX = x;
          windowY = y;
        } else if (type === RECORD.SETVIEWPORTORGEX) {
          viewportX = x;
          viewportY = y;
        } else {
          if (x === 0 || y === 0) return null;
          if (type === RECORD.SETWINDOWEXTEX) {
            windowWidth = x;
            windowHeight = y;
          } else {
            viewportWidth = x;
            viewportHeight = y;
          }
        }
        break;
      }
      case RECORD.CREATEBRUSHINDIRECT: {
        if (size < BRUSH_RECORD_BYTES) return null;
        const style = unsigned(SECOND_VALUE_OFFSET);
        if (style !== BRUSH_SOLID && style !== BRUSH_NULL) return null;
        const color = [
          ...bytes.subarray(
            offset + BRUSH_COLOR_OFFSET,
            offset + BRUSH_COLOR_OFFSET + RGB_CHANNEL_COUNT,
          ),
        ]
          .map((channel) => channel.toString(HEX_RADIX).padStart(2, '0'))
          .join('');
        brushes.set(unsigned(VALUE_OFFSET), style === BRUSH_NULL ? 'none' : `#${color}`);
        break;
      }
      case RECORD.SELECTOBJECT: {
        if (size < VALUE_RECORD_BYTES) return null;
        const handle = unsigned(VALUE_OFFSET);
        const selected =
          handle >= STOCK_OBJECT_FLAG
            ? STOCK_BRUSHES[handle - STOCK_OBJECT_FLAG]
            : brushes.get(handle);
        if (selected === undefined) return null;
        brush = selected;
        break;
      }
      case RECORD.DELETEOBJECT:
        if (size < VALUE_RECORD_BYTES) return null;
        brushes.delete(unsigned(VALUE_OFFSET));
        break;
      case RECORD.SETPOLYFILLMODE:
        if (
          size < VALUE_RECORD_BYTES ||
          (unsigned(VALUE_OFFSET) !== FILL_ALTERNATE && unsigned(VALUE_OFFSET) !== FILL_WINDING)
        )
          return null;
        fillRule = unsigned(VALUE_OFFSET) === FILL_WINDING ? 'nonzero' : 'evenodd';
        break;
      case RECORD.BEGINPATH:
        path = [];
        recordingPath = true;
        break;
      case RECORD.MOVETOEX:
        if (size < PAIR_RECORD_BYTES || !recordingPath) return null;
        path.push(`M${point(signed(VALUE_OFFSET), signed(SECOND_VALUE_OFFSET))}`);
        break;
      case RECORD.POLYLINETO:
      case RECORD.POLYLINETO16:
      case RECORD.POLYBEZIERTO:
      case RECORD.POLYBEZIERTO16: {
        if (size < POINTS_OFFSET || !recordingPath || path.length === 0) return null;
        const short = type === RECORD.POLYLINETO16 || type === RECORD.POLYBEZIERTO16;
        const bezier = type === RECORD.POLYBEZIERTO || type === RECORD.POLYBEZIERTO16;
        const coordinateBytes = short ? SHORT_BYTES : WORD_BYTES;
        const pointBytes = coordinateBytes * 2;
        const count = unsigned(POINT_COUNT_OFFSET);
        if (
          count > Math.floor((size - POINTS_OFFSET) / pointBytes) ||
          (bezier && count % BEZIER_POINT_COUNT !== 0)
        )
          return null;
        for (let index = 0; index < count; index++) {
          const at = POINTS_OFFSET + index * pointBytes;
          const x = short ? record.getInt16(at, true) : signed(at);
          const y = short
            ? record.getInt16(at + coordinateBytes, true)
            : signed(at + coordinateBytes);
          const command = bezier ? (index % BEZIER_POINT_COUNT === 0 ? 'C' : ' ') : 'L';
          path.push(`${command}${point(x, y)}`);
        }
        break;
      }
      case RECORD.CLOSEFIGURE:
        if (!recordingPath || path.length === 0) return null;
        path.push('Z');
        break;
      case RECORD.ENDPATH:
        if (!recordingPath) return null;
        recordingPath = false;
        break;
      case RECORD.FILLPATH:
        if (size < BOUNDS_RECORD_BYTES || recordingPath || path.length === 0) return null;
        drawings.push(
          `<path d="${path.join('')}" fill="${brush}" fill-rule="${fillRule}"${clip}/>`,
        );
        path = [];
        break;
      case RECORD.SELECTCLIPPATH: {
        if (
          size < VALUE_RECORD_BYTES ||
          unsigned(VALUE_OFFSET) !== REGION_COPY ||
          recordingPath ||
          path.length === 0
        )
          return null;
        const id = `emf-clip-${clipCount++}`;
        definitions.push(
          `<clipPath id="${id}"><path d="${path.join('')}" clip-rule="${fillRule}"/></clipPath>`,
        );
        clip = ` clip-path="url(#${id})"`;
        path = [];
        break;
      }
      case RECORD.EXTSELECTCLIPRGN:
        if (
          size < PAIR_RECORD_BYTES ||
          unsigned(VALUE_OFFSET) !== 0 ||
          unsigned(SECOND_VALUE_OFFSET) !== REGION_COPY
        )
          return null;
        clip = '';
        break;
      case RECORD.COMMENT:
        if (size < VALUE_RECORD_BYTES || unsigned(VALUE_OFFSET) > size - VALUE_RECORD_BYTES)
          return null;
        // EMF+ comments can contain drawing commands absent from the GDI records.
        if (
          unsigned(VALUE_OFFSET) >= WORD_BYTES &&
          unsigned(SECOND_VALUE_OFFSET) === EMFPLUS_COMMENT
        )
          return null;
        break;
      // Text/bitmap settings and the brush origin do not affect solid-filled paths.
      case RECORD.SETBKMODE:
      case RECORD.SETSTRETCHBLTMODE:
      case RECORD.SETTEXTALIGN:
        if (size < VALUE_RECORD_BYTES) return null;
        break;
      case RECORD.SETBRUSHORGEX:
        if (size < PAIR_RECORD_BYTES) return null;
        break;
      default:
        return null;
    }
    offset += size;
  }
  return null;
};
