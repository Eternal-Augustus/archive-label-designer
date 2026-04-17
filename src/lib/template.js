const PAPER_WIDTH_MM = 210;
const PAPER_HEIGHT_MM = 297;
const PX_PER_MM = 96 / 25.4;
const DEFAULT_MIN_FONT_SIZE = 9;

let rowSeed = 1;

export const FONT_FAMILY_OPTIONS = [
  {
    value: "songti",
    label: "宋体",
    stack: '"SimSun", "Songti SC", serif'
  },
  {
    value: "yahei",
    label: "微软雅黑",
    stack: '"Microsoft YaHei", "PingFang SC", sans-serif'
  },
  {
    value: "heiti",
    label: "黑体",
    stack: '"SimHei", "Heiti SC", sans-serif'
  }
];

export const OVERFLOW_MODE_OPTIONS = [
  {
    value: "clip",
    label: "单行裁切"
  },
  {
    value: "wrap",
    label: "自动换行"
  },
  {
    value: "shrink",
    label: "自动缩字"
  }
];

const FONT_STACK_MAP = Object.fromEntries(
  FONT_FAMILY_OPTIONS.map((item) => [item.value, item.stack])
);

const ROW_MATCHERS = [
  ["编号", ["流水", "编号", "序号"]],
  ["姓名", ["姓名"]],
  ["学院/专业", ["学院", "专业"]],
  ["班级", ["班级"]],
  ["学号", ["学号"]],
  ["身份证", ["身份证", "证件号"]],
  ["状态", ["调档", "档案", "状态"]]
];

const DEFAULT_TEMPLATE_ROWS = [
  {
    id: 1,
    label: "编号",
    sourceType: "serialCode",
    fieldKey: "",
    text: "",
    fontSize: 20,
    fontWeight: 700,
    fontFamily: "songti",
    align: "center",
    overflowMode: "shrink",
    yearText: "25",
    degreeText: "博",
    serialFieldKey: "field_1",
    serialPadLength: 3
  },
  {
    id: 2,
    label: "姓名",
    sourceType: "field",
    fieldKey: "field_4",
    text: "",
    fontSize: 19,
    fontWeight: 700,
    fontFamily: "songti",
    align: "center",
    overflowMode: "shrink",
    yearText: "25",
    degreeText: "博",
    serialFieldKey: "",
    serialPadLength: 3
  },
  {
    id: 4,
    label: "班级",
    sourceType: "field",
    fieldKey: "field_7",
    text: "",
    fontSize: 17,
    fontWeight: 500,
    fontFamily: "songti",
    align: "center",
    overflowMode: "shrink",
    yearText: "25",
    degreeText: "博",
    serialFieldKey: "",
    serialPadLength: 3
  },
  {
    id: 5,
    label: "学号",
    sourceType: "field",
    fieldKey: "field_5",
    text: "",
    fontSize: 15,
    fontWeight: 700,
    fontFamily: "yahei",
    align: "center",
    overflowMode: "shrink",
    yearText: "25",
    degreeText: "博",
    serialFieldKey: "",
    serialPadLength: 3
  },
  {
    id: 7,
    label: "状态",
    sourceType: "field",
    fieldKey: "field_27",
    text: "调档",
    fontSize: 15,
    fontWeight: 700,
    fontFamily: "songti",
    align: "center",
    overflowMode: "shrink",
    yearText: "25",
    degreeText: "博",
    serialFieldKey: "",
    serialPadLength: 3
  }
];

function nextRowId() {
  const id = rowSeed;
  rowSeed += 1;
  return id;
}

function touchRowSeed(id) {
  const numericId = Number(id);

  if (Number.isFinite(numericId)) {
    rowSeed = Math.max(rowSeed, Math.trunc(numericId) + 1);
  }
}

export function clampNumber(value, min, max, fallback) {
  const numericValue = Number(value);

  if (!Number.isFinite(numericValue)) {
    return fallback;
  }

  return Math.min(max, Math.max(min, numericValue));
}

export function createDefaultLayout() {
  return {
    columns: 4,
    rows: 3,
    gapMm: 0,
    pagePaddingTopMm: 6,
    pagePaddingRightMm: 6,
    pagePaddingBottomMm: 6,
    pagePaddingLeftMm: 6,
    cellPaddingMm: 0,
    showCutLines: true
  };
}

export function normalizeLayout(layout = {}) {
  const defaults = createDefaultLayout();

  return {
    columns: clampNumber(layout.columns, 1, 6, defaults.columns),
    rows: clampNumber(layout.rows, 1, 12, defaults.rows),
    gapMm: clampNumber(layout.gapMm, 0, 10, defaults.gapMm),
    pagePaddingTopMm: clampNumber(layout.pagePaddingTopMm, 0, 30, defaults.pagePaddingTopMm),
    pagePaddingRightMm: clampNumber(layout.pagePaddingRightMm, 0, 30, defaults.pagePaddingRightMm),
    pagePaddingBottomMm: clampNumber(layout.pagePaddingBottomMm, 0, 30, defaults.pagePaddingBottomMm),
    pagePaddingLeftMm: clampNumber(layout.pagePaddingLeftMm, 0, 30, defaults.pagePaddingLeftMm),
    cellPaddingMm: clampNumber(layout.cellPaddingMm, 0, 10, defaults.cellPaddingMm),
    showCutLines: typeof layout.showCutLines === "boolean" ? layout.showCutLines : defaults.showCutLines
  };
}

export function calculateCellMetrics(layout) {
  const safeLayout = normalizeLayout(layout);
  const effectiveGapMm = safeLayout.showCutLines ? 0 : safeLayout.gapMm;
  const innerWidth =
    PAPER_WIDTH_MM -
    safeLayout.pagePaddingLeftMm -
    safeLayout.pagePaddingRightMm -
    effectiveGapMm * (safeLayout.columns - 1);
  const innerHeight =
    PAPER_HEIGHT_MM -
    safeLayout.pagePaddingTopMm -
    safeLayout.pagePaddingBottomMm -
    effectiveGapMm * (safeLayout.rows - 1);

  return {
    widthMm: innerWidth / safeLayout.columns,
    heightMm: innerHeight / safeLayout.rows
  };
}

export function getFontStack(fontFamily) {
  return FONT_STACK_MAP[fontFamily] || FONT_STACK_MAP.songti;
}

export function getTemplateRowFlex(row) {
  if (row.sourceType === "serialCode") {
    return 0.96;
  }

  if (row.label === "姓名") {
    return 1.12;
  }

  if (row.label === "状态") {
    return 0.96;
  }

  return 1;
}

export function normalizeTemplateRow(row = {}) {
  const resolvedId = Number.isFinite(Number(row.id)) ? Number(row.id) : nextRowId();
  touchRowSeed(resolvedId);

  return {
    id: resolvedId,
    label: String(row.label || ""),
    sourceType: ["field", "text", "serialCode"].includes(row.sourceType)
      ? row.sourceType
      : "field",
    fieldKey: String(row.fieldKey || ""),
    text: String(row.text || ""),
    fontSize: clampNumber(row.fontSize, 10, 28, 15),
    fontWeight: [400, 500, 600, 700].includes(Number(row.fontWeight))
      ? Number(row.fontWeight)
      : 500,
    fontFamily: FONT_STACK_MAP[row.fontFamily] ? row.fontFamily : "songti",
    align: ["left", "center", "right"].includes(row.align) ? row.align : "center",
    overflowMode: OVERFLOW_MODE_OPTIONS.some((option) => option.value === row.overflowMode)
      ? row.overflowMode
      : "clip",
    yearText: String(row.yearText || "25"),
    degreeText: String(row.degreeText || "博"),
    serialFieldKey: String(row.serialFieldKey || ""),
    serialPadLength: clampNumber(row.serialPadLength, 1, 6, 3)
  };
}

export function createTemplateRow(overrides = {}) {
  const { id: _ignoredId, ...restOverrides } = overrides;

  return normalizeTemplateRow({
    id: nextRowId(),
    label: "",
    sourceType: "field",
    fieldKey: "",
    text: "",
    fontSize: 15,
    fontWeight: 500,
    fontFamily: "songti",
    align: "center",
    overflowMode: "clip",
    yearText: "25",
    degreeText: "博",
    serialFieldKey: "",
    serialPadLength: 3,
    ...restOverrides
  });
}

export function createDefaultTemplateRows() {
  return DEFAULT_TEMPLATE_ROWS.map((row) => normalizeTemplateRow(row));
}

export function moveTemplateRow(rows, fromIndex, toIndex) {
  const nextRows = [...rows];
  const [targetRow] = nextRows.splice(fromIndex, 1);
  nextRows.splice(toIndex, 0, targetRow);
  return nextRows;
}

export function chunkRecords(records, size) {
  if (!size || size <= 0) {
    return [];
  }

  const chunks = [];

  for (let index = 0; index < records.length; index += size) {
    chunks.push(records.slice(index, index + size));
  }

  return chunks;
}

function findMatchingColumn(label, columns) {
  const matcher = ROW_MATCHERS.find(([matcherLabel]) => matcherLabel === label);

  if (!matcher) {
    return null;
  }

  return (
    columns.find((column) => matcher[1].some((keyword) => column.label.includes(keyword))) || null
  );
}

export function bindTemplateRowsToColumns(rows, columns) {
  return rows.map((row) => {
    if (row.sourceType === "text") {
      return row;
    }

    if (row.sourceType === "serialCode") {
      if (columns.some((column) => column.key === row.serialFieldKey)) {
        return row;
      }

      const serialColumn = findMatchingColumn("编号", columns);

      return serialColumn
        ? normalizeTemplateRow({
            ...row,
            serialFieldKey: serialColumn.key
          })
        : row;
    }

    if (columns.some((column) => column.key === row.fieldKey)) {
      return row;
    }

    const column = findMatchingColumn(row.label, columns);

    if (!column) {
      return row;
    }

    return normalizeTemplateRow({
      ...row,
      fieldKey: column.key
    });
  });
}

export function buildSampleRecord(columns) {
  const sample = {};

  columns.forEach((column, index) => {
    sample[column.key] = `${column.label}${index + 1}`;
  });

  sample.__index = 1;
  return sample;
}

function formatSerialValue(rawValue, padLength, fallbackIndex) {
  const digits = String(rawValue ?? "").replace(/\D/g, "");

  if (digits) {
    return digits.slice(-padLength).padStart(padLength, "0");
  }

  return String(fallbackIndex || 1).padStart(padLength, "0");
}

function getRecordValue(record, key) {
  if (!record || !key) {
    return "";
  }

  const value = record[key];

  return value == null ? "" : String(value).trim();
}

function findFallbackFieldKey(row, columnsMap) {
  if (!row?.label || !columnsMap?.values) {
    return "";
  }

  const normalizedLabel = String(row.label).trim();

  for (const column of columnsMap.values()) {
    if (
      column.label === normalizedLabel ||
      column.label.includes(normalizedLabel) ||
      normalizedLabel.includes(column.label)
    ) {
      return column.key;
    }
  }

  return "";
}

function getCharacterWidthRatio(character) {
  if (!character) {
    return 0;
  }

  if (character === " ") {
    return 0.3;
  }

  if (/[\u4e00-\u9fff\u3400-\u4dbf]/u.test(character)) {
    return 1;
  }

  if (/[A-Z]/.test(character)) {
    return 0.68;
  }

  if (/[a-z]/.test(character)) {
    return 0.56;
  }

  if (/[0-9]/.test(character)) {
    return 0.58;
  }

  return 0.4;
}

function measureLineWidthUnits(line = "") {
  return [...String(line)].reduce((sum, character) => sum + getCharacterWidthRatio(character), 0);
}

function estimateWrappedLineCount(text, fontSize, availableWidthPx) {
  const normalizedText = String(text || "");
  const lines = normalizedText.split(/\r?\n/);

  return lines.reduce((count, line) => {
    const widthPx = measureLineWidthUnits(line) * fontSize;
    return count + Math.max(1, Math.ceil(widthPx / Math.max(availableWidthPx, 1)));
  }, 0);
}

function getTextAreaMetrics(row, layout, templateRows) {
  const safeLayout = normalizeLayout(layout);
  const cellMetrics = calculateCellMetrics(safeLayout);
  const totalFlex = Math.max(
    1,
    (Array.isArray(templateRows) ? templateRows : []).reduce(
      (sum, currentRow) => sum + getTemplateRowFlex(currentRow),
      0
    )
  );
  const rowFlex = getTemplateRowFlex(row);
  const widthPx =
    Math.max(18, cellMetrics.widthMm - safeLayout.cellPaddingMm * 2) * PX_PER_MM - 14;
  const heightPx = (cellMetrics.heightMm * PX_PER_MM * rowFlex) / totalFlex - 8;

  return {
    widthPx: Math.max(40, widthPx),
    heightPx: Math.max(18, heightPx)
  };
}

function resolveShrinkFontSize(row, text, layout, templateRows) {
  const textValue = String(text || "");

  if (!textValue) {
    return row.fontSize;
  }

  const { widthPx, heightPx } = getTextAreaMetrics(row, layout, templateRows);
  const lineUnits = measureLineWidthUnits(textValue.replace(/\r?\n/g, ""));
  let nextFontSize = row.fontSize;

  while (nextFontSize > DEFAULT_MIN_FONT_SIZE) {
    const estimatedWidth = lineUnits * nextFontSize;
    const estimatedHeight = nextFontSize * 1.14;

    if (estimatedWidth <= widthPx && estimatedHeight <= heightPx) {
      break;
    }

    nextFontSize -= 1;
  }

  return Math.max(DEFAULT_MIN_FONT_SIZE, nextFontSize);
}

function resolveWrapFontSize(row, text, layout, templateRows) {
  const textValue = String(text || "");

  if (!textValue) {
    return row.fontSize;
  }

  const { widthPx, heightPx } = getTextAreaMetrics(row, layout, templateRows);
  let nextFontSize = row.fontSize;

  while (nextFontSize > DEFAULT_MIN_FONT_SIZE) {
    const lineCount = estimateWrappedLineCount(textValue, nextFontSize, widthPx);
    const estimatedHeight = lineCount * nextFontSize * 1.22;

    if (estimatedHeight <= heightPx) {
      break;
    }

    nextFontSize -= 1;
  }

  return Math.max(DEFAULT_MIN_FONT_SIZE, nextFontSize);
}

export function getTemplateRowStyle(row, text, layout, templateRows) {
  let fontSize = row.fontSize;

  if (row.overflowMode === "shrink") {
    fontSize = resolveShrinkFontSize(row, text, layout, templateRows);
  }

  if (row.overflowMode === "wrap") {
    fontSize = resolveWrapFontSize(row, text, layout, templateRows);
  }

  return {
    flex: `${getTemplateRowFlex(row)} 1 0`,
    minHeight: 0,
    fontSize: `${fontSize}px`,
    fontWeight: row.fontWeight,
    fontFamily: getFontStack(row.fontFamily),
    textAlign: row.align,
    justifyContent:
      row.align === "left" ? "flex-start" : row.align === "right" ? "flex-end" : "center",
    whiteSpace: row.overflowMode === "wrap" ? "pre-wrap" : "nowrap",
    overflowWrap: row.overflowMode === "wrap" ? "anywhere" : "normal",
    wordBreak: row.overflowMode === "wrap" ? "break-word" : "normal",
    textOverflow: row.overflowMode === "clip" ? "ellipsis" : "clip",
    lineHeight: row.overflowMode === "wrap" ? 1.22 : 1.12
  };
}

export function resolveTemplateRowText(row, record, columnsMap) {
  if (!record) {
    return "";
  }

  if (row.sourceType === "text") {
    return row.text || "";
  }

  if (row.sourceType === "serialCode") {
    const rawSerial = row.serialFieldKey ? record?.[row.serialFieldKey] : record?.__index;
    const serialValue = formatSerialValue(rawSerial, row.serialPadLength, record?.__index);
    const yearText = row.yearText || "25";
    const degreeText = row.degreeText || "博";
    return `${yearText}-${degreeText}-${serialValue}`;
  }

  const directValue = getRecordValue(record, row.fieldKey);
  const fallbackFieldKey = directValue ? "" : findFallbackFieldKey(row, columnsMap);
  const fallbackValue = fallbackFieldKey ? getRecordValue(record, fallbackFieldKey) : "";
  const resolvedValue = directValue || fallbackValue;
  const displayFieldKey = row.fieldKey || fallbackFieldKey;
  const columnLabel = displayFieldKey ? columnsMap.get(displayFieldKey)?.label || "" : "";

  if (resolvedValue) {
    return resolvedValue;
  }

  if (!displayFieldKey) {
    return columnLabel || "未映射";
  }

  return "";
}

function convertBlocksToRows(templateBlocks = []) {
  return [...templateBlocks]
    .sort((left, right) => {
      if (left.y === right.y) {
        return left.x - right.x;
      }

      return left.y - right.y;
    })
    .map((block) =>
      createTemplateRow({
        label: block.label || "",
        sourceType: block.sourceType === "text" ? "text" : "field",
        fieldKey: block.fieldKey || "",
        text: block.text || "",
        fontSize: block.fontSize,
        fontWeight: block.fontWeight,
        fontFamily: "songti",
        align: block.align
      })
    );
}

export function normalizeTemplatePayload(payload = {}) {
  const templateRows = Array.isArray(payload.templateRows)
    ? payload.templateRows.map((row) => normalizeTemplateRow(row))
    : Array.isArray(payload.templateBlocks)
      ? convertBlocksToRows(payload.templateBlocks)
      : createDefaultTemplateRows();

  return {
    version: 4,
    layout: normalizeLayout(payload.layout || {}),
    templateRows
  };
}

export function buildTemplatePayload(layout, templateRows) {
  return {
    version: 4,
    savedAt: new Date().toISOString(),
    layout: normalizeLayout(layout),
    templateRows: templateRows.map((row) => normalizeTemplateRow(row))
  };
}
