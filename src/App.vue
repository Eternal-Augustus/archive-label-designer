<script setup>
import { computed, reactive, ref, watch } from "vue";
import ncepuLogoUrl from "../华北电力大学-logo.svg";
import FieldPalette from "./components/FieldPalette.vue";
import PrintCenterDialog from "./components/PrintCenterDialog.vue";
import PrintPreview from "./components/PrintPreview.vue";
import TemplateCanvas from "./components/TemplateCanvas.vue";
import {
  FONT_FAMILY_OPTIONS,
  OVERFLOW_MODE_OPTIONS,
  bindTemplateRowsToColumns,
  buildSampleRecord,
  buildTemplatePayload,
  chunkRecords,
  createDefaultLayout,
  createDefaultTemplateRows,
  createTemplateRow,
  moveTemplateRow,
  normalizeLayout,
  normalizeTemplatePayload,
  normalizeTemplateRow
} from "./lib/template";

const desktopBridge = window.desktopBridge;
const isElectronRuntime = Boolean(desktopBridge?.chooseWorkbook);

const workbook = reactive({
  fileName: "",
  sheetName: "",
  rowCount: 0,
  columns: [],
  rows: []
});

const state = reactive({
  loading: false,
  busyMessage: "",
  busyDetail: "",
  busyProgress: null,
  feedback: "",
  error: ""
});
const printCenter = reactive({
  open: false,
  loadingPrinters: false,
  printerError: "",
  selectedPrinterName: "",
  printers: [],
  submitting: false
});

const layout = reactive(createDefaultLayout());
const templateRows = ref(createDefaultTemplateRows());
const selectedRowId = ref(templateRows.value[0]?.id || null);
const currentPage = ref(1);

const previewLayout = computed(() => normalizeLayout(layout));
const selectedRow = computed(
  () => templateRows.value.find((row) => row.id === selectedRowId.value) || null
);
const columnsMap = computed(() => new Map(workbook.columns.map((column) => [column.key, column])));
const sampleRecord = computed(() => workbook.rows[0] || buildSampleRecord(workbook.columns));
const records = computed(() => workbook.rows);
const cellsPerPage = computed(() => previewLayout.value.columns * previewLayout.value.rows);
const pages = computed(() => chunkRecords(records.value, cellsPerPage.value));
const pageCount = computed(() => Math.max(1, pages.value.length));

function createDemoWorkbook() {
  const columns = [
    { key: "field_1", label: "流水号" },
    { key: "field_4", label: "姓名" },
    { key: "field_5", label: "学号" },
    { key: "field_7", label: "班级" },
    { key: "field_27", label: "状态" }
  ];

  const names = [
    "陈晓",
    "田宇轩",
    "丁浩",
    "四浩明",
    "杜思源",
    "何天鹏",
    "张谦",
    "张笑凡",
    "杨琳",
    "闻宪",
    "曲梓涵",
    "张耀中"
  ];

  const rows = Array.from({ length: 12 }, (_value, index) => ({
    __rowNumber: index + 2,
    __index: index + 1,
    field_1: String(index + 1).padStart(4, "0"),
    field_4: names[index],
    field_5: `12025210${String(index + 1).padStart(4, "0")}`,
    field_7: "研电2501",
    field_27: "调档"
  }));

  return {
    fileName: "浏览器演示数据.xlsx",
    sheetName: "Sheet1",
    rowCount: rows.length,
    columns,
    rows
  };
}

watch(pageCount, (value) => {
  if (currentPage.value > value) {
    currentPage.value = value;
  }
});

watch(
  templateRows,
  (rows) => {
    if (!rows.some((row) => row.id === selectedRowId.value)) {
      selectedRowId.value = rows[0]?.id || null;
    }
  },
  {
    deep: true
  }
);

function setFeedback(message = "", error = "") {
  state.feedback = message;
  state.error = error;
}

function setBusyState({ loading, message = "", detail = "", progress = null }) {
  state.loading = loading;
  state.busyMessage = message;
  state.busyDetail = detail;
  state.busyProgress = progress;
}

function assignLayout(sourceLayout) {
  Object.assign(layout, normalizeLayout(sourceLayout));
}

function applyWorkbook(result) {
  workbook.fileName = result.fileName;
  workbook.sheetName = result.sheetName;
  workbook.rowCount = result.rowCount;
  workbook.columns = result.columns;
  workbook.rows = result.rows;
  templateRows.value = bindTemplateRowsToColumns(templateRows.value, workbook.columns);
  selectedRowId.value = templateRows.value[0]?.id || null;
  currentPage.value = 1;
}

function pickDefaultPrinter(printers) {
  return printers.find((printer) => printer.isDefault)?.name || printers[0]?.name || "";
}

if (!isElectronRuntime) {
  applyWorkbook(createDemoWorkbook());
  setFeedback("已载入演示数据，可先调整模板结构、分页和打印样式。");
}

async function chooseWorkbook() {
  if (!isElectronRuntime) {
    setFeedback("", "当前环境仅支持页面预览，无法直接读取本地 Excel。");
    return;
  }

  setBusyState({
    loading: true,
    message: "正在读取 Excel..."
  });
  setFeedback();

  try {
    const result = await desktopBridge.chooseWorkbook();

    if (!result) {
      return;
    }

    applyWorkbook(result);
    setFeedback(`已载入 ${result.fileName}，共 ${result.rowCount} 条数据。`);
  } catch (error) {
    setFeedback("", error.message || "读取 Excel 失败。");
  } finally {
    setBusyState({
      loading: false
    });
  }
}

async function loadPrinters() {
  if (!desktopBridge?.listPrinters) {
    printCenter.printerError = "当前环境仅支持页面预览，无法读取系统输出设备。";
    return;
  }

  printCenter.loadingPrinters = true;
  printCenter.printerError = "";

  try {
    const printers = await desktopBridge.listPrinters();
    printCenter.printers = Array.isArray(printers) ? printers : [];

    if (!printCenter.printers.some((printer) => printer.name === printCenter.selectedPrinterName)) {
      printCenter.selectedPrinterName = pickDefaultPrinter(printCenter.printers);
    }
  } catch (error) {
    printCenter.printerError = error.message || "读取系统打印机失败。";
  } finally {
    printCenter.loadingPrinters = false;
  }
}

async function openPrintCenter() {
  if (!isElectronRuntime) {
    setFeedback("", "当前环境仅支持页面预览，无法调用系统打印。");
    return;
  }

  printCenter.open = true;
  printCenter.printerError = "";
  await loadPrinters();
}

function closePrintCenter() {
  if (printCenter.submitting) {
    return;
  }

  printCenter.open = false;
}

function updatePreviewPage(nextPage) {
  currentPage.value = nextPage;
}

function insertRow(row, targetIndex = null) {
  const nextRows = [...templateRows.value];

  if (targetIndex == null || targetIndex < 0 || targetIndex > nextRows.length) {
    nextRows.push(row);
  } else {
    nextRows.splice(targetIndex, 0, row);
  }

  templateRows.value = nextRows;
  selectedRowId.value = row.id;
}

function getSuggestedOverflowMode(label = "") {
  if (["学院", "专业"].some((keyword) => label.includes(keyword))) {
    return "wrap";
  }

  if (["姓名", "学号", "身份证", "编号", "班级"].some((keyword) => label.includes(keyword))) {
    return "shrink";
  }

  return "clip";
}

function createFieldRow(column) {
  return createTemplateRow({
    label: column.label,
    sourceType: "field",
    fieldKey: column.key,
    fontSize: ["学号", "身份证"].some((keyword) => column.label.includes(keyword)) ? 14 : 15,
    fontWeight: column.label.includes("姓名") ? 700 : 500,
    fontFamily: ["学号", "身份证"].some((keyword) => column.label.includes(keyword))
      ? "yahei"
      : "songti",
    overflowMode: getSuggestedOverflowMode(column.label)
  });
}

function addFieldRow(column, targetIndex = null) {
  insertRow(createFieldRow(column), targetIndex);
}

function addStaticRow(targetIndex = null) {
  insertRow(
    createTemplateRow({
      label: "静态文本",
      sourceType: "text",
      text: "请输入固定文案",
      fontSize: 15,
      fontWeight: 500,
      fontFamily: "songti",
      overflowMode: "clip"
    }),
    targetIndex
  );
}

function addSerialRow(targetIndex = null) {
  insertRow(
    createTemplateRow({
      label: "编号",
      sourceType: "serialCode",
      fontSize: 14,
      fontWeight: 700,
      fontFamily: "songti",
      overflowMode: "shrink",
      yearText: "25",
      degreeText: "博",
      serialPadLength: 3
    }),
    targetIndex
  );
}

function restoreDefaultTemplate() {
  const defaults = createDefaultTemplateRows();
  templateRows.value = workbook.columns.length
    ? bindTemplateRowsToColumns(defaults, workbook.columns)
    : defaults;
  selectedRowId.value = templateRows.value[0]?.id || null;
  setFeedback("已恢复默认模板。");
}

function selectRow(rowId) {
  selectedRowId.value = rowId;
}

function patchRow(rowId, patch) {
  templateRows.value = templateRows.value.map((row) =>
    row.id === rowId ? normalizeTemplateRow({ ...row, ...patch }) : row
  );
}

function updateSelectedRow(key, value) {
  if (!selectedRow.value) {
    return;
  }

  patchRow(selectedRow.value.id, {
    [key]: value
  });
}

function handleMoveRow({ fromRowId, toRowId }) {
  const fromIndex = templateRows.value.findIndex((row) => row.id === fromRowId);

  if (fromIndex < 0) {
    return;
  }

  const toIndex =
    toRowId == null
      ? templateRows.value.length - 1
      : templateRows.value.findIndex((row) => row.id === toRowId);

  if (toIndex < 0 || fromIndex === toIndex) {
    return;
  }

  templateRows.value = moveTemplateRow(templateRows.value, fromIndex, toIndex);
}

function handleDropColumn({ column, targetRowId }) {
  const targetIndex =
    targetRowId == null
      ? templateRows.value.length
      : templateRows.value.findIndex((row) => row.id === targetRowId);

  addFieldRow(column, targetIndex < 0 ? null : targetIndex);
}

function removeSelectedRow() {
  if (!selectedRow.value) {
    return;
  }

  templateRows.value = templateRows.value.filter((row) => row.id !== selectedRow.value.id);
}

function duplicateSelectedRow() {
  if (!selectedRow.value) {
    return;
  }

  insertRow(
    createTemplateRow({
      ...selectedRow.value,
      label: `${selectedRow.value.label || "字段"} 副本`
    }),
    templateRows.value.findIndex((row) => row.id === selectedRow.value.id) + 1
  );
}

async function saveTemplate() {
  if (!desktopBridge?.saveTemplate) {
    setFeedback("", "当前环境仅支持页面预览，无法保存模板文件。");
    return;
  }

  setBusyState({
    loading: true,
    message: "正在保存模板..."
  });
  setFeedback();

  try {
    const savedPath = await desktopBridge.saveTemplate(
      buildTemplatePayload(previewLayout.value, templateRows.value)
    );

    if (savedPath) {
      setFeedback(`模板已保存到 ${savedPath}`);
    }
  } catch (error) {
    setFeedback("", error.message || "保存模板失败。");
  } finally {
    setBusyState({
      loading: false
    });
  }
}

async function loadTemplate() {
  if (!desktopBridge?.loadTemplate) {
    setFeedback("", "当前环境仅支持页面预览，无法加载模板文件。");
    return;
  }

  setBusyState({
    loading: true,
    message: "正在读取模板..."
  });
  setFeedback();

  try {
    const payload = await desktopBridge.loadTemplate();

    if (!payload) {
      return;
    }

    const normalized = normalizeTemplatePayload(payload);
    assignLayout(normalized.layout);
    templateRows.value = workbook.columns.length
      ? bindTemplateRowsToColumns(normalized.templateRows, workbook.columns)
      : normalized.templateRows;
    selectedRowId.value = templateRows.value[0]?.id || null;

    setFeedback("模板已载入。");
  } catch (error) {
    setFeedback("", error.message || "加载模板失败。");
  } finally {
    setBusyState({
      loading: false
    });
  }
}

async function printLabels() {
  if (!desktopBridge?.printLabels) {
    setFeedback("", "当前环境仅支持页面预览，无法调用系统打印。");
    return;
  }

  if (!printCenter.selectedPrinterName) {
    printCenter.printerError = "请先在打印与输出窗口中选择一个设备。";
    return;
  }

  const selectedPrinter = printCenter.printers.find(
    (printer) => printer.name === printCenter.selectedPrinterName
  );
  const isVirtualPrinter = Boolean(selectedPrinter?.isVirtual);
  let hintTimer = null;
  let longWaitTimer = null;

  printCenter.submitting = true;
  setBusyState({
    loading: true,
    message: isVirtualPrinter ? "正在准备 PDF 输出..." : "正在发送打印任务...",
    detail: isVirtualPrinter
      ? "系统即将接管保存流程，请在稍后弹出的 Windows 窗口中确认文件名和保存位置。"
      : `正在提交到 ${selectedPrinter?.displayName || printCenter.selectedPrinterName}。`,
    progress: 36
  });
  setFeedback();

  if (isVirtualPrinter) {
    hintTimer = setTimeout(() => {
      setBusyState({
        loading: true,
        message: "正在等待系统确认...",
        detail: "如果保存窗口尚未出现，请切换到 Windows 打印或保存界面继续完成。",
        progress: 52
      });
    }, 4000);

    longWaitTimer = setTimeout(() => {
      setBusyState({
        loading: true,
        message: "系统仍在处理中...",
        detail: "待 Windows 完成保存确认后，当前窗口会自动恢复可操作状态。",
        progress: 60
      });
    }, 15000);
  }

  try {
    await desktopBridge.printLabels({
      deviceName: printCenter.selectedPrinterName,
      silent: !isVirtualPrinter
    });
    setBusyState({
      loading: true,
      message: "任务已交给系统",
      detail: isVirtualPrinter
        ? "请在 Windows 保存窗口中完成 PDF 输出。"
        : "系统正在将版面发送到所选打印机。",
      progress: 100
    });
    setFeedback(
      isVirtualPrinter
        ? "已进入系统 PDF 输出流程。"
        : `已提交到 ${selectedPrinter?.displayName || printCenter.selectedPrinterName}。`
    );
  } catch (error) {
    const rawMessage = error.message || "打印失败。";

    if (/Print job canceled/i.test(rawMessage)) {
      setFeedback("", "打印已取消。");
    } else {
      setFeedback("", rawMessage);
    }
  } finally {
    if (hintTimer) {
      clearTimeout(hintTimer);
    }

    if (longWaitTimer) {
      clearTimeout(longWaitTimer);
    }

    setTimeout(() => {
      setBusyState({
        loading: false
      });
    }, 500);
    printCenter.submitting = false;
  }
}
</script>

<template>
  <div class="app-shell">
    <aside class="left-rail">
      <div class="brand-card">
        <div class="brand-header">
          <img class="brand-logo" :src="ncepuLogoUrl" alt="华北电力大学校徽" />
          <div class="brand-title-block">
            <p class="eyebrow">NCEPU Archive Workflow</p>
            <h1>档案标签排版器</h1>
            <div class="brand-tags">
              <span>离线运行</span>
              <span>A4 标签</span>
              <span>Excel 导入</span>
            </div>
          </div>
        </div>
        <p class="brand-copy">
          面向档案整理场景的离线排版工具。导入 Excel，调整单格模板，核对分页，再交给系统打印或
          Microsoft Print to PDF。
        </p>
        <p v-if="!isElectronRuntime" class="runtime-note">
          当前处于页面预览模式，可查看模板结构与分页效果；Excel 读取、模板文件存取和系统打印需在桌面客户端中完成。
        </p>
      </div>

      <section class="panel import-panel">
        <div class="section-heading">
          <h2>1. 导入数据</h2>
          <p>读取学生信息 Excel，系统会识别表头，并用于字段绑定与默认模板匹配。</p>
        </div>

        <button
          class="primary-button"
          type="button"
          :disabled="!isElectronRuntime"
          @click="chooseWorkbook"
        >
          选择 Excel 文件
        </button>

        <div class="dataset-summary">
          <div>
            <span>文件</span>
            <strong>{{ workbook.fileName || "尚未导入" }}</strong>
          </div>
          <div>
            <span>工作表</span>
            <strong>{{ workbook.sheetName || "-" }}</strong>
          </div>
          <div>
            <span>记录数</span>
            <strong>{{ workbook.rowCount || 0 }}</strong>
          </div>
        </div>

        <p v-if="state.feedback" class="feedback success">{{ state.feedback }}</p>
        <p v-if="state.error" class="feedback error">{{ state.error }}</p>
      </section>

      <FieldPalette
        :columns="workbook.columns"
        @add-field-row="addFieldRow"
        @add-static-row="addStaticRow"
        @add-serial-row="addSerialRow"
      />

      <section class="panel layout-panel">
        <div class="section-heading">
          <h2>2. 页面排版</h2>
          <p>默认采用 4 x 3 布局，可按标签纸尺寸继续微调页边距、间距和边线。</p>
        </div>

        <div class="form-grid compact">
          <label>
            <span>列数</span>
            <input v-model.number="layout.columns" type="number" min="1" max="6" />
          </label>
          <label>
            <span>行数</span>
            <input v-model.number="layout.rows" type="number" min="1" max="12" />
          </label>
          <label>
            <span>页边距上(mm)</span>
            <input v-model.number="layout.pagePaddingTopMm" type="number" min="0" max="30" step="0.5" />
          </label>
          <label>
            <span>页边距右(mm)</span>
            <input v-model.number="layout.pagePaddingRightMm" type="number" min="0" max="30" step="0.5" />
          </label>
          <label>
            <span>页边距下(mm)</span>
            <input v-model.number="layout.pagePaddingBottomMm" type="number" min="0" max="30" step="0.5" />
          </label>
          <label>
            <span>页边距左(mm)</span>
            <input v-model.number="layout.pagePaddingLeftMm" type="number" min="0" max="30" step="0.5" />
          </label>
          <label>
            <span>格间距(mm)</span>
            <input v-model.number="layout.gapMm" type="number" min="0" max="10" step="0.2" />
          </label>
          <label>
            <span>单格内边距(mm)</span>
            <input v-model.number="layout.cellPaddingMm" type="number" min="0" max="10" step="0.2" />
          </label>
        </div>

        <label class="switch-row">
          <input v-model="layout.showCutLines" type="checkbox" />
          <span>显示标签格边线</span>
        </label>
      </section>
    </aside>

    <main class="workspace">
      <section class="panel template-panel">
        <div class="section-heading">
          <h2>3. 行单元模板</h2>
          <p>每个字段都会以整行单元的方式填满单格，拖动即可调整上下顺序，阅读方式更接近表格。</p>
        </div>

        <div class="template-layout">
          <TemplateCanvas
            :rows="templateRows"
            :selected-row-id="selectedRowId"
            :sample-record="sampleRecord"
            :columns-map="columnsMap"
            :layout="previewLayout"
            @select-row="selectRow"
            @move-row="handleMoveRow"
            @drop-column="handleDropColumn"
            @add-static-row="addStaticRow"
          />

          <div class="editor-panel">
            <div class="editor-header">
              <div>
                <h3>当前行设置</h3>
                <p class="editor-tip">保留常用排版项：字段来源、字体、字号、字重、对齐、溢出处理和编号规则。</p>
              </div>

              <div class="editor-actions">
                <button class="ghost-button" type="button" @click="addStaticRow()">
                  添加静态行
                </button>
                <button class="ghost-button" type="button" @click="addSerialRow()">
                  添加编号行
                </button>
                <button class="ghost-button" type="button" @click="duplicateSelectedRow">
                  复制当前行
                </button>
                <button class="ghost-button danger" type="button" @click="removeSelectedRow">
                  删除当前行
                </button>
              </div>
            </div>

            <div v-if="selectedRow" class="form-grid">
              <label>
                <span>行标题</span>
                <input
                  :value="selectedRow.label"
                  type="text"
                  @input="updateSelectedRow('label', $event.target.value)"
                />
              </label>
              <label>
                <span>内容类型</span>
                <select
                  :value="selectedRow.sourceType"
                  @change="updateSelectedRow('sourceType', $event.target.value)"
                >
                  <option value="field">Excel 字段</option>
                  <option value="text">静态文本</option>
                  <option value="serialCode">编号合成</option>
                </select>
              </label>

              <label v-if="selectedRow.sourceType === 'field'">
                <span>绑定字段</span>
                <select
                  :value="selectedRow.fieldKey"
                  @change="updateSelectedRow('fieldKey', $event.target.value)"
                >
                  <option value="">请选择字段</option>
                  <option
                    v-for="column in workbook.columns"
                    :key="column.key"
                    :value="column.key"
                  >
                    {{ column.label }}
                  </option>
                </select>
              </label>

              <template v-if="selectedRow.sourceType === 'text'">
                <label>
                  <span>静态文本</span>
                  <input
                    :value="selectedRow.text"
                    type="text"
                    @input="updateSelectedRow('text', $event.target.value)"
                  />
                </label>
              </template>

              <template v-if="selectedRow.sourceType === 'serialCode'">
                <label>
                  <span>入学年份</span>
                  <input
                    :value="selectedRow.yearText"
                    type="text"
                    placeholder="例如 25"
                    @input="updateSelectedRow('yearText', $event.target.value)"
                  />
                </label>
                <label>
                  <span>培养层次简称</span>
                  <input
                    :value="selectedRow.degreeText"
                    type="text"
                    placeholder="博 / 硕"
                    @input="updateSelectedRow('degreeText', $event.target.value)"
                  />
                </label>
                <label>
                  <span>流水号字段</span>
                  <select
                    :value="selectedRow.serialFieldKey"
                    @change="updateSelectedRow('serialFieldKey', $event.target.value)"
                  >
                    <option value="">按数据顺序自动编号</option>
                    <option
                      v-for="column in workbook.columns"
                      :key="column.key"
                      :value="column.key"
                    >
                      {{ column.label }}
                    </option>
                  </select>
                </label>
                <label>
                  <span>流水号位数</span>
                  <input
                    :value="selectedRow.serialPadLength"
                    type="number"
                    min="1"
                    max="6"
                    step="1"
                    @input="updateSelectedRow('serialPadLength', Number($event.target.value))"
                  />
                </label>
              </template>

              <label>
                <span>字体</span>
                <select
                  :value="selectedRow.fontFamily"
                  @change="updateSelectedRow('fontFamily', $event.target.value)"
                >
                  <option
                    v-for="font in FONT_FAMILY_OPTIONS"
                    :key="font.value"
                    :value="font.value"
                  >
                    {{ font.label }}
                  </option>
                </select>
              </label>
              <label>
                <span>字号(px)</span>
                <input
                  :value="selectedRow.fontSize"
                  type="number"
                  min="10"
                  max="28"
                  step="1"
                  @input="updateSelectedRow('fontSize', Number($event.target.value))"
                />
              </label>
              <label>
                <span>字重</span>
                <select
                  :value="selectedRow.fontWeight"
                  @change="updateSelectedRow('fontWeight', Number($event.target.value))"
                >
                  <option :value="400">常规</option>
                  <option :value="500">中等</option>
                  <option :value="700">加粗</option>
                </select>
              </label>
              <label>
                <span>对齐</span>
                <select
                  :value="selectedRow.align"
                  @change="updateSelectedRow('align', $event.target.value)"
                >
                  <option value="left">左对齐</option>
                  <option value="center">居中</option>
                  <option value="right">右对齐</option>
                </select>
              </label>
              <label>
                <span>溢出处理</span>
                <select
                  :value="selectedRow.overflowMode"
                  @change="updateSelectedRow('overflowMode', $event.target.value)"
                >
                  <option
                    v-for="option in OVERFLOW_MODE_OPTIONS"
                    :key="option.value"
                    :value="option.value"
                  >
                    {{ option.label }}
                  </option>
                </select>
              </label>
            </div>

            <div v-if="!selectedRow" class="empty-note">
              先点击一个字段行，或从左侧字段库拖进来一行。
            </div>
          </div>
        </div>
      </section>

      <section class="toolbar-panel">
        <div class="toolbar-cluster">
          <button
            class="ghost-button"
            type="button"
            :disabled="!isElectronRuntime"
            @click="loadTemplate"
          >
            载入模板
          </button>
          <button
            class="ghost-button"
            type="button"
            :disabled="!isElectronRuntime"
            @click="saveTemplate"
          >
            保存模板
          </button>
          <button class="ghost-button" type="button" @click="restoreDefaultTemplate">
            恢复默认
          </button>
        </div>

        <div class="pager">
          <button
            class="ghost-button"
            type="button"
            :disabled="currentPage <= 1"
            @click="currentPage -= 1"
          >
            上一页
          </button>
          <span>第 {{ currentPage }} / {{ pageCount }} 页</span>
          <button
            class="ghost-button"
            type="button"
            :disabled="currentPage >= pageCount"
            @click="currentPage += 1"
          >
            下一页
          </button>
        </div>

        <div class="toolbar-actions">
          <button
            class="primary-button"
            type="button"
            :disabled="!isElectronRuntime"
            @click="openPrintCenter"
          >
            打印与输出
          </button>
        </div>

        <p class="toolbar-note">
          建议先在应用内核对分页、字号和边线，再发送到系统打印或 Microsoft Print to PDF。
        </p>
      </section>

      <PrintPreview
        :layout="previewLayout"
        :template-rows="templateRows"
        :pages="pages"
        :columns-map="columnsMap"
        :current-page="currentPage"
      />
    </main>

    <PrintCenterDialog
      :open="printCenter.open"
      :printers="printCenter.printers"
      :selected-printer-name="printCenter.selectedPrinterName"
      :loading-printers="printCenter.loadingPrinters"
      :printer-error="printCenter.printerError"
      :is-submitting="printCenter.submitting"
      :layout="previewLayout"
      :template-rows="templateRows"
      :pages="pages"
      :columns-map="columnsMap"
      :current-page="currentPage"
      @close="closePrintCenter"
      @refresh-printers="loadPrinters"
      @update:selected-printer-name="printCenter.selectedPrinterName = $event"
      @update:current-page="updatePreviewPage"
      @print="printLabels"
    />

    <div v-if="state.loading" class="busy-mask" data-print-exclude="true">
      <div class="busy-card">
        <strong>{{ state.busyMessage }}</strong>
        <span>{{ state.busyDetail || "请稍候，整个过程都在本机离线完成。" }}</span>
        <div v-if="typeof state.busyProgress === 'number'" class="busy-progress">
          <div
            class="busy-progress-bar"
            :style="{ width: `${Math.max(0, Math.min(100, state.busyProgress))}%` }"
          ></div>
        </div>
        <small v-if="typeof state.busyProgress === 'number'" class="busy-progress-text">
          {{ Math.round(state.busyProgress) }}%
        </small>
      </div>
    </div>
  </div>
</template>
