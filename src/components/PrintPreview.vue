<script setup>
import {
  getTemplateRowStyle,
  resolveTemplateRowText
} from "../lib/template";

function renderRowText(row, record, columnsMap) {
  return resolveTemplateRowText(row, record, columnsMap);
}

function getRowStyle(row, record, columnsMap, layout, templateRows) {
  return getTemplateRowStyle(
    row,
    renderRowText(row, record, columnsMap),
    layout,
    templateRows
  );
}

function getCellClass(cellIndex, layout) {
  const zeroBasedIndex = cellIndex - 1;
  const columnIndex = zeroBasedIndex % layout.columns;
  const rowIndex = Math.floor(zeroBasedIndex / layout.columns);

  return {
    "cell-last-column": columnIndex === layout.columns - 1,
    "cell-last-row": rowIndex === layout.rows - 1
  };
}

function buildPrintRows(page, layout) {
  return Array.from({ length: layout.rows }, (_unused, rowIndex) =>
    Array.from(
      { length: layout.columns },
      (_value, columnIndex) => page[rowIndex * layout.columns + columnIndex] || null
    )
  );
}

defineProps({
  layout: {
    type: Object,
    required: true
  },
  templateRows: {
    type: Array,
    required: true
  },
  pages: {
    type: Array,
    required: true
  },
  columnsMap: {
    type: Object,
    required: true
  },
  currentPage: {
    type: Number,
    default: 1
  },
  showHeading: {
    type: Boolean,
    default: true
  },
  showMeta: {
    type: Boolean,
    default: true
  },
  renderPrintStage: {
    type: Boolean,
    default: true
  },
  bare: {
    type: Boolean,
    default: false
  }
});
</script>

<template>
  <section class="preview-shell" :class="{ panel: !bare, 'preview-shell-bare': bare }">
    <div v-if="showHeading" class="section-heading">
      <h2>版面预览</h2>
      <p>预览区与实际输出共用同一套模板数据，用来核对分页、边线和文字摆放。</p>
    </div>

    <div v-if="showMeta" class="preview-meta">
      <span>每页 {{ layout.columns * layout.rows }} 人</span>
      <span>共 {{ pages.length || 1 }} 页</span>
      <span>默认 4 x 3</span>
    </div>

    <div class="screen-preview">
      <div
        v-for="(page, pageIndex) in pages.length ? [pages[currentPage - 1]] : [[]]"
        :key="pageIndex"
        class="preview-paper"
        :style="{
          '--grid-columns': layout.columns,
          '--grid-rows': layout.rows,
          '--grid-gap': layout.showCutLines ? '1px' : `${layout.gapMm}mm`,
          '--page-padding-top': `${layout.pagePaddingTopMm}mm`,
          '--page-padding-right': `${layout.pagePaddingRightMm}mm`,
          '--page-padding-bottom': `${layout.pagePaddingBottomMm}mm`,
          '--page-padding-left': `${layout.pagePaddingLeftMm}mm`,
          '--cell-padding': `${layout.cellPaddingMm}mm`
        }"
      >
        <div class="preview-grid" :class="{ outlined: layout.showCutLines }">
          <div
            v-for="cellIndex in layout.columns * layout.rows"
            :key="cellIndex"
            class="preview-cell"
            :class="getCellClass(cellIndex, layout)"
          >
            <div class="preview-sheet">
              <div
                v-for="row in templateRows"
                :key="row.id"
                class="preview-sheet-row"
                :style="getRowStyle(row, page[cellIndex - 1], columnsMap, layout, templateRows)"
              >
                <span class="preview-sheet-row-text">
                  {{ renderRowText(row, page[cellIndex - 1], columnsMap) }}
                </span>
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>

    <div v-if="renderPrintStage" class="print-stage" aria-hidden="true">
      <div
        v-for="(page, pageIndex) in pages.length ? pages : [[]]"
        :key="pageIndex"
        class="print-paper"
        :style="{
          '--grid-columns': layout.columns,
          '--grid-rows': layout.rows,
          '--grid-gap': layout.showCutLines ? '1px' : `${layout.gapMm}mm`,
          '--page-padding-top': `${layout.pagePaddingTopMm}mm`,
          '--page-padding-right': `${layout.pagePaddingRightMm}mm`,
          '--page-padding-bottom': `${layout.pagePaddingBottomMm}mm`,
          '--page-padding-left': `${layout.pagePaddingLeftMm}mm`,
          '--cell-padding': `${layout.cellPaddingMm}mm`
        }"
      >
        <table v-if="layout.showCutLines" class="print-grid-table">
          <tbody>
            <tr
              v-for="(printRow, rowIndex) in buildPrintRows(page, layout)"
              :key="`print-row-${pageIndex}-${rowIndex}`"
              class="print-grid-row"
            >
              <td
                v-for="(record, columnIndex) in printRow"
                :key="`print-cell-${pageIndex}-${rowIndex}-${columnIndex}`"
                class="print-grid-cell"
              >
                <div class="preview-sheet">
                  <div
                    v-for="row in templateRows"
                    :key="row.id"
                    class="preview-sheet-row"
                    :style="getRowStyle(row, record, columnsMap, layout, templateRows)"
                  >
                    <span class="preview-sheet-row-text">
                      {{ renderRowText(row, record, columnsMap) }}
                    </span>
                  </div>
                </div>
              </td>
            </tr>
          </tbody>
        </table>

        <div v-else class="preview-grid">
          <div
            v-for="cellIndex in layout.columns * layout.rows"
            :key="cellIndex"
            class="preview-cell"
          >
            <div class="preview-sheet">
              <div
                v-for="row in templateRows"
                :key="row.id"
                class="preview-sheet-row"
                :style="getRowStyle(row, page[cellIndex - 1], columnsMap, layout, templateRows)"
              >
                <span class="preview-sheet-row-text">
                  {{ renderRowText(row, page[cellIndex - 1], columnsMap) }}
                </span>
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  </section>
</template>
