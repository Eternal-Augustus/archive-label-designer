<script setup>
import { computed } from "vue";
import {
  calculateCellMetrics,
  getTemplateRowStyle,
  resolveTemplateRowText
} from "../lib/template";

const props = defineProps({
  rows: {
    type: Array,
    required: true
  },
  selectedRowId: {
    type: Number,
    default: null
  },
  sampleRecord: {
    type: Object,
    required: true
  },
  columnsMap: {
    type: Object,
    required: true
  },
  layout: {
    type: Object,
    required: true
  }
});

const emit = defineEmits([
  "select-row",
  "move-row",
  "drop-column",
  "add-static-row"
]);

const cellMetrics = computed(() => calculateCellMetrics(props.layout));
const canvasStyle = computed(() => ({
  "--cell-aspect-ratio": `${cellMetrics.value.widthMm} / ${cellMetrics.value.heightMm}`
}));

function renderRowText(row) {
  return resolveTemplateRowText(row, props.sampleRecord, props.columnsMap);
}

function getRowStyle(row) {
  return getTemplateRowStyle(
    row,
    renderRowText(row),
    props.layout,
    props.rows
  );
}

function handleDragStart(event, row) {
  event.dataTransfer.effectAllowed = "move";
  event.dataTransfer.setData("application/x-template-row-id", String(row.id));
}

function handleDrop(event, targetRowId = null) {
  event.preventDefault();

  const rowId = event.dataTransfer.getData("application/x-template-row-id");

  if (rowId) {
    emit("move-row", {
      fromRowId: Number(rowId),
      toRowId: targetRowId
    });
    return;
  }

  const rawColumn = event.dataTransfer.getData("application/x-archive-column");

  if (!rawColumn) {
    return;
  }

  try {
    const column = JSON.parse(rawColumn);
    emit("drop-column", {
      column,
      targetRowId
    });
  } catch (_error) {
    // Ignore invalid drops.
  }
}
</script>

<template>
  <div class="template-dropzone">
    <div class="mini-sheet" :style="canvasStyle">
      <div
        v-for="row in rows"
        :key="row.id"
        class="mini-sheet-row"
        :class="{ active: row.id === selectedRowId }"
        :style="getRowStyle(row)"
        draggable="true"
        @click="emit('select-row', row.id)"
        @dragstart="handleDragStart($event, row)"
        @dragover.prevent
        @drop="handleDrop($event, row.id)"
      >
        <span class="mini-sheet-row-text">{{ renderRowText(row) }}</span>
      </div>

      <button
        v-if="!rows.length"
        type="button"
        class="mini-sheet-empty"
        @click="emit('add-static-row')"
        @dragover.prevent
        @drop="handleDrop($event, null)"
      >
        <strong>拖入字段开始排版</strong>
        <span>字段会自动铺满整行，阅读方式更接近表格单元格。</span>
      </button>
    </div>
  </div>
</template>
