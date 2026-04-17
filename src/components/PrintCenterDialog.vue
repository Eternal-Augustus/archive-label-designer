<script setup>
import { computed } from "vue";
import PrintPreview from "./PrintPreview.vue";

const props = defineProps({
  open: {
    type: Boolean,
    default: false
  },
  printers: {
    type: Array,
    default: () => []
  },
  selectedPrinterName: {
    type: String,
    default: ""
  },
  loadingPrinters: {
    type: Boolean,
    default: false
  },
  printerError: {
    type: String,
    default: ""
  },
  isSubmitting: {
    type: Boolean,
    default: false
  },
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
  }
});

const emit = defineEmits([
  "close",
  "refresh-printers",
  "update:selectedPrinterName",
  "update:currentPage",
  "print"
]);

const selectedPrinter = computed(
  () => props.printers.find((printer) => printer.name === props.selectedPrinterName) || null
);
const pageCount = computed(() => Math.max(1, props.pages.length));
const printerHint = computed(() => {
  if (!selectedPrinter.value) {
    return "先选择一个输出设备，再发送当前版面。";
  }

  if (selectedPrinter.value.isVirtual) {
    return "当前是虚拟打印机。选择 Microsoft Print to PDF 后，Windows 会继续让您确认保存位置。";
  }

  return "确认版面无误后，当前内容会直接发送到所选打印机。";
});

function updateSelectedPrinter(event) {
  emit("update:selectedPrinterName", event.target.value);
}

function updatePage(nextPage) {
  const safePage = Math.max(1, Math.min(pageCount.value, nextPage));
  emit("update:currentPage", safePage);
}
</script>

<template>
  <div v-if="open" class="print-center-mask" data-print-exclude="true">
    <section class="print-center-window">
      <aside class="print-center-sidebar">
        <div class="print-center-header">
          <div>
            <h2>打印与输出</h2>
            <p>先在这里核对分页和边框效果，再决定发送到打印机或 PDF 输出。</p>
          </div>

          <button class="ghost-button" type="button" @click="emit('close')">关闭</button>
        </div>

        <label class="print-center-field">
          <span>输出设备</span>
          <select
            :value="selectedPrinterName"
            :disabled="loadingPrinters || isSubmitting"
            @change="updateSelectedPrinter"
          >
            <option value="">请选择设备</option>
            <option
              v-for="printer in printers"
              :key="printer.name"
              :value="printer.name"
            >
              {{ printer.displayName }}{{ printer.isDefault ? "（默认）" : "" }}
            </option>
          </select>
        </label>

        <div class="print-center-actions">
          <button
            class="ghost-button"
            type="button"
            :disabled="loadingPrinters || isSubmitting"
            @click="emit('refresh-printers')"
          >
            刷新设备
          </button>

          <button
            class="primary-button"
            type="button"
            :disabled="!selectedPrinterName || loadingPrinters || isSubmitting"
            @click="emit('print')"
          >
            {{ isSubmitting ? "正在提交..." : "发送到系统" }}
          </button>
        </div>

        <p v-if="loadingPrinters" class="print-center-note">
          正在读取系统输出设备...
        </p>
        <p v-else class="print-center-note">
          {{ printerHint }}
        </p>
        <p v-if="printerError" class="feedback error">{{ printerError }}</p>

        <div class="print-center-pager">
          <button
            class="ghost-button"
            type="button"
            :disabled="currentPage <= 1"
            @click="updatePage(currentPage - 1)"
          >
            上一页
          </button>
          <span>第 {{ currentPage }} / {{ pageCount }} 页</span>
          <button
            class="ghost-button"
            type="button"
            :disabled="currentPage >= pageCount"
            @click="updatePage(currentPage + 1)"
          >
            下一页
          </button>
        </div>
      </aside>

      <div class="print-center-preview">
        <PrintPreview
          bare
          :show-heading="false"
          :show-meta="false"
          :render-print-stage="false"
          :layout="layout"
          :template-rows="templateRows"
          :pages="pages"
          :columns-map="columnsMap"
          :current-page="currentPage"
        />
      </div>
    </section>
  </div>
</template>
