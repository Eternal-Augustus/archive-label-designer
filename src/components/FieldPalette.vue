<script setup>
defineProps({
  columns: {
    type: Array,
    default: () => []
  }
});

const emit = defineEmits(["add-field-row", "add-static-row", "add-serial-row"]);

function handleDragStart(event, column) {
  event.dataTransfer.effectAllowed = "copy";
  event.dataTransfer.setData("application/x-archive-column", JSON.stringify(column));
}
</script>

<template>
  <section class="panel field-palette">
    <div class="section-heading">
      <h2>字段库</h2>
      <p>从 Excel 表头生成，点击或拖动即可插入为一整行单元。</p>
    </div>

    <div class="field-actions">
      <button class="ghost-button" type="button" @click="emit('add-static-row')">
        添加静态行
      </button>
      <button class="ghost-button" type="button" @click="emit('add-serial-row')">
        添加编号行
      </button>
    </div>

    <div v-if="columns.length" class="field-list">
      <button
        v-for="column in columns"
        :key="column.key"
        type="button"
        class="field-pill"
        draggable="true"
        @click="emit('add-field-row', column)"
        @dragstart="handleDragStart($event, column)"
      >
        <span>{{ column.label }}</span>
        <small>{{ column.key }}</small>
      </button>
    </div>

    <div v-else class="empty-note">
      先导入 Excel，字段库会在这里出现。
    </div>
  </section>
</template>
