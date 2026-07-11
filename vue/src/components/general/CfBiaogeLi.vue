<template>
  <FunctionCard title="拆成多个工作表" :single-file="true" file-accept=".xlsx" :show-params="true" param-label="拆分列字母（如 A）" param-placeholder="输入列字母" api-endpoint="/api/general/split-to-sheets" param-name="column" />
</template>
<script>import FunctionCard from '@/components/FunctionCard.vue'; export default { name: 'CfBiaogeLi', components: { FunctionCard } }</script>
