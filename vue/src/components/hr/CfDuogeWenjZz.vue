<template>
  <FunctionCard title="拆分两次_在职离职" :single-file="true" file-accept=".xlsx" :show-params="true" param-label="拆分列字母（多个用逗号分隔，如 A,B）" param-placeholder="A,B" api-endpoint="/api/hr/split-to-files" param-name="columns" />
</template>
<script>import FunctionCard from '@/components/FunctionCard.vue'; export default { name: 'CfDuogeWenjZz', components: { FunctionCard } }</script>
