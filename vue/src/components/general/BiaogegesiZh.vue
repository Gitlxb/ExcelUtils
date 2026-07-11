<template>
  <FunctionCard title="表格格式转换" description="上传ZIP包，批量将.xls/.csv等转为.xlsx" :single-file="true" file-accept=".zip" :show-params="false" api-endpoint="/api/general/convert-format" />
</template>
<script>import FunctionCard from '@/components/FunctionCard.vue'; export default { name: 'BiaogegesiZh', components: { FunctionCard } }</script>
