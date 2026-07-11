<template>
  <FunctionCard title="PDF转Excel" description="提取PDF文本内容并转换为Excel" :single-file="true" file-accept=".pdf" :show-params="false" api-endpoint="/api/general/pdf-to-excel" />
</template>
<script>import FunctionCard from '@/components/FunctionCard.vue'; export default { name: 'ZhPdfExcel', components: { FunctionCard } }</script>
