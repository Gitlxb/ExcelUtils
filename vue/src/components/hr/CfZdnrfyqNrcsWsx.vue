<template>
  <FunctionCard title="按拆分参数配置拆分" description="Excel中需包含'拆分参数'工作表（A列公司名，B列分组编号）" :single-file="true" file-accept=".xlsx" :show-params="true" param-label="拆分列字母（如 E）" param-placeholder="E" api-endpoint="/api/hr/split-by-config" param-name="column" />
</template>
<script>import FunctionCard from '@/components/FunctionCard.vue'; export default { name: 'CfZdnrfyqNrcsWsx', components: { FunctionCard } }</script>
