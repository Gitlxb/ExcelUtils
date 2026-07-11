<template>
  <FunctionCard
    title="供应商合并"
    description="上传ZIP包（内含多个.xlsx），合并指定工作表"
    :single-file="true"
    file-accept=".zip"
    :show-params="true"
    param-label="工作表名称"
    param-placeholder="Sheet1"
    api-endpoint="/api/finance/merge-supplier"
    param-name="sheetName"
  />
</template>
<script>
import FunctionCard from '@/components/FunctionCard.vue'
export default { name: 'CwGysQzy', components: { FunctionCard } }
</script>
