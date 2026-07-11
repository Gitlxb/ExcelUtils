<template>
  <FunctionCard
    title="开关账筛选"
    description="根据系统日期自动计算范围（20号前3个月，20号后2个月）"
    :single-file="true"
    file-accept=".xlsx"
    :show-params="true"
    param-label="日期列名（表头文字）"
    param-placeholder="输入日期列的名称"
    api-endpoint="/api/finance/date-filter"
    param-name="dateColumn"
  />
</template>
<script>
import FunctionCard from '@/components/FunctionCard.vue'
export default { name: 'CwKgzSx', components: { FunctionCard } }
</script>
