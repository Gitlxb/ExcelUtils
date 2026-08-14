<template>
  <div class="gzbcf-wrap">
    <div class="gzbcf-toggle">
      <span class="gzbcf-label">工资表拆分模式：</span>
      <el-radio-group v-model="mode">
        <el-radio-button value="enhanced">增强版</el-radio-button>
        <el-radio-button value="simple">简化版</el-radio-button>
      </el-radio-group>
    </div>
    <!-- 复用已有两套实现，仅在此统一切换 -->
    <CfGzbcfHyp v-if="mode === 'enhanced'" />
    <CfGzbcfHypSimple v-else />
  </div>
</template>

<script>
import CfGzbcfHyp from './CfGzbcfHyp.vue'
import CfGzbcfHypSimple from './CfGzbcfHypSimple.vue'

export default {
  name: 'HrGzbcfSplit',
  components: { CfGzbcfHyp, CfGzbcfHypSimple },
  data() {
    return {
      mode: 'enhanced'
    }
  }
}
</script>

<style scoped>
.gzbcf-wrap { padding: 0; }
.gzbcf-toggle {
  margin-bottom: 12px;
  display: flex;
  align-items: center;
}
.gzbcf-label { font-size: 14px; color: #606266; margin-right: 8px; }
</style>
