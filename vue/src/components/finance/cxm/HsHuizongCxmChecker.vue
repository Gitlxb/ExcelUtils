<template>
  <el-card class="func-card">
    <div class="page-container">
    <div class="cxm-page-header">
      <h2 class="cxm-page-title">核对</h2>
      <p class="cxm-page-desc">
        上传发放记录表和供应商对账单，自动汇总并核对数据。支持月份/类型筛选、多工作表处理、核对结果预览与下载。
      </p>
    </div>

    <!-- 区域A：发放记录表 -->
    <el-divider content-position="left">区域A：发放记录表</el-divider>
    <el-row :gutter="16" style="margin-bottom: 12px">
      <el-col :span="12">
        <div class="cxm-upload-label">发放记录表</div>
        <el-upload
          ref="uploadDisbursement"
          class="cxm-upload"
          drag
          :auto-upload="false"
          :limit="1"
          :on-change="handleDisbursementFileChange"
          accept=".xlsx,.xls"
        >
          <el-icon class="cxm-upload-icon"><UploadFilled /></el-icon>
          <div>将文件拖到此处，或 <em>点击选择</em></div>
          <div v-if="disbursementFile" class="cxm-file-name">{{ disbursementFile.name }}</div>
        </el-upload>
      </el-col>
      <el-col :span="12">
        <div class="cxm-upload-label">工作表（可多选）</div>
        <div v-if="disbursementSheets.length" class="cxm-sheet-checkbox-group">
          <el-checkbox
            v-for="s in disbursementSheets"
            :key="s"
            v-model="selectedDisbursementSheets"
            :label="s"
            class="cxm-sheet-checkbox"
            @change="onSheetSelectionChange"
          >{{ s }}</el-checkbox>
        </div>
        <div v-else class="cxm-sheet-empty">请先选择文件</div>

        <!-- 参数选择区域 -->
        <div v-if="selectedDisbursementSheets.length && !detectingParams" style="margin-top: 12px; border: 1px solid #e4e7ed; border-radius: 6px; padding: 12px; background: #fafbfc;">
          <div style="font-weight: 600; margin-bottom: 10px; color: #409eff;">参数设置</div>

          <!-- 打款金额列选择 -->
          <div style="margin-bottom: 10px">
            <div class="cxm-upload-label">打款金额列</div>
            <el-radio-group v-model="colH" size="small">
              <el-radio-button :label="2">打款金额（C列）</el-radio-button>
              <el-radio-button :label="7">打款金额（H列）</el-radio-button>
            </el-radio-group>
          </div>

          <!-- 月份多选 -->
          <div v-if="detectedParams.months.length" style="margin-bottom: 10px">
            <div class="cxm-upload-label">月份（可多选）</div>
            <el-checkbox-group v-model="selectedMonths" size="small">
              <el-checkbox-button
                v-for="m in detectedParams.months"
                :key="m"
                :label="m"
              >{{ m }}</el-checkbox-button>
            </el-checkbox-group>
          </div>

          <!-- 打款类型多选 -->
          <div style="margin-bottom: 0">
            <div class="cxm-upload-label">打款类型（可多选）</div>
            <el-checkbox-group v-model="selectedTypes" size="small">
              <el-checkbox-button label="对公打款">对公打款</el-checkbox-button>
              <el-checkbox-button label="代发打款">代发打款</el-checkbox-button>
              <el-checkbox-button label="刘先锋现金卡支付">刘先锋现金卡支付</el-checkbox-button>
            </el-checkbox-group>
          </div>

          <div v-if="!detectedParams.months.length" style="color: #909399; font-size: 12px; margin-top: 6px;">
            未检测到月份筛选条件，将处理全部数据
          </div>
        </div>

        <div v-else-if="selectedDisbursementSheets.length && detectingParams" style="margin-top: 12px; text-align: center; padding: 16px; color: #409eff;">
          <span>正在检测参数...</span>
        </div>

        <div style="margin-top: 12px">
          <el-button
            type="primary"
            :loading="disbursementProcessing"
            :disabled="!canProcessDisbursement"
            @click="processDisbursement"
          >
            {{ disbursementProcessing ? '汇总中...' : '汇总发放记录' }}
          </el-button>
        </div>

        <div v-if="disbursementFileName" style="margin-top: 12px">
          <el-alert type="success" :closable="false" show-icon>
            <template #default>
              发放记录汇总表：{{ disbursementFileName }}
              <el-button type="primary" size="small" style="margin-left: 12px" @click="downloadFile(disbursementFileName)">下载</el-button>
            </template>
          </el-alert>
        </div>
      </el-col>
    </el-row>

    <!-- 区域B：供应商对账单 -->
    <el-divider content-position="left">区域B：供应商对账单</el-divider>
    <el-row :gutter="16" style="margin-bottom: 12px">
      <el-col :span="12">
        <div class="cxm-upload-label">供应商对账单</div>
        <el-upload
          ref="uploadStatement"
          class="cxm-upload"
          drag
          :auto-upload="false"
          :limit="1"
          :on-change="handleStatementFileChange"
          accept=".xlsx,.xls"
        >
          <el-icon class="cxm-upload-icon"><UploadFilled /></el-icon>
          <div>将文件拖到此处，或 <em>点击选择</em></div>
          <div v-if="statementFile" class="cxm-file-name">{{ statementFile.name }}</div>
        </el-upload>
      </el-col>
      <el-col :span="12">
        <div class="cxm-upload-label">工作表（单选）</div>
        <el-select v-model="selectedStatementSheet" placeholder="请选择工作表" style="width: 100%" :disabled="!statementSheets.length">
          <el-option v-for="s in statementSheets" :key="s" :label="s" :value="s" />
        </el-select>

        <div style="margin-top: 8px">
          <el-button
            type="primary"
            :loading="statementProcessing"
            :disabled="!statementFile || !selectedStatementSheet"
            @click="processStatement"
          >
            {{ statementProcessing ? '汇总中...' : '汇总对账单' }}
          </el-button>
        </div>

        <div v-if="statementFileName" style="margin-top: 12px">
          <el-alert type="success" :closable="false" show-icon>
            <template #default>
              对账单汇总表：{{ statementFileName }}
              <el-button type="primary" size="small" style="margin-left: 12px" @click="downloadFile(statementFileName)">下载</el-button>
            </template>
          </el-alert>
        </div>
      </el-col>
    </el-row>

    <!-- 核对操作 -->
    <el-divider content-position="left">核对操作</el-divider>
    <div style="margin-bottom: 16px; text-align: center">
      <el-button
        type="success"
        size="large"
        :loading="reconcileProcessing"
        :disabled="!disbursementFileName || !statementFileName"
        @click="startReconcile"
      >
        {{ reconcileProcessing ? '核对中...' : '开始核对' }}
      </el-button>
    </div>

    <div v-if="reconcileFileName" style="margin-bottom: 16px">
      <el-alert type="success" :closable="false" show-icon>
        <template #default>
          核对结果：{{ reconcileFileName }} （含3个Sheet：发放记录汇总 / 对账单汇总 / 核对结果）
          <el-button type="primary" size="small" style="margin-left: 12px" @click="downloadFile(reconcileFileName)">下载核对结果</el-button>
        </template>
      </el-alert>
    </div>

    <!-- 数据预览 -->
    <el-divider content-position="left">数据预览</el-divider>
    <el-table
      v-if="previewData.length"
      :data="previewData"
      border
      size="small"
      max-height="400px"
      style="width: 100%; margin-bottom: 16px"
    >
      <el-table-column v-for="col in previewColumns" :key="col" :prop="col" :label="col" min-width="100" align="center">
        <template #default="{ row }">
          <span :class="{ 'cxm-matched': col === '状态' && row[col] === '已核对', 'cxm-pending': col === '状态' && row[col] === '待核对' }">{{ row[col] }}</span>
        </template>
      </el-table-column>
    </el-table>
    <el-empty v-else description="暂无预览数据" style="margin: 20px 0" />

    <!-- 状态日志 -->
    <el-divider content-position="left">状态信息</el-divider>
    <div class="cxm-log-box">
      <div v-for="(log, idx) in logs" :key="idx" class="cxm-log-line">{{ log }}</div>
      <div v-if="!logs.length" class="cxm-log-line cxm-log-empty">欢迎使用核对功能！请上传发放记录表和对账单开始处理。</div>
    </div>
    </div>
  </el-card>
</template>

<script>
import axios from 'axios'

export default {
  name: 'HsHuizongCxmChecker',
  data() {
    return {
      // 区域A
      disbursementFile: null,
      disbursementSheets: [],
      selectedDisbursementSheets: [],
      colH: '',
      selectedMonths: [],
      selectedTypes: [],
      disbursementProcessing: false,
      disbursementFileName: '',
      detectingParams: false,
      detectedParams: {
        columns: [],
        months: [],
        types: []
      },
      // 区域B
      statementFile: null,
      statementSheets: [],
      selectedStatementSheet: '',
      statementProcessing: false,
      statementFileName: '',
      // 核对
      reconcileProcessing: false,
      reconcileFileName: '',
      // 预览
      previewData: [],
      previewColumns: [],
      // 日志
      logs: []
    }
  },
  computed: {
    canProcessDisbursement() {
      return this.disbursementFile
             && this.selectedDisbursementSheets.length
             && (this.colH === 2 || this.colH === 7)
    }
  },
  watch: {
    disbursementFile(val) {
      if (val) this.loadDisbursementSheets(val)
      else {
        this.disbursementSheets = []
        this.selectedDisbursementSheets = []
        this.resetDetectedParams()
      }
    },
    statementFile(val) {
      if (val) this.loadStatementSheets(val)
      else { this.statementSheets = []; this.selectedStatementSheet = '' }
    }
  },
  methods: {
    addLog(msg) {
      const t = new Date().toLocaleTimeString('zh-CN', { hour: '2-digit', minute: '2-digit', second: '2-digit' })
      this.logs.push(`[${t}] ${msg}`)
    },
    resetDetectedParams() {
      this.detectedParams = { months: [], types: [] }
      this.colH = ''
      this.selectedMonths = []
      this.selectedTypes = []
    },
    handleDisbursementFileChange(file) {
      this.disbursementFile = file.raw
    },
    async loadDisbursementSheets(file) {
      const formData = new FormData()
      formData.append('file', file)
      try {
        const res = await axios.post('/api/finance/summary/sheets', formData)
        if (res.data.success) {
          this.disbursementSheets = res.data.sheets || []
          this.selectedDisbursementSheets = []
          this.resetDetectedParams()
          this.addLog(`读取发放记录表工作表成功，共 ${this.disbursementSheets.length} 个`)
        } else {
          this.$message.warning(res.data.message || '读取工作表失败')
        }
      } catch (e) {
        this.$message.error('读取工作表失败：' + e.message)
      }
    },
    onSheetSelectionChange() {
      if (this.selectedDisbursementSheets.length > 0) {
        this.detectParams()
      } else {
        this.resetDetectedParams()
      }
    },
    async detectParams() {
      if (!this.disbursementFile || !this.selectedDisbursementSheets.length) return
      this.detectingParams = true
      this.resetDetectedParams()
      const formData = new FormData()
      formData.append('file', this.disbursementFile)
      this.selectedDisbursementSheets.forEach(s => {
        formData.append('sheetNames', s)
      })
      try {
        const res = await axios.post('/api/finance/check/detect-params', formData)
        if (res.data.success) {
          this.detectedParams.months = res.data.months || []
          this.detectedParams.types = res.data.types || []
          this.selectedMonths = []
          this.selectedTypes = []
          this.addLog(`参数检测完成（${this.selectedDisbursementSheets.length}个工作表）：${this.detectedParams.months.length} 个月份`)
        } else {
          this.$message.warning(res.data.message || '参数检测失败')
        }
      } catch (e) {
        this.$message.error('参数检测失败：' + e.message)
        this.addLog('参数检测失败：' + e.message)
      } finally {
        this.detectingParams = false
      }
    },
    async processDisbursement() {
      if (!this.canProcessDisbursement) {
        this.$message.warning('请完成参数设置')
        return
      }
      this.disbursementProcessing = true
      this.addLog('开始汇总发放记录表...')
      const formData = new FormData()
      formData.append('file', this.disbursementFile)
      formData.append('sheetName', this.selectedDisbursementSheets[0])
      formData.append('colH', this.colH)
      if (this.selectedMonths.length) {
        formData.append('selectedMonths', this.selectedMonths.join(','))
      }
      if (this.selectedTypes.length) {
        formData.append('selectedTypes', this.selectedTypes.join(','))
      }
      try {
        const res = await axios.post('/api/finance/check/disbursement', formData)
        if (res.data.success) {
          this.disbursementFileName = res.data.fileName
          this.previewColumns = res.data.previewColumns || []
          this.previewData = res.data.previewData || []
          this.addLog('发放记录表汇总完成：' + res.data.fileName)
          this.$message.success('发放记录表汇总完成')
        } else {
          this.$message.error(res.data.message || '汇总失败')
        }
      } catch (e) {
        this.$message.error('汇总失败：' + e.message)
        this.addLog('汇总失败：' + e.message)
      } finally {
        this.disbursementProcessing = false
      }
    },
    handleStatementFileChange(file) {
      this.statementFile = file.raw
    },
    async loadStatementSheets(file) {
      const formData = new FormData()
      formData.append('file', file)
      try {
        const res = await axios.post('/api/finance/summary/sheets', formData)
        if (res.data.success) {
          this.statementSheets = res.data.sheets || []
          this.selectedStatementSheet = ''
          this.addLog(`读取对账单工作表成功，共 ${this.statementSheets.length} 个`)
        } else {
          this.$message.warning(res.data.message || '读取工作表失败')
        }
      } catch (e) {
        this.$message.error('读取工作表失败：' + e.message)
      }
    },
    async processStatement() {
      if (!this.statementFile || !this.selectedStatementSheet) {
        this.$message.warning('请选择文件和工作表')
        return
      }
      this.statementProcessing = true
      this.addLog('开始汇总对账单...')
      const formData = new FormData()
      formData.append('file', this.statementFile)
      formData.append('sheetName', this.selectedStatementSheet)
      try {
        const res = await axios.post('/api/finance/check/statement', formData)
        if (res.data.success) {
          this.statementFileName = res.data.fileName
          this.previewColumns = res.data.previewColumns || []
          this.previewData = res.data.previewData || []
          this.addLog('对账单汇总完成：' + res.data.fileName)
          this.$message.success('对账单汇总完成')
        } else {
          this.$message.error(res.data.message || '汇总失败')
        }
      } catch (e) {
        this.$message.error('汇总失败：' + e.message)
        this.addLog('汇总失败：' + e.message)
      } finally {
        this.statementProcessing = false
      }
    },
    async startReconcile() {
      if (!this.disbursementFileName || !this.statementFileName) {
        this.$message.warning('请先完成发放记录汇总和对账单汇总')
        return
      }
      this.reconcileProcessing = true
      this.addLog('开始核对...')
      const formData = new FormData()
      formData.append('disbursementFileName', this.disbursementFileName)
      formData.append('statementFileName', this.statementFileName)
      try {
        const res = await axios.post('/api/finance/check/reconcile', formData)
        if (res.data.success) {
          this.reconcileFileName = res.data.fileName
          this.previewColumns = res.data.previewColumns || []
          this.previewData = res.data.previewData || []
          this.addLog('核对完成：' + res.data.fileName)
          this.$message.success('核对完成')
        } else {
          this.$message.error(res.data.message || '核对失败')
        }
      } catch (e) {
        this.$message.error('核对失败：' + e.message)
        this.addLog('核对失败：' + e.message)
      } finally {
        this.reconcileProcessing = false
      }
    },
    downloadFile(fileName) {
      window.open('/api/finance/download/' + fileName, '_blank')
    }
  }
}
</script>

<style scoped>
.page-container {
  width: 100%;
  padding: 10px;
}

.cxm-page-header {
  padding: 16px 0 8px;
}

.cxm-page-title {
  font-size: 18px;
  font-weight: 700;
  color: #303133;
  margin: 0 0 8px 0;
}

.cxm-page-desc {
  font-size: 13px;
  color: #606266;
  line-height: 1.7;
  margin: 0;
}

.cxm-upload {
  width: 100%;
  margin-bottom: 12px;
}

.cxm-upload-label {
  font-weight: 600;
  margin-bottom: 6px;
  color: #606266;
}

.cxm-file-name {
  margin-top: 6px;
  font-size: 12px;
  color: #409eff;
}

.cxm-sheet-checkbox-group {
  max-height: 200px;
  overflow-y: auto;
  border: 1px solid #dcdfe6;
  border-radius: 4px;
  padding: 8px;
}

.cxm-sheet-checkbox {
  display: block;
  margin-left: 0 !important;
  margin-bottom: 4px;
}

.cxm-sheet-empty {
  color: #909399;
  font-size: 14px;
  text-align: center;
  padding: 20px 0;
}

.cxm-log-box {
  background: #f5f7fa;
  border: 1px solid #e4e7ed;
  border-radius: 4px;
  padding: 10px;
  max-height: 200px;
  overflow-y: auto;
  font-family: 'Consolas', 'Monaco', monospace;
  font-size: 13px;
}

.cxm-log-line {
  padding: 2px 0;
  color: #606266;
}

.cxm-log-empty {
  color: #909399;
  font-style: italic;
}

.cxm-matched {
  color: #67c23a;
  font-weight: 600;
}

.cxm-pending {
  color: #e6a23c;
  font-weight: 600;
}
</style>
