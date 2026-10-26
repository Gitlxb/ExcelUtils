<template>
  <el-card class="func-card">
    <div class="page-container">
    <div class="cxm-page-header">
      <h2 class="cxm-page-title">拆分汇总比较</h2>
      <p class="cxm-page-desc">
        提供代工费发放记录表的数据处理、核对与追加功能。支持多工作表批量处理、月份/类型筛选、核对月份/介绍人条件配置。
      </p>
    </div>

    <!-- 步骤1：上传发放记录表 + 数据处理 -->
    <el-divider content-position="left">步骤1：数据处理（代工费发放记录表）</el-divider>
    <el-row :gutter="16" style="margin-bottom:12px">
      <el-col :span="12">
        <div class="cxm-upload-label">代工费发放记录表</div>
        <el-upload
          ref="uploadRecord"
          class="cxm-upload"
          drag
          :auto-upload="false"
          :limit="1"
          :on-change="handleRecordFileChange"
          accept=".xlsx,.xls"
        >
          <el-icon class="cxm-upload-icon"><UploadFilled /></el-icon>
          <div>将文件拖到此处，或 <em>点击选择</em></div>
          <div v-if="recordFile" class="cxm-file-name">{{ recordFile.name }}</div>
        </el-upload>
      </el-col>
      <el-col :span="12">
        <div class="cxm-upload-label">工作表（可多选）</div>
        <!-- 多选工作表 checkbox 组 -->
        <div v-if="recordSheets.length" class="cxm-sheet-checkbox-group">
          <el-checkbox
            v-for="s in recordSheets"
            :key="s"
            v-model="selectedSheets"
            :label="s"
            class="cxm-sheet-checkbox"
          >{{ s }}</el-checkbox>
        </div>
        <div v-else class="cxm-sheet-empty">请先选择文件</div>
        <div style="margin-top:8px">
          <el-button type="success" :loading="detecting || processing" :disabled="!recordFile || !selectedSheets.length" @click="startDetect">
            {{ detecting ? '检测中...' : (processing ? '处理中...' : '检测参数') }}
          </el-button>
          <el-button style="margin-left:8px" @click="refreshPage">
            <el-icon><Refresh /></el-icon> 刷新
          </el-button>
        </div>
        <div v-if="summaryFileName" style="margin-top:12px">
          <el-alert type="success" :closable="false" show-icon>
            <template #default>
              供应商核算汇总：{{ summaryFileName }}
              <el-button type="primary" size="small" style="margin-left:12px" @click="downloadFile(summaryFileName)">下载</el-button>
            </template>
          </el-alert>
        </div>
      </el-col>
    </el-row>

    <!-- 参数选择弹窗（类型/月份/金额列）-->
    <el-dialog v-model="showParamDialog" title="选择处理参数" width="520px" :close-on-click-modal="false">
      <el-form label-width="100px">
        <el-form-item v-if="detectedTypes.length" label="打款类型">
          <el-checkbox-group v-model="selectedTypes">
            <el-checkbox v-for="t in detectedTypes" :key="t" :label="t">{{ t }}</el-checkbox>
          </el-checkbox-group>
        </el-form-item>
        <el-form-item v-if="detectedMonths.length" label="月份">
          <el-checkbox-group v-model="selectedMonths">
            <el-checkbox v-for="m in detectedMonths" :key="m" :label="m">{{ m }}</el-checkbox>
          </el-checkbox-group>
        </el-form-item>
        <el-form-item v-if="detectedAmountCols.length >= 2" label="打款金额列">
          <el-radio-group v-model="selectedAmountColIdx">
            <el-radio v-for="col in detectedAmountCols" :key="col.index" :label="col.index">
              打款金额（{{ col.colLetter }}列）
            </el-radio>
          </el-radio-group>
        </el-form-item>
      </el-form>
      <template #footer>
        <el-button @click="showParamDialog = false">取消</el-button>
        <el-button type="primary" :loading="processing" :disabled="!canConfirmParam" @click="confirmParamDialog">
          {{ processing ? '处理中...' : '确定处理' }}
        </el-button>
      </template>
    </el-dialog>

    <!-- 步骤2：上传汇总表 + 核对 -->
    <el-divider content-position="left">步骤2：数据核对（代工费汇总表）</el-divider>
    <el-row :gutter="16" style="margin-bottom:12px">
      <el-col :span="12">
        <div class="cxm-upload-label">代工费汇总表</div>
        <el-upload
          ref="uploadCheck"
          class="cxm-upload"
          drag
          :auto-upload="false"
          :limit="1"
          :on-change="handleCheckFileChange"
          accept=".xlsx,.xls"
        >
          <el-icon class="cxm-upload-icon"><UploadFilled /></el-icon>
          <div>将文件拖到此处，或 <em>点击选择</em></div>
          <div v-if="checkFile" class="cxm-file-name">{{ checkFile.name }}</div>
        </el-upload>
      </el-col>
      <el-col :span="12">
        <div class="cxm-upload-label">工作表</div>
        <el-select v-model="checkSheet" placeholder="请选择工作表" style="width:100%" :disabled="!checkSheets.length">
          <el-option v-for="s in checkSheets" :key="s" :label="s" :value="s" />
        </el-select>
        <div style="margin-top:8px">
          <el-button type="primary" :loading="checkPreparing" :disabled="!checkFile || !checkSheet" @click="prepareCheck">
            {{ checkPreparing ? '检测中...' : '检测月份/介绍人' }}
          </el-button>
        </div>
        <div v-if="checkFileName" style="margin-top:12px">
          <el-alert type="success" :closable="false" show-icon>
            <template #default>
              核对结果：{{ checkFileName }}
              <el-button type="primary" size="small" style="margin-left:12px" @click="downloadFile(checkFileName)">下载</el-button>
            </template>
          </el-alert>
        </div>
      </el-col>
    </el-row>

    <!-- 核对条件选择弹窗（月份单选 + 介绍人多选）-->
    <el-dialog v-model="showCheckDialog" title="核对条件" width="900px" :close-on-click-modal="false">
      <el-row :gutter="16">
        <!-- 左侧：月份单选 -->
        <el-col :span="8">
          <div class="cxm-dialog-section-title">月份（单选）</div>
          <div class="cxm-dialog-content">
            <el-radio-group v-model="checkSelectedMonth" class="cxm-month-group">
              <el-radio v-for="m in checkMonths" :key="m" :label="m">{{ m }}</el-radio>
            </el-radio-group>
          </div>
        </el-col>
        <!-- 右侧：介绍人多选（按月份联动） -->
        <el-col :span="16">
          <div class="cxm-dialog-section-title">
            介绍人（多选，{{ checkSelectedMonth }}）
            <el-input v-model="checkPersonSearch" placeholder="搜索介绍人" size="small" style="width:160px;margin-left:8px" />
            <el-button size="small" style="margin-left:6px" @click="checkSelectAll">全选</el-button>
            <el-button size="small" @click="checkDeselectAll">全不选</el-button>
          </div>
          <div class="cxm-dialog-content">
            <div class="cxm-person-grid">
              <el-checkbox v-for="p in filteredCheckPersons" :key="p" v-model="checkSelectedPersons" :label="p" class="cxm-person-item">
                {{ p }}
              </el-checkbox>
            </div>
          </div>
          <div class="cxm-person-count">显示 {{ filteredCheckPersons.length }} / {{ currentMonthPersons.length }} 个</div>
        </el-col>
      </el-row>
      <template #footer>
        <el-button @click="showCheckDialog = false">取消</el-button>
        <el-button type="primary" :loading="checkProcessing" :disabled="!checkSelectedMonth || !checkSelectedPersons.length" @click="executeCheck">
          {{ checkProcessing ? '核对中...' : '确定核对' }}
        </el-button>
      </template>
    </el-dialog>

    <!-- 步骤3：追加到目标表 -->
    <el-divider content-position="left">步骤3：追加到目标表</el-divider>
    <el-row :gutter="16" style="margin-bottom:16px">
      <el-col :span="12">
        <div class="cxm-upload-label">追加到目标表</div>
        <el-upload
          ref="uploadTarget"
          class="cxm-upload"
          drag
          :auto-upload="false"
          :limit="1"
          :on-change="handleTargetFileChange"
          accept=".xlsx"
        >
          <el-icon class="cxm-upload-icon"><UploadFilled /></el-icon>
          <div>将文件拖到此处，或 <em>点击选择</em></div>
          <div v-if="targetFile" class="cxm-file-name">{{ targetFile.name }}</div>
        </el-upload>
      </el-col>
      <el-col :span="12">
        <div style="display:flex;align-items:flex-end">
          <el-button type="warning" :loading="appending" :disabled="!canAppend" @click="appendToTargetFn">
            {{ appending ? '追加中...' : '追加到目标表' }}
          </el-button>
          <el-button style="margin-left:8px" @click="resetAll">重置</el-button>
        </div>
        <div v-if="appendFileName" style="margin-top:12px">
          <el-alert type="success" :closable="false" show-icon>
            <template #default>
              追加结果：{{ appendFileName }}
              <el-button type="primary" size="small" style="margin-left:12px" @click="downloadFile(appendFileName)">下载</el-button>
            </template>
          </el-alert>
        </div>
      </el-col>
    </el-row>

    <!-- 数据预览 -->
    <el-divider content-position="left">数据预览</el-divider>
    <el-table
      v-if="previewData.length"
      :data="previewData"
      border
      size="small"
      max-height="400px"
      style="width:100%;margin-bottom:16px"
    >
      <el-table-column v-for="col in previewColumns" :key="col" :prop="col" :label="col" min-width="100" align="center">
        <template #default="{ row }">
          <span :class="{ 'cxm-pending': col === '状态' && row[col] === '待核对' }">{{ row[col] }}</span>
        </template>
      </el-table-column>
    </el-table>
    <el-empty v-else description="暂无预览数据" style="margin:20px 0" />

    <!-- 状态日志 -->
    <el-divider content-position="left">状态信息</el-divider>
    <div class="cxm-log-box">
      <div v-for="(log, idx) in logs" :key="idx" class="cxm-log-line">{{ log }}</div>
      <div v-if="!logs.length" class="cxm-log-line cxm-log-empty">欢迎使用核算汇总工具！请选择代工费发放记录表开始处理。</div>
    </div>
    </div>
  </el-card>
</template>

<script>
import axios from 'axios'
import { postAsync } from '../../../utils/asyncTask'

export default {
  name: 'HsHuizongCxm',
  data() {
    return {
      // 步骤1
      recordFile: null,
      recordSheets: [],
      selectedSheets: [],
      detecting: false,
      processing: false,
      detectedTypes: [],
      detectedMonths: [],
      detectedAmountCols: [],
      showParamDialog: false,
      selectedTypes: [],
      selectedMonths: [],
      selectedAmountColIdx: -1,
      // 步骤2
      checkFile: null,
      checkSheet: '',
      checkSheets: [],
      checkPreparing: false,
      checkProcessing: false,
      checkMonths: [],
      checkMonthPersonMap: {},
      checkPersonSearch: '',
      showCheckDialog: false,
      checkSelectedMonth: '',
      checkSelectedPersons: [],
      // 步骤3
      targetFile: null,
      appending: false,
      // 结果
      summaryFileName: '',
      checkFileName: '',
      appendFileName: '',
      downloadFileName: '',
      previewData: [],
      previewColumns: [],
      // 日志
      logs: []
    }
  },
  computed: {
    canConfirmParam() {
      if (this.detectedTypes.length && !this.selectedTypes.length) return false
      if (this.detectedMonths.length && !this.selectedMonths.length) return false
      if (this.detectedAmountCols.length >= 2 && this.selectedAmountColIdx < 0) return false
      return true
    },
    canAppend() {
      return (this.summaryFileName || this.checkFileName) && this.targetFile
    },
    currentMonthPersons() {
      if (!this.checkSelectedMonth) return []
      return this.checkMonthPersonMap[this.checkSelectedMonth] || []
    },
    filteredCheckPersons() {
      const persons = this.currentMonthPersons
      if (!this.checkPersonSearch) return persons
      const q = this.checkPersonSearch.toLowerCase()
      return persons.filter(p => p.toLowerCase().includes(q))
    }
  },
  watch: {
    recordFile(val) {
      if (val) this.loadRecordSheets(val)
      else { this.recordSheets = []; this.selectedSheets = [] }
    },
    checkFile(val) {
      if (val) this.loadCheckSheets(val)
      else { this.checkSheets = []; this.checkSheet = '' }
    },
    checkSelectedMonth() {
      this.checkSelectedPersons = []
    }
  },
  methods: {
    addLog(msg) {
      const t = new Date().toLocaleTimeString('zh-CN', { hour: '2-digit', minute: '2-digit', second: '2-digit' })
      this.logs.push(`[${t}] ${msg}`)
    },
    refreshPage() {
      this.detecting = false
      this.processing = false
      this.checkPreparing = false
      this.checkProcessing = false
      this.appending = false
      this.showParamDialog = false
      this.showCheckDialog = false
      this.addLog('页面已刷新（已停止所有加载状态，操作数据已保留）')
    },
    handleRecordFileChange(file) {
      this.recordFile = file.raw
    },
    async loadRecordSheets(file) {
      const formData = new FormData()
      formData.append('file', file)
      try {
        const res = await axios.post('/api/finance/summary/sheets', formData)
        if (res.data.success) {
          this.recordSheets = res.data.sheets || []
          this.selectedSheets = []
          this.addLog(`读取工作表成功，共 ${this.recordSheets.length} 个`)
        } else {
          this.$message.warning(res.data.message || '读取工作表失败')
        }
      } catch (err) {
        this.$message.error(err.response?.data?.message || err.message)
      }
    },
    async startDetect() {
      if (!this.recordFile || !this.selectedSheets.length) {
        this.$message.warning('请先选择文件和工作表')
        return
      }
      this.detecting = true
      this.addLog(`开始检测参数，共 ${this.selectedSheets.length} 个工作表...`)
      try {
        const formData = new FormData()
        formData.append('file', this.recordFile)
        formData.append('sheetNames', this.selectedSheets.join(','))
        const res = await axios.post('/api/finance/summary/batch-detect', formData)
        if (!res.data.success) {
          this.$message.error(res.data.message || '检测失败')
          this.detecting = false
          return
        }
        this.detectedTypes = res.data.types || []
        this.detectedMonths = res.data.months || []
        this.detectedAmountCols = res.data.amountColumns || []
        this.addLog(`检测完成：类型=${this.detectedTypes.join(',') || '无'}，月份=${this.detectedMonths.join(',') || '无'}，金额列=${this.detectedAmountCols.length}个`)
        this.selectedTypes = []
        this.selectedMonths = []
        this.selectedAmountColIdx = this.detectedAmountCols.length >= 2 ? this.detectedAmountCols[0].index : -1
        this.showParamDialog = true
      } catch (err) {
        this.$message.error(err.response?.data?.message || err.message)
        this.addLog(`检测错误：${err.message}`)
      } finally {
        this.detecting = false
      }
    },
    async confirmParamDialog() {
      this.showParamDialog = false
      this.processing = true
      this.addLog(`开始批量处理 ${this.selectedSheets.length} 个工作表...`)
      try {
        const formData = new FormData()
        formData.append('file', this.recordFile)
        formData.append('sheetNames', this.selectedSheets.join(','))
        formData.append('types', this.selectedTypes.join(','))
        formData.append('months', this.selectedMonths.join(','))
        formData.append('amountCol', this.selectedAmountColIdx)
        const res = await postAsync('/api/finance/summary/batch', formData)
        if (res.data.success) {
          this.summaryFileName = res.data.fileName
          this.downloadFileName = res.data.fileName
          this.addLog(`批量处理完成：${res.data.message}`)
          if (res.data.previewColumns && res.data.previewColumns.length) {
            this.previewColumns = res.data.previewColumns
            this.previewData = res.data.previewData || []
          }
          this.$message.success('数据处理完成！')
        } else {
          this.$message.error(res.data.message || '处理失败')
          this.addLog(`错误：${res.data.message}`)
        }
      } catch (err) {
        this.$message.error(err.response?.data?.message || err.message)
        this.addLog(`处理错误：${err.message}`)
      } finally {
        this.processing = false
      }
    },
    handleCheckFileChange(file) {
      this.checkFile = file.raw
    },
    async loadCheckSheets(file) {
      const formData = new FormData()
      formData.append('file', file)
      try {
        const res = await axios.post('/api/finance/summary/sheets', formData)
        if (res.data.success) {
          this.checkSheets = res.data.sheets || []
          if (this.checkSheets.length) this.checkSheet = this.checkSheets[0]
        }
      } catch (err) {
        this.$message.error(err.response?.data?.message || err.message)
      }
    },
    async prepareCheck() {
      if (!this.checkFile || !this.checkSheet) {
        this.$message.warning('请先选择文件和工作表')
        return
      }
      this.checkPreparing = true
      this.addLog('开始检测月份和介绍人...')
      try {
        const formData = new FormData()
        formData.append('checkFile', this.checkFile)
        formData.append('checkSheet', this.checkSheet)
        const res = await axios.post('/api/finance/summary/check/prepare', formData)
        if (res.data.success) {
          this.checkMonths = res.data.months || []
          this.checkMonthPersonMap = res.data.monthPersonMap || {}
          if (this.checkMonths.length) this.checkSelectedMonth = this.checkMonths[0]
          this.checkSelectedPersons = []
          this.checkPersonSearch = ''
          this.showCheckDialog = true
          const totalPersons = Object.values(this.checkMonthPersonMap).flat().length
          this.addLog(`检测完成：月份=${this.checkMonths.join(',')}，介绍人共${totalPersons}个`)
        } else {
          this.$message.error(res.data.message || '检测失败')
        }
      } catch (err) {
        this.$message.error(err.response?.data?.message || err.message)
      } finally {
        this.checkPreparing = false
      }
    },
    checkSelectAll() {
      this.checkSelectedPersons = [...this.filteredCheckPersons]
    },
    checkDeselectAll() {
      this.checkSelectedPersons = []
    },
    async executeCheck() {
      if (!this.checkSelectedMonth || !this.checkSelectedPersons.length) {
        this.$message.warning('请选择月份和介绍人')
        return
      }
      this.checkProcessing = true
      this.addLog(`执行核对：月份=${this.checkSelectedMonth}，介绍人=${this.checkSelectedPersons.length}个...`)
      try {
        const formData = new FormData()
        if (this.summaryFileName) {
          formData.append('refFileName', this.summaryFileName)
        }
        formData.append('checkFile', this.checkFile)
        formData.append('checkSheet', this.checkSheet)
        formData.append('selectedMonth', this.checkSelectedMonth)
        formData.append('selectedPersons', this.checkSelectedPersons.join(','))
        const res = await postAsync('/api/finance/summary/check', formData)
        if (res.data.success) {
          this.checkFileName = res.data.fileName
          this.downloadFileName = res.data.fileName
          this.addLog(`核对完成：${res.data.message}`)
          if (res.data.previewColumns && res.data.previewColumns.length) {
            this.previewColumns = res.data.previewColumns
            this.previewData = res.data.previewData || []
          }
          this.$message.success('核对完成！')
          this.showCheckDialog = false
        } else {
          this.$message.error(res.data.message || '核对失败')
        }
      } catch (err) {
        this.$message.error(err.response?.data?.message || err.message)
        this.addLog(`核对错误：${err.message}`)
      } finally {
        this.checkProcessing = false
      }
    },
    handleTargetFileChange(file) {
      this.targetFile = file.raw
    },
    async appendToTargetFn() {
      this.appending = true
      this.addLog('开始追加到目标表...')
      try {
        const formData = new FormData()
        if (this.summaryFileName) formData.append('summaryFileName', this.summaryFileName)
        if (this.checkFileName) formData.append('checkFileName', this.checkFileName)
        formData.append('targetFile', this.targetFile)
        const res = await postAsync('/api/finance/summary/append', formData)
        if (res.data.success) {
          this.appendFileName = res.data.fileName
          this.downloadFileName = res.data.fileName
          this.addLog(`追加完成：${res.data.message}`)
          if (res.data.skipped && res.data.skipped.length) {
            this.$message.warning(res.data.message)
          } else {
            this.$message.success('追加完成！')
          }
        } else {
          this.$message.error(res.data.message || '追加失败')
        }
      } catch (err) {
        this.$message.error(err.response?.data?.message || err.message)
        this.addLog(`追加错误：${err.message}`)
      } finally {
        this.appending = false
      }
    },
    resetAll() {
      this.recordFile = null
      this.recordSheets = []
      this.selectedSheets = []
      this.checkFile = null
      this.checkSheet = ''
      this.checkSheets = []
      this.checkMonths = []
      this.checkMonthPersonMap = {}
      this.checkSelectedMonth = ''
      this.checkSelectedPersons = []
      this.checkPersonSearch = ''
      this.targetFile = null
      this.summaryFileName = ''
      this.checkFileName = ''
      this.appendFileName = ''
      this.downloadFileName = ''
      this.previewData = []
      this.previewColumns = []
      this.logs = []
      this.$refs.uploadRecord?.clearFiles()
      this.$refs.uploadCheck?.clearFiles()
      this.$refs.uploadTarget?.clearFiles()
      this.addLog('已重置所有数据')
    },
    downloadFile(fileName) {
      const url = '/api/finance/download/' + encodeURIComponent(fileName)
      const link = document.createElement('a')
      link.href = url
      link.download = fileName
      document.body.appendChild(link)
      link.click()
      document.body.removeChild(link)
    }
  }
}
</script>

<style scoped>
.page-container {
  width: 100%;
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
  min-height: 100px;
}

.cxm-upload .el-upload-dragger {
  width: 100%;
  min-height: 100px;
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
}

.cxm-upload-icon {
  font-size: 36px;
  color: #409EFF;
  margin-bottom: 8px;
}

.cxm-upload-label {
  font-size: 13px;
  font-weight: 600;
  color: #606266;
  margin-bottom: 6px;
}

.cxm-file-name {
  font-size: 12px;
  color: #409EFF;
  margin-top: 4px;
  word-break: break-all;
}

.cxm-sheet-checkbox-group {
  max-height: 180px;
  overflow-y: auto;
  border: 1px solid #dcdfe6;
  border-radius: 4px;
  padding: 8px 12px;
  background: #f5f7fa;
}

.cxm-sheet-checkbox {
  display: block;
  margin-right: 0 !important;
  height: 28px;
  line-height: 28px;
  font-size: 12px;
}

.cxm-sheet-empty {
  color: #909399;
  font-size: 12px;
  padding: 8px 0;
}

.cxm-dialog-section-title {
  font-size: 13px;
  font-weight: 600;
  color: #303133;
  margin-bottom: 8px;
  display: flex;
  align-items: center;
  height: 32px;
  flex-shrink: 0;
}

.cxm-dialog-content {
  max-height: 320px;
  overflow-y: auto;
}

.cxm-month-group {
  display: flex;
  flex-direction: column;
  gap: 4px;
}

.cxm-month-group .el-radio {
  margin-right: 0 !important;
  height: 28px;
  line-height: 28px;
}

.cxm-person-grid {
  display: grid;
  grid-template-columns: repeat(4, 1fr);
  gap: 4px 12px;
}

.cxm-person-item {
  font-size: 12px;
  margin-right: 0 !important;
  height: 28px;
  line-height: 28px;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.cxm-person-count {
  font-size: 12px;
  color: #909399;
  margin-top: 6px;
}

.cxm-log-box {
  background: #f5f7fa;
  border: 1px solid #e4e7ed;
  border-radius: 4px;
  padding: 10px 14px;
  max-height: 200px;
  overflow-y: auto;
  font-family: Consolas, 'Courier New', monospace;
  font-size: 12px;
  line-height: 1.6;
}

.cxm-log-line {
  color: #303133;
  white-space: pre-wrap;
  word-break: break-all;
}

.cxm-log-empty {
  color: #909399;
  font-style: italic;
}

.cxm-pending {
  color: #e6a23c;
  font-weight: 600;
}
</style>
