<template>
  <el-card class="func-card">
    <div class="hbgs-page-header">
      <h2 class="hbgs-page-title">合并个税</h2>
      <p class="hbgs-page-desc">
        递归合并个税统计表，按手续费规则计算银行手续费。上传规则表和数据文件夹后，自动处理并生成汇总表。
      </p>
    </div>

    <!-- 步骤1：上传规则表 -->
    <el-divider content-position="left">步骤1：上传手续费规则表</el-divider>
    <el-row :gutter="16" style="margin-bottom:12px">
      <el-col :span="12">
        <div class="hbgs-upload-label">手续费规则表（.xlsx/.xls）</div>
        <el-upload
          ref="uploadRules"
          class="hbgs-upload"
          drag
          :auto-upload="false"
          :limit="1"
          :on-change="handleRulesChange"
          :on-exceed="onExceed"
          accept=".xlsx,.xls"
        >
          <el-icon class="hbgs-upload-icon"><UploadFilled /></el-icon>
          <div>将规则表拖到此处，或 <em>点击选择</em></div>
          <div v-if="rulesFile" class="hbgs-file-name">{{ rulesFile.name }}</div>
        </el-upload>
      </el-col>
      <el-col :span="12">
        <div class="hbgs-upload-label">规则表预检</div>
        <el-button type="primary" :loading="previewing" :disabled="!rulesFile" @click="previewRules">
          {{ previewing ? '预检中...' : '预检规则表' }}
        </el-button>
        <div v-if="rulesPreview" class="hbgs-preview-info">
          <el-tag type="success" size="large">
            {{ rulesPreview.worksheet }}，共 {{ rulesPreview.ruleCount }} 条规则
          </el-tag>
        </div>
        <div v-if="rulesError" class="hbgs-error">{{ rulesError }}</div>
      </el-col>
    </el-row>

    <!-- 步骤2：上传数据文件夹 -->
    <el-divider content-position="left">步骤2：上传数据文件夹</el-divider>
    <el-row :gutter="16" style="margin-bottom:12px">
      <el-col :span="12">
        <div class="hbgs-upload-label">包含个税统计表的文件夹（递归子目录所有 Excel）</div>
        <div class="hbgs-folder-upload-area" @click="triggerFolderInput">
          <input
            ref="folderInput"
            type="file"
            webkitdirectory
            multiple
            accept=".xlsx,.xls"
            style="display:none"
            @change="handleFolderChange"
          />
          <el-icon class="hbgs-upload-icon"><FolderOpened /></el-icon>
          <div>点击此处 <em>选择文件夹</em></div>
          <div v-if="dataFileCount > 0" class="hbgs-file-name">
            已选择 {{ dataFileCount }} 个 Excel 文件
          </div>
        </div>
      </el-col>
      <el-col :span="12">
        <div class="hbgs-upload-label">操作</div>
        <el-button
          type="primary"
          :loading="processing"
          :disabled="!rulesFile || !dataFiles.length"
          @click="startMerge"
          size="large"
        >
          {{ processing ? '处理中...' : '开始合并' }}
        </el-button>
        <el-button @click="clearAll" :disabled="processing">清空</el-button>
        <div v-if="resultFileName" class="hbgs-download-area">
          <el-alert type="success" :closable="false" show-icon>
            <template #default>
              <span>合并完成：{{ resultFileName }}</span>
              <el-button type="primary" size="small" style="margin-left:12px" @click="downloadFile">
                <el-icon><Download /></el-icon> 下载
              </el-button>
            </template>
          </el-alert>
        </div>
      </el-col>
    </el-row>

    <!-- 统计结果 -->
    <div v-if="mergeStats" class="hbgs-stats">
      <el-descriptions :column="4" border size="small">
        <el-descriptions-item label="扫描文件">{{ mergeStats.totalFiles }}</el-descriptions-item>
        <el-descriptions-item label="有效数据行">{{ mergeStats.totalRows }}</el-descriptions-item>
        <el-descriptions-item label="跳过文件">{{ mergeStats.skipped }}</el-descriptions-item>
      </el-descriptions>
    </div>

    <!-- 警告提示 -->
    <div v-if="mergeWarnings && mergeWarnings.length" class="hbgs-warnings">
      <el-alert
        v-for="(w, i) in mergeWarnings"
        :key="i"
        type="warning"
        :closable="false"
        show-icon
        style="margin-bottom: 6px"
      >
        <template #title>{{ w }}</template>
      </el-alert>
    </div>

    <!-- 进度 -->
    <div v-if="processing" class="hbgs-progress">
      <el-progress :percentage="progressPercent" :status="progressStatus" />
      <span class="hbgs-progress-text">{{ progressText }}</span>
    </div>

    <!-- 日志区 -->
    <div v-if="logs.length" class="hbgs-logs">
      <h4>处理日志：</h4>
      <div class="hbgs-log-box">
        <div v-for="(line, i) in logs" :key="i" class="hbgs-log-line">{{ line }}</div>
      </div>
    </div>

    <!-- 错误 -->
    <div v-if="mergeError" class="hbgs-error-area">
      <el-alert :title="mergeError" type="error" :closable="false" show-icon />
    </div>

  </el-card>
</template>

<script>
import axios from 'axios'

export default {
  name: 'HbgsCard',
  data() {
    return {
      rulesFile: null,
      rulesPreview: null,
      rulesError: '',
      previewing: false,

      dataFiles: [],
      dataFileCount: 0,

      processing: false,
      progressPercent: 0,
      progressStatus: '',
      progressText: '',

      resultFileName: '',
      mergeStats: null,
      mergeWarnings: [],
      mergeError: '',
      logs: []
    }
  },
  methods: {
    // ========== 规则表 ==========

    handleRulesChange(file) {
      this.rulesFile = file?.raw || null
      this.rulesPreview = null
      this.rulesError = ''
    },
    onExceed() {
      this.$message.warning('只能选择1个文件')
    },

    previewRules() {
      if (!this.rulesFile) {
        this.$message.warning('请先选择规则表文件')
        return
      }
      this.previewing = true
      this.rulesError = ''
      this.rulesPreview = null

      const formData = new FormData()
      formData.append('rulesFile', this.rulesFile)

      axios.post('/api/finance/hbgs/preview-rules', formData)
        .then(res => {
          if (res.data.success) {
            this.rulesPreview = {
              worksheet: res.data.worksheet || '',
              ruleCount: res.data.ruleCount || 0
            }
            if (res.data.logs) {
              this.logs = res.data.logs.concat(this.logs)
            }
          } else {
            this.rulesError = res.data.message || '预检失败'
          }
        })
        .catch(err => {
          this.rulesError = err.response?.data?.message || err.message
        })
        .finally(() => {
          this.previewing = false
        })
    },

    // ========== 文件夹上传 ==========

    triggerFolderInput() {
      this.$refs.folderInput.click()
    },
    handleFolderChange(e) {
      const files = e.target.files
      if (files && files.length > 0) {
        this.dataFiles = Array.from(files).filter(f => {
          const name = f.name.toLowerCase()
          return (name.endsWith('.xlsx') || name.endsWith('.xls')) && !f.name.startsWith('~$')
        })
        this.dataFileCount = this.dataFiles.length
      } else {
        this.dataFiles = []
        this.dataFileCount = 0
      }
    },

    // ========== 合并 ==========

    startMerge() {
      if (!this.rulesFile) {
        this.$message.warning('请先选择手续费规则表')
        return
      }
      if (!this.dataFiles.length) {
        this.$message.warning('请先选择数据文件夹')
        return
      }

      this.processing = true
      this.progressPercent = 0
      this.progressStatus = ''
      this.progressText = '准备处理...'
      this.resultFileName = ''
      this.mergeStats = null
      this.mergeWarnings = []
      this.mergeError = ''
      this.logs = ['开始合并...']

      // 模拟进度
      const timer = setInterval(() => {
        if (this.progressPercent < 90) {
          this.progressPercent += 5
          this.progressText = `正在处理... (${Math.min(this.progressPercent, 99)}%)`
        }
      }, 600)

      const formData = new FormData()
      formData.append('rulesFile', this.rulesFile)
      this.dataFiles.forEach(f => formData.append('dataFiles', f))

      axios.post('/api/finance/hbgs/merge', formData, {
        timeout: 600000  // 10分钟超时（大文件可能需要较长时间）
      })
        .then(res => {
          clearInterval(timer)
          this.progressPercent = 100
          if (res.data.success) {
            this.progressStatus = 'success'
            this.progressText = '处理完成'
            this.resultFileName = res.data.fileName
            this.mergeStats = {
              totalFiles: res.data.totalFiles,
              totalRows: res.data.totalRows,
              skipped: res.data.skipped
            }
            this.mergeWarnings = res.data.warnings || []
          } else {
            this.progressStatus = 'exception'
            this.progressText = '处理失败'
            this.mergeError = res.data.message || '处理失败'
          }
          if (res.data.logs) {
            this.logs = res.data.logs
          }
        })
        .catch(err => {
          clearInterval(timer)
          this.progressStatus = 'exception'
          this.progressText = '处理失败'
          this.mergeError = err.response?.data?.message || err.message || '网络错误'
        })
        .finally(() => {
          this.processing = false
        })
    },

    downloadFile() {
      if (this.resultFileName) {
        window.location.href = `/api/finance/download/${this.resultFileName}`
      }
    },

    clearAll() {
      this.rulesFile = null
      this.rulesPreview = null
      this.rulesError = ''
      this.dataFiles = []
      this.dataFileCount = 0
      this.resultFileName = ''
      this.mergeStats = null
      this.mergeWarnings = []
      this.mergeError = ''
      this.logs = []
      this.progressPercent = 0
      this.progressStatus = ''
      this.progressText = ''
      this.$refs.uploadRules?.clearFiles()
      if (this.$refs.folderInput) this.$refs.folderInput.value = ''
    }
  }
}
</script>

<style scoped src="../../styles/components/HbgsCard.css"></style>
<style scoped src="../../styles/components/FuncPageHeader.css"></style>
