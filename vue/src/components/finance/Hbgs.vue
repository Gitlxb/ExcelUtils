<template>
  <el-card class="func-card">
    <div class="hbgs-page-header">
      <h2 class="hbgs-page-title">合并个税</h2>
      <p class="hbgs-page-desc">
        递归合并个税统计表，按手续费规则计算银行手续费。只需选择数据文件夹（其中须包含
        <strong>收取手续费银行信息.xlsx</strong>），系统自动识别规则表并生成汇总表。
      </p>
    </div>

    <!-- 唯一输入：数据文件夹（规则表从目录内自动识别，不再单独上传） -->
    <el-divider content-position="left">选择数据文件夹</el-divider>
    <el-row :gutter="16" style="margin-bottom:12px">
      <el-col :span="12">
        <div class="hbgs-upload-label">
          包含个税统计表的文件夹（递归子目录所有 Excel，须含「收取手续费银行信息.xlsx」）
        </div>
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
            已选择 {{ dataFileCount }} 个数据文件
          </div>
        </div>
        <div v-if="rulesFileName" class="hbgs-preview-info">
          <el-tag type="success" size="large">已识别规则表：{{ rulesFileName }}</el-tag>
        </div>
        <div v-if="folderChecked && !rulesFileName" class="hbgs-error">
          所选目录需要包含「收取手续费银行信息.xlsx」，未检测到该文件，已停止处理。
        </div>
      </el-col>
      <el-col :span="12">
        <div class="hbgs-upload-label">操作</div>
        <el-button
          type="primary"
          :loading="processing"
          :disabled="!canStart"
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
        <el-descriptions-item label="规则条数">{{ mergeStats.ruleCount }}</el-descriptions-item>
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
import { postAsync } from '../../utils/asyncTask'

/** 规则表固定文件名（去扩展名后比较） */
const RULES_FILE_BASE = '收取手续费银行信息'
/** 服务端单次请求上限（application.properties: max-request-size=300MB） */
const MAX_REQUEST_MB = 300

export default {
  name: 'HbgsCard',
  data() {
    return {
      /** 目录内全部 Excel（含规则表本身，后端负责识别并剔除） */
      dataFiles: [],
      /** 展示用：排除规则表后的数据文件数 */
      dataFileCount: 0,
      /** 已识别到的规则表文件名，空串表示未找到 */
      rulesFileName: '',
      /** 是否已进行过一次目录选择（用于控制"未包含规则表"提示的显示） */
      folderChecked: false,

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
  computed: {
    canStart() {
      return this.dataFiles.length > 0 && !!this.rulesFileName
    }
  },
  methods: {
    // ========== 目录选择 ==========
    triggerFolderInput() {
      this.$refs.folderInput.click()
    },

    /** 文件名是否为规则表（忽略扩展名大小写，兼容 .xls） */
    isRulesFileName(name) {
      if (!name) return false
      const dot = name.lastIndexOf('.')
      if (dot <= 0) return false
      const base = name.slice(0, dot).trim()
      const ext = name.slice(dot + 1).toLowerCase()
      return (ext === 'xlsx' || ext === 'xls') && base === RULES_FILE_BASE
    },

    handleFolderChange(e) {
      const files = e.target.files
      this.folderChecked = false
      this.rulesFileName = ''

      if (!files || files.length === 0) {
        this.dataFiles = []
        this.dataFileCount = 0
        return
      }

      const all = Array.from(files).filter(f => {
        const name = f.name.toLowerCase()
        return (name.endsWith('.xlsx') || name.endsWith('.xls')) && !f.name.startsWith('~$')
      })

      const rules = all.find(f => this.isRulesFileName(f.name))
      this.rulesFileName = rules ? rules.name : ''
      this.dataFiles = all
      this.dataFileCount = Math.max(0, all.length - (rules ? 1 : 0))
      this.folderChecked = true

      // 允许重复选择同一目录（否则 change 不触发）
      e.target.value = ''

      if (!rules) {
        this.$message.error('所选择的目录需要包含「' + RULES_FILE_BASE + '.xlsx」，已停止处理。')
        return
      }

      // 体积预警：当前服务端单次请求上限 50MB
      const totalMb = all.reduce((s, f) => s + (f.size || 0), 0) / 1048576
      if (totalMb > MAX_REQUEST_MB * 0.9) {
        this.$message.warning(
          `所选目录合计约 ${totalMb.toFixed(1)} MB，接近/超过上传上限 ${MAX_REQUEST_MB} MB，可能上传失败`
        )
      }
    },

    // ========== 合并 ==========
    startMerge() {
      if (!this.dataFiles.length) {
        this.$message.warning('请先选择数据文件夹')
        return
      }
      if (!this.rulesFileName) {
        this.$message.error('所选择的目录需要包含「' + RULES_FILE_BASE + '.xlsx」，已停止处理。')
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

      // 模拟进度（后端异步任务，真实进度待后续接入轮询）
      const timer = setInterval(() => {
        if (this.progressPercent < 90) {
          this.progressPercent += 5
          this.progressText = `正在处理... (${Math.min(this.progressPercent, 99)}%)`
        }
      }, 600)

      const formData = new FormData()
      // 规则表随目录一起上传，由后端按文件名识别并剔除
      this.dataFiles.forEach(f => formData.append('dataFiles', f))

      postAsync('/api/finance/hbgs/merge', formData, {
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
              skipped: res.data.skipped,
              ruleCount: res.data.ruleCount
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
      this.dataFiles = []
      this.dataFileCount = 0
      this.rulesFileName = ''
      this.folderChecked = false
      this.resultFileName = ''
      this.mergeStats = null
      this.mergeWarnings = []
      this.mergeError = ''
      this.logs = []
      this.progressPercent = 0
      this.progressStatus = ''
      this.progressText = ''
      if (this.$refs.folderInput) this.$refs.folderInput.value = ''
    }
  }
}
</script>

<style scoped src="../../styles/components/HbgsCard.css"></style>
<style scoped src="../../styles/components/FuncPageHeader.css"></style>
