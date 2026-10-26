<template>
  <el-card class="func-card">
    <template #header>
      <div class="card-header">
        <span>{{ title }}</span>
        <el-tag v-if="description" type="info" size="small">{{ description }}</el-tag>
      </div>
    </template>

    <!-- 文件上传 -->
    <el-upload
      ref="uploadRef"
      drag
      :accept="fileAccept"
      :auto-upload="false"
      :show-file-list="true"
      :limit="fileLimit"
      :on-change="handleFileChange"
      :on-exceed="handleExceed"
    >
      <el-icon class="el-icon--upload"><UploadFilled /></el-icon>
      <div class="el-upload__text">
        拖拽文件到此处或 <em>点击上传</em>
      </div>
      <template #tip>
        <div class="el-upload__tip">
          支持格式：{{ fileAccept }} | {{ multiFile ? '可同时选择多个文件' : '每次选择一个文件' }}
        </div>
      </template>
    </el-upload>

    <!-- 参数输入 -->
    <div v-if="showParams" class="params-section">
      <div v-if="!multiParams" class="param-row">
        <el-input
          v-model="paramValue"
          :placeholder="paramPlaceholder"
          clearable
          style="width:300px"
        >
          <template #prepend>{{ paramLabel }}</template>
        </el-input>
      </div>
      <div v-else class="param-row">
        <el-input
          v-model="paramValue"
          :placeholder="paramPlaceholder"
          clearable
          style="width:220px;margin-right:10px"
        >
          <template #prepend>{{ paramLabel }}</template>
        </el-input>
        <el-input
          v-model="paramValue2"
          :placeholder="paramPlaceholder2"
          clearable
          style="width:220px"
        >
          <template #prepend>{{ paramLabel2 }}</template>
        </el-input>
      </div>
    </div>

    <!-- 操作按钮 -->
    <div class="action-section">
      <el-button
        type="primary"
        :loading="loading"
        :disabled="!files.length"
        @click="startProcess"
      >
        {{ loading ? '处理中...' : '开始处理' }}
      </el-button>
      <el-button
        v-if="restoredTaskId && !loading && !downloadUrl"
        size="small"
        style="margin-left:8px"
        @click="resumePolling"
      >恢复查看上次任务</el-button>
    </div>

    <!-- 进度 -->
    <div v-if="loading" class="progress-section">
      <el-progress :percentage="progressPercent" :status="progressStatus" />
      <div class="process-hint" v-if="loading && !errorMsg">
        任务已在后台处理，可切换到其他页面，完成后回来自动恢复结果
      </div>
    </div>

    <!-- 下载 -->
    <div v-if="downloadUrl" class="result-section">
      <el-alert title="处理完成！" type="success" :closable="false" show-icon />
      <el-button type="primary" @click="downloadFile" class="download-btn">
        <el-icon><Download /></el-icon> 下载处理后的文件
      </el-button>
      <div v-if="resultExtras" class="result-extras" v-html="resultExtras"></div>
    </div>

    <!-- 错误 -->
    <div v-if="errorMsg" class="result-section">
      <el-alert :title="errorMsg" type="error" :closable="false" show-icon />
    </div>
  </el-card>
</template>

<script>
import axios from 'axios'
import { rememberTask, forgetTask, getActiveTasks } from '../utils/deviceId'

export default {
  name: 'FunctionCard',
  props: {
    title: { type: String, required: true },
    description: { type: String, default: '' },
    singleFile: { type: Boolean, default: true },
    multiFile: { type: Boolean, default: false },
    fileAccept: { type: String, default: '.xlsx' },
    showParams: { type: Boolean, default: false },
    paramLabel: { type: String, default: '' },
    paramPlaceholder: { type: String, default: '' },
    paramLabel2: { type: String, default: '' },
    paramPlaceholder2: { type: String, default: '' },
    multiParams: { type: Boolean, default: false },
    apiEndpoint: { type: String, required: true },
    paramName: { type: String, default: '' },
    paramName2: { type: String, default: '' },
    extraParams: { type: String, default: '' }
  },
  emits: ['process'],
  data() {
    return {
      files: [],
      paramValue: '',
      paramValue2: '',
      loading: false,
      progressPercent: 0,
      progressStatus: '',
      downloadUrl: '',
      errorMsg: '',
      fileName: '',
      resultExtras: '',        // 定制页面展示用的附加结果（如预览表/统计），按需填充
      pollTimer: null,
      activeTaskId: null,      // 当前轮询中的任务
      restoredTaskId: null     // 进入页面时从 localStorage 找到的本工具任务
    }
  },
  computed: {
    fileLimit() {
      return this.multiFile ? 0 : (this.singleFile ? 1 : 0)
    },
    downloadBase() {
      // 从 apiEndpoint 提取 controller 路径
      const match = this.apiEndpoint.match(/^(\/api\/[^/]+)/)
      return match ? match[1] : '/api'
    }
  },
  mounted() {
    this.tryRestoreTask()
  },
  beforeUnmount() {
    this.stopPolling()   // 只停轮询，不取消任务（后台继续）
  },
  methods: {
    handleFileChange(uploadFile) {
      this.files = this.uploadRef?.uploadFiles?.map(f => f.raw) || [uploadFile.raw]
      this.errorMsg = ''
      this.downloadUrl = ''
      this.restoredTaskId = null
      this.stopPolling()
    },
    handleExceed() {
      this.$message.warning('最多只能选择1个文件')
    },

    /** 进入页面时：若本工具曾有未完成任务（localStorage 记录），尝试恢复 */
    async tryRestoreTask() {
      const mine = getActiveTasks().filter(id => id)
      if (!mine.length) return
      // 逐个查询后端状态，找到属于本 apiEndpoint 的进行中/已完成未清理任务
      for (const taskId of mine) {
        try {
          const res = await axios.get('/api/tasks/' + taskId)
          const task = res.data && res.data.task
          if (!task) { forgetTask(taskId); continue }
          // 匹配：task.apiName === apiEndpoint 去前缀后的路径，或 apiEndpoint 含 task.apiName
          const api = this.apiEndpoint.replace(/^\//, '')
          if (api !== task.apiName && !api.endsWith(task.apiName)) continue
          if (task.status === 'SUCCESS') {
            // 已完成 → 直接展示下载
            this.applyResult(task)
            forgetTask(taskId)
          } else if (task.status === 'RUNNING' || task.status === 'PENDING') {
            this.restoredTaskId = taskId
          } else if (task.status === 'FAILED') {
            this.errorMsg = task.message || '处理失败'
            this.loading = false
            forgetTask(taskId)
          }
        } catch (e) { /* 任务已过期/无权 → 忽略 */ }
      }
    },

    resumePolling() {
      if (this.restoredTaskId) {
        this.startPolling(this.restoredTaskId)
      }
    },

    startProcess() {
      if (!this.files.length) {
        this.$message.warning('请先选择文件')
        return
      }
      this.loading = true
      this.progressPercent = 0
      this.progressStatus = ''
      this.downloadUrl = ''
      this.errorMsg = ''
      this.restoredTaskId = null
      // 模拟进度（真实百分比由任务中心后续增强；当前仅示进度感）
      const timer = setInterval(() => {
        if (this.loading && this.progressPercent < 90) this.progressPercent += 8
      }, 800)
      const stopFakeProgress = () => { clearInterval(timer) }

      const formData = new FormData()
      if (this.multiFile) {
        this.files.forEach(f => formData.append('file', f))
      } else {
        formData.append('file', this.files[0])
      }
      if (this.showParams && this.paramValue) {
        formData.append(this.paramName, this.paramValue)
        if (this.multiParams && this.paramValue2) {
          formData.append(this.paramName2, this.paramValue2)
        }
      }
      if (this.extraParams) {
        this.extraParams.split('&').forEach(p => {
          const [k, v] = p.split('=')
          formData.append(k, v)
        })
      }

      axios.post(this.apiEndpoint, formData).then(res => {
        stopFakeProgress()
        // 异步任务模式：返回 { taskId } → 轮询
        if (res.data && res.data.taskId) {
          this.activeTaskId = res.data.taskId
          rememberTask(res.data.taskId)
          this.startPolling(res.data.taskId)
        } else if (res.data && res.data.success && res.data.fileName) {
          // 兼容：未改造端点（理论上 FunctionCard 工具已全异步）直接给结果
          this.progressPercent = 100
          this.progressStatus = 'success'
          this.fileName = res.data.fileName
          this.downloadUrl = `${this.downloadBase}/download/${res.data.fileName}`
          this.loading = false
        } else {
          this.loading = false
          this.progressStatus = 'exception'
          this.errorMsg = (res.data && res.data.message) || '处理失败'
        }
      }).catch(err => {
        stopFakeProgress()
        this.loading = false
        this.progressStatus = 'exception'
        this.errorMsg = err.response?.data?.message || err.message
      })
    },

    /** 每 1.5s 轮询任务状态；SUCCESS → 展示下载；FAILED → 错误 */
    startPolling(taskId) {
      this.stopPolling()
      const poll = async () => {
        try {
          const res = await axios.get('/api/tasks/' + taskId)
          const task = res.data && res.data.task
          if (!task) { this.finishPollingError('任务不存在或已过期'); return }
          if (task.status === 'SUCCESS') {
            this.applyResult(task)
            forgetTask(taskId)
            this.stopPolling()
          } else if (task.status === 'FAILED') {
            this.applyFailed(task)
            forgetTask(taskId)
            this.stopPolling()
          }
          // RUNNING/PENDING → 继续轮询
        } catch (e) {
          // 网络瞬断不终止，下次轮询重试
        }
      }
      poll() // 立即查一次
      this.pollTimer = setInterval(poll, 1500)
    },

    stopPolling() {
      if (this.pollTimer) {
        clearInterval(this.pollTimer)
        this.pollTimer = null
      }
    },

    /** 成功：fileName 拼下载 URL + 附带 result（预览/统计）留给定制页面通过事件透出 */
    applyResult(task) {
      const result = task.result || {}
      this.loading = false
      this.progressPercent = 100
      this.progressStatus = 'success'
      this.restoredTaskId = null
      this.activeTaskId = null
      if (result.fileName) {
        this.fileName = result.fileName
        const module = task.module || this.downloadBase.replace('/api/', '')
        this.downloadUrl = `/api/${module}/download/${encodeURIComponent(result.fileName)}`
      } else {
        this.downloadUrl = ''
      }
      // 附带业务字段（logs/warnings/preview 等）通过 process 事件交给定制页
      const payload = Object.assign({}, result)
      if (Object.keys(payload).some(k => k !== 'fileName' && k !== 'success' && k !== 'message')) {
        this.resultExtras = ''
        this.$emit('process', { taskId: task.taskId, result: payload })
      }
    },

    applyFailed(task) {
      this.loading = false
      this.progressStatus = 'exception'
      this.progressPercent = 100
      this.restoredTaskId = null
      this.activeTaskId = null
      this.errorMsg = task.message || '处理失败'
      // 失败也透出 result（可能含 logs/warnings）
      if (task.result) {
        this.$emit('process', { taskId: task.taskId, result: task.result })
      }
    },

    finishPollingError(msg) {
      this.loading = false
      this.progressStatus = 'exception'
      this.errorMsg = msg
    },

    downloadFile() {
      if (this.downloadUrl) window.location.href = this.downloadUrl
    }
  }
}
</script>

<style scoped src="../styles/components/FunctionCard.css"></style>
