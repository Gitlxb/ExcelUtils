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
    </div>

    <!-- 进度 -->
    <div v-if="loading" class="progress-section">
      <el-progress :percentage="progressPercent" :status="progressStatus" />
    </div>

    <!-- 下载 -->
    <div v-if="downloadUrl" class="result-section">
      <el-alert title="处理完成！" type="success" :closable="false" show-icon />
      <el-button type="primary" @click="downloadFile" class="download-btn">
        <el-icon><Download /></el-icon> 下载处理后的文件
      </el-button>
    </div>

    <!-- 错误 -->
    <div v-if="errorMsg" class="result-section">
      <el-alert :title="errorMsg" type="error" :closable="false" show-icon />
    </div>
  </el-card>
</template>

<script>
import axios from 'axios'

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
      fileName: ''
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
  methods: {
    handleFileChange(uploadFile) {
      this.files = this.uploadRef?.uploadFiles?.map(f => f.raw) || [uploadFile.raw]
      this.errorMsg = ''
      this.downloadUrl = ''
    },
    handleExceed() {
      this.$message.warning('最多只能选择1个文件')
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

      // 模拟进度
      const timer = setInterval(() => {
        if (this.progressPercent < 90) this.progressPercent += 10
      }, 500)

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
      // 添加额外固定参数
      if (this.extraParams) {
        this.extraParams.split('&').forEach(p => {
          const [k, v] = p.split('=')
          formData.append(k, v)
        })
      }

      // FormData 上传不能手动设置 Content-Type，浏览器会自动设置并带上 boundary
      axios.post(this.apiEndpoint, formData).then(res => {
        clearInterval(timer)
        this.progressPercent = 100
        if (res.data.success) {
          this.progressStatus = 'success'
          this.fileName = res.data.fileName
          this.downloadUrl = `${this.downloadBase}/download/${this.fileName}`
        } else {
          this.progressStatus = 'exception'
          this.errorMsg = res.data.message || '处理失败'
        }
      }).catch(err => {
        clearInterval(timer)
        this.progressStatus = 'exception'
        this.errorMsg = err.response?.data?.message || err.message
      }).finally(() => {
        this.loading = false
      })
    },
    downloadFile() {
      if (this.downloadUrl) window.location.href = this.downloadUrl
    }
  }
}
</script>

<style scoped src="../styles/components/FunctionCard.css"></style>
