<template>
  <el-card class="func-card">
    <template #header>
      <div class="card-header">
        <span>批量批注工具</span>
        <el-tag type="info" size="small">添加批注 / 拆分批注</el-tag>
      </div>
    </template>

    <el-upload
      ref="uploadRef"
      drag
      accept=".xlsx"
      :auto-upload="false"
      :show-file-list="true"
      :limit="1"
      :on-change="handleFileChange"
      :on-exceed="onExceed"
    >
      <el-icon class="el-icon--upload"><UploadFilled /></el-icon>
      <div class="el-upload__text">拖拽或点击上传Excel文件</div>
    </el-upload>

    <div class="action-section">
      <el-button
        type="success"
        :loading="loadingAdd"
        :disabled="!file"
        @click="addComments"
      >
        <el-icon><EditPen /></el-icon> 批量添加批注
        <div class="btn-desc">将B~K列内容合并写入A列批注</div>
      </el-button>
      <el-button
        type="warning"
        :loading="loadingSplit"
        :disabled="!file"
        @click="splitComments"
      >
        <el-icon><DocumentCopy /></el-icon> 批量拆分批注
        <div class="btn-desc">将A列批注拆开写入右侧列</div>
      </el-button>
    </div>

    <div v-if="loadingAdd || loadingSplit" class="progress-section">
      <el-progress :percentage="progressPercent" :status="progressStatus" />
    </div>

    <div v-if="downloadUrl" class="result-section">
      <el-alert title="处理完成！" type="success" :closable="false" show-icon />
      <el-button type="primary" @click="downloadFile" class="download-btn">
        <el-icon><Download /></el-icon> 下载处理后的文件
      </el-button>
    </div>

    <div v-if="errorMsg" class="result-err">{{ errorMsg }}</div>
  </el-card>
</template>

<script>
import axios from 'axios'

export default {
  name: 'CommentCard',
  data() {
    return {
      file: null,
      loadingAdd: false,
      loadingSplit: false,
      progressPercent: 0,
      progressStatus: '',
      downloadUrl: '',
      errorMsg: ''
    }
  },
  methods: {
    handleFileChange() {
      this.$nextTick(() => {
        const upload = this.$refs.uploadRef
        const files = upload?.uploadFiles
        this.file = files?.length ? files[0].raw : null
        this.errorMsg = ''
        this.downloadUrl = ''
      })
    },
    onExceed() { this.$message.warning('只能选1个文件') },
    addComments() {
      this.exec('/api/finance/comment/add', 'loadingAdd')
    },
    splitComments() {
      this.exec('/api/finance/comment/split', 'loadingSplit')
    },
    exec(endpoint, loadingKey) {
      if (!this.file) { this.$message.warning('请先选择文件'); return }
      this[loadingKey] = true
      this.errorMsg = ''
      this.downloadUrl = ''
      this.progressPercent = 0
      this.progressStatus = ''

      const timer = setInterval(() => {
        if (this.progressPercent < 90) this.progressPercent += 10
      }, 500)

      const formData = new FormData()
      formData.append('file', this.file)

      // FormData 上传不能手动设置 Content-Type，浏览器会自动设置并带上 boundary
      axios.post(endpoint, formData).then(res => {
        clearInterval(timer)
        this.progressPercent = 100
        if (res.data.success) {
          this.progressStatus = 'success'
          this.downloadUrl = `/api/finance/download/${res.data.fileName}`
        } else {
          this.progressStatus = 'exception'
          this.errorMsg = res.data.message || '处理失败'
        }
      }).catch(err => {
        clearInterval(timer)
        this.progressStatus = 'exception'
        this.errorMsg = err.response?.data?.message || err.message
      }).finally(() => {
        this[loadingKey] = false
      })
    },
    downloadFile() {
      if (this.downloadUrl) window.location.href = this.downloadUrl
    }
  }
}
</script>

<style scoped src="../../styles/components/CommentCard.css"></style>
