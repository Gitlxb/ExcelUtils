<template>
  <el-card class="func-card">
    <div class="page-container">
    <div class="func-page-header">
      <h2 class="func-page-title">预支处理</h2>
      <p class="func-page-desc">上传工人预支申请 Excel 文件，自动处理并下载结果。</p>
    </div>

    <div class="tb-section-title">上传文件</div>
    <el-upload
      ref="uploadFee"
      class="tb-upload"
      drag
      :auto-upload="false"
      :limit="1"
      :on-change="handleFeeFileChange"
      accept=".xlsx"
    >
      <el-icon class="tb-upload-icon"><UploadFilled /></el-icon>
      <div>将文件拖到此处，或 <em>点击选择</em></div>
      <div v-if="feeFile" class="tb-file-name">{{ feeFile.name }}</div>
    </el-upload>
    <div class="tb-section-title" style="margin-top: 16px">操作</div>
    <el-button type="success" :loading="loading" @click="submitFeeSplit">开始处理</el-button>
    <el-button style="margin-left: 8px" @click="refreshFee">
      <el-icon><Refresh /></el-icon> 刷新
    </el-button>
    <el-card v-if="feeDownloadName" style="margin-top: 16px">
      <span>处理完成：{{ feeDownloadName }}</span>
      <el-button type="primary" style="margin-left: 12px" @click="downloadFile(feeDownloadName)">下载文件</el-button>
    </el-card>

    <el-alert
      v-if="message && !feeDownloadName"
      :title="message"
      :type="messageType"
      show-icon
      closable
      style="margin-top: 16px"
      @close="message = ''"
    />
    </div>
  </el-card>
</template>

<script>
import { postAsync } from '../../../utils/asyncTask'

export default {
  name: 'CwgjxYuzhi',
  data() {
    return {
      loading: false,
      message: '',
      messageType: 'info',
      feeFile: null,
      feeDownloadName: ''
    }
  },
  methods: {
    handleFeeFileChange(file) {
      this.feeFile = file.raw
    },
    showMessage(text, type = 'info') {
      this.message = text
      this.messageType = type
    },
    submitFeeSplit() {
      if (!this.feeFile) {
        this.showMessage('请先选择文件', 'warning')
        return
      }
      const formData = new FormData()
      formData.append('file', this.feeFile)
      this.loading = true
      this.message = ''
      postAsync('/api/finance/toolbox/fee-split', formData).then(res => {
        if (res.data.success) {
          this.showMessage(res.data.message, 'success')
          this.feeDownloadName = res.data.fileName
        } else {
          this.showMessage(res.data.message || '处理失败', 'error')
        }
      }).catch(err => {
        this.showMessage(err.response?.data?.message || err.message, 'error')
      }).finally(() => {
        this.loading = false
      })
    },
    downloadFile(fileName) {
      const url = '/api/finance/download/' + encodeURIComponent(fileName)
      const link = document.createElement('a')
      link.href = url
      link.download = fileName
      document.body.appendChild(link)
      link.click()
      document.body.removeChild(link)
    },
    refreshFee() {
      this.loading = false
      this.showMessage('预支处理已刷新', 'success')
    }
  }
}
</script>

<style scoped>
.page-container {
  width: 100%;
}

/* 功能页面标题头部 */
.func-page-header {
  margin-bottom: 20px;
  padding: 16px 20px;
  background: linear-gradient(135deg, #f0f7ff 0%, #e8f4e8 100%);
  border-radius: 8px;
  border-left: 4px solid #409EFF;
}

.func-page-header .func-page-title {
  margin: 0 0 8px;
  font-size: 18px;
  font-weight: 600;
  color: #303133;
}

.func-page-header .func-page-desc {
  margin: 0;
  font-size: 13px;
  color: #606266;
  line-height: 1.6;
}

.tb-section-title {
  font-size: 14px;
  font-weight: 600;
  color: #303133;
  margin-bottom: 8px;
  padding-left: 8px;
  border-left: 3px solid #409EFF;
}

.tb-upload { width: 100%; }
.tb-upload .el-upload-dragger {
  width: 100% !important;
  height: 100px;
}

.tb-upload-icon {
  font-size: 36px;
  color: #409EFF;
  margin-bottom: 4px;
}

.tb-file-name {
  font-size: 12px;
  color: #67C23A;
  margin-top: 4px;
}
</style>
