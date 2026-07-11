<template>
  <el-card class="func-card">
    <div class="page-container">
    <div class="func-page-header">
      <h2 class="func-page-title">数据匹配</h2>
      <p class="func-page-desc">
        每月绩效考核表的数据匹配处理。选择本月原始表和上月已核对表，自动进行数据匹配生成结果。
      </p>
    </div>

    <el-form label-width="140px">
      <el-form-item label="本月原始表">
        <el-upload
          ref="uploadSource"
          drag
          :auto-upload="false"
          :limit="1"
          :on-change="handleSourceChange"
          accept=".xlsx,.xls"
        >
          <el-icon class="el-icon--upload"><UploadFilled /></el-icon>
          <div class="el-upload__text">将文件拖到此处，或 <em>点击选择</em></div>
        </el-upload>
      </el-form-item>
      <el-form-item label="上月（已核对）">
        <el-upload
          ref="uploadPrev"
          drag
          :auto-upload="false"
          :limit="1"
          :on-change="handlePrevChange"
          accept=".xlsx,.xls"
        >
          <el-icon class="el-icon--upload"><UploadFilled /></el-icon>
          <div class="el-upload__text">将文件拖到此处，或 <em>点击选择</em></div>
        </el-upload>
      </el-form-item>
      <el-form-item>
        <el-button type="success" :loading="loading" @click="submitSjpp">开始生成</el-button>
      </el-form-item>
    </el-form>
    <el-alert
      title="说明：程序会自动识别以'代工费-' / '百事通报备名单-' / '员工信息汇总-'开头的工作表。"
      type="info"
      :closable="false"
      style="margin-top: 12px"
    />

    <!-- 下载区域 -->
    <el-card v-if="downloadFileName" style="margin-top: 16px">
      <span>处理完成：{{ downloadFileName }}</span>
      <el-button type="primary" style="margin-left: 12px" @click="downloadFile(downloadFileName)">
        下载文件
      </el-button>
    </el-card>

    <el-alert
      v-if="message && !downloadFileName"
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
import axios from 'axios'

export default {
  name: 'JxkhSjpp',
  data() {
    return {
      loading: false,
      message: '',
      messageType: 'info',
      downloadFileName: '',
      sourceFile: null,
      prevFile: null
    }
  },
  methods: {
    handleSourceChange(file) {
      this.sourceFile = file.raw
    },
    handlePrevChange(file) {
      this.prevFile = file.raw
    },
    showMessage(text, type = 'info') {
      this.message = text
      this.messageType = type
    },
    submitSjpp() {
      if (!this.sourceFile) {
        this.showMessage('请选择本月原始表', 'warning')
        return
      }
      const formData = new FormData()
      formData.append('sourceFile', this.sourceFile)
      if (this.prevFile) formData.append('prevFile', this.prevFile)
      this.post('/api/finance/jxkh/sjpp', formData)
    },
    post(endpoint, formData) {
      this.loading = true
      this.message = ''
      this.downloadFileName = ''
      axios.post(endpoint, formData).then(res => {
        if (res.data.success) {
          this.showMessage(res.data.message, 'success')
          this.downloadFileName = res.data.fileName
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
    }
  }
}
</script>

<style scoped>
.page-container {
  width: 100%;
}

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
</style>
