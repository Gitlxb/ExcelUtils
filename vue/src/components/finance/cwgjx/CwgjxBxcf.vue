<template>
  <el-card class="func-card">
    <div class="page-container">
    <div class="func-page-header">
      <h2 class="func-page-title">报销拆分</h2>
      <p class="func-page-desc">上传报销表格，选择拆分列后自动拆分生成结果文件。</p>
    </div>

    <div class="tb-section-title">步骤1：上传文件</div>
    <el-upload
      ref="uploadExpense"
      class="tb-upload"
      drag
      :auto-upload="false"
      :limit="1"
      :on-change="handleExpenseFileChange"
      accept=".xlsx"
    >
      <el-icon class="tb-upload-icon"><UploadFilled /></el-icon>
      <div>将文件拖到此处，或 <em>点击选择</em></div>
      <div v-if="formExpense.file" class="tb-file-name">{{ formExpense.file.name }}</div>
    </el-upload>
    <div class="tb-section-title" style="margin-top: 16px">步骤2：处理预览</div>
    <el-button type="primary" :loading="expensePreviewing" :disabled="!formExpense.file" @click="submitExpensePreview">处理并预览列名</el-button>
    <el-card v-if="expensePreviewCols.length > 0" style="margin-top: 12px">
      <template #header>处理后列名（共 {{ expensePreviewCols.length }} 列）</template>
      <div style="max-height:200px;overflow:auto">
        <el-tag v-for="(col, idx) in expensePreviewCols" :key="idx" style="margin:2px 4px">{{ idx + 1 }}. {{ col }}</el-tag>
      </div>
    </el-card>
    <div class="tb-section-title" style="margin-top: 16px">步骤3：选择拆分列并执行</div>
    <el-form label-position="top" style="margin-top: 8px">
      <el-form-item label="拆分列">
        <el-select v-model="formExpense.splitColumn" placeholder="请选择拆分列" style="width:100%">
          <el-option v-for="(col, idx) in expensePreviewCols" :key="idx" :label="col" :value="col" />
        </el-select>
      </el-form-item>
    </el-form>
    <div class="tb-section-title">操作</div>
    <el-button type="success" :loading="loading" :disabled="!formExpense.splitColumn" @click="submitExpenseSplit">执行拆分</el-button>
    <el-button style="margin-left: 8px" @click="refreshExpense">
      <el-icon><Refresh /></el-icon> 刷新
    </el-button>
    <el-card v-if="expenseDownloadName" style="margin-top: 16px">
      <span>处理完成：{{ expenseDownloadName }}</span>
      <el-button type="primary" style="margin-left: 12px" @click="downloadFile(expenseDownloadName)">下载文件</el-button>
    </el-card>

    <el-alert
      v-if="message && !expenseDownloadName"
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
import { postAsync } from '../../../utils/asyncTask'

export default {
  name: 'CwgjxBxcf',
  data() {
    return {
      loading: false,
      message: '',
      messageType: 'info',
      formExpense: {
        file: null,
        splitColumn: ''
      },
      expensePreviewing: false,
      expensePreviewCols: [],
      expenseDownloadName: ''
    }
  },
  methods: {
    handleExpenseFileChange(file) {
      this.formExpense.file = file.raw
    },
    showMessage(text, type = 'info') {
      this.message = text
      this.messageType = type
    },
    submitExpensePreview() {
      if (!this.formExpense.file) {
        this.showMessage('请先选择文件', 'warning')
        return
      }
      this.expensePreviewing = true
      this.message = ''
      this.expensePreviewCols = []
      const formData = new FormData()
      formData.append('file', this.formExpense.file)
      axios.post('/api/finance/toolbox/expense-preview', formData).then(res => {
        if (res.data.success) {
          this.expensePreviewCols = res.data.columns || []
          this.showMessage('预览处理完成，共 ' + this.expensePreviewCols.length + ' 列', 'success')
        } else {
          this.showMessage(res.data.message || '预览处理失败', 'error')
        }
      }).catch(err => {
        this.showMessage(err.response?.data?.message || err.message, 'error')
      }).finally(() => {
        this.expensePreviewing = false
      })
    },
    submitExpenseSplit() {
      if (!this.formExpense.file) {
        this.showMessage('请先选择文件', 'warning')
        return
      }
      if (!this.formExpense.splitColumn) {
        this.showMessage('请先选择拆分列', 'warning')
        return
      }
      const formData = new FormData()
      formData.append('file', this.formExpense.file)
      formData.append('splitColumn', this.formExpense.splitColumn)
      this.loading = true
      this.message = ''
      postAsync('/api/finance/toolbox/expense-split', formData).then(res => {
        if (res.data.success) {
          this.showMessage(res.data.message, 'success')
          this.expenseDownloadName = res.data.fileName
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
    refreshExpense() {
      this.loading = false
      this.expensePreviewing = false
      this.expensePreviewCols = []
      this.formExpense.splitColumn = ''
      this.showMessage('报销拆分已刷新', 'success')
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
