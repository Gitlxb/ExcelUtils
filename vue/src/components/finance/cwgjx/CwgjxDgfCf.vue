<template>
  <el-card class="func-card">
    <div class="page-container">
    <div class="func-page-header">
      <h2 class="func-page-title">代工费拆分_批量处理</h2>
      <p class="func-page-desc">批量处理代工费表格，结合发放记录自动拆分并汇总。支持多文件批量上传。</p>
    </div>

    <div class="tb-section-title">上传文件</div>
    <el-row :gutter="16">
      <el-col :span="12">
        <div class="tb-upload-label">
          代工费表格（支持多选）
          <el-tag v-if="formLabor.files1.length > 0" size="small" type="info" style="margin-left: 8px">{{ formLabor.files1.length }} 个文件</el-tag>
        </div>
        <el-upload
          ref="uploadLabor1"
          class="tb-upload"
          drag
          :auto-upload="false"
          :multiple="true"
          :on-change="(f) => handleLaborFileChange('file1', f)"
          accept=".xlsx"
        >
          <el-icon class="tb-upload-icon"><UploadFilled /></el-icon>
          <div>将文件拖到此处，或 <em>点击选择</em></div>
        </el-upload>
        <el-form label-position="top" style="margin-top: 8px">
          <el-form-item label="工作表名">
            <el-select v-model="formLabor.sheet1Name" placeholder="请选择工作表" style="width:100%" :disabled="laborSheets1.length === 0">
              <el-option v-for="s in laborSheets1" :key="s" :label="s" :value="s" />
            </el-select>
          </el-form-item>
        </el-form>
      </el-col>
      <el-col :span="12">
        <div class="tb-upload-label">发放记录表格</div>
        <el-upload
          ref="uploadLabor2"
          class="tb-upload"
          drag
          :auto-upload="false"
          :limit="1"
          :on-change="(f) => handleLaborFileChange('file2', f)"
          accept=".xlsx"
        >
          <el-icon class="tb-upload-icon"><UploadFilled /></el-icon>
          <div>将文件拖到此处，或 <em>点击选择</em></div>
          <div v-if="formLabor.file2" class="tb-file-name">{{ formLabor.file2.name }}</div>
        </el-upload>
        <el-form label-position="top" style="margin-top: 8px">
          <el-form-item label="工作表A（拆分写入）">
            <el-select v-model="formLabor.sheet2Name" placeholder="请选择工作表" style="width:100%">
              <el-option v-for="s in laborSheets2" :key="s" :label="s" :value="s" />
            </el-select>
          </el-form-item>
          <el-form-item label="工作表B（汇总参考）">
            <el-select v-model="formLabor.sheetBName" placeholder="请选择工作表" style="width:100%">
              <el-option v-for="s in laborSheets2" :key="s" :label="s" :value="s" />
            </el-select>
          </el-form-item>
        </el-form>
      </el-col>
    </el-row>
    <div class="tb-section-title" style="margin-top: 16px">操作</div>
    <el-button type="success" :loading="loading" @click="submitLaborSplit">开始拆分</el-button>
    <el-button style="margin-left: 8px" @click="refreshLabor">
      <el-icon><Refresh /></el-icon> 刷新
    </el-button>
    <el-card v-if="laborDownloadName" style="margin-top: 16px">
      <div style="margin-bottom: 8px">处理完成，请下载以下文件：</div>
      <div style="display:flex;gap:12px;flex-wrap:wrap;margin-bottom:8px">
        <el-tag type="primary">发放记录表（含拆分结果）</el-tag>
        <el-button type="primary" size="small" @click="downloadFile(laborDownloadName)">下载文件</el-button>
      </div>
      <div v-for="(name, idx) in laborSummaryFileNames" :key="idx" style="display:flex;gap:12px;flex-wrap:wrap;margin-bottom:8px">
        <el-tag type="success">{{ name }}</el-tag>
        <el-button type="success" size="small" @click="downloadFile(name)">下载文件</el-button>
      </div>
    </el-card>
    <!-- 代工费拆分操作日志 -->
    <el-card v-if="laborLogs.length > 0" header="处理日志" style="margin-top: 16px">
      <div class="labor-log-box">
        <div v-for="(log, idx) in laborLogs" :key="idx" class="labor-log-line">{{ log }}</div>
      </div>
    </el-card>

    <el-alert
      v-if="message && !laborDownloadName"
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
  name: 'CwgjxDgfCf',
  data() {
    return {
      loading: false,
      message: '',
      messageType: 'info',
      formLabor: {
        files1: [],
        file2: null,
        sheet1Name: '',
        sheet2Name: '',
        sheetBName: ''
      },
      laborSheets1: [],
      laborSheets2: [],
      laborLogs: [],
      laborSplitting: false,
      laborDownloadName: '',
      laborSummaryFileNames: []
    }
  },
  methods: {
    handleLaborFileChange(key, file) {
      if (key === 'file1') {
        const raw = file.raw
        if (!this.formLabor.files1.find(f => f.name === raw.name && f.size === raw.size)) {
          this.formLabor.files1.push(raw)
        }
        if (this.formLabor.files1.length === 1) {
          this.loadSheets(raw, 'laborSheets1', 'sheet1Name')
        }
      } else {
        this.formLabor.file2 = file.raw
        this.loadSheets(file.raw, 'laborSheets2', 'sheet2Name')
      }
    },
    removeLaborFile(idx) {
      this.formLabor.files1.splice(idx, 1)
      if (this.formLabor.files1.length === 0) {
        this.laborSheets1 = []
        this.formLabor.sheet1Name = ''
      }
    },
    loadSheets(file, listKey, selectKey) {
      const formData = new FormData()
      formData.append('file', file)
      axios.post('/api/finance/toolbox/sheets', formData).then(res => {
        if (res.data.success) {
          this[listKey] = res.data.sheets || []
          if (this[listKey].length > 0 && !this[selectKey]) {
            this[selectKey] = this[listKey][0]
          }
        } else {
          this.showMessage(res.data.message || '读取工作表失败', 'warning')
        }
      }).catch(err => {
        this.showMessage(err.response?.data?.message || err.message, 'error')
      })
    },
    showMessage(text, type = 'info') {
      this.message = text
      this.messageType = type
    },
    submitLaborSplit() {
      if (this.formLabor.files1.length === 0 || !this.formLabor.file2) {
        this.showMessage('请先选择代工费表格（至少1个）和发放记录表格', 'warning')
        return
      }
      if (!this.formLabor.sheet1Name || !this.formLabor.sheet2Name || !this.formLabor.sheetBName) {
        this.showMessage('请先填写工作表名称', 'warning')
        return
      }
      this.laborLogs = []
      this.laborSplitting = true
      this.addLaborLog('代工费表格：共 ' + this.formLabor.files1.length + ' 个文件')
      this.formLabor.files1.forEach((f, i) => {
        this.addLaborLog('  [' + (i + 1) + '] ' + f.name)
      })
      this.addLaborLog('发放记录表格：' + this.formLabor.file2.name)
      this.addLaborLog('  工作簿（A）：' + this.formLabor.sheet2Name)
      this.addLaborLog('  工作簿（B）：' + this.formLabor.sheetBName)
      this.addLaborLog('')

      const formData = new FormData()
      this.formLabor.files1.forEach(f => formData.append('file1', f))
      formData.append('file2', this.formLabor.file2)
      formData.append('sheet1Name', this.formLabor.sheet1Name)
      formData.append('sheet2Name', this.formLabor.sheet2Name)
      formData.append('sheetBName', this.formLabor.sheetBName)

      this.loading = true
      this.message = ''
      axios.post('/api/finance/toolbox/labor-split-batch', formData).then(res => {
        if (res.data.success) {
          this.addLaborLog('  批量处理完成，共处理 ' + this.formLabor.files1.length + ' 个文件')
          this.addLaborLog('  汇总记录总数：' + (res.data.summaryCount || '未知'))
          this.addLaborLog('  拆分行总数：' + (res.data.splitCount || '未知'))
          this.addLaborLog('')
          const step4Details = res.data.step4Details || []
          step4Details.forEach(detail => this.addLaborLog(detail))
          const step5Details = res.data.step5Details || []
          step5Details.forEach(detail => this.addLaborLog(detail))
          const step5Result = res.data.step5Result || []
          step5Result.forEach(line => this.addLaborLog(line))
          this.addLaborLog('')
          this.addLaborLog('全部处理完毕！')
          this.showMessage(res.data.message, 'success')
          this.laborDownloadName = res.data.fileName
          this.laborSummaryFileNames = res.data.summaryFileNames || []
        } else {
          this.addLaborLog('[错误] ' + (res.data.message || '处理失败'))
          this.showMessage(res.data.message || '处理失败', 'error')
        }
      }).catch(err => {
        const msg = err.response?.data?.message || err.message
        this.addLaborLog('[错误] ' + msg)
        this.showMessage(msg, 'error')
      }).finally(() => {
        this.loading = false
        this.laborSplitting = false
      })
    },
    addLaborLog(msg) {
      this.laborLogs.push(msg)
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
    refreshLabor() {
      this.loading = false
      this.laborSplitting = false
      this.showMessage('代工费拆分已刷新', 'success')
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

.tb-upload-label {
  font-size: 13px;
  color: #606266;
  margin-bottom: 6px;
}

.tb-file-name {
  font-size: 12px;
  color: #67C23A;
  margin-top: 4px;
}

.labor-log-box {
  max-height: 300px;
  overflow-y: auto;
  background: #f5f7fa;
  border-radius: 4px;
  padding: 12px;
  font-family: 'Consolas', 'Monaco', monospace;
  font-size: 13px;
  line-height: 1.8;
}

.labor-log-line {
  color: #303133;
  white-space: pre-wrap;
  word-break: break-all;
}

.labor-log-line:nth-child(odd) {
  background: rgba(64, 158, 255, 0.04);
}
</style>
