<template>
  <el-card class="func-card">
    <div class="page-container">
    <div class="func-page-header">
      <h2 class="func-page-title">稳岗率分析</h2>
      <p class="func-page-desc">
        根据员工信息汇总表，按月分析项目驻场、项目经理、稳岗率、奖励汇总、基准数，生成5张工作表。
      </p>
    </div>

    <el-form label-width="140px">
      <el-form-item label="Excel文件">
        <el-upload
          ref="uploadWgl"
          drag
          :auto-upload="false"
          :limit="1"
          :on-change="handleWglFileChange"
          accept=".xlsx,.xls"
        >
          <el-icon class="el-icon--upload"><UploadFilled /></el-icon>
          <div class="el-upload__text">将文件拖到此处，或 <em>点击选择</em></div>
        </el-upload>
      </el-form-item>
      <el-form-item label="工作表名称">
        <el-select v-model="formWgl.sheetName" placeholder="请选择工作表" style="width: 220px">
          <el-option v-for="s in wglSheets" :key="s" :label="s" :value="s" />
        </el-select>
      </el-form-item>
      <el-form-item label="起始月份">
        <el-config-provider :locale="customLocale">
          <el-date-picker
            v-model="formWgl.startMonth"
            type="month"
            format="YYYY/MM"
            value-format="YYYY/MM"
            placeholder="选择起始月份"
            style="width: 220px"
          />
        </el-config-provider>
      </el-form-item>
      <el-form-item label="结束月份">
        <el-config-provider :locale="customLocale">
          <el-date-picker
            v-model="formWgl.endMonth"
            type="month"
            format="YYYY/MM"
            value-format="YYYY/MM"
            placeholder="选择结束月份"
            style="width: 220px"
          />
        </el-config-provider>
      </el-form-item>
      <el-form-item label="保留原始工作表">
        <el-switch v-model="formWgl.copyOriginal" active-text="是" inactive-text="否" />
      </el-form-item>
      <el-form-item>
        <el-button type="success" :loading="loading" @click="submitWgl">开始分析</el-button>
      </el-form-item>
    </el-form>

    <!-- 操作日志 -->
    <div v-if="wglLogs.length" class="log-panel">
      <div class="log-title">处理日志</div>
      <div v-for="(log, idx) in wglLogs" :key="idx" class="log-line">{{ log }}</div>
    </div>

    <el-alert
      title="说明：生成项目驻场、项目经理、稳岗率、奖励汇总、基准数 5张工作表。"
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
import { postAsync } from '../../../utils/asyncTask'
import zhCn from 'element-plus/dist/locale/zh-cn.mjs'

const customLocale = {
  ...zhCn,
  el: {
    ...zhCn.el,
    datepicker: {
      ...zhCn.el.datepicker,
      month1: '1月',
      month2: '2月',
      month3: '3月',
      month4: '4月',
      month5: '5月',
      month6: '6月',
      month7: '7月',
      month8: '8月',
      month9: '9月',
      month10: '10月',
      month11: '11月',
      month12: '12月',
      months: {
        ...zhCn.el.datepicker.months,
        jan: '1月', feb: '2月', mar: '3月', apr: '4月',
        may: '5月', jun: '6月', jul: '7月', aug: '8月',
        sep: '9月', oct: '10月', nov: '11月', dec: '12月'
      }
    }
  }
}

export default {
  name: 'JxkhWgl',
  data() {
    return {
      loading: false,
      message: '',
      messageType: 'info',
      downloadFileName: '',
      wglFile: null,
      wglSheets: [],
      wglLogs: [],
      customLocale: customLocale,
      formWgl: {
        sheetName: '',
        startMonth: '',
        endMonth: '',
        copyOriginal: true
      }
    }
  },
  methods: {
    handleWglFileChange(file) {
      this.wglFile = file.raw
      this.wglSheets = []
      this.formWgl.sheetName = ''
      if (file.raw) {
        this.loadWglSheets(file.raw)
      }
    },
    loadWglSheets(file) {
      const formData = new FormData()
      formData.append('file', file)
      axios.post('/api/finance/toolbox/sheets', formData).then(res => {
        if (res.data.success) {
          this.wglSheets = res.data.sheets || []
          if (this.wglSheets.length > 0) {
            this.formWgl.sheetName = this.wglSheets[0]
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
    submitWgl() {
      if (!this.wglFile) {
        this.showMessage('请选择Excel文件', 'warning')
        return
      }
      if (!this.formWgl.startMonth || !this.formWgl.endMonth) {
        this.showMessage('请填写起始和结束月份', 'warning')
        return
      }
      this.wglLogs = []
      const formData = new FormData()
      formData.append('file', this.wglFile)
      if (this.formWgl.sheetName) formData.append('sheetName', this.formWgl.sheetName)
      formData.append('startMonth', this.formWgl.startMonth)
      formData.append('endMonth', this.formWgl.endMonth)
      formData.append('copyOriginal', this.formWgl.copyOriginal)
      this.post('/api/finance/jxkh/wgl', formData, (res) => {
        if (res.data.logs && res.data.logs.length) {
          this.wglLogs = res.data.logs
        }
      })
    },
    post(endpoint, formData, onSuccess) {
      this.loading = true
      this.message = ''
      this.downloadFileName = ''
      postAsync(endpoint, formData).then(res => {
        if (res.data.success) {
          this.showMessage(res.data.message, 'success')
          this.downloadFileName = res.data.fileName
          if (typeof onSuccess === 'function') onSuccess(res)
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

.log-panel {
  margin-top: 12px;
  padding: 10px 14px;
  background: #f5f7fa;
  border: 1px solid #e4e7ed;
  border-radius: 4px;
  font-size: 13px;
  color: #606266;
  max-height: 300px;
  overflow-y: auto;
}

.log-title {
  font-weight: bold;
  margin-bottom: 6px;
  color: #303133;
}

.log-line {
  line-height: 1.8;
  word-break: break-all;
}
</style>
