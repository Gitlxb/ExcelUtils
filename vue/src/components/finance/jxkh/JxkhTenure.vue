<template>
  <el-card class="func-card">
    <div class="page-container">
      <div class="func-page-header">
        <h2 class="func-page-title">员工在职月份计算</h2>
        <p class="func-page-desc">
          根据员工信息表（状态 / 姓名 / 身份证号码 / 入职日期 / 离职日期），按人计算在职区间并扣除短离职，
          逐月展开「在职月份 / 在职天数 / 总在职天数」，生成单张工作表「在职月份结果」。
        </p>
      </div>

      <el-form label-width="140px">
        <el-form-item label="Excel文件">
          <el-upload
            ref="uploadTenure"
            drag
            :auto-upload="false"
            :limit="1"
            :on-change="handleTenureFileChange"
            accept=".xlsx,.xls"
          >
            <el-icon class="el-icon--upload"><UploadFilled /></el-icon>
            <div class="el-upload__text">将文件拖到此处，或 <em>点击选择</em></div>
          </el-upload>
        </el-form-item>

        <el-form-item label="工作表名称">
          <el-select v-model="formTenure.sheetName" placeholder="请选择工作表" style="width: 220px">
            <el-option v-for="s in tenureSheets" :key="s" :label="s" :value="s" />
          </el-select>
        </el-form-item>

        <el-form-item label="起始日期">
          <el-config-provider :locale="customLocale">
            <el-date-picker
              v-model="formTenure.startDate"
              type="date"
              format="YYYY-MM-DD"
              value-format="YYYY-MM-DD"
              placeholder="如 2026-07-01"
              style="width: 220px"
              clearable
            />
          </el-config-provider>
        </el-form-item>

        <el-form-item label="截止日期">
          <el-config-provider :locale="customLocale">
            <el-date-picker
              v-model="formTenure.endDate"
              type="date"
              format="YYYY-MM-DD"
              value-format="YYYY-MM-DD"
              placeholder="如 2026-07-31"
              style="width: 220px"
              clearable
            />
          </el-config-provider>
        </el-form-item>

        <el-form-item label="合并总在职天数">
          <el-switch v-model="formTenure.mergeTotalDays" active-text="是" inactive-text="否" />
        </el-form-item>

        <el-form-item>
          <el-button type="success" :loading="loading" @click="submitTenure">开始计算</el-button>
        </el-form-item>
      </el-form>

      <!-- 操作日志 -->
      <div v-if="tenureLogs.length" class="log-panel">
        <div class="log-title">处理日志</div>
        <div v-for="(log, idx) in tenureLogs" :key="idx" class="log-line">{{ log }}</div>
      </div>

      <el-alert
        title="说明：输出单张工作表『在职月份结果』，含原表列 + 在职月份 + 在职天数 + 总在职天数（按人合并）。"
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

function currentYearMonth() {
  const d = new Date()
  const m = String(d.getMonth() + 1).padStart(2, '0')
  return d.getFullYear() + '-' + m + '-01'
}

function currentMonthEnd() {
  const d = new Date()
  const lastDay = new Date(d.getFullYear(), d.getMonth() + 1, 0).getDate()
  const m = String(d.getMonth() + 1).padStart(2, '0')
  return d.getFullYear() + '-' + m + '-' + String(lastDay).padStart(2, '0')
}

export default {
  name: 'JxkhTenure',
  data() {
    return {
      loading: false,
      message: '',
      messageType: 'info',
      downloadFileName: '',
      tenureFile: null,
      tenureSheets: [],
      tenureLogs: [],
      customLocale: customLocale,
      formTenure: {
        sheetName: '',
        startDate: currentYearMonth(),
        endDate: currentMonthEnd(),
        mergeTotalDays: false
      }
    }
  },
  methods: {
    handleTenureFileChange(file) {
      this.tenureFile = file.raw
      this.tenureSheets = []
      this.formTenure.sheetName = ''
      if (file.raw) {
        this.loadTenureSheets(file.raw)
      }
    },
    loadTenureSheets(file) {
      const formData = new FormData()
      formData.append('file', file)
      axios.post('/api/finance/toolbox/sheets', formData).then(res => {
        if (res.data.success) {
          this.tenureSheets = res.data.sheets || []
          if (this.tenureSheets.length > 0) {
            this.formTenure.sheetName = this.tenureSheets[0]
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
    submitTenure() {
      if (!this.tenureFile) {
        this.showMessage('请选择Excel文件', 'warning')
        return
      }
      if (!this.formTenure.startDate) {
        this.showMessage('请选择起始日期', 'warning')
        return
      }
      if (!this.formTenure.endDate) {
        this.showMessage('请选择截止日期', 'warning')
        return
      }
      if (this.formTenure.startDate > this.formTenure.endDate) {
        this.showMessage('起始日期不能晚于截止日期', 'warning')
        return
      }
      this.tenureLogs = []
      const formData = new FormData()
      formData.append('file', this.tenureFile)
      if (this.formTenure.sheetName) formData.append('sheetName', this.formTenure.sheetName)
      formData.append('startDate', this.formTenure.startDate)
      formData.append('endDate', this.formTenure.endDate)
      formData.append('mergeTotalDays', this.formTenure.mergeTotalDays)
      this.post('/api/finance/jxkh/tenure', formData, (res) => {
        if (res.data.logs && res.data.logs.length) {
          this.tenureLogs = res.data.logs
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

<style scoped src="../../../styles/components/JxkhTenure.css"></style>
