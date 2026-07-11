<template>
  <el-card class="func-card">
    <!-- 页面标题区域 -->
    <div class="func-page-header">
      <h2 class="func-page-title">银行流水专业修改</h2>
      <p class="func-page-desc">
        基于配置表的多规则银行流水处理工具。支持多种内容格式的自动识别、字段映射和数据转换，适用于复杂银行流水的批量修改场景。
      </p>
    </div>

    <!-- 步骤1：上传文件并检测配置 -->
    <div v-if="step === 1" class="step-content">
      <h4>步骤1：上传Excel文件并读取配置</h4>
      <el-upload
        ref="upload1"
        drag
        accept=".xlsx"
        :auto-upload="false"
        :show-file-list="true"
        :limit="1"
        :on-change="handleFile1"
        :on-exceed="onExceed"
      >
        <el-icon class="el-icon--upload"><UploadFilled /></el-icon>
        <div class="el-upload__text">拖拽或点击上传Excel文件</div>
      </el-upload>
      <el-button type="primary" :loading="loading" :disabled="!file1" @click="detectConfig">读取配置</el-button>
      <div v-if="detectError" class="result-err">{{ detectError }}</div>
    </div>

    <!-- 步骤2：选择参数 -->
    <div v-if="step === 2" class="step-content">
      <h4>步骤2：配置处理参数</h4>
      <el-form label-width="160px" class="param-form">
        <el-form-item label="往来款公司（AS列）" required>
          <el-select v-model="selectedAs" placeholder="请选择">
            <el-option v-for="opt in asOptions" :key="opt" :label="opt" :value="opt" />
          </el-select>
        </el-form-item>
        <el-form-item label="替换'锦途'的公司（J列）" required>
          <el-select v-model="selectedCompany" placeholder="请选择">
            <el-option v-for="opt in jOptions" :key="opt" :label="opt" :value="opt" />
          </el-select>
        </el-form-item>
        <el-form-item label="内部转账筛选名称">
          <el-input v-model="inputNameFilter" placeholder="输入用于匹配收(付)方名称的内容" />
        </el-form-item>
        <el-form-item label="银行账号后四位">
          <el-input v-model="inputSummaryText" placeholder="用于摘要输出的后四位" />
        </el-form-item>
        <el-form-item label="锦途代发月薪摘要">
          <el-input v-model="inputSummary" placeholder="例如：锦途代发x月xx员工月薪" />
        </el-form-item>
      </el-form>
      <el-button type="primary" :loading="loading" @click="processFile">开始处理</el-button>
      <el-button @click="step = 1">重新上传</el-button>
    </div>

    <!-- 步骤3：结果 -->
    <div v-if="step === 3" class="step-content">
      <div v-if="downloadUrl" class="result-section">
        <el-alert title="处理完成！" type="success" :closable="false" show-icon />
        <el-button type="primary" @click="downloadFile" class="download-btn">
          <el-icon><Download /></el-icon> 下载处理后的文件
        </el-button>
      </div>
      <div v-if="errorMsg" class="result-err">{{ errorMsg }}</div>
      <el-button @click="reset">重新开始</el-button>
    </div>

    <!-- 进度 -->
    <div v-if="loading" class="progress-section">
      <el-progress :percentage="progressPercent" :status="progressStatus" />
    </div>
  </el-card>
</template>

<script>
import axios from 'axios'

export default {
  name: 'BankFlowCard',
  data() {
    return {
      step: 1,
      file1: null,
      asOptions: [],
      jOptions: [],
      selectedAs: '',
      selectedCompany: '',
      inputNameFilter: '',
      inputSummaryText: '',
      inputSummary: '',
      loading: false,
      progressPercent: 0,
      progressStatus: '',
      downloadUrl: '',
      errorMsg: '',
      detectError: ''
    }
  },
  methods: {
    handleFile1() {
      this.$nextTick(() => {
        const upload = this.$refs.upload1
        const files = upload?.uploadFiles
        this.file1 = files?.length ? files[0].raw : null
        this.detectError = ''
      })
    },
    onExceed() { this.$message.warning('只能选1个文件') },
    detectConfig() {
      if (!this.file1) { this.$message.warning('请先选择文件'); return }
      this.loading = true
      this.detectError = ''
      const formData = new FormData()
      formData.append('file', this.file1)
      // FormData 上传不能手动设置 Content-Type，浏览器会自动设置并带上 boundary
      axios.post('/api/finance/bank-flow/detect', formData).then(res => {
        if (res.data.success) {
          this.asOptions = res.data.asOptions || []
          this.jOptions = res.data.jOptions || []
          this.step = 2
          if (this.asOptions.length) this.selectedAs = this.asOptions[0]
          if (this.jOptions.length) this.selectedCompany = this.jOptions[0]
        } else {
          this.detectError = res.data.message || '检测失败'
        }
      }).catch(err => {
        this.detectError = err.response?.data?.message || err.message
      }).finally(() => { this.loading = false })
    },
    processFile() {
      if (!this.selectedAs || !this.selectedCompany) {
        this.$message.warning('请选择往来款公司和替换公司名称')
        return
      }
      this.loading = true
      this.errorMsg = ''
      this.downloadUrl = ''

      const timer = setInterval(() => {
        if (this.progressPercent < 90) this.progressPercent += 10
      }, 500)

      const formData = new FormData()
      formData.append('file', this.file1)
      formData.append('selectedAs', this.selectedAs)
      formData.append('selectedCompany', this.selectedCompany)
      if (this.inputNameFilter) formData.append('inputNameFilter', this.inputNameFilter)
      if (this.inputSummaryText) formData.append('inputSummaryText', this.inputSummaryText)
      if (this.inputSummary) formData.append('inputSummary', this.inputSummary)

      // FormData 上传不能手动设置 Content-Type，浏览器会自动设置并带上 boundary
      axios.post('/api/finance/bank-flow/process', formData).then(res => {
        clearInterval(timer)
        this.progressPercent = 100
        if (res.data.success) {
          this.progressStatus = 'success'
          this.downloadUrl = `/api/finance/download/${res.data.fileName}`
          this.step = 3
        } else {
          this.progressStatus = 'exception'
          this.errorMsg = res.data.message || '处理失败'
          this.step = 3
        }
      }).catch(err => {
        clearInterval(timer)
        this.progressStatus = 'exception'
        this.errorMsg = err.response?.data?.message || err.message
        this.step = 3
      }).finally(() => { this.loading = false })
    },
    downloadFile() {
      if (this.downloadUrl) window.location.href = this.downloadUrl
    },
    reset() {
      this.step = 1
      this.file1 = null
      this.asOptions = []
      this.jOptions = []
      this.selectedAs = ''
      this.selectedCompany = ''
      this.inputNameFilter = ''
      this.inputSummaryText = ''
      this.inputSummary = ''
      this.downloadUrl = ''
      this.errorMsg = ''
      this.progressPercent = 0
      this.progressStatus = ''
      this.$refs.upload1?.clearFiles?.()
    }
  }
}
</script>

<style scoped src="../../styles/components/BankFlowCard.css"></style>

<style scoped src="../../styles/components/FuncPageHeader.css"></style>
