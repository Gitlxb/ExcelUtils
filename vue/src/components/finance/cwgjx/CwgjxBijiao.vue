<template>
  <el-card class="func-card">
    <div class="page-container">
    <div class="func-page-header">
      <h2 class="func-page-title">数据比较</h2>
      <p class="func-page-desc">上传申请表和明细表，按身份信息自动比对数据。</p>
    </div>

    <div class="tb-section-title">上传文件</div>
    <el-row :gutter="16">
      <el-col :span="12">
        <div class="tb-upload-label">文件1（申请表）</div>
        <el-upload
          ref="uploadFile1"
          class="tb-upload"
          drag
          :auto-upload="false"
          :limit="1"
          :on-change="(f) => handleFileChange('file1', f, 'compare')"
          accept=".xlsx"
        >
          <el-icon class="tb-upload-icon"><UploadFilled /></el-icon>
          <div>将文件拖到此处，或 <em>点击选择</em></div>
          <div v-if="formCompare.file1" class="tb-file-name">{{ formCompare.file1.name }}</div>
        </el-upload>
        <el-form label-position="top" style="margin-top: 8px">
          <el-form-item label="工作表">
            <el-select v-model="formCompare.sheet1Name" placeholder="请选择工作表" style="width:100%">
              <el-option v-for="s in compareSheets1" :key="s" :label="s" :value="s" />
            </el-select>
          </el-form-item>
        </el-form>
      </el-col>
      <el-col :span="12">
        <div class="tb-upload-label">文件2（明细表）</div>
        <el-upload
          ref="uploadFile2"
          class="tb-upload"
          drag
          :auto-upload="false"
          :limit="1"
          :on-change="(f) => handleFileChange('file2', f, 'compare')"
          accept=".xlsx"
        >
          <el-icon class="tb-upload-icon"><UploadFilled /></el-icon>
          <div>将文件拖到此处，或 <em>点击选择</em></div>
          <div v-if="formCompare.file2" class="tb-file-name">{{ formCompare.file2.name }}</div>
        </el-upload>
        <el-form label-position="top" style="margin-top: 8px">
          <el-form-item label="工作表">
            <el-select v-model="formCompare.sheet2Name" placeholder="请选择工作表" style="width:100%">
              <el-option v-for="s in compareSheets2" :key="s" :label="s" :value="s" />
            </el-select>
          </el-form-item>
        </el-form>
      </el-col>
    </el-row>
    <div class="tb-section-title" style="margin-top: 16px">筛选条件</div>
    <el-form :inline="true" label-width="80px">
      <el-form-item label="开始日期">
        <el-date-picker v-model="formCompare.startDate" type="date" placeholder="开始日期" value-format="YYYY-MM-DD" style="width:150px" />
      </el-form-item>
      <el-form-item label="结束日期">
        <el-date-picker v-model="formCompare.endDate" type="date" placeholder="结束日期" value-format="YYYY-MM-DD" style="width:150px" />
      </el-form-item>
    </el-form>
    <div class="tb-section-title">操作</div>
    <el-button type="success" :loading="loading" @click="submitDataCompare">开始比较</el-button>
    <el-button style="margin-left: 8px" @click="refreshCompare">
      <el-icon><Refresh /></el-icon> 刷新
    </el-button>
    <!-- 数据比较 预览 + 下载 -->
    <div v-if="comparePreview" style="margin-top: 16px">
      <el-card header="对比结果预览">
        <el-table :data="comparePreviewRows" border size="small" max-height="400px" style="width:100%">
          <el-table-column prop="name" label="姓名" min-width="80" />
          <el-table-column prop="amountLeft" label="预支金额" min-width="90" align="right">
            <template #default="{ row }">{{ row.amountLeft.toFixed(2) }}</template>
          </el-table-column>
          <el-table-column prop="idCard" label="身份证号" min-width="160" />
          <el-table-column prop="result" label="比较结果" min-width="90" align="center">
            <template #default="{ row }">
              <el-tag :type="row.result === 'TRUE' ? 'success' : 'danger'" size="small">{{ row.result }}</el-tag>
            </template>
          </el-table-column>
          <el-table-column prop="diff" label="差额" min-width="80" align="right">
            <template #default="{ row }">{{ row.diff.toFixed(2) }}</template>
          </el-table-column>
          <el-table-column prop="nameRight" label="姓名(右)" min-width="80" />
          <el-table-column prop="amountRight" label="预支数额(右)" min-width="100" align="right">
            <template #default="{ row }">{{ row.amountRight.toFixed(2) }}</template>
          </el-table-column>
          <el-table-column prop="remark" label="备注" min-width="120" />
          <el-table-column prop="unit" label="运营中心" min-width="90" />
        </el-table>
        <el-descriptions :column="3" border style="margin-top: 12px">
          <el-descriptions-item label="总记录数">{{ comparePreview.totalCount }}</el-descriptions-item>
          <el-descriptions-item label="相符(TRUE)"><span style="color:green">{{ comparePreview.trueCount }}</span></el-descriptions-item>
          <el-descriptions-item label="不符(FALSE)"><span style="color:red">{{ comparePreview.falseCount }}</span></el-descriptions-item>
          <el-descriptions-item label="申请表总金额">{{ comparePreview.sumLeft.toFixed(2) }}</el-descriptions-item>
          <el-descriptions-item label="差额总计">{{ comparePreview.sumDiff.toFixed(2) }}</el-descriptions-item>
          <el-descriptions-item label="明细表总金额">{{ comparePreview.sumRight.toFixed(2) }}</el-descriptions-item>
        </el-descriptions>
        <el-alert
          v-if="comparePreview.falseCount > 0"
          :title="'有 ' + comparePreview.falseCount + ' 条记录不符，请检查'"
          type="warning"
          :closable="false"
          style="margin-top: 12px"
        />
        <el-button type="primary" style="margin-top: 12px" @click="downloadFile(compareFileName)">下载完整结果</el-button>
      </el-card>
    </div>

    <el-alert
      v-if="message && !comparePreview"
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
  name: 'CwgjxBijiao',
  data() {
    return {
      loading: false,
      message: '',
      messageType: 'info',
      formCompare: {
        file1: null,
        file2: null,
        sheet1Name: '',
        sheet2Name: '',
        startDate: '',
        endDate: ''
      },
      compareSheets1: [],
      compareSheets2: [],
      comparePreview: null,
      comparePreviewRows: [],
      compareFileName: ''
    }
  },
  methods: {
    handleFileChange(key, file, type) {
      const raw = file.raw
      if (type === 'compare') {
        this.formCompare[key] = raw
        const target = key === 'file1' ? 'compareSheets1' : 'compareSheets2'
        const sheetKey = key === 'file1' ? 'sheet1Name' : 'sheet2Name'
        this.loadSheets(raw, target, sheetKey)
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
    submitDataCompare() {
      if (!this.formCompare.file1 || !this.formCompare.file2) {
        this.showMessage('请先选择两个文件', 'warning')
        return
      }
      if (!this.formCompare.sheet1Name || !this.formCompare.sheet2Name) {
        this.showMessage('请先选择工作表', 'warning')
        return
      }
      this.loading = true
      this.message = ''
      const formData = new FormData()
      formData.append('file1', this.formCompare.file1)
      formData.append('file2', this.formCompare.file2)
      formData.append('sheet1Name', this.formCompare.sheet1Name)
      formData.append('sheet2Name', this.formCompare.sheet2Name)
      formData.append('startDate', this.formCompare.startDate)
      formData.append('endDate', this.formCompare.endDate)
      postAsync('/api/finance/toolbox/data-compare', formData, {
        headers: { 'Content-Type': 'multipart/form-data' }
      }).then(res => {
        if (res.data.success) {
          this.showMessage(res.data.message, 'success')
          this.comparePreview = res.data.preview
          this.compareFileName = res.data.fileName
          this.comparePreviewRows = res.data.previewRows || []
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
    refreshCompare() {
      this.loading = false
      this.showMessage('数据比较已刷新', 'success')
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
</style>
