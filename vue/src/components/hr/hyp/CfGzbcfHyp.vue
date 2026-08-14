<template>
  <el-card class="func-card">
    <div class="page-container">
    <div class="func-page-header">
      <h2 class="func-page-title">代工费拆分（增强版）</h2>
      <p class="func-page-desc">
        根据代工费发放记录表和对应的发放记录表，按收款人姓名拆分各供应商的打款金额。增强版：完整功能，含社保、公积金、总计行。
      </p>
    </div>

    <!-- 1. 上传文件 -->
    <div class="section">
      <div class="section-title">1. 上传文件</div>
      <el-upload
        ref="uploadRef"
        drag
        accept=".xlsx,.xlsm"
        :auto-upload="false"
        :limit="1"
        :on-change="handleFileChange"
        :on-exceed="handleExceed"
        class="gzbcf-upload"
      >
        <el-icon class="el-icon--upload"><UploadFilled /></el-icon>
        <div class="el-upload__text">
          拖拽文件到此处或 <em>点击上传</em>
        </div>
        <div class="el-upload__tip">支持 .xlsx / .xlsm 格式</div>
      </el-upload>
      <div v-if="fileName" class="file-name">已选择：{{ fileName }}</div>
      <el-button
        v-if="fileRaw && !sheetList.length"
        type="primary"
        @click="loadSheets"
        :loading="loadingSheets"
        style="margin-top:12px"
      >识别工作表</el-button>
    </div>

    <!-- 2. 选择工作表 -->
    <div class="section">
      <div class="section-title">2. 选择工作表</div>
      <div class="form-row">
        <el-select
          v-model="selectedSheet"
          placeholder="请选择工作表"
          style="width:260px"
          :disabled="!sheetList.length"
          @change="onSheetChange"
        >
          <el-option v-for="s in sheetList" :key="s" :label="s" :value="s" />
        </el-select>
        <el-button
          type="primary"
          :disabled="!selectedSheet"
          @click="loadHeaders"
          :loading="loadingHeaders"
          size="small"
        >识别表头</el-button>
      </div>
      <div v-if="!sheetList.length" class="hint-text">请先上传文件并识别工作表</div>
    </div>

    <!-- 3. 选择拆分依据列 -->
    <div class="section">
      <div class="section-title">3. 选择拆分依据列</div>
      <div class="form-row">
        <el-select
          v-model="selectedHeader"
          placeholder="请选择拆分依据列（第2行表头）"
          style="width:320px"
          :disabled="!headerList.length"
        >
          <el-option v-for="h in headerList" :key="h" :label="h" :value="h" />
        </el-select>
      </div>
      <div v-if="headerList.length" class="preview-hint">
        已识别第2行表头，共 <strong>{{ headerList.length }}</strong> 个字段
      </div>
      <div v-else class="hint-text">请先选择工作表并识别表头</div>
    </div>

    <!-- 4. 选择拆分方式 -->
    <div class="section">
      <div class="section-title">4. 选择拆分方式</div>
      <div class="form-row">
        <el-button
          type="primary"
          @click="doSplit"
          :loading="loadingSplit"
          :disabled="!selectedHeader"
        >按列直接拆分</el-button>
        <el-button
          type="warning"
          @click="openCombineDialog"
          :disabled="!selectedHeader"
          :loading="loadingValues || loadingCombine"
        >选择拆分</el-button>
      </div>
      <div v-if="groupList.length" class="group-summary">
        已设置 <strong>{{ groupList.length }}</strong> 个分组，共覆盖 <strong>{{ groupedValueCount }}</strong> 个值
      </div>
      <div v-if="!selectedHeader" class="hint-text">请先选择拆分依据列</div>
    </div>

    <!-- 选择拆分弹窗 -->
    <el-dialog
      v-model="showCombineDialog"
      title="选择拆分 — 自定义分组"
      width="700px"
      top="8vh"
      :close-on-click-modal="false"
      :before-close="closeCombineDialog"
    >
      <div class="combine-body">
        <!-- 左侧：可选值列表 -->
        <div class="combine-left">
          <div class="combine-panel-title">
            可选值
            <span class="combine-count">({{ availableValues.length }} 个)</span>
          </div>
          <div class="combine-value-list">
            <el-checkbox-group v-model="selectedValues">
              <el-checkbox
                v-for="v in availableValues"
                :key="v"
                :label="v"
                class="combine-checkbox"
              >{{ v }}</el-checkbox>
            </el-checkbox-group>
          </div>
          <el-button
            type="primary"
            size="small"
            :disabled="!selectedValues.length"
            @click="createGroup"
            style="margin-top:8px;width:100%"
          >创建分组（选中 {{ selectedValues.length }} 个）</el-button>
        </div>

        <!-- 右侧：已建分组 -->
        <div class="combine-right">
          <div class="combine-panel-title">
            已建分组
            <span class="combine-count">({{ groupList.length }} 个)</span>
          </div>
          <div class="combine-group-list">
            <div
              v-for="(g, idx) in groupList"
              :key="g.id"
              class="combine-group-item"
            >
              <el-tag
                v-for="v in g.values"
                :key="v"
                closable
                @close="removeValueFromGroup(idx, v)"
                size="small"
                style="margin:2px 4px 2px 0"
              >{{ v }}</el-tag>
              <el-button
                type="danger"
                size="small"
                text
                @click="removeGroup(idx)"
                style="margin-left:auto"
              >删除分组</el-button>
            </div>
            <div v-if="!groupList.length" class="hint-text" style="padding:20px;text-align:center">
              暂无分组，请从左侧选择值并创建
            </div>
          </div>
        </div>
      </div>

      <template #footer>
        <el-button @click="closeCombineDialog">取消</el-button>
        <el-button type="primary" @click="confirmCombine">确认分组并拆分</el-button>
      </template>
    </el-dialog>

    <!-- 结果区域 -->
    <div v-if="resultMsg || resultOk" class="section result-section">
      <el-alert v-if="resultOk" title="工资表拆分完成！" type="success" :closable="false" show-icon />
      <el-alert v-if="!resultOk && resultMsg" :title="resultMsg" type="error" :closable="false" show-icon />
      <div v-if="resultOk" class="download-area">
        <el-button type="primary" @click="downloadFile">下载拆分结果（ZIP）</el-button>
      </div>
      <el-button text @click="resetAll" style="margin-top:12px">重新开始</el-button>
    </div>
    </div>
  </el-card>
</template>

<script>
import axios from 'axios'
import { UploadFilled } from '@element-plus/icons-vue'

const API_BASE = '/api/hr/split-salary'

export default {
  name: 'CfGzbcfHyp',
  components: {
    UploadFilled
  },
  data() {
    return {
      fileRaw: null,
      fileName: '',
      sheetList: [],
      selectedSheet: '',
      loadingSheets: false,
      headerList: [],
      selectedHeader: '',
      loadingHeaders: false,
      loadingSplit: false,
      resultOk: false,
      resultMsg: '',
      downloadFileName: '',
      showCombineDialog: false,
      loadingValues: false,
      loadingCombine: false,
      allColumnValues: [],
      availableValues: [],
      selectedValues: [],
      groupList: [],
      groupIdCounter: 0
    }
  },
  computed: {
    groupedValueCount() {
      const set = new Set()
      this.groupList.forEach(g => g.values.forEach(v => set.add(v)))
      return set.size
    }
  },
  methods: {
    handleFileChange(uploadFile) {
      this.fileRaw = uploadFile.raw
      this.fileName = uploadFile.name
      this.sheetList = []
      this.selectedSheet = ''
      this.headerList = []
      this.selectedHeader = ''
      this.resultOk = false
      this.resultMsg = ''
      this.groupList = []
      this.allColumnValues = []
      this.availableValues = []
      this.selectedValues = []
      this.loadingCombine = false
    },
    handleExceed() {
      this.$message.warning('最多只能选择1个文件')
    },
    async loadSheets() {
      if (!this.fileRaw) return
      this.loadingSheets = true
      this.sheetList = []
      this.selectedSheet = ''
      this.headerList = []
      this.selectedHeader = ''
      try {
        const formData = new FormData()
        formData.append('file', this.fileRaw)
        const res = await axios.post(API_BASE + '/sheets', formData)
        if (res.data.success) {
          this.sheetList = res.data.sheets || []
          if (this.sheetList.length > 0) {
            this.selectedSheet = this.sheetList[0]
          }
        } else {
          this.$message.error(res.data.message || '获取工作表失败')
        }
      } catch (err) {
        this.$message.error(err.response?.data?.message || err.message || '获取工作表失败')
      } finally {
        this.loadingSheets = false
      }
    },
    onSheetChange() {
      this.headerList = []
      this.selectedHeader = ''
      this.resultOk = false
      this.resultMsg = ''
      this.groupList = []
      this.allColumnValues = []
      this.availableValues = []
      this.selectedValues = []
    },
    async loadHeaders() {
      if (!this.fileRaw || !this.selectedSheet) return
      this.loadingHeaders = true
      this.headerList = []
      this.selectedHeader = ''
      try {
        const formData = new FormData()
        formData.append('file', this.fileRaw)
        formData.append('sheetName', this.selectedSheet)
        formData.append('headerRow', 2)
        const res = await axios.post(API_BASE + '/headers', formData)
        if (res.data.success) {
          this.headerList = res.data.headers || []
          if (this.headerList.length === 0) {
            this.$message.warning('未识别到表头，请确认第2行是否包含列名')
          }
        } else {
          this.$message.error(res.data.message || '获取表头失败')
        }
      } catch (err) {
        this.$message.error(err.response?.data?.message || err.message || '获取表头失败')
      } finally {
        this.loadingHeaders = false
      }
    },
    async doSplit() {
      if (!this.fileRaw || !this.selectedSheet || !this.selectedHeader) return
      this.loadingSplit = true
      this.resultOk = false
      this.resultMsg = ''
      try {
        await this.executeSplit()
      } finally {
        this.loadingSplit = false
      }
    },
    async doSplitWithGroups() {
      if (!this.fileRaw || !this.selectedSheet || !this.selectedHeader) return
      this.loadingCombine = true
      this.resultOk = false
      this.resultMsg = ''
      try {
        await this.executeSplit()
      } finally {
        this.loadingCombine = false
      }
    },
    async executeSplit() {
      const formData = new FormData()
      formData.append('file', this.fileRaw)
      formData.append('sheetName', this.selectedSheet)
      formData.append('splitColName', this.selectedHeader)
      if (this.groupList.length > 0) {
        const groupsArr = this.groupList.map(g => g.values)
        const json = JSON.stringify(groupsArr)
        formData.append('groups', btoa(unescape(encodeURIComponent(json))))
      }
      const res = await axios.post(API_BASE, formData)
      if (res.data.success) {
        this.resultOk = true
        this.resultMsg = res.data.message || '工资表拆分完成'
        this.downloadFileName = res.data.fileName || ''
      } else {
        this.resultOk = false
        this.resultMsg = res.data.message || '处理失败'
      }
    },
    async loadColumnValues() {
      if (!this.fileRaw || !this.selectedSheet || !this.selectedHeader) return
      this.loadingValues = true
      try {
        const formData = new FormData()
        formData.append('file', this.fileRaw)
        formData.append('sheetName', this.selectedSheet)
        formData.append('colName', this.selectedHeader)
        const res = await axios.post(API_BASE + '/values', formData)
        if (res.data.success) {
          this.allColumnValues = res.data.values || []
          this.availableValues = [...this.allColumnValues]
        } else {
          this.$message.error(res.data.message || '获取列值失败')
        }
      } catch (err) {
        this.$message.error(err.response?.data?.message || err.message || '获取列值失败')
      } finally {
        this.loadingValues = false
      }
    },
    async openCombineDialog() {
      if (!this.selectedHeader) return
      await this.loadColumnValues()
      this.showCombineDialog = true
    },
    closeCombineDialog() {
      this.showCombineDialog = false
      this.selectedValues = []
    },
    createGroup() {
      if (!this.selectedValues.length) return
      const group = {
        id: ++this.groupIdCounter,
        values: [...this.selectedValues]
      }
      this.groupList.push(group)
      this.availableValues = this.availableValues.filter(v => !this.selectedValues.includes(v))
      this.selectedValues = []
    },
    removeGroup(idx) {
      const removed = this.groupList[idx]
      removed.values.forEach(v => {
        if (this.allColumnValues.includes(v) && !this.availableValues.includes(v)) {
          this.availableValues.push(v)
        }
      })
      this.availableValues.sort((a, b) => this.allColumnValues.indexOf(a) - this.allColumnValues.indexOf(b))
      this.groupList.splice(idx, 1)
    },
    removeValueFromGroup(groupIdx, val) {
      const group = this.groupList[groupIdx]
      const valIdx = group.values.indexOf(val)
      if (valIdx > -1) {
        group.values.splice(valIdx, 1)
        this.availableValues.push(val)
        this.availableValues.sort((a, b) => this.allColumnValues.indexOf(a) - this.allColumnValues.indexOf(b))
        if (!group.values.length) {
          this.groupList.splice(groupIdx, 1)
        }
      }
    },
    confirmCombine() {
      this.showCombineDialog = false
      this.selectedValues = []
      if (this.groupList.length > 0) {
        this.doSplitWithGroups()
      }
    },
    downloadFile() {
      if (this.downloadFileName) {
        window.location.href = '/api/hr/download/' + this.downloadFileName
      }
    },
    resetAll() {
      this.fileRaw = null
      this.fileName = ''
      this.sheetList = []
      this.selectedSheet = ''
      this.headerList = []
      this.selectedHeader = ''
      this.resultOk = false
      this.resultMsg = ''
      this.downloadFileName = ''
      this.showCombineDialog = false
      this.loadingCombine = false
      this.allColumnValues = []
      this.availableValues = []
      this.selectedValues = []
      this.groupList = []
      this.groupIdCounter = 0
      this.$refs.uploadRef?.clearFiles()
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

.section {
  position: relative;
  padding: 20px 0;
  border-bottom: 1px dashed #e4e7ed;
  transition: background 0.3s, box-shadow 0.3s;
}

.section:last-child {
  border-bottom: none;
}

.section:hover {
  background: #fafbfc;
  /* 用 box-shadow 模拟背景外扩到父容器 padding 区域，
     避免 margin 0 -24px 造成横向溢出导致底部出现横向滚动条 */
  box-shadow: 0 0 0 24px #fafbfc;
  border-radius: 8px;
}

.section-title {
  font-size: 15px;
  font-weight: bold;
  color: #303133;
  margin-bottom: 14px;
  display: flex;
  align-items: center;
  gap: 6px;
}

.section-title::before {
  content: '';
  display: inline-block;
  width: 4px;
  height: 16px;
  background: #e67e22;
  border-radius: 2px;
}

.gzbcf-upload {
  width: 100%;
}

.gzbcf-upload .el-upload-dragger {
  width: 100%;
  border: 2px dashed #dcdfe6;
  border-radius: 8px;
  background: #fafafa;
  transition: all 0.3s;
}

.gzbcf-upload .el-upload-dragger:hover {
  border-color: #e67e22;
  background: #fef9f3;
}

.gzbcf-upload .el-upload-dragger .el-icon--upload {
  font-size: 48px;
  color: #c0c4cc;
  transition: color 0.3s;
}

.gzbcf-upload .el-upload-dragger:hover .el-icon--upload {
  color: #e67e22;
}

.gzbcf-upload .el-upload__text {
  color: #606266;
  font-size: 14px;
}

.gzbcf-upload .el-upload__text em {
  color: #e67e22;
  font-style: normal;
  font-weight: bold;
}

.gzbcf-upload .el-upload__tip {
  color: #909399;
  font-size: 12px;
}

.file-name {
  margin-top: 10px;
  padding: 8px 12px;
  background: #f0f9ff;
  border-left: 3px solid #409eff;
  border-radius: 4px;
  color: #409eff;
  font-size: 13px;
  word-break: break-all;
}

.form-row {
  display: flex;
  align-items: center;
  flex-wrap: wrap;
  gap: 10px;
  margin-bottom: 8px;
}

.hint-text {
  font-size: 13px;
  color: #c0c4cc;
  margin-top: 6px;
  padding-left: 2px;
}

.preview-hint {
  font-size: 13px;
  color: #909399;
  margin-top: 10px;
  padding: 8px 12px;
  background: #f5f7fa;
  border-radius: 4px;
}

.preview-hint strong {
  color: #e67e22;
  font-size: 15px;
}

.result-section {
  background: linear-gradient(135deg, #f5f7fa 0%, #fef9f3 100%);
  border-radius: 8px;
  padding: 20px;
  margin-top: 12px;
  border: 1px solid #f0e6d6;
}

.result-section .el-alert {
  border-radius: 6px;
}

.download-area {
  margin-top: 16px;
}

.download-area .el-button--primary {
  background-color: #27ae60;
  border-color: #27ae60;
  padding: 10px 24px;
  font-size: 14px;
}

.download-area .el-button--primary:hover {
  background-color: #2ecc71;
  border-color: #2ecc71;
}

.combine-body {
  display: flex;
  gap: 16px;
  min-height: 350px;
  max-height: 55vh;
}

.combine-left,
.combine-right {
  flex: 1;
  display: flex;
  flex-direction: column;
}

.combine-left {
  border-right: 1px solid #e4e7ed;
  padding-right: 12px;
}

.combine-right {
  padding-left: 12px;
}

.combine-panel-title {
  font-size: 14px;
  font-weight: 600;
  color: #303133;
  margin-bottom: 8px;
  padding-bottom: 8px;
  border-bottom: 1px solid #ebeef5;
}

.combine-count {
  font-weight: 400;
  color: #909399;
  font-size: 12px;
}

.combine-value-list {
  flex: 1;
  overflow-y: auto;
  border: 1px solid #dcdfe6;
  border-radius: 6px;
  padding: 10px 14px;
  background: #fafafa;
}

.combine-checkbox {
  display: flex !important;
  margin-bottom: 6px;
}

.combine-group-list {
  flex: 1;
  overflow-y: auto;
  border: 1px solid #dcdfe6;
  border-radius: 6px;
  padding: 10px;
  background: #fafafa;
}

.combine-group-item {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  padding: 8px 10px;
  margin-bottom: 8px;
  border: 1px solid #e6a23c;
  border-radius: 6px;
  background: #fdf6ec;
}

.combine-group-item:last-child {
  margin-bottom: 0;
}

.group-summary {
  margin-top: 8px;
  font-size: 13px;
  color: #e6a23c;
}

.group-summary strong {
  color: #303133;
}

@media (max-width: 768px) {
  .page-container {
    max-width: 100%;
  }

  .form-row {
    flex-direction: column;
    align-items: stretch;
  }

  .form-row .el-select {
    width: 100% !important;
  }

  .section:hover {
    /* 移动端 16px 同样用 box-shadow 模拟，避免横向溢出 */
    box-shadow: 0 0 0 16px #fafbfc;
  }
}
</style>
