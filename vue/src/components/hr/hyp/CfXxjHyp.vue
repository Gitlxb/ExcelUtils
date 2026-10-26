<template>
  <el-card class="xxj-card">
    <div class="xxj-page">
      <div class="xxj-header">
        <h2 class="xxj-header-title">孝心金拆分</h2>
        <p class="xxj-header-desc">
          按拆分列将孝心金表拆分为独立的 Excel 文件。自动识别第 2 行表头，保留标题行、表头行与表尾汇总公式。
          同一名称若不连续出现（如两段「浙江锦途」），会分别输出为独立文件。
        </p>
      </div>

      <!-- 1. 上传文件 -->
      <div class="xxj-section">
        <div class="xxj-section-title">1. 上传文件</div>
        <el-upload
          ref="uploadRef"
          drag
          accept=".xlsx,.xlsm"
          :auto-upload="false"
          :limit="1"
          :on-change="handleFileChange"
          :on-exceed="handleExceed"
          class="xxj-upload"
        >
          <el-icon class="el-icon--upload"><UploadFilled /></el-icon>
          <div class="el-upload__text">拖拽文件到此处或 <em>点击上传</em></div>
          <div class="el-upload__tip">支持 .xlsx / .xlsm 格式</div>
        </el-upload>
        <div v-if="fileName" class="xxj-file-name">已选择：{{ fileName }}</div>
        <el-button
          v-if="fileRaw && !sheetList.length"
          type="primary"
          class="xxj-mt-12"
          @click="loadSheets"
          :loading="loadingSheets"
        >识别工作表</el-button>
      </div>

      <!-- 2. 选择工作表 -->
      <div class="xxj-section">
        <div class="xxj-section-title">2. 选择工作表</div>
        <div class="xxj-form-row">
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
        <div v-if="!sheetList.length" class="xxj-hint">请先上传文件并识别工作表</div>
      </div>

      <!-- 3. 选择拆分列 -->
      <div class="xxj-section">
        <div class="xxj-section-title">3. 选择拆分列</div>
        <div class="xxj-form-row">
          <el-select
            v-model="selectedHeader"
            placeholder="请选择拆分列（第2行表头）"
            style="width:300px"
            :disabled="!headerList.length"
            @change="loadSegments"
          >
            <el-option v-for="h in headerList" :key="h" :label="h" :value="h" />
          </el-select>
          <span v-if="loadingValues" class="xxj-hint">正在识别分组…</span>
        </div>
        <div v-if="!headerList.length" class="xxj-hint">请先选择工作表并识别表头</div>
      </div>

      <!-- 4. 分组预览 -->
      <div class="xxj-section">
        <div class="xxj-section-title">4. 分组预览</div>
        <div v-if="segmentRows.length" class="xxj-table-wrap">
          <el-table :data="segmentRows" size="small" max-height="320" stripe>
            <el-table-column prop="value" label="拆分值" min-width="150" show-overflow-tooltip />
            <el-table-column prop="rangeText" label="源文件行范围" min-width="130" align="center" />
            <el-table-column prop="count" label="行数" min-width="70" align="center" />
            <el-table-column label="金额合计" min-width="150" align="right">
              <template #default="{ row }">
                <span class="xxj-amount">{{ row.amountText }}</span>
              </template>
            </el-table-column>
          </el-table>
        </div>
        <div v-if="segmentRows.length" class="xxj-preview-hint">
          共 <strong>{{ segmentRows.length }}</strong> 个分组 / {{ totalRows }} 行<span v-if="hasAmount"> / 金额合计 <strong>{{ totalAmountText }}</strong></span>
        </div>
        <div v-else class="xxj-hint">选择拆分列后自动识别分组</div>
      </div>

      <!-- 5. 选择拆分方式 -->
      <div class="xxj-section">
        <div class="xxj-section-title">5. 选择拆分方式</div>
        <div class="xxj-form-row">
          <el-button
            type="primary"
            @click="doSplit"
            :loading="loadingSplit"
            :disabled="!segmentRows.length"
          >直接拆分</el-button>
          <el-button
            type="warning"
            @click="openCombineDialog"
            :disabled="!segmentRows.length"
            :loading="loadingCombine"
          >组合拆分</el-button>
        </div>
        <div v-if="groupList.length" class="xxj-group-summary">
          已设置 <strong>{{ groupList.length }}</strong> 个分组，共覆盖 <strong>{{ groupedSegmentCount }}</strong> 个分组项
        </div>
        <div v-if="!segmentRows.length" class="xxj-hint">请先选择拆分列</div>
      </div>

      <!-- 组合拆分弹窗 -->
      <el-dialog
        v-model="showCombineDialog"
        title="组合拆分 — 自定义分组"
        width="720px"
        top="8vh"
        class="xxj-dialog"
        :close-on-click-modal="false"
        :before-close="closeCombineDialog"
      >
        <div class="xxj-combine-body">
          <div class="xxj-combine-left">
            <div class="xxj-combine-panel-title">
              可选分组
              <span class="xxj-combine-count">（{{ availableSegments.length }} 个）</span>
            </div>
            <div class="xxj-combine-list">
              <el-checkbox-group v-model="selectedIndexes">
                <el-checkbox
                  v-for="s in availableSegments"
                  :key="s.index"
                  :label="s.index"
                  class="xxj-combine-checkbox"
                >{{ segLabel(s) }}</el-checkbox>
              </el-checkbox-group>
            </div>
            <el-button
              type="primary"
              size="small"
              class="xxj-mt-8-full"
              :disabled="!selectedIndexes.length"
              @click="createGroup"
            >创建分组（选中 {{ selectedIndexes.length }} 个）</el-button>
          </div>

          <div class="xxj-combine-right">
            <div class="xxj-combine-panel-title">
              已建分组
              <span class="xxj-combine-count">（{{ groupList.length }} 个）</span>
            </div>
            <div class="xxj-combine-group-list">
              <div v-for="(g, idx) in groupList" :key="g.id" class="xxj-combine-group-item">
                <el-tag
                  v-for="i in g.indexes"
                  :key="i"
                  closable
                  size="small"
                  class="xxj-group-tag"
                  @close="removeSegmentFromGroup(idx, i)"
                >{{ segLabelByIndex(i) }}</el-tag>
                <el-button
                  type="danger"
                  size="small"
                  text
                  class="xxj-group-del"
                  @click="removeGroup(idx)"
                >删除分组</el-button>
              </div>
              <div v-if="!groupList.length" class="xxj-combine-empty">
                暂无分组，请从左侧选择分组并创建
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
      <div v-if="resultMsg || resultOk" class="xxj-section xxj-result">
        <el-alert v-if="resultOk" title="孝心金拆分完成！" type="success" :closable="false" show-icon />
        <el-alert v-if="!resultOk && resultMsg" :title="resultMsg" type="error" :closable="false" show-icon />
        <div v-if="resultOk" class="xxj-download-area">
          <el-button type="primary" @click="downloadFile">下载拆分结果（ZIP）</el-button>
        </div>
        <el-button text class="xxj-mt-12" @click="resetAll">重新开始</el-button>
      </div>
    </div>
  </el-card>
</template>

<script>
import axios from 'axios'
import { postAsync } from '../../../utils/asyncTask'
import { UploadFilled } from '@element-plus/icons-vue'

const API_BASE = '/api/hr/split-xxj'

export default {
  name: 'CfXxjHyp',
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
      loadingValues: false,
      loadingSplit: false,
      loadingCombine: false,
      segments: [],
      hasAmount: false,
      resultOk: false,
      resultMsg: '',
      downloadFileName: '',
      showCombineDialog: false,
      availableSegments: [],
      selectedIndexes: [],
      groupList: [],
      groupIdCounter: 0
    }
  },
  computed: {
    segmentRows() {
      return this.segments.map(s => ({
        index: s.index,
        value: s.value,
        count: s.count,
        // 行范围由后端算好：单处为「第5-48行」，同名值跨多处为「第5-48行 等2段」
        rangeText: s.rangeText || `第${s.startRow}-${s.endRow}行`,
        amountText: this.fmtAmount(s.amount)
      }))
    },
    totalRows() {
      return this.segments.reduce((n, s) => n + (s.count || 0), 0)
    },
    totalAmountText() {
      const total = this.segments.reduce((n, s) => n + (s.amount || 0), 0)
      return this.fmtAmount(total)
    },
    groupedSegmentCount() {
      const set = new Set()
      this.groupList.forEach(g => g.indexes.forEach(i => set.add(i)))
      return set.size
    }
  },
  methods: {
    fmtAmount(v) {
      const n = Number(v || 0)
      return n.toLocaleString('zh-CN', { maximumFractionDigits: 2 })
    },
    segLabel(seg) {
      const range = seg.rangeText || `第${seg.startRow}-${seg.endRow}行`
      return `${seg.value}（${range}）`
    },
    segLabelByIndex(index) {
      const seg = this.segments.find(s => s.index === index)
      return seg ? seg.value : `#${index}`
    },
    handleFileChange(uploadFile) {
      this.fileRaw = uploadFile.raw
      this.fileName = uploadFile.name
      this.resetDownstream()
    },
    handleExceed() {
      this.$message.warning('最多只能选择1个文件')
    },
    resetDownstream() {
      this.sheetList = []
      this.selectedSheet = ''
      this.headerList = []
      this.selectedHeader = ''
      this.segments = []
      this.hasAmount = false
      this.resultOk = false
      this.resultMsg = ''
      this.downloadFileName = ''
      this.groupList = []
      this.availableSegments = []
      this.selectedIndexes = []
      this.showCombineDialog = false
    },
    async loadSheets() {
      if (!this.fileRaw) return
      this.loadingSheets = true
      this.sheetList = []
      this.selectedSheet = ''
      this.headerList = []
      this.selectedHeader = ''
      this.segments = []
      try {
        const formData = new FormData()
        formData.append('file', this.fileRaw)
        const res = await axios.post(API_BASE + '/sheets', formData)
        if (res.data.success) {
          this.sheetList = res.data.sheets || []
          if (this.sheetList.length > 0) this.selectedSheet = this.sheetList[0]
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
      this.segments = []
      this.hasAmount = false
      this.resultOk = false
      this.resultMsg = ''
      this.groupList = []
      this.availableSegments = []
      this.selectedIndexes = []
    },
    async loadHeaders() {
      if (!this.fileRaw || !this.selectedSheet) return
      this.loadingHeaders = true
      this.headerList = []
      this.selectedHeader = ''
      this.segments = []
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
    async loadSegments() {
      if (!this.fileRaw || !this.selectedSheet || !this.selectedHeader) return
      this.loadingValues = true
      this.segments = []
      this.hasAmount = false
      this.resultOk = false
      this.resultMsg = ''
      this.groupList = []
      this.availableSegments = []
      this.selectedIndexes = []
      try {
        const formData = new FormData()
        formData.append('file', this.fileRaw)
        formData.append('sheetName', this.selectedSheet)
        formData.append('colName', this.selectedHeader)
        const res = await axios.post(API_BASE + '/values', formData)
        if (res.data.success) {
          this.segments = res.data.segments || []
          this.hasAmount = this.segments.some(s => Number(s.amount) > 0)
          if (this.segments.length === 0) {
            this.$message.warning('未识别到可拆分的数据分组')
          }
        } else {
          this.$message.error(res.data.message || '识别分组失败')
        }
      } catch (err) {
        this.$message.error(err.response?.data?.message || err.message || '识别分组失败')
      } finally {
        this.loadingValues = false
      }
    },
    async doSplit() {
      if (!this.fileRaw || !this.selectedSheet || !this.selectedHeader) return
      this.loadingSplit = true
      this.resultOk = false
      this.resultMsg = ''
      try {
        await this.executeSplit(false)
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
        await this.executeSplit(true)
      } finally {
        this.loadingCombine = false
      }
    },
    async executeSplit(withGroups) {
      const formData = new FormData()
      formData.append('file', this.fileRaw)
      formData.append('sheetName', this.selectedSheet)
      formData.append('splitColName', this.selectedHeader)
      if (withGroups && this.groupList.length > 0) {
        const groupsArr = this.groupList.map(g => g.indexes)
        const json = JSON.stringify(groupsArr)
        formData.append('groups', btoa(unescape(encodeURIComponent(json))))
      }
      const res = await postAsync(API_BASE, formData)
      if (res.data.success) {
        this.resultOk = true
        this.resultMsg = res.data.message || '孝心金拆分完成'
        this.downloadFileName = res.data.fileName || ''
      } else {
        this.resultOk = false
        this.resultMsg = res.data.message || '处理失败'
      }
    },
    openCombineDialog() {
      if (!this.segments.length) return
      this.availableSegments = [...this.segments]
      this.selectedIndexes = []
      this.groupList = []
      this.groupIdCounter = 0
      this.showCombineDialog = true
    },
    closeCombineDialog() {
      this.showCombineDialog = false
      this.selectedIndexes = []
    },
    refreshAvailable() {
      const used = new Set()
      this.groupList.forEach(g => g.indexes.forEach(i => used.add(i)))
      this.availableSegments = this.segments.filter(s => !used.has(s.index))
    },
    createGroup() {
      if (!this.selectedIndexes.length) return
      this.groupList.push({ id: ++this.groupIdCounter, indexes: [...this.selectedIndexes] })
      this.selectedIndexes = []
      this.refreshAvailable()
    },
    removeGroup(idx) {
      this.groupList.splice(idx, 1)
      this.refreshAvailable()
    },
    removeSegmentFromGroup(groupIdx, index) {
      const group = this.groupList[groupIdx]
      const i = group.indexes.indexOf(index)
      if (i === -1) return
      group.indexes.splice(i, 1)
      if (!group.indexes.length) this.groupList.splice(groupIdx, 1)
      this.refreshAvailable()
    },
    confirmCombine() {
      if (!this.groupList.length) {
        this.$message.warning('请先创建至少一个分组')
        return
      }
      this.showCombineDialog = false
      this.selectedIndexes = []
      this.doSplitWithGroups()
    },
    downloadFile() {
      if (this.downloadFileName) {
        window.location.href = '/api/hr/download/' + this.downloadFileName
      }
    },
    resetAll() {
      this.fileRaw = null
      this.fileName = ''
      this.resetDownstream()
      this.$refs.uploadRef?.clearFiles()
    }
  }
}
</script>

<style scoped src="../../../styles/components/XxjSplitCard.css"></style>
