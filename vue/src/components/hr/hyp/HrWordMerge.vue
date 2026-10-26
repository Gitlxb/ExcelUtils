<template>
  <el-card class="func-card">
    <div class="page-container">
      <div class="func-page-header">
        <h2 class="func-page-title">Word 文档合并（分类版）</h2>
        <p class="func-page-desc">
          上传评标办法文档提取分类关键词，上传候选文档后智能搜索（精准 → 模糊两级），
          按分类层级合并为一个统一文档，文件间自动插入 A4 分节符。
        </p>
      </div>

      <el-row :gutter="16">
        <!-- 左栏：操作主流程 -->
        <el-col :xs="24" :sm="24" :md="16" :lg="16" class="main-col">
          <!-- 第1步 -->
          <el-divider content-position="left">第1步：选择评标办法文档（可多选）</el-divider>
          <el-upload
            drag multiple :auto-upload="false" accept=".docx"
            :file-list="standardList"
            :on-change="onStandardChange" :on-remove="onStandardRemove"
          >
            <el-icon class="el-icon--upload"><UploadFilled /></el-icon>
            <div class="el-upload__text">将评标办法文档拖到此处，或 <em>点击选择</em></div>
          </el-upload>

          <!-- 第2步 -->
          <el-divider content-position="left">第2步：选择搜索文件夹</el-divider>
          <div class="dir-upload" @click="$refs.dirInput.click()">
            <el-icon class="el-icon--upload"><FolderOpened /></el-icon>
            <div class="el-upload__text">
              点击此处 <em>选择文件夹</em>（仅收录 .docx）
            </div>
            <div v-if="candidateCount" class="dir-count">已收录 {{ candidateCount }} 个文件</div>
            <input
              ref="dirInput"
              type="file"
              webkitdirectory
              accept=".docx"
              style="display:none"
              @change="onDirectoryChange"
            />
          </div>

          <!-- 第3步 -->
          <el-divider content-position="left">第3步：提取分类关键词（可编辑）</el-divider>
          <el-button type="primary" :disabled="standardList.length === 0" :loading="loading" @click="extract">
            提取关键词
          </el-button>
          <el-alert
            v-if="!standardList.length"
            type="info" :closable="false" show-icon
            title="请先在第1步选择至少一个评标办法文档（.docx）"
            style="margin-top:10px"
          />
          <div v-if="categories.length" class="cat-edit-wrap">
            <div v-for="(cat, idx) in categories" :key="'cat'+idx" class="cat-edit">
              <div class="cat-edit-title">
                <strong>{{ cat.parent }}</strong>
                <span class="cat-edit-count">（{{ cat.keywords.length }} 个关键词）</span>
              </div>
              <el-input
                type="textarea"
                :rows="Math.max(2, Math.min(cat.keywords.length + 1, 8))"
                v-model="keywordText[idx]"
                placeholder="每行一个关键词"
              />
            </div>
          </div>

          <!-- 第4步 -->
          <el-divider content-position="left">第4步：搜索文件（精准 → 模糊）</el-divider>
          <el-button
            type="primary" :disabled="categories.length === 0 || candidateFiles.length === 0"
            :loading="loading" @click="search"
          >
            搜索文件
          </el-button>
          <span v-if="categories.length && !candidateFiles.length" class="hint">请先在第2步选择搜索文件夹</span>

          <!-- 第5步 -->
          <el-divider content-position="left">第5步：搜索结果（勾选要合并的文件）</el-divider>
          <el-alert
            v-if="!categories.length"
            type="info" :closable="false" show-icon
            title="尚无搜索结果。请先提取关键词并点击「搜索文件」。"
            style="margin-bottom:10px"
          />
          <div v-for="(cat, idx) in categories" :key="'res'+idx" class="result-cat">
            <div class="result-cat-title">
              [{{ cat.parent }}] — 已勾选 {{ selectedCount(idx) }} / {{ cat.matchedFiles.length }} 个文件
              <el-tag v-if="preciseCount(idx)" size="small" type="success" effect="plain" class="title-tag">精确 {{ preciseCount(idx) }}</el-tag>
              <el-tag v-if="fuzzyCount(idx)" size="small" type="warning" effect="plain" class="title-tag">模糊 {{ fuzzyCount(idx) }}</el-tag>
            </div>
            <div v-if="cat.matchedFiles.length === 0" class="result-empty">（未匹配到文件）</div>
            <el-checkbox
              v-for="(f, fi) in cat.matchedFiles" :key="fi"
              v-model="selected[idx][fi]"
              :label="f.fileName" border size="small" class="result-item"
            >
              <span class="result-label">{{ f.fileName }}</span>
              <el-tag v-if="f.matchType === 'fuzzy'" size="small" type="warning" effect="plain" class="result-tag">模糊</el-tag>
              <el-tag v-else-if="f.matchType === 'precise'" size="small" type="success" effect="plain" class="result-tag">精确</el-tag>
            </el-checkbox>
          </div>

          <!-- 第6步 -->
          <el-divider content-position="left">第6步：合并</el-divider>
          <el-button type="success" :loading="loading" :disabled="!hasSelected()" @click="merge">
            ▶ 开始合并
          </el-button>

          <el-card v-if="downloadFileName" style="margin-top:16px">
            <span>合并完成：{{ downloadFileName }}</span>
            <el-button type="primary" style="margin-left:12px" @click="downloadFile">下载文件</el-button>
          </el-card>

          <el-alert
            v-if="message" :title="message" :type="messageType" show-icon closable
            style="margin-top:16px" @close="message=''"
          />
        </el-col>

        <!-- 右栏：操作日志 -->
        <el-col :xs="24" :sm="24" :md="8" :lg="8">
          <div class="log-col">
            <div class="log-head">
              <span class="log-title">操作日志</span>
              <div class="log-head-actions">
                <span class="log-count">{{ logs.length }} 条</span>
                <el-button size="small" text type="primary" @click="clearLogs">清空</el-button>
              </div>
            </div>
            <div class="log-panel">
              <div v-for="(l, i) in logs" :key="i" class="log-line">{{ l }}</div>
              <div v-if="!logs.length" class="log-empty">暂无日志，开始操作后此处显示详细过程</div>
            </div>
          </div>
        </el-col>
      </el-row>
    </div>
  </el-card>
</template>

<script>
import axios from 'axios'
import { postAsync } from '../../../utils/asyncTask'
import { UploadFilled, FolderOpened } from '@element-plus/icons-vue'

export default {
  name: 'HrWordMerge',
  components: { UploadFilled, FolderOpened },
  data() {
    return {
      loading: false,
      message: '',
      messageType: 'info',
      downloadFileName: '',
      standardList: [],
      candidateList: [],
      standardFiles: [],
      candidateFiles: [],
      categories: [],          // [{parent, type, keywords, matchedFiles}]
      keywordText: [],          // 索引对齐 categories：可编辑关键词文本
      selected: [],             // 索引对齐 categories：每项为 boolean[]（按 matchedFiles）
      logs: []
    }
  },
  computed: {
    candidateCount() {
      return this.candidateFiles.length
    }
  },
  methods: {
    log(msg) {
      this.logs.push(msg)
      if (this.logs.length > 500) this.logs.shift()
    },
    clearLogs() {
      this.logs = []
    },

    // ---- 上传处理 ----
    onStandardChange(file, fileList) {
      this.standardList = fileList
      this.standardFiles = fileList.map(f => f.raw).filter(Boolean)
    },
    onStandardRemove(file, fileList) {
      this.standardList = fileList
      this.standardFiles = fileList.map(f => f.raw).filter(Boolean)
    },

    // ---- 目录选择 ----
    onDirectoryChange(e) {
      const allFiles = Array.from(e.target.files)
      const docxFiles = allFiles.filter(f => f.name.toLowerCase().endsWith('.docx'))
      const skipped = allFiles.filter(f => !f.name.toLowerCase().endsWith('.docx'))
      this.candidateFiles = docxFiles
      this.candidateList = docxFiles.map(f => ({ name: f.name }))
      this.log('✓ 已选择文件夹，共收录 ' + docxFiles.length + ' 个 .docx 文件'
        + (skipped.length > 0 ? '（跳过 ' + skipped.length + ' 个非 .docx 文件）' : ''))
      if (skipped.length > 0) {
        this.log('  跳过的文件：' + skipped.map(f => f.name).join(', '))
      }
      e.target.value = ''
    },

    // ---- 第3步：提取关键词 ----
    async extract() {
      if (this.standardFiles.length === 0) return
      this.loading = true
      this.message = ''
      this.log('开始提取分类关键词...')
      const fd = new FormData()
      this.standardFiles.forEach(f => fd.append('files', f))
      try {
        const res = await axios.post('/api/hr/word-merge/extract', fd)
        if (res.data.success) {
          this.categories = res.data.categories || []
          this.keywordText = this.categories.map(c => (c.keywords || []).join('\n'))
          this.selected = []
          this.log('✓ 共提取 ' + this.categories.length + ' 个类别')
          this.message = res.data.message
          this.messageType = 'success'
        } else {
          this.message = res.data.message || '提取失败'
          this.messageType = 'error'
          this.log('✗ ' + this.message)
        }
      } catch (e) {
        this.message = '提取失败：' + (e.response?.data?.message || e.message)
        this.messageType = 'error'
        this.log('✗ ' + this.message)
      } finally {
        this.loading = false
      }
    },

    // 将编辑框中的关键词同步回 categories
    syncKeywords() {
      this.categories.forEach((c, idx) => {
        const text = (this.keywordText[idx] || '').split('\n')
          .map(s => s.trim()).filter(s => s.length > 0)
        c.keywords = text
      })
    },

    // ---- 第4步：搜索 ----
    async search() {
      if (this.categories.length === 0 || this.candidateFiles.length === 0) return
      this.syncKeywords()
      this.loading = true
      this.message = ''
      this.log('开始搜索（精准 → 模糊）...')
      const fd = new FormData()
      fd.append('categoriesJson', JSON.stringify(this.categories.map(c => ({
        parent: c.parent, type: c.type, keywords: c.keywords
      }))))
      this.candidateFiles.forEach(f => fd.append('names', f.name))
      try {
        const res = await axios.post('/api/hr/word-merge/search', fd)
        if (res.data.success) {
          this.categories = res.data.categories || []
          // 初始化勾选状态（默认全选）
          this.selected = this.categories.map(c => c.matchedFiles.map(() => true))
          const total = this.categories.reduce((s, c) => s + c.matchedFiles.length, 0)
          this.log('✓ 搜索完成，共匹配 ' + total + ' 个文件')
          this.categories.forEach(c => {
            const p = c.matchedFiles.filter(f => f.matchType === 'precise').length
            const z = c.matchedFiles.filter(f => f.matchType === 'fuzzy').length
            this.log('  · ' + c.parent + '：精确 ' + p + ' / 模糊 ' + z)
          })
          this.message = res.data.message
          this.messageType = 'success'
        } else {
          this.message = res.data.message || '搜索失败'
          this.messageType = 'error'
          this.log('✗ ' + this.message)
        }
      } catch (e) {
        this.message = '搜索失败：' + (e.response?.data?.message || e.message)
        this.messageType = 'error'
        this.log('✗ ' + this.message)
      } finally {
        this.loading = false
      }
    },

    selectedCount(idx) {
      const arr = this.selected[idx] || []
      return arr.filter(Boolean).length
    },
    preciseCount(idx) {
      const list = this.categories[idx]?.matchedFiles || []
      return list.filter(f => f.matchType === 'precise').length
    },
    fuzzyCount(idx) {
      const list = this.categories[idx]?.matchedFiles || []
      return list.filter(f => f.matchType === 'fuzzy').length
    },
    hasSelected() {
      return this.selected.some(arr => Array.isArray(arr) && arr.some(Boolean))
    },

    // ---- 第6步：合并 ----
    async merge() {
      if (!this.hasSelected()) return
      this.syncKeywords()
      this.loading = true
      this.message = ''
      this.downloadFileName = ''
      this.log('开始分类合并...')
      const fd = new FormData()
      const payload = this.categories.map((c, idx) => ({
        parent: c.parent,
        type: c.type,
        keywords: c.keywords,
        matchedFiles: c.matchedFiles.filter((f, fi) => this.selected[idx][fi])
      }))
      fd.append('categoriesJson', JSON.stringify(payload))
      this.candidateFiles.forEach(f => fd.append('files', f))
      try {
        const res = await postAsync('/api/hr/word-merge', fd)
        if (res.data.success) {
          this.downloadFileName = res.data.fileName
          this.log('✓ 合并完成 → ' + this.downloadFileName)
          if (Array.isArray(res.data.logs) && res.data.logs.length) {
            this.logs.push(...res.data.logs)
          }
          this.message = res.data.message
          this.messageType = 'success'
        } else {
          this.message = res.data.message || '合并失败'
          this.messageType = 'error'
          this.log('✗ ' + this.message)
        }
      } catch (e) {
        this.message = '合并失败：' + (e.response?.data?.message || e.message)
        this.messageType = 'error'
        this.log('✗ ' + this.message)
      } finally {
        this.loading = false
      }
    },

    downloadFile() {
      if (this.downloadFileName) {
        window.location.href = '/api/hr/download/' + encodeURIComponent(this.downloadFileName)
      }
    }
  }
}
</script>

<style scoped>
.page-container { padding: 4px; }
.func-page-header { margin-bottom: 8px; }
.func-page-title { margin: 0 0 4px 0; font-size: 20px; }
.func-page-desc { margin: 0; color: #888; font-size: 13px; }
.cat-edit-wrap { margin-top: 12px; }
.dir-upload {
  border: 2px dashed #dcdfe6; border-radius: 6px; padding: 30px 20px;
  text-align: center; cursor: pointer; transition: border-color .2s;
}
.dir-upload:hover { border-color: #409EFF; }
.dir-count { margin-top: 8px; color: #67C23A; font-weight: 600; font-size: 14px; }
.cat-edit { margin-bottom: 12px; padding: 10px 12px; border: 1px dashed #dcdfe6; border-radius: 4px; }
.cat-edit-title { margin-bottom: 6px; }
.cat-edit-count { color: #888; font-size: 12px; }
.hint { color: #888; font-size: 12px; margin-left: 10px; }
.result-cat { margin-bottom: 14px; }
.result-cat-title { font-weight: 600; margin-bottom: 6px; }
.title-tag { margin-left: 8px; font-weight: 400; }
.result-empty { color: #aaa; font-size: 12px; }
.result-item { margin: 4px 8px 4px 0; }
.result-label { margin-right: 6px; }
.result-tag { margin-left: 2px; }
.main-col { min-width: 0; }
.log-col {
  position: sticky; top: 16px;
  display: flex; flex-direction: column;
  max-height: calc(100vh - 40px);
}
.log-head {
  display: flex; align-items: center; justify-content: space-between;
  padding: 8px 12px; background: #2c3e50; color: #fff;
  border-radius: 4px 4px 0 0; font-size: 13px; font-weight: 600;
}
.log-head-actions { display: flex; align-items: center; gap: 8px; }
.log-count { color: #9fb3c8; font-weight: 400; font-size: 12px; }
.log-panel {
  background: #1f2d3d; color: #c0d0e0; font-family: Consolas, monospace;
  font-size: 12px; padding: 10px; border-radius: 0 0 4px 4px;
  flex: 1; overflow-y: auto; max-height: calc(100vh - 96px);
}
.log-line { line-height: 1.6; white-space: pre-wrap; word-break: break-all; }
.log-empty { color: #6b7c8f; font-style: italic; line-height: 1.8; }
@media (max-width: 991px) {
  .log-col { position: static; max-height: none; margin-top: 12px; }
  .log-panel { max-height: 260px; }
}
</style>
