<template>
  <el-card class="func-card">
    <div class="page-container">
    <div class="func-page-header">
      <h2 class="func-page-title">合并表格</h2>
      <p class="func-page-desc">
        绩效考核表的合并工具。支持处理模式（填月份列）、通用合并（原样堆叠）和多表合并（删列+标注）三种模式。
      </p>
    </div>

    <!-- 模式切换 -->
    <el-radio-group v-model="hbbgMode" @change="onHbbgModeChange" style="margin-bottom:12px">
      <el-radio-button value="process">处理模式（填月份列）</el-radio-button>
      <el-radio-button value="generic">通用合并（原样堆叠）</el-radio-button>
      <el-radio-button value="custom">多表合并（删列+标注）</el-radio-button>
    </el-radio-group>

    <!-- 文件选择 -->
    <el-upload
      ref="uploadHbbg"
      drag
      :auto-upload="false"
      multiple
      :on-change="handleHbbgFilesChange"
      accept=".xlsx,.xls"
    >
      <el-icon class="el-icon--upload"><UploadFilled /></el-icon>
      <div class="el-upload__text">将文件拖到此处，或 <em>点击选择</em></div>
    </el-upload>

    <!-- 处理模式参数 -->
    <el-form v-if="hbbgMode === 'process'" label-width="140px" style="margin-top:12px">
      <el-form-item label="目标列字母">
        <el-input v-model="formHbbg.targetCol" placeholder="如 AU" style="width: 120px" />
      </el-form-item>
    </el-form>

    <!-- 通用合并模式参数 -->
    <el-form v-if="hbbgMode === 'generic'" label-width="140px" style="margin-top:12px">
      <el-form-item label="文件间空行数">
        <el-input-number v-model="formHbbg.gapRows" :min="0" :max="100" style="width:120px" />
        <span style="margin-left:8px;color:#888;font-size:12px">行（0=不插空行）</span>
      </el-form-item>
      <el-form-item label="保留表头文件">
        <el-select v-model="formHbbg.keepHeaderFile" placeholder="（不指定）" clearable style="width:320px" :disabled="hbbgFiles.length === 0">
          <el-option v-for="(f, idx) in hbbgFiles" :key="idx" :label="f.name" :value="f.name" />
        </el-select>
        <span style="margin-left:8px;color:#888;font-size:12px">不选则全部保留表头</span>
      </el-form-item>
      <el-form-item label="跳过表头行数">
        <el-input-number v-model="formHbbg.skipRows" :min="0" :max="100" style="width:120px" />
        <span style="margin-left:8px;color:#888;font-size:12px">行（其余文件跳过前N行）</span>
      </el-form-item>

      <!-- 列移动规则（可选） -->
      <el-divider content-position="left" style="margin:16px 0 12px 0">列移动设置（可选）</el-divider>
      <div v-for="(rule, idx) in formHbbg.shiftRules" :key="idx" style="margin-bottom:12px;padding:12px;border:1px dashed #dcdfe6;border-radius:4px;">
        <el-form-item label="目标文件" style="margin-bottom:8px">
          <el-select
            v-model="rule.fileNames"
            placeholder="选择文件（可多选）"
            multiple
            collapse-tags
            collapse-tags-tooltip
            style="width:100%"
          >
            <el-option v-for="(f, fidx) in hbbgFiles" :key="fidx" :label="f.name" :value="f.name" />
          </el-select>
        </el-form-item>
        <el-form-item label="移动列" style="margin-bottom:8px">
          <el-select v-model="rule.colLetter" placeholder="选择列" style="width:120px">
            <el-option v-for="col in availableCols" :key="col" :label="col + '列'" :value="col" />
          </el-select>
        </el-form-item>
        <el-form-item label="右移格数" style="margin-bottom:8px">
          <el-input-number v-model="rule.shiftCount" :min="1" :max="20" style="width:120px" />
          <span style="margin-left:8px;color:#888;font-size:12px">格（整列数据向右移动N格，原位置留空）</span>
          <el-button type="danger" size="small" style="margin-left:12px" @click="formHbbg.shiftRules.splice(idx, 1)">删除此规则</el-button>
        </el-form-item>
      </div>
      <el-button v-if="hbbgFiles.length > 0" size="small" @click="addShiftRule">+ 添加列移动规则</el-button>
      <div v-if="formHbbg.shiftRules.length > 0" style="margin-top:8px;color:#888;font-size:12px">
        说明：合并前，对指定文件的某列数据整体向右移动N格，原位置（含表头）留空
      </div>
    </el-form>

    <!-- 多表合并模式参数 -->
    <el-form v-if="hbbgMode === 'custom'" label-width="120px" style="margin-top:12px">
      <el-form-item label="删除起始列">
        <el-input v-model="formHbbg.deleteColLetter" placeholder="如 AT" style="width: 120px" />
        <span style="margin-left:8px;color:#888;font-size:12px">含该列及之后所有列</span>
      </el-form-item>
      <el-form-item label="公司命名列">
        <el-input v-model="formHbbg.sheetNameColLetter" placeholder="如 AR" style="width: 120px" />
        <span style="margin-left:8px;color:#888;font-size:12px">列（所有数据行在该列写入工作表名，用于追溯来源）</span>
      </el-form-item>
      <el-form-item label="表头行数">
        <el-input-number v-model="formHbbg.headerRows" :min="1" :max="100" style="width:120px" />
        <span style="margin-left:8px;color:#888;font-size:12px">行（首个工作表保留全部，其余跳过前N行）</span>
      </el-form-item>

      <!-- 工作表勾选 -->
      <el-form-item label="勾选工作表">
        <div style="width:100%">
          <el-button-group style="margin-bottom:8px">
            <el-button size="small" @click="selectAllSheets">全选</el-button>
            <el-button size="small" @click="clearAllSheets">清空</el-button>
          </el-button-group>
          <el-tree
            ref="sheetTree"
            :data="sheetTreeData"
            show-checkbox
            node-key="key"
            :props="{ label: 'label', children: 'children' }"
            @check-change="onSheetCheckChange"
            style="border:1px solid #dcdfe6;border-radius:4px;padding:8px;max-height:260px;overflow-y:auto"
          >
            <template #default="{ node, data }">
              <span style="font-size:13px">
                <span v-if="node.level === 1" style="font-weight:600">📁 {{ data.label }}</span>
                <span v-else>{{ data.label }}<span v-if="data.hidden" style="color:#909399;font-size:11px">（隐藏）</span></span>
              </span>
            </template>
          </el-tree>
        </div>
      </el-form-item>
    </el-form>

    <el-button type="success" :loading="loading" @click="submitHbbg" style="margin-top:12px">
      开始合并
    </el-button>
    <el-progress v-if="hbbgProgress > 0 && hbbgProgress < 100" :percentage="Math.round(hbbgProgress)" style="margin-top:8px" />
    <div v-if="processingHint" style="margin-top:8px;color:#e6a23c;font-size:13px">
      {{ processingHint }}
    </div>

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

export default {
  name: 'JxkhHbbg',
  data() {
    return {
      loading: false,
      message: '',
      messageType: 'info',
      downloadFileName: '',
      hbbgFiles: [],
      hbbgMode: 'process',
      hbbgProgress: 0,
      processingHint: '',
      formHbbg: {
        targetCol: '',
        gapRows: 0,
        keepHeaderFile: '',
        skipRows: 1,
        shiftRules: [],
        deleteColLetter: '',
        sheetNameColLetter: '',
        headerRows: 5
      },
      sheetTreeData: []
    }
  },
  computed: {
    availableCols() {
      return 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.split('')
    }
  },
  methods: {
    handleHbbgFilesChange(file, fileList) {
      this.hbbgFiles = fileList.map(f => f.raw)
      if (this.hbbgMode === 'custom') {
        this.loadSheetsPreview()
      }
    },
    showMessage(text, type = 'info') {
      this.message = text
      this.messageType = type
    },
    onHbbgModeChange() {
      this.downloadFileName = ''
      this.message = ''
      this.hbbgProgress = 0
      this.processingHint = ''
      this.formHbbg.shiftRules = []
      if (this.hbbgMode === 'custom' && this.hbbgFiles.length > 0) {
        this.loadSheetsPreview()
      }
    },
    addShiftRule() {
      this.formHbbg.shiftRules.push({
        fileNames: [],
        colLetter: 'B',
        shiftCount: 1
      })
    },
    submitHbbg() {
      // 检测列移动规则中是否有重复文件
      if (this.formHbbg.shiftRules.length > 0) {
        const fileCount = {}
        const duplicateFiles = []
        for (const rule of this.formHbbg.shiftRules) {
          if (!rule.fileNames || rule.fileNames.length === 0) continue
          for (const fname of rule.fileNames) {
            fileCount[fname] = (fileCount[fname] || 0) + 1
          }
        }
        for (const [fileName, count] of Object.entries(fileCount)) {
          if (count > 1) duplicateFiles.push(fileName)
        }
        if (duplicateFiles.length > 0) {
          this.$message({
            message: `注意：以下文件设置了多条列移动规则（将按顺序依次执行）：${duplicateFiles.join('、')}`,
            type: 'warning',
            duration: 6000,
            showClose: true
          })
        }
      }

      if (this.hbbgFiles.length === 0) {
        this.showMessage('请选择至少一个Excel文件', 'warning')
        return
      }
      const formData = new FormData()
      this.hbbgFiles.forEach(f => formData.append('files', f))

      if (this.hbbgMode === 'process') {
        if (!this.formHbbg.targetCol) {
          this.showMessage('请填写目标列字母', 'warning')
          return
        }
        formData.append('targetCol', this.formHbbg.targetCol)
        this.post('/api/finance/jxkh/hbbg', formData)
      } else if (this.hbbgMode === 'custom') {
        if (!this.formHbbg.deleteColLetter || !this.formHbbg.sheetNameColLetter) {
          this.showMessage('请填写删除起始列和公司命名列', 'warning')
          return
        }
        const checkedNodes = this.$refs.sheetTree.getCheckedNodes(true)
        const sheetNodes = checkedNodes.filter(n => !n.children || n.children.length === 0)
        if (sheetNodes.length === 0) {
          this.showMessage('请至少勾选一个工作表', 'warning')
          return
        }
        const selections = sheetNodes.map(n => ({ fileName: n.fileName, sheetName: n.sheetName }))
        formData.append('sheetSelections', JSON.stringify(selections))
        formData.append('deleteColLetter', this.formHbbg.deleteColLetter)
        formData.append('sheetNameColLetter', this.formHbbg.sheetNameColLetter)
        formData.append('headerRows', this.formHbbg.headerRows)
        this.hbbgProgress = 5
        this.processingHint = '正在处理，大文件（如 27MB）解析与合并可能需要 1~2 分钟，请勿关闭页面...'
        this.post('/api/finance/jxkh/hbbg-custom', formData, () => {
          this.hbbgProgress = 100
          this.processingHint = ''
        })
      } else {
        formData.append('gapRows', this.formHbbg.gapRows)
        formData.append('keepHeaderFile', this.formHbbg.keepHeaderFile || '')
        formData.append('skipRows', this.formHbbg.skipRows)
        if (this.formHbbg.shiftRules.length > 0) {
          formData.append('shiftRules', JSON.stringify(this.formHbbg.shiftRules))
        }
        this.hbbgProgress = 10
        this.post('/api/finance/jxkh/hbbg-generic', formData, () => { this.hbbgProgress = 100 })
      }
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
          this.processingHint = ''
          this.hbbgProgress = 0
          this.showMessage(res.data.message || '处理失败', 'error')
        }
      }).catch(err => {
        this.processingHint = ''
        this.hbbgProgress = 0
        this.showMessage(err.response?.data?.message || err.message, 'error')
      }).finally(() => {
        this.loading = false
      })
    },
    loadSheetsPreview() {
      if (this.hbbgFiles.length === 0) { this.sheetTreeData = []; return }
      const formData = new FormData()
      this.hbbgFiles.forEach(f => formData.append('files', f))
      axios.post('/api/finance/jxkh/hbbg-sheets', formData).then(res => {
        if (res.data.success) {
          const sheets = res.data.sheets
          this.sheetTreeData = Object.entries(sheets).map(([fileName, sheetList]) => ({
            key: 'file_' + fileName,
            label: fileName,
            children: sheetList.map(s => ({
              key: fileName + '|||' + s.name,
              label: s.name,
              fileName: fileName,
              sheetName: s.name,
              hidden: s.hidden
            }))
          }))
        }
      }).catch(err => {
        this.showMessage('读取工作表列表失败：' + (err.response?.data?.message || err.message), 'error')
      })
    },
    selectAllSheets() {
      const checkKeys = []
      for (const node of this.sheetTreeData) {
        if (node.children) {
          for (const child of node.children) {
            checkKeys.push(child.key)
          }
        }
      }
      this.$refs.sheetTree.setCheckedKeys(checkKeys)
    },
    clearAllSheets() {
      this.$refs.sheetTree.setCheckedKeys([])
    },
    onSheetCheckChange() {
      // 树勾选状态已由 el-tree 管理，此钩子仅作预留
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
