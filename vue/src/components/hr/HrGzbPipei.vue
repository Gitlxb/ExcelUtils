<template>
  <el-card class="func-card">
    <!-- 页面标题区域 -->
    <div class="func-page-header">
      <h2 class="func-page-title">工资表匹配</h2>
      <p class="func-page-desc">
        工资表数据的交叉匹配与核对工具。支持甲方匹配内部、内部匹配甲方等多种匹配模式，自动标红差异单元格，快速定位数据不一致之处。
      </p>
    </div>

    <el-tabs v-model="activeMode" type="border-card">
      <!-- 模式A：甲方匹配内部 -->
      <el-tab-pane label="模式A：甲方匹配内部" name="modeA">
        <el-steps :active="modeAStep" finish-status="success" simple>
          <el-step title="上传文件与配置" />
          <el-step title="字段映射" />
          <el-step title="执行匹配" />
        </el-steps>

        <!-- 步骤0：上传文件 + 选择工作簿和表头（左右并排） -->
        <div v-if="modeAStep === 0" class="step-body">
          <div class="config-two-panels">
            <!-- 甲方配置 -->
            <div class="config-panel">
              <div class="config-title">甲方工资表</div>
              <el-upload
                drag
                :auto-upload="false"
                :on-change="(f) => handleFileA(f, 'jia')"
                accept=".xlsx,.xlsm"
                style="margin-bottom: 12px"
              >
                <el-icon class="el-icon--upload"><UploadFilled /></el-icon>
                <div class="el-upload__text">
                  {{ jiaFile ? jiaFile.name : '拖拽或点击上传甲方工资表' }}
                </div>
              </el-upload>

              <div class="form-row" v-if="jiaSheets.length">
                <el-select
                  v-model="jiaSheet"
                  placeholder="选择工作表"
                  @change="onJiaSheetChange"
                  style="width: 100%"
                >
                  <el-option
                    v-for="s in jiaSheets"
                    :key="s"
                    :label="s"
                    :value="s"
                  />
                </el-select>
              </div>

              <div class="form-row" v-if="jiaPreviewRows.length">
                <div class="header-label">表头行：</div>
                <el-radio-group v-model="jiaHeaderRow">
                  <el-radio
                    v-for="r in jiaPreviewRows"
                    :key="r.rowNum"
                    :label="r.rowNum"
                  >
                    第 {{ r.rowNum }} 行
                    <span class="preview-text">{{ formatPreviewRow(r.cells) }}</span>
                  </el-radio>
                </el-radio-group>
              </div>

              <div v-if="jiaPreviewRows.length" class="preview-box">
                <el-table :data="jiaPreviewRows" size="small" border style="width: 100%">
                  <el-table-column prop="rowNum" label="行号" width="60" />
                  <el-table-column
                    v-for="(_, idx) in jiaPreviewRows[0].cells"
                    :key="idx"
                    :label="'列' + (idx + 1)"
                    show-overflow-tooltip
                  >
                    <template #default="scope">{{ scope.row.cells[idx] }}</template>
                  </el-table-column>
                </el-table>
              </div>
            </div>

            <!-- 内部配置 -->
            <div class="config-panel">
              <div class="config-title">内部工资表</div>
              <el-upload
                drag
                :auto-upload="false"
                :on-change="(f) => handleFileA(f, 'nei')"
                accept=".xlsx,.xlsm"
                style="margin-bottom: 12px"
              >
                <el-icon class="el-icon--upload"><UploadFilled /></el-icon>
                <div class="el-upload__text">
                  {{ neiFile ? neiFile.name : '拖拽或点击上传内部工资表' }}
                </div>
              </el-upload>

              <div class="form-row" v-if="neiSheets.length">
                <el-select
                  v-model="neiSheet"
                  placeholder="选择工作表"
                  @change="onNeiSheetChange"
                  style="width: 100%"
                >
                  <el-option
                    v-for="s in neiSheets"
                    :key="s"
                    :label="s"
                    :value="s"
                  />
                </el-select>
              </div>

              <div class="form-row" v-if="neiPreviewRows.length">
                <div class="header-label">表头行：</div>
                <el-radio-group v-model="neiHeaderRow">
                  <el-radio
                    v-for="r in neiPreviewRows"
                    :key="r.rowNum"
                    :label="r.rowNum"
                  >
                    第 {{ r.rowNum }} 行
                    <span class="preview-text">{{ formatPreviewRow(r.cells) }}</span>
                  </el-radio>
                </el-radio-group>
              </div>

              <div v-if="neiPreviewRows.length" class="preview-box">
                <el-table :data="neiPreviewRows" size="small" border style="width: 100%">
                  <el-table-column prop="rowNum" label="行号" width="60" />
                  <el-table-column
                    v-for="(_, idx) in neiPreviewRows[0].cells"
                    :key="idx"
                    :label="'列' + (idx + 1)"
                    show-overflow-tooltip
                  >
                    <template #default="scope">{{ scope.row.cells[idx] }}</template>
                  </el-table-column>
                </el-table>
              </div>
            </div>
          </div>

          <div class="step-actions">
            <el-button
              type="primary"
              :disabled="!jiaFile || !neiFile || !jiaSheet || !jiaHeaderRow || !neiSheet || !neiHeaderRow"
              @click="confirmAConfig"
            >
              确定
            </el-button>
          </div>
        </div>

        <!-- 步骤1：字段映射 -->
        <div v-if="modeAStep === 1" class="step-body">
          <div class="key-config">
            <div class="form-row">
              <span class="header-label">甲方关键列：</span>
              <el-select v-model="jiaKey" placeholder="选择关键列（用于行匹配）" style="width: 220px">
                <el-option v-for="h in jiaHeaders" :key="h" :label="h" :value="h" />
              </el-select>
            </div>
            <div class="form-row">
              <span class="header-label">内部关键列：</span>
              <el-select v-model="neiKey" placeholder="选择关键列（用于行匹配）" style="width: 220px">
                <el-option v-for="h in neiHeaders" :key="h" :label="h" :value="h" />
              </el-select>
            </div>
          </div>
          <div class="mapping-box">
            <div class="mapping-col">
              <div class="mapping-title">甲方字段</div>
              <div class="field-list">
                <div
                  v-for="h in jiaHeaders"
                  :key="h"
                  :class="['field-item', selectedJiaField === h ? 'active' : '']"
                  @click="selectedJiaField = h"
                >
                  {{ h }}
                </div>
              </div>
            </div>
            <div class="mapping-op">
              <el-button
                type="primary"
                :disabled="!selectedJiaField || !selectedNeiField"
                @click="addMapping"
              >
                添加映射 →
              </el-button>
              <el-button type="success" plain @click="autoMatch">
                智能匹配
              </el-button>
            </div>
            <div class="mapping-col">
              <div class="mapping-title">内部字段</div>
              <div class="field-list">
                <div
                  v-for="h in neiHeaders"
                  :key="h"
                  :class="['field-item', selectedNeiField === h ? 'active' : '']"
                  @click="selectedNeiField = h"
                >
                  {{ h }}
                </div>
              </div>
            </div>
            <div class="mapping-result">
              <div class="mapping-title">已添加映射（按此顺序匹配）</div>
              <div class="field-list">
                <div
                  v-for="(m, idx) in fieldMappings"
                  :key="idx"
                  class="field-item mapping-item"
                >
                  <span>{{ m[0] }} ↔ {{ m[1] }}</span>
                  <el-button
                    type="danger"
                    link
                    size="small"
                    @click="removeMapping(idx)"
                  >
                    删除
                  </el-button>
                </div>
              </div>
              <el-button
                v-if="fieldMappings.length"
                type="warning"
                link
                size="small"
                @click="fieldMappings = []"
              >
                清空全部
              </el-button>
            </div>
          </div>
          <div class="step-actions">
            <el-button @click="resetModeA">上一步</el-button>
            <el-button
              type="primary"
              :disabled="!fieldMappings.length"
              @click="modeAStep = 2"
            >
              下一步
            </el-button>
          </div>
        </div>

        <!-- 步骤2：执行 -->
        <div v-if="modeAStep === 2" class="step-body">
          <el-alert
            title="准备执行匹配"
            type="info"
            :closable="false"
            show-icon
          >
            <div>甲方：{{ jiaSheet }}（表头第{{ jiaHeaderRow }}行）</div>
            <div>内部：{{ neiSheet }}（表头第{{ neiHeaderRow }}行）</div>
            <div>关键列：甲方「{{ jiaKey }}」 ↔ 内部「{{ neiKey }}」</div>
            <div>映射数：{{ fieldMappings.length }} 组</div>
          </el-alert>
          <div class="step-actions">
            <el-button @click="modeAStep = 1">上一步</el-button>
            <el-button type="primary" :loading="loading" @click="executeModeA">
              开始匹配
            </el-button>
          </div>
        </div>
      </el-tab-pane>

      <!-- 模式B：内部互相匹配 -->
      <el-tab-pane label="模式B：内部互相匹配" name="modeB">
        <el-steps :active="modeBStep" finish-status="success" simple>
          <el-step title="上传文件与配置" />
          <el-step title="字段映射" />
          <el-step title="执行匹配" />
        </el-steps>

        <!-- 步骤0：上传文件 + 选择工作簿和表头（左右并排） -->
        <div v-if="modeBStep === 0" class="step-body">
          <div class="upload-single-center" style="margin-bottom: 16px">
            <el-upload
              drag
              :auto-upload="false"
              :on-change="handleFileB"
              accept=".xlsx,.xlsm"
              style="width: 400px"
            >
              <el-icon class="el-icon--upload"><UploadFilled /></el-icon>
              <div class="el-upload__text">
                {{ bFile ? bFile.name : '拖拽或点击上传内部工资表（含多个工作表）' }}
              </div>
            </el-upload>
          </div>

          <div v-if="bFile" class="config-two-panels">
            <!-- 工作簿A -->
            <div class="config-panel">
              <div class="config-title">工作簿 A</div>
              <div class="form-row">
                <el-select
                  v-model="sheetA"
                  placeholder="选择工作表A"
                  @change="onSheetAChange"
                  style="width: 100%"
                >
                  <el-option
                    v-for="s in bSheets"
                    :key="s"
                    :label="s"
                    :value="s"
                  />
                </el-select>
              </div>
              <div class="form-row" v-if="previewA.length">
                <div class="header-label">表头行：</div>
                <el-radio-group v-model="headerA">
                  <el-radio
                    v-for="r in previewA"
                    :key="r.rowNum"
                    :label="r.rowNum"
                  >
                    第 {{ r.rowNum }} 行
                    <span class="preview-text">{{ formatPreviewRow(r.cells) }}</span>
                  </el-radio>
                </el-radio-group>
              </div>
              <div v-if="previewA.length" class="preview-box">
                <el-table :data="previewA" size="small" border style="width: 100%">
                  <el-table-column prop="rowNum" label="行号" width="60" />
                  <el-table-column
                    v-for="(_, idx) in previewA[0].cells"
                    :key="idx"
                    :label="'列' + (idx + 1)"
                    show-overflow-tooltip
                  >
                    <template #default="scope">{{ scope.row.cells[idx] }}</template>
                  </el-table-column>
                </el-table>
              </div>
            </div>

            <!-- 工作簿B -->
            <div class="config-panel">
              <div class="config-title">工作簿 B</div>
              <div class="form-row">
                <el-select
                  v-model="sheetB"
                  placeholder="选择工作表B"
                  @change="onSheetBChange"
                  style="width: 100%"
                >
                  <el-option
                    v-for="s in bSheets"
                    :key="s"
                    :label="s"
                    :value="s"
                  />
                </el-select>
              </div>
              <div class="form-row" v-if="previewB.length">
                <div class="header-label">表头行：</div>
                <el-radio-group v-model="headerB">
                  <el-radio
                    v-for="r in previewB"
                    :key="r.rowNum"
                    :label="r.rowNum"
                  >
                    第 {{ r.rowNum }} 行
                    <span class="preview-text">{{ formatPreviewRow(r.cells) }}</span>
                  </el-radio>
                </el-radio-group>
              </div>
              <div v-if="previewB.length" class="preview-box">
                <el-table :data="previewB" size="small" border style="width: 100%">
                  <el-table-column prop="rowNum" label="行号" width="60" />
                  <el-table-column
                    v-for="(_, idx) in previewB[0].cells"
                    :key="idx"
                    :label="'列' + (idx + 1)"
                    show-overflow-tooltip
                  >
                    <template #default="scope">{{ scope.row.cells[idx] }}</template>
                  </el-table-column>
                </el-table>
              </div>
            </div>
          </div>

          <div class="step-actions">
            <el-button
              type="primary"
              :disabled="!bFile || !sheetA || !headerA || !sheetB || !headerB"
              @click="confirmBConfig"
            >
              确定
            </el-button>
          </div>
        </div>

        <!-- 步骤1：字段映射 -->
        <div v-if="modeBStep === 1" class="step-body">
          <div class="key-config">
            <div class="form-row">
              <span class="header-label">工作簿A关键列：</span>
              <el-select v-model="keyA" placeholder="选择关键列（用于行匹配）" style="width: 220px">
                <el-option v-for="h in headersA" :key="h" :label="h" :value="h" />
              </el-select>
            </div>
            <div class="form-row">
              <span class="header-label">工作簿B关键列：</span>
              <el-select v-model="keyB" placeholder="选择关键列（用于行匹配）" style="width: 220px">
                <el-option v-for="h in headersB" :key="h" :label="h" :value="h" />
              </el-select>
            </div>
          </div>
          <div class="mapping-box">
            <div class="mapping-col">
              <div class="mapping-title">工作表A字段</div>
              <div class="field-list">
                <div
                  v-for="h in headersA"
                  :key="h"
                  :class="['field-item', selectedAField === h ? 'active' : '']"
                  @click="selectedAField = h"
                >
                  {{ h }}
                </div>
              </div>
            </div>
            <div class="mapping-op">
              <el-button
                type="primary"
                :disabled="!selectedAField || !selectedBField"
                @click="addMappingB"
              >
                添加映射 →
              </el-button>
              <el-button type="success" plain @click="autoMatchB">
                智能匹配
              </el-button>
            </div>
            <div class="mapping-col">
              <div class="mapping-title">工作表B字段</div>
              <div class="field-list">
                <div
                  v-for="h in headersB"
                  :key="h"
                  :class="['field-item', selectedBField === h ? 'active' : '']"
                  @click="selectedBField = h"
                >
                  {{ h }}
                </div>
              </div>
            </div>
            <div class="mapping-result">
              <div class="mapping-title">已添加映射</div>
              <div class="field-list">
                <div
                  v-for="(m, idx) in fieldMappingsB"
                  :key="idx"
                  class="field-item mapping-item"
                >
                  <span>{{ m[0] }} ↔ {{ m[1] }}</span>
                  <el-button
                    type="danger"
                    link
                    size="small"
                    @click="removeMappingB(idx)"
                  >
                    删除
                  </el-button>
                </div>
              </div>
              <el-button
                v-if="fieldMappingsB.length"
                type="warning"
                link
                size="small"
                @click="fieldMappingsB = []"
              >
                清空全部
              </el-button>
            </div>
          </div>
          <div class="step-actions">
            <el-button @click="resetModeB">上一步</el-button>
            <el-button
              type="primary"
              :disabled="!fieldMappingsB.length"
              @click="modeBStep = 2"
            >
              下一步
            </el-button>
          </div>
        </div>

        <!-- 步骤2：执行 -->
        <div v-if="modeBStep === 2" class="step-body">
          <el-alert
            title="准备执行匹配"
            type="info"
            :closable="false"
            show-icon
          >
            <div>工作表A：{{ sheetA }}（表头第{{ headerA }}行）</div>
            <div>工作表B：{{ sheetB }}（表头第{{ headerB }}行）</div>
            <div>关键列：A「{{ keyA }}」 ↔ B「{{ keyB }}」</div>
            <div>映射数：{{ fieldMappingsB.length }} 组</div>
          </el-alert>
          <div class="step-actions">
            <el-button @click="modeBStep = 1">上一步</el-button>
            <el-button type="primary" :loading="loading" @click="executeModeB">
              开始匹配
            </el-button>
          </div>
        </div>
      </el-tab-pane>
    </el-tabs>

    <div v-if="downloadUrl" class="result-section">
      <el-alert
        :title="`处理完成！共匹配 ${compareRows} 人，发现 ${diffCount} 个差异单元格`"
        type="success"
        :closable="false"
        show-icon
      />
      <div v-if="totalInfo" class="result-detail">
        <div class="detail-line">
          <span>基准表数据行：{{ totalInfo.baseDataRows }} 人{{ totalInfo.baseHasTotal ? '（含合计）' : '' }}</span>
          <span>目标表数据行：{{ totalInfo.targetDataRows }} 人{{ totalInfo.targetHasTotal ? '（含合计）' : '' }}</span>
        </div>
        <div class="detail-line" v-if="totalInfo.baseHasTotal || totalInfo.targetHasTotal">
          <span v-if="totalInfo.participated" class="ok">合计行参与对比（两边数据行数相等）</span>
          <span v-else class="warn">合计行未参与对比（仅一侧有合计，或两边数据行数不等）</span>
        </div>
        <div class="detail-line" v-if="unmatchedKeys.length">
          <span class="warn">基准表未匹配人员（{{ unmatchedKeys.length }} 人）：{{ unmatchedKeys.join('、') }}</span>
        </div>
        <div class="detail-line" v-if="totalInfo.extraTargetKeys && totalInfo.extraTargetKeys.length">
          <span class="warn">目标表多出未参与匹配人员（{{ totalInfo.extraTargetKeys.length }} 人）：{{ totalInfo.extraTargetKeys.join('、') }}</span>
        </div>
        <div class="detail-line" v-if="!unmatchedKeys.length && !(totalInfo.extraTargetKeys && totalInfo.extraTargetKeys.length)">
          <span class="ok">所有人员均已完成匹配</span>
        </div>
      </div>
      <el-button type="primary" @click="downloadFile" class="download-btn">
        <el-icon><Download /></el-icon> 下载匹配结果
      </el-button>
    </div>
    <div v-if="errorMsg" class="result-section">
      <el-alert :title="errorMsg" type="error" :closable="false" show-icon />
    </div>
  </el-card>
</template>

<script>
import axios from 'axios'

export default {
  name: 'GzbPipeiCard',
  data() {
    return {
      activeMode: 'modeA',
      loading: false,
      downloadUrl: '',
      errorMsg: '',
      compareRows: 0,
      diffCount: 0,
      unmatchedKeys: [],
      totalInfo: null,

      // 模式A
      modeAStep: 0,
      jiaFile: null,
      jiaSheets: [],
      jiaSheet: '',
      jiaHeaderRow: null,
      jiaPreviewRows: [],
      jiaHeaders: [],
      jiaKey: '',
      selectedJiaField: '',
      neiFile: null,
      neiSheets: [],
      neiSheet: '',
      neiHeaderRow: null,
      neiPreviewRows: [],
      neiHeaders: [],
      neiKey: '',
      selectedNeiField: '',
      fieldMappings: [],

      // 模式B
      modeBStep: 0,
      bFile: null,
      bSheets: [],
      sheetA: '',
      headerA: null,
      previewA: [],
      headersA: [],
      keyA: '',
      selectedAField: '',
      sheetB: '',
      headerB: null,
      previewB: [],
      headersB: [],
      keyB: '',
      selectedBField: '',
      fieldMappingsB: []
    }
  },
  methods: {
    // ---------- 通用方法 ----------

    formatPreviewRow(cells) {
      const vals = []
      for (let i = 0; i < Math.min(cells.length, 12); i++) {
        const v = cells[i]
        if (v !== null && v !== undefined && String(v).trim() !== '') {
          vals.push(`列${i + 1}:[${v}]`)
        }
      }
      return vals.length ? vals.join('  ') : '(空行)'
    },

    async uploadAndGetSheets(file, target) {
      const formData = new FormData()
      formData.append('file', file)
      const res = await axios.post('/api/hr/gzb-pipei/sheets', formData)
      if (!res.data.success) throw new Error(res.data.message)
      if (target === 'jia') {
        this.jiaSheets = res.data.sheets
        this.jiaSheet = res.data.sheets[0] || ''
      } else if (target === 'nei') {
        this.neiSheets = res.data.sheets
        this.neiSheet = res.data.sheets[0] || ''
      } else if (target === 'b') {
        this.bSheets = res.data.sheets
      }
    },

    async getPreview(file, sheetName, target) {
      const formData = new FormData()
      formData.append('file', file)
      formData.append('sheetName', sheetName)
      const res = await axios.post('/api/hr/gzb-pipei/preview', formData)
      if (!res.data.success) throw new Error(res.data.message)
      if (target === 'jia') this.jiaPreviewRows = res.data.rows
      else if (target === 'nei') this.neiPreviewRows = res.data.rows
      else if (target === 'A') this.previewA = res.data.rows
      else if (target === 'B') this.previewB = res.data.rows
    },

    async getHeaders(file, sheetName, headerRow) {
      const formData = new FormData()
      formData.append('file', file)
      formData.append('sheetName', sheetName)
      formData.append('headerRow', headerRow)
      const res = await axios.post('/api/hr/gzb-pipei/headers', formData)
      if (!res.data.success) throw new Error(res.data.message)
      return res.data.headers
    },

    resetResult() {
      this.downloadUrl = ''
      this.errorMsg = ''
      this.compareRows = 0
      this.diffCount = 0
      this.unmatchedKeys = []
      this.totalInfo = null
    },

    // 智能推荐关键列：优先含「姓名/名字/名称」的字段，否则取第一项
    autoPickKey(headers) {
      if (!headers || !headers.length) return ''
      const hit = headers.find((h) => h.includes('姓名') || h.includes('名字') || h.includes('名称'))
      return hit || headers[0]
    },

    // ---------- 模式A ----------

    async handleFileA(uploadFile, target) {
      const raw = uploadFile.raw
      this.resetResult()
      if (target === 'jia') {
        this.jiaFile = raw
        this.jiaSheets = []
        this.jiaSheet = ''
        this.jiaHeaderRow = null
        this.jiaPreviewRows = []
        this.jiaHeaders = []
        await this.uploadAndGetSheets(raw, 'jia')
        if (this.jiaSheet) {
          await this.getPreview(raw, this.jiaSheet, 'jia')
          if (this.jiaPreviewRows.length) {
            this.jiaHeaderRow = this.jiaPreviewRows[0].rowNum
          }
        }
      } else {
        this.neiFile = raw
        this.neiSheets = []
        this.neiSheet = ''
        this.neiHeaderRow = null
        this.neiPreviewRows = []
        this.neiHeaders = []
        await this.uploadAndGetSheets(raw, 'nei')
        if (this.neiSheet) {
          await this.getPreview(raw, this.neiSheet, 'nei')
          if (this.neiPreviewRows.length) {
            this.neiHeaderRow = this.neiPreviewRows[0].rowNum
          }
        }
      }
    },

    async onJiaSheetChange() {
      this.jiaPreviewRows = []
      this.jiaHeaderRow = null
      if (this.jiaSheet && this.jiaFile) {
        await this.getPreview(this.jiaFile, this.jiaSheet, 'jia')
        if (this.jiaPreviewRows.length) {
          this.jiaHeaderRow = this.jiaPreviewRows[0].rowNum
        }
      }
    },

    async onNeiSheetChange() {
      this.neiPreviewRows = []
      this.neiHeaderRow = null
      if (this.neiSheet && this.neiFile) {
        await this.getPreview(this.neiFile, this.neiSheet, 'nei')
        if (this.neiPreviewRows.length) {
          this.neiHeaderRow = this.neiPreviewRows[0].rowNum
        }
      }
    },

    async confirmAConfig() {
      this.jiaHeaders = await this.getHeaders(
        this.jiaFile,
        this.jiaSheet,
        this.jiaHeaderRow
      )
      this.neiHeaders = await this.getHeaders(
        this.neiFile,
        this.neiSheet,
        this.neiHeaderRow
      )
      this.jiaKey = this.autoPickKey(this.jiaHeaders)
      this.neiKey = this.autoPickKey(this.neiHeaders)
      this.modeAStep = 1
    },

    resetModeA() {
      this.modeAStep = 0
      this.jiaHeaders = []
      this.neiHeaders = []
      this.jiaKey = ''
      this.neiKey = ''
      this.selectedJiaField = ''
      this.selectedNeiField = ''
      this.fieldMappings = []
    },

    addMapping() {
      const exists = this.fieldMappings.some(
        (m) => m[0] === this.selectedJiaField && m[1] === this.selectedNeiField
      )
      if (exists) return
      this.fieldMappings.push([this.selectedJiaField, this.selectedNeiField])
      this.selectedJiaField = ''
      this.selectedNeiField = ''
    },

    autoMatch() {
      let count = 0
      for (const h1 of this.jiaHeaders) {
        if (this.neiHeaders.includes(h1)) {
          const exists = this.fieldMappings.some((m) => m[0] === h1 && m[1] === h1)
          if (!exists) {
            this.fieldMappings.push([h1, h1])
            count++
          }
        }
      }
      this.$message.success(`智能匹配完成，自动建立 ${count} 组映射`)
    },

    removeMapping(idx) {
      this.fieldMappings.splice(idx, 1)
    },

    async executeModeA() {
      this.loading = true
      this.resetResult()
      try {
        const formData = new FormData()
        formData.append('jiaFile', this.jiaFile)
        formData.append('neiFile', this.neiFile)
        formData.append('jiaSheet', this.jiaSheet)
        formData.append('neiSheet', this.neiSheet)
        formData.append('jiaHeaderRow', this.jiaHeaderRow)
        formData.append('neiHeaderRow', this.neiHeaderRow)
        formData.append('jiaKey', this.jiaKey)
        formData.append('neiKey', this.neiKey)
        formData.append('fieldMapping', JSON.stringify(this.fieldMappings))
        const res = await axios.post('/api/hr/gzb-pipei/mode-a', formData)
        if (res.data.success) {
          this.downloadUrl = `/api/hr/download/${res.data.fileName}`
          this.compareRows = res.data.compareRows
          this.diffCount = res.data.diffCount
          this.unmatchedKeys = res.data.unmatchedKeys || []
          this.totalInfo = res.data.totalInfo || null
        } else {
          this.errorMsg = res.data.message
        }
      } catch (e) {
        this.errorMsg = e.response?.data?.message || e.message
      } finally {
        this.loading = false
      }
    },

    // ---------- 模式B ----------

    async handleFileB(uploadFile) {
      const raw = uploadFile.raw
      this.resetResult()
      this.bFile = raw
      this.bSheets = []
      this.sheetA = ''
      this.headerA = null
      this.previewA = []
      this.headersA = []
      this.sheetB = ''
      this.headerB = null
      this.previewB = []
      this.headersB = []
      await this.uploadAndGetSheets(raw, 'b')
    },

    async onSheetAChange() {
      this.previewA = []
      this.headerA = null
      if (this.sheetA && this.bFile) {
        await this.getPreview(this.bFile, this.sheetA, 'A')
        if (this.previewA.length) {
          this.headerA = this.previewA[0].rowNum
        }
      }
    },

    async onSheetBChange() {
      this.previewB = []
      this.headerB = null
      if (this.sheetB && this.bFile) {
        await this.getPreview(this.bFile, this.sheetB, 'B')
        if (this.previewB.length) {
          this.headerB = this.previewB[0].rowNum
        }
      }
    },

    async confirmBConfig() {
      this.headersA = await this.getHeaders(this.bFile, this.sheetA, this.headerA)
      this.headersB = await this.getHeaders(this.bFile, this.sheetB, this.headerB)
      this.keyA = this.autoPickKey(this.headersA)
      this.keyB = this.autoPickKey(this.headersB)
      this.modeBStep = 1
    },

    resetModeB() {
      this.modeBStep = 0
      this.headersA = []
      this.headersB = []
      this.keyA = ''
      this.keyB = ''
      this.selectedAField = ''
      this.selectedBField = ''
      this.fieldMappingsB = []
    },

    addMappingB() {
      const exists = this.fieldMappingsB.some(
        (m) => m[0] === this.selectedAField && m[1] === this.selectedBField
      )
      if (exists) return
      this.fieldMappingsB.push([this.selectedAField, this.selectedBField])
      this.selectedAField = ''
      this.selectedBField = ''
    },

    autoMatchB() {
      let count = 0
      for (const h of this.headersA) {
        if (this.headersB.includes(h)) {
          const exists = this.fieldMappingsB.some((m) => m[0] === h && m[1] === h)
          if (!exists) {
            this.fieldMappingsB.push([h, h])
            count++
          }
        }
      }
      this.$message.success(`智能匹配完成，自动建立 ${count} 组映射`)
    },

    removeMappingB(idx) {
      this.fieldMappingsB.splice(idx, 1)
    },

    async executeModeB() {
      this.loading = true
      this.resetResult()
      try {
        const formData = new FormData()
        formData.append('file', this.bFile)
        formData.append('sheetA', this.sheetA)
        formData.append('sheetB', this.sheetB)
        formData.append('headerA', this.headerA)
        formData.append('headerB', this.headerB)
        formData.append('keyA', this.keyA)
        formData.append('keyB', this.keyB)
        formData.append('fieldMapping', JSON.stringify(this.fieldMappingsB))
        const res = await axios.post('/api/hr/gzb-pipei/mode-b', formData)
        if (res.data.success) {
          this.downloadUrl = `/api/hr/download/${res.data.fileName}`
          this.compareRows = res.data.compareRows
          this.diffCount = res.data.diffCount
          this.unmatchedKeys = res.data.unmatchedKeys || []
          this.totalInfo = res.data.totalInfo || null
        } else {
          this.errorMsg = res.data.message
        }
      } catch (e) {
        this.errorMsg = e.response?.data?.message || e.message
      } finally {
        this.loading = false
      }
    },

    // ---------- 通用 ----------

    downloadFile() {
      if (this.downloadUrl) window.location.href = this.downloadUrl
    }
  }
}
</script>

<style scoped src="../../styles/components/GzbPipeiCard.css"></style>

<style scoped src="../../styles/components/FuncPageHeader.css"></style>
