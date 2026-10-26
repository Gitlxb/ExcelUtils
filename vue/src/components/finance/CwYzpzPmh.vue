<template>
  <el-card class="func-card">
    <!-- 页面标题区域 -->
    <div class="func-page-header">
      <h2 class="func-page-title">预支平账</h2>
      <p class="func-page-desc">
        代工费预支数据的核对与平账处理工具。通过比对预支表与发放记录表，自动计算差异并生成平账结果，支持多轮次顺序执行。
      </p>
    </div>

    <!-- 步骤进度条 -->
    <el-steps :active="currentStep" align-center finish-status="success" style="margin-bottom:20px">
      <el-step title="核对" description="上传文件+选择工作簿+执行核对" />
      <el-step title="平账" description="执行平账处理" />
      <el-step title="保存" description="下载结果或继续下一轮" />
    </el-steps>

    <!-- 步骤 ①：核对 -->
    <div v-if="currentStep === 0" class="step-content">
      <h3>① 核对（第{{ yzpzRound }}轮）</h3>
      <p>上传预支明细表和工资表，一键完成：准备 → 复制姓名 → 姓名匹配 → 处理扣款</p>

      <!-- ======== 主文件区域 ======== -->
      <el-divider content-position="left">主文件（xxxx年预支明细）</el-divider>
      <el-upload
        ref="uploadMain"
        drag
        accept=".xlsx"
        :auto-upload="false"
        :show-file-list="true"
        :limit="1"
        :disabled="mainUploadDisabled"
        :on-change="onMainFileSelected"
        :on-exceed="onExceed"
      >
        <el-icon class="el-icon--upload"><UploadFilled /></el-icon>
        <div class="el-upload__text">拖拽或点击上传主文件（xxxx年预支明细.xlsx）</div>
      </el-upload>
      <!-- 主文件工作簿 + 单位 -->
      <div class="param-row" v-if="sheetOptions.length > 0" style="margin-top:10px">
        <el-select v-model="mainSheet" placeholder="选择主文件工作簿" style="width:220px;margin-right:10px" @change="onMainSheetChange">
          <el-option v-for="s in sheetOptions" :key="s" :label="s" :value="s" />
        </el-select>
        <!-- 单位多选（续轮时保留已选值） -->
        <el-select v-model="selectedUnits" multiple collapse-tags filterable default-first-option
                   placeholder="选择单位（可选）" style="width:300px">
          <template #prefix>
            <span style="color:#909399;margin-right:4px">单位</span>
          </template>
          <el-option v-for="u in detectedUnitOptions" :key="u" :label="u" :value="u" />
        </el-select>
        <el-tag v-if="detectedUnitOptions.length === 0 && mainSheet" type="info" size="small" style="margin-left:8px">检测单位中...</el-tag>
      </div>

      <!-- ======== 工资表区域 ======== -->
      <el-divider content-position="left">工资表（包含"需扣回、预支扣回、预支扣款"）</el-divider>
      <el-upload
        ref="uploadDed"
        drag
        accept=".xlsx"
        :auto-upload="false"
        :show-file-list="true"
        :limit="1"
        :disabled="dedUploadDisabled"
        :on-change="onDedFileSelected"
        :on-exceed="onExceed"
      >
        <el-icon class="el-icon--upload"><UploadFilled /></el-icon>
        <div class="el-upload__text">拖拽或点击上传工资表（包含"需扣回、预支扣回、预支扣款"）</div>
      </el-upload>
      <!-- 工资表工作簿 + 列类型 -->
      <div class="param-row" v-if="dedSheetOptions.length > 0" style="margin-top:10px">
        <el-select v-model="dedSheet" placeholder="选择工资表工作簿" style="width:220px;margin-right:10px" @change="detectColTypes">
          <el-option v-for="s in dedSheetOptions" :key="s" :label="s" :value="s" />
        </el-select>
        <el-select v-model="colType" placeholder="选择列类型" style="width:180px" :disabled="colTypeOptions.length === 0">
          <el-option v-for="c in colTypeOptions" :key="c" :label="c" :value="c" />
        </el-select>
        <el-tag v-if="colTypeOptions.length === 0 && dedSheet" type="info" size="small" style="margin-left:8px">检测列类型中...</el-tag>
      </div>

      <!-- 执行按钮 -->
      <el-button type="primary" :loading="loading" :disabled="!canExecCheck" @click="execCheck" style="margin-top:16px">
        【核对】
      </el-button>

      <div v-if="resultMsg" :class="resultOk ? 'result-ok' : 'result-err'" style="margin-top:10px">{{ resultMsg }}</div>
    </div>

    <!-- 步骤 ②：平账 -->
    <div v-if="currentStep === 1" class="step-content">
      <h3>② 平账（第{{ yzpzRound }}轮）</h3>
      <p>执行平账处理（拆分并标记）</p>

      <!-- 主文件（上一步结果） -->
      <div class="file-section">
        <div class="file-label">主文件（上一步核对结果）：</div>
        <el-button size="small" @click="usePrevResult">{{ prevFileName || '点击加载上一步结果文件' }}</el-button>
      </div>

      <div class="param-row" v-if="salarySheetOptions.length > 0">
        <el-select v-model="salarySheet" placeholder="选择工资表" style="width:200px;margin-right:10px">
          <el-option v-for="s in salarySheetOptions" :key="s" :label="s" :value="s" />
        </el-select>
        <el-select v-model="yzpzMonth" placeholder="选择月份" style="width:180px">
          <el-option v-for="m in monthOptions" :key="m" :label="m" :value="m" />
        </el-select>
      </div>

      <el-button type="primary" :loading="loading" :disabled="!mainFile || !salarySheet || !yzpzMonth" @click="execBalance">
        【平账】
      </el-button>

      <div v-if="resultMsg" :class="resultOk ? 'result-ok' : 'result-err'">{{ resultMsg }}</div>
    </div>

    <!-- 步骤 ③：保存 -->
    <div v-if="currentStep === 2" class="step-content">
      <h3>③ 保存（第{{ yzpzRound }}轮完成）</h3>
      <p>所有步骤已完成，下载最终文件，或继续下一轮核对平账</p>
      <el-alert title="预支平账处理完成！" type="success" :closable="false" show-icon />
      <div class="result-section">
        <el-button type="primary" @click="doDownloadFinal" class="download-btn">
          <el-icon><Download /></el-icon> 下载最终文件
        </el-button>
      </div>
      <div style="margin-top:12px">
        <el-button type="success" @click="continueNextRound">进行第{{ yzpzRound + 1 }}轮核对平账</el-button>
        <el-button @click="resetAll" style="margin-left:10px">重新开始</el-button>
      </div>
    </div>

    <!-- 多单位选择弹窗 -->
    <el-dialog v-model="unitDialogVisible" title="选择单位" width="360px" :close-on-click-modal="false">
      <p style="margin-bottom:12px">检测到多个单位，请选择要筛选的单位：</p>
      <el-checkbox-group v-model="unitChecked">
        <el-checkbox v-for="u in unitOptions" :key="u" :label="u">{{ u }}</el-checkbox>
      </el-checkbox-group>
      <template #footer>
        <el-button @click="unitDialogVisible = false">取消</el-button>
        <el-button type="primary" @click="confirmUnits" :disabled="unitChecked.length === 0">确定</el-button>
      </template>
    </el-dialog>

    <!-- 操作日志 -->
    <div class="log-wrapper">
      <div class="log-header" @click="showLog = !showLog">
        <span>操作日志</span>
        <el-icon :class="{ 'is-rotate': showLog }"><ArrowDown /></el-icon>
      </div>
      <div v-if="showLog" class="log-body">
        <el-input type="textarea" v-model="logText" readonly :rows="6" />
      </div>
    </div>
  </el-card>
</template>

<script>
import axios from 'axios'
import { postAsync } from '../../utils/asyncTask'
import { ArrowDown, UploadFilled, Download } from '@element-plus/icons-vue'

export default {
  components: { ArrowDown, UploadFilled, Download },
  name: 'YzpzCard',
  data() {
    return {
      currentStep: 0,
      // 文件
      mainFile: null,    // 主文件（xxxx年预支明细）
      dedFile: null,     // 工资表（扣款文件）
      prevFileName: '',
      isNextRound: false,  // 是否为续轮（第二轮及以后）
      // 工作表选项
      sheetOptions: [],      // 主文件工作簿列表
      dedSheetOptions: [],   // 工资表工作簿列表
      salarySheetOptions: [], // 平账步骤工资表列表
      // 选中值
      mainSheet: '',        // 主文件工作簿
      dedSheet: '',         // 工资表工作簿
      salarySheet: '',      // 平账步骤选中的工资表
      colType: '',          // 列类型
      colTypeOptions: [],
      selectedUnits: [],     // 已选单位（续轮保留）
      detectedUnitOptions: [],// 检测到的单位列表
      yzpzRound: 1,
      yzpzMonth: '',
      monthOptions: ['1月','2月','3月','4月','5月','6月','7月','8月','9月','10月','11月','12月'],
      // 状态
      loading: false,
      resultMsg: '',
      resultOk: false,
      showLog: false,
      logs: [],
      logText: '',
      // 多单位弹窗
      unitDialogVisible: false,
      unitOptions: [],
      unitChecked: []
    }
  },
  computed: {
    // 续轮时禁止重新上传主文件和工资表
    mainUploadDisabled() { return this.isNextRound && this.mainFile !== null },
    dedUploadDisabled() { return this.isNextRound && this.dedFile !== null },
    // 步骤1"核对"按钮是否可用
    canExecCheck() {
      if (this.currentStep !== 0) return false
      // 必须有主文件和工资表
      if (!this.mainFile || !this.dedFile) return false
      // 必须选了主文件工作簿
      if (!this.mainSheet) return false
      // 必须选了工资表工作簿和列类型
      if (!this.dedSheet || !this.colType) return false
      return true
    }
  },
  created() {
    this.restoreState();
  },
  beforeUnmount() {
    this.saveState();
  },
  methods: {
    // ==================== 状态持久化 ====================
    saveState() {
      const st = {
        currentStep: this.currentStep,
        yzpzRound: this.yzpzRound,
        mainSheet: this.mainSheet,
        dedSheet: this.dedSheet,
        salarySheet: this.salarySheet,
        colType: this.colType,
        selectedUnits: this.selectedUnits,
        yzpzMonth: this.yzpzMonth,
        isNextRound: this.isNextRound,
        prevFileName: sessionStorage.getItem('yzpz_last_file') || this.prevFileName,
        showLog: this.showLog,
        logs: this.logs,
        resultMsg: this.resultMsg,
        resultOk: this.resultOk
      };
      sessionStorage.setItem('yzpz_card_state', JSON.stringify(st));
    },
    restoreState() {
      try {
        const raw = sessionStorage.getItem('yzpz_card_state');
        if (!raw) return;
        const st = JSON.parse(raw);
        if (st.currentStep !== undefined) this.currentStep = st.currentStep;
        if (st.yzpzRound !== undefined) this.yzpzRound = st.yzpzRound;
        if (st.mainSheet !== undefined) this.mainSheet = st.mainSheet;
        if (st.dedSheet !== undefined) this.dedSheet = st.dedSheet;
        if (st.salarySheet !== undefined) this.salarySheet = st.salarySheet;
        if (st.colType !== undefined) this.colType = st.colType;
        if (st.selectedUnits !== undefined) this.selectedUnits = st.selectedUnits;
        if (st.yzpzMonth !== undefined) this.yzpzMonth = st.yzpzMonth;
        if (st.isNextRound !== undefined) this.isNextRound = st.isNextRound;
        if (st.prevFileName) this.prevFileName = st.prevFileName;
        if (st.showLog !== undefined) this.showLog = st.showLog;
        if (st.logs) { this.logs = st.logs; this.logText = st.logs.join('\n'); }
        if (st.resultMsg !== undefined) this.resultMsg = st.resultMsg;
        if (st.resultOk !== undefined) this.resultOk = st.resultOk;
      } catch (e) { /* ignore */ }
    },

    // ==================== 文件选择 ====================
    onMainFileSelected(file) {
      this.mainFile = file?.raw || null
      // 非续轮时清空单位选择
      if (!this.isNextRound) {
        this.selectedUnits = []
        this.detectedUnitOptions = []
      }
      if (this.mainFile) {
        this.detectSheets(this.mainFile, 'sheetOptions', 'mainSheet')
      }
    },
    // 主文件工作簿选择变化事件
    onMainSheetChange() {
      if (this.mainSheet && this.mainFile) {
        this.detectMainUnits(this.mainFile, this.mainSheet)
      }
    },
    onDedFileSelected(file) {
      this.dedFile = file?.raw || null
      if (this.dedFile) {
        this.detectSheets(this.dedFile, 'dedSheetOptions', 'dedSheet')
      }
    },
    onExceed() { this.$message.warning('只能选1个文件') },

    // ==================== 工作表检测 ====================
    async detectSheets(file, optionsKey, modelKey) {
      if (!file) return Promise.resolve()
      const formData = new FormData()
      formData.append('file', file)
      try {
        const res = await axios.post('/api/finance/toolbox/sheets', formData)
        if (res.data.success) {
          this[optionsKey] = res.data.sheets || []
          if (this[optionsKey].length > 0) {
            // 不再默认选中第一个工作簿，等待用户手动选择
            this[modelKey] = ''
            this.addLog(`检测到 ${this[optionsKey].length} 个工作表：${this[optionsKey].join(', ')}`)
            // 注意：单位检测将在用户选择工作簿后通过 @change 事件触发
          } else {
            this.$message.warning('未检测到工作表')
            this.addLog('? 工作表检测返回空列表')
          }
        } else {
          this.$message.error('工作表检测失败：' + (res.data.message || '未知错误'))
          this.addLog('? 工作表检测失败：' + (res.data.message || '未知错误'))
        }
      } catch (err) {
        const msg = err.response?.data?.message || err.message || '请求失败'
        this.$message.error('工作表检测失败：' + msg)
        this.addLog('? 工作表检测异常：' + msg)
        this[optionsKey] = []
      }
    },

    // ==================== 加载上一步结果 ====================
    async usePrevResult() {
      const fileName = sessionStorage.getItem('yzpz_last_file')
      if (!fileName) {
        this.$message.warning('没有上一步结果文件')
        return
      }
      try {
        this.addLog('正在加载上一步结果文件...')
        const response = await fetch(`/api/finance/download/${fileName}`)
        const blob = await response.blob()
        const file = new File([blob], fileName, {
          type: blob.type || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        })
        this.mainFile = file
        this.prevFileName = fileName
        // 步骤1=平账界面：检测工资表选项
        if (this.currentStep === 1) {
          await this.detectSheets(file, 'salarySheetOptions', 'salarySheet')
        }
        // 更新上传组件显示
        if (this.$refs.uploadMain) {
          this.$refs.uploadMain.uploadFiles = [{ name: fileName, raw: file, uid: Date.now(), status: 'success', percentage: 100 }]
        }
        this.addLog(`? 已加载上一步结果文件：${fileName}`)
      } catch (e) {
        this.addLog(`? 加载失败：${e.message}`)
        this.$message.error('加载上一步文件失败')
      }
    },

    // ==================== 保存结果文件 ====================
    saveResultFile(fileName) {
      sessionStorage.setItem('yzpz_last_file', fileName)
      this.prevFileName = fileName
    },

    // ==================== 步骤 ①：【核对】按钮 ====================
    async execCheck() {
      if (!this.canExecCheck) return
      this.loading = true
      this.resultMsg = ''
      this.resultOk = false
      this.addLog(`开始执行第${this.yzpzRound}轮核对（准备+复制姓名+姓名匹配+处理扣款）...`)

      const formData = new FormData()
      formData.append('file', this.mainFile)
      formData.append('checkFile', this.dedFile)  // 工资表即核对文件
      formData.append('dedFile', this.dedFile)
      formData.append('sheet', this.mainSheet)
      formData.append('checkSheet', this.dedSheet)
      formData.append('dedSheet', this.dedSheet)
      formData.append('round', this.yzpzRound)
      if (this.selectedUnits.length > 0) formData.append('units', this.selectedUnits.join(','))
      if (this.colType) formData.append('columnType', this.colType)

      try {
        const res = await axios.post('/api/finance/yzpz/prepare-and-check', formData)
        const data = res.data

        // 多单位检测
        if (data.needSelectUnits) {
          this.unitOptions = data.units || []
          this.unitChecked = [...this.unitOptions]
          this.unitDialogVisible = true
          this.addLog(`检测到 ${this.unitOptions.length} 个单位，等待用户选择...`)
          this.loading = false
          return
        }

        if (data.success) {
          this.resultMsg = data.message || '核对完成'
          this.resultOk = true
          this.addLog(`? 核对完成：${data.message || ''}`)
          if (data.fileName) {
            this.saveResultFile(data.fileName)
            this.addLog(`  输出文件：${data.fileName}`)
          }
          this.currentStep = 1 // 自动进入步骤2（平账）
          this.saveState();   // 持久化状态
          // 自动加载上一步结果文件，无需用户手动点击
          try {
            await this.usePrevResult()
          } catch (e) {
            this.addLog('? 自动加载结果文件失败，请手动点击加载：' + e.message)
          }
        } else {
          this.resultMsg = data.message || '核对失败'
          this.addLog(`? 核对失败：${data.message || ''}`)
        }
      } catch (err) {
        const msg = err.response?.data?.message || err.message
        this.resultMsg = msg
        this.addLog(`? 核对异常：${msg}`)
      } finally {
        this.loading = false
      }
    },

    confirmUnits() {
      if (this.unitChecked.length === 0) {
        this.$message.warning('请至少选择一个单位')
        return
      }
      this.selectedUnits = [...this.unitChecked]
      this.unitDialogVisible = false
      this.addLog(`? 已选择单位：${this.selectedUnits.join(',')}`)
      this.execCheck()
    },

    // 检测扣款文件列类型
    async detectColTypes() {
      if (!this.dedFile || !this.dedSheet) return
      this.colTypeOptions = []
      this.colType = ''
      this.addLog('正在检测扣款文件列类型...')
      const formData = new FormData()
      formData.append('dedFile', this.dedFile)
      formData.append('dedSheet', this.dedSheet)
      try {
        const res = await axios.post('/api/finance/yzpz/detect-columns', formData)
        if (res.data.success) {
          this.colTypeOptions = res.data.columns || []
          if (this.colTypeOptions.length === 1) {
            this.colType = this.colTypeOptions[0]
            this.addLog(`? 自动选择列类型：${this.colType}`)
          } else if (this.colTypeOptions.length > 1) {
            this.addLog(`? 检测到 ${this.colTypeOptions.length} 种列类型：${this.colTypeOptions.join(', ')}`)
          } else {
            this.addLog('? 未检测到有效的扣款列类型')
          }
        } else {
          this.addLog('? 列类型检测失败：' + res.data.message)
        }
      } catch (err) {
        this.addLog('? 列类型检测异常：' + (err.response?.data?.message || err.message))
      }
    },

    // 检测主文件中的单位
    async detectMainUnits(file, sheet) {
      if (!file || !sheet) return Promise.resolve()
      // 续轮时保留已选单位，只更新可选列表
      const prevSelected = this.isNextRound ? [...this.selectedUnits] : []
      this.detectedUnitOptions = []
      if (!this.isNextRound) this.selectedUnits = []
      this.addLog('正在检测工资表单位...')
      const formData = new FormData()
      formData.append('file', file)
      formData.append('sheet', sheet)
      try {
        const res = await axios.post('/api/finance/yzpz/detect-units', formData)
        if (res.data.success) {
          this.detectedUnitOptions = res.data.units || []
          this.addLog(`? 检测到 ${this.detectedUnitOptions.length} 个单位：${this.detectedUnitOptions.join(', ')}`)
          // 续轮时恢复之前选中的单位
          if (this.isNextRound && prevSelected.length > 0) {
            const stillAvailable = prevSelected.filter(u => this.detectedUnitOptions.includes(u))
            this.selectedUnits = stillAvailable
            this.addLog(`? 续轮保留单位选择：${stillAvailable.join(',')}`)
          }
        } else {
          this.addLog('? 单位检测失败：' + res.data.message)
        }
      } catch (err) {
        this.addLog('? 单位检测异常：' + (err.response?.data?.message || err.message))
      }
    },

    // ==================== 步骤 ②：【平账】按钮 ====================
    async execBalance() {
      if (this.currentStep !== 1) return
      this.loading = true
      this.resultMsg = ''
      this.resultOk = false
      this.addLog(`开始执行第${this.yzpzRound}轮平账...`)

      const formData = new FormData()
      formData.append('file', this.mainFile)
      formData.append('salarySheet', this.salarySheet)
      formData.append('month', this.yzpzMonth)
      formData.append('round', this.yzpzRound)
      if (this.selectedUnits.length > 0) formData.append('units', this.selectedUnits.join(','))

      try {
        const res = await postAsync('/api/finance/yzpz/balance', formData)
        if (res.data.success) {
          this.resultMsg = res.data.message || '平账完成'
          this.resultOk = true
          this.addLog(`? 平账完成：${res.data.message || ''}`)
          if (res.data.fileName) {
            this.saveResultFile(res.data.fileName)
            this.addLog(`  输出文件：${res.data.fileName}`)
          }
          this.currentStep = 2 // 自动进入步骤3（保存）
          this.saveState();
        } else {
          this.resultMsg = res.data.message || '平账失败'
          this.addLog(`? 平账失败：${res.data.message || ''}`)
        }
      } catch (err) {
        const msg = err.response?.data?.message || err.message
        this.resultMsg = msg
        this.addLog(`? 平账异常：${msg}`)
      } finally {
        this.loading = false
      }
    },

    // ==================== 步骤 ③：【保存】按钮 ====================
    doDownloadFinal() {
      const fileName = sessionStorage.getItem('yzpz_last_file')
      if (fileName) {
        window.location.href = `/api/finance/download/${fileName}`
      } else {
        this.$message.warning('没有可下载的文件')
      }
    },

    resetAll() {
      this.currentStep = 0
      this.mainFile = null
      this.dedFile = null
      this.prevFileName = ''
      this.isNextRound = false
      this.sheetOptions = []
      this.dedSheetOptions = []
      this.salarySheetOptions = []
      this.mainSheet = ''
      this.dedSheet = ''
      this.salarySheet = ''
      this.colType = ''
      this.colTypeOptions = []
      this.selectedUnits = []
      this.detectedUnitOptions = []
      this.yzpzRound = 1
      this.yzpzMonth = ''
      this.resultMsg = ''
      this.resultOk = false
      this.logs = []
      this.logText = ''
      // 清空上传组件
      if (this.$refs.uploadMain) this.$refs.uploadMain.uploadFiles = []
      if (this.$refs.uploadDed) this.$refs.uploadDed.uploadFiles = []
      sessionStorage.removeItem('yzpz_last_file')
      sessionStorage.removeItem('yzpz_card_state')
    },

    // ==================== 进行下一轮核对平账 ====================
    async continueNextRound() {
      try {
        await this.$confirm(
          `确定要进行第 ${this.yzpzRound + 1} 轮核对平账吗？`,
          '确认',
          { confirmButtonText: '确定', cancelButtonText: '取消', type: 'warning' }
        )
      } catch (e) {
        return // 用户取消
      }
      this.yzpzRound++
      this.addLog(`准备进行第${this.yzpzRound}轮核对平账...`)
      try {
        const fileName = sessionStorage.getItem('yzpz_last_file')
        if (!fileName) {
          this.$message.warning('没有上一轮的结果文件')
          return
        }

        this.addLog(`正在加载上一轮结果文件：${fileName}`)
        const response = await fetch(`/api/finance/download/${fileName}`)
        const blob = await response.blob()
        const file = new File([blob], fileName, {
          type: blob.type || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        })

        this.mainFile = file
        this.prevFileName = fileName
        this.isNextRound = true  // 标记为续轮
        this.yzpzMonth = ''     // 重置月份，避免保留上一轮值

        // 续轮时保留工资表文件（与第一轮相同），不清除 dedFile 和 dedSheetOptions
        // 只清除工作表选择和列类型，让用户自行选择
        this.dedSheet = ''
        this.colType = ''
        // 保留 dedFile、dedSheetOptions、colTypeOptions，不重新上传工资表

        // 回到步骤0（核对界面）
        this.currentStep = 0

        // 检测主文件工作表（需重新选择）
        await this.detectSheets(file, 'sheetOptions', 'mainSheet')

        // 清空主文件工作簿选择，让用户重新选择
        this.mainSheet = ''

        // 更新上传组件显示（禁用状态）
        if (this.$refs.uploadMain) {
          this.$refs.uploadMain.uploadFiles = [{ name: fileName, raw: file, uid: Date.now(), status: 'success', percentage: 100 }]
        }
        // 保留工资表上传组件显示（文件不变，与第一轮相同）
        if (this.$refs.uploadDed && this.dedFile) {
          const dedName = this.dedFile.name || '工资表.xlsx'
          this.$refs.uploadDed.uploadFiles = [{ name: dedName, raw: this.dedFile, uid: Date.now() + 1, status: 'success', percentage: 100 }]
        }

        this.addLog(`? 已加载上一轮结果，当前轮次：第${this.yzpzRound}轮`)
        this.saveState();
        this.$message.success('已加载上一轮结果，请选择主文件工作簿，并选择工资表工作簿和列类型后执行核对')
      } catch (e) {
        this.addLog(`? 加载失败：${e.message}`)
        this.$message.error('加载上一轮文件失败')
      }
    },

    // ==================== 日志 ====================
    addLog(msg) {
      const now = new Date().toLocaleTimeString()
      this.logs.push(`[${now}] ${msg}`)
      this.logText = this.logs.join('\n')
      this.$nextTick(() => {
        const ta = this.$el.querySelector('.log-body textarea')
        if (ta) ta.scrollTop = ta.scrollHeight
      })
    }
  }
}
</script>

<style scoped src="../../styles/components/YzpzCard.css"></style>

<style scoped src="../../styles/components/FuncPageHeader.css"></style>
