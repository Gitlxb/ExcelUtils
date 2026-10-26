<template>
  <el-container class="app-container">
    <!-- 左侧导航 -->
    <el-aside width="220px" class="app-sidebar">
      <div class="sidebar-header">
        <img src="/logo.png" alt="logo" class="sidebar-logo">
        <h2>Excel处理工具</h2>
      </div>
      <el-menu
        :default-active="activeMenu"
        class="sidebar-menu"
        @select="handleMenuSelect"
      >
        <el-sub-menu index="general">
          <template #title>
            <el-icon><FolderOpened /></el-icon><span>通用功能</span>
          </template>
          <el-menu-item index="/general/split-sheets">拆成多个工作表</el-menu-item>
          <el-menu-item index="/general/split-files">拆成多个Excel文件</el-menu-item>
          <el-menu-item index="/general/extract">文件夹文件提取</el-menu-item>
          <el-menu-item index="/general/convert-format">表格格式转换</el-menu-item>
          <el-menu-item index="/general/pdf-excel">PDF转Excel</el-menu-item>
        </el-sub-menu>

        <el-sub-menu index="finance">
          <template #title>
            <el-icon><Money /></el-icon><span>财务功能</span>
          </template>
          <el-menu-item index="/finance/date-filter">开关账筛选</el-menu-item>
          <el-menu-item index="/finance/merge-supplier">供应商合并</el-menu-item>
          <el-menu-item index="/finance/comment">批量批注工具</el-menu-item>
          <el-menu-item index="/finance/bank-flow">银行流水修改</el-menu-item>
          <el-menu-item index="/finance/yzpz">预支平账</el-menu-item>
          <el-sub-menu index="finance-toolbox">
            <template #title><span>财务工具箱</span></template>
            <el-menu-item index="/finance/toolbox/fee-split">预支处理</el-menu-item>
            <el-menu-item index="/finance/toolbox/data-compare">数据比较</el-menu-item>
            <el-menu-item index="/finance/toolbox/expense-split">报销拆分</el-menu-item>
            <el-menu-item index="/finance/toolbox/labor-split">代工费拆分_批量处理</el-menu-item>
          </el-sub-menu>
          <el-sub-menu index="finance-jxkh">
            <template #title><span>绩效考核</span></template>
            <el-menu-item index="/finance/jxkh/sjpp">数据匹配</el-menu-item>
            <el-menu-item index="/finance/jxkh/wgl">稳岗率分析</el-menu-item>
            <el-menu-item index="/finance/jxkh/hbbg">合并表格</el-menu-item>
            <el-menu-item index="/finance/jxkh/tenure">在职月份计算</el-menu-item>
          </el-sub-menu>
          <el-sub-menu index="finance-cxm">
            <template #title><span>核算汇总</span></template>
            <el-menu-item index="/finance/cxm/summary">拆分汇总比较</el-menu-item>
            <el-menu-item index="/finance/cxm/checker">核对</el-menu-item>
          </el-sub-menu>
          <el-menu-item index="/finance/hbgs">合并个税</el-menu-item>
        </el-sub-menu>

        <el-sub-menu index="hr">
          <template #title>
            <el-icon><User /></el-icon><span>人事功能</span>
          </template>
          <el-menu-item index="/hr/split-files">拆分两次_在职离职</el-menu-item>
          <el-menu-item index="/hr/split-config">按拆分参数配置拆分</el-menu-item>
          <el-sub-menu index="hr-doc-toolbox">
            <template #title><span>文档处理工具箱</span></template>
            <el-menu-item index="/hr/gzbcf-split">工资表拆分</el-menu-item>
            <el-menu-item index="/hr/xxj-split">孝心金拆分</el-menu-item>
            <el-menu-item index="/hr/word-merge">Word文档合并</el-menu-item>
          </el-sub-menu>
          <el-menu-item index="/hr/gzb-pipei">工资表匹配</el-menu-item>
        </el-sub-menu>
      </el-menu>

      <!-- 任务中心入口（底部固定） -->
      <div class="task-center-entry" @click="openTaskCenter">
        <el-badge :value="runningCount" :hidden="runningCount === 0" :max="99" class="task-badge">
          <el-icon :size="20"><List /></el-icon>
        </el-badge>
        <span class="task-entry-text">任务中心</span>
        <span v-if="runningCount > 0" class="task-entry-running">{{ runningCount }} 处理中</span>
      </div>
    </el-aside>

    <!-- 任务中心抽屉 -->
    <el-drawer v-model="taskDrawerVisible" title="任务中心" size="420px" :append-to-body="true">
      <div v-if="taskList.length === 0" class="task-empty">暂无任务</div>
      <div v-for="t in taskList" :key="t.taskId" class="task-item">
        <div class="task-item-head">
          <span class="task-item-name">{{ t.displayName || t.apiName }}</span>
          <el-tag :type="statusTag(t.status)" size="small">{{ statusText(t.status) }}</el-tag>
        </div>
        <div class="task-item-meta">
          <span v-if="t.originalFileName" class="task-item-file">{{ t.originalFileName }}</span>
          <span class="task-item-time">{{ fmtTime(t.createTime) }}</span>
        </div>
        <div v-if="t.status === 'FAILED' && t.message" class="task-item-err">{{ t.message }}</div>
        <div v-if="t.status === 'SUCCESS' && t.result && t.result.fileName" class="task-item-done">
          结果：{{ t.result.fileName }}
          <el-button type="primary" size="small" @click="downloadTaskResult(t)">下载</el-button>
        </div>
        <el-divider v-if="t !== taskList[taskList.length - 1]" style="margin:8px 0" />
      </div>
    </el-drawer>

    <!-- 右侧主内容 -->
    <el-main class="app-main">
      <!-- 功能页面提示（首页不显示） -->
      <el-alert
        v-if="$route.path !== '/'"
        type="warning"
        :closable="true"
        show-icon
        style="margin-bottom: 16px"
      >
        <template #title>
          <span>上传Excel文件前，请先<strong>关闭已打开的Excel文件</strong>，否则可能导致上传失败</span>
        </template>
      </el-alert>

      <router-view ref="pageView" />
    </el-main>

    <!-- 悬浮AI聊天机器人 -->
    <AiChatBot />
  </el-container>
</template>

<script>
import axios from 'axios'
import AiChatBot from './components/AiChatBot.vue'

export default {
  name: 'App',
  components: { AiChatBot },
  data() {
    return {
      taskDrawerVisible: false,
      taskList: [],
      runningCount: 0,
      taskPollTimer: null,
      taskDrawerTimer: null
    }
  },
  computed: {
    activeMenu() {
      return this.$route.path
    }
  },
  mounted() {
    this.refreshTaskBadge()
    // 后台每 5s 刷新一次角标（仅刷新计数，开销小）
    this.taskPollTimer = setInterval(this.refreshTaskBadge, 5000)
  },
  beforeUnmount() {
    if (this.taskPollTimer) clearInterval(this.taskPollTimer)
    if (this.taskDrawerTimer) clearInterval(this.taskDrawerTimer)
  },
  methods: {
    handleMenuSelect(index) {
      if (this.$route.path !== index) {
        this.$router.push(index)
      }
    },

    async refreshTaskBadge() {
      try {
        const res = await axios.get('/api/tasks')
        const list = (res.data && res.data.tasks) || []
        this.runningCount = list.filter(t => t.status === 'RUNNING' || t.status === 'PENDING').length
        if (this.taskDrawerVisible) this.taskList = list
      } catch (e) { /* 忽略 */ }
    },

    openTaskCenter() {
      this.taskDrawerVisible = true
      this.refreshTaskBadge()
      // 抽屉打开期间每 3s 刷新列表
      if (this.taskDrawerTimer) clearInterval(this.taskDrawerTimer)
      this.taskDrawerTimer = setInterval(this.refreshTaskBadge, 3000)
    },

    statusText(s) {
      return { PENDING: '排队中', RUNNING: '处理中', SUCCESS: '已完成', FAILED: '失败' }[s] || s
    },
    statusTag(s) {
      return { PENDING: 'info', RUNNING: 'warning', SUCCESS: 'success', FAILED: 'danger' }[s] || 'info'
    },
    fmtTime(ms) {
      if (!ms) return ''
      const d = new Date(ms)
      const p = n => (n < 10 ? '0' + n : '' + n)
      return `${p(d.getMonth() + 1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}`
    },
    downloadTaskResult(t) {
      if (t && t.result && t.result.fileName) {
        const module = t.module || 'general'
        window.location.href = `/api/${module}/download/${encodeURIComponent(t.result.fileName)}`
      }
    }
  }
}
</script>

<style>
body { margin: 0; font-family: 'Microsoft YaHei', sans-serif; }
.app-container { height: 100vh; }
.app-sidebar {
  background: #304156;
  color: #fff;
  overflow-y: auto;
}
.sidebar-header {
  padding: 16px 20px;
  text-align: center;
  border-bottom: 1px solid rgba(255,255,255,0.1);
}
.sidebar-logo {
  width: 36px;
  height: 36px;
  border-radius: 6px;
  margin-bottom: 8px;
  vertical-align: middle;
}
.sidebar-header h2 {
  margin: 0;
  font-size: 18px;
  color: #fff;
}
.sidebar-menu {
  border-right: none;
  background: transparent;
}
.sidebar-menu .el-menu-item,
.sidebar-menu .el-sub-menu__title {
  color: #bfcbd9 !important;
}
.sidebar-menu .el-sub-menu .el-menu {
  background: #1f2d3d !important;
}
.sidebar-menu .el-sub-menu .el-menu-item {
  color: #bfcbd9 !important;
  background: transparent !important;
}
.sidebar-menu .el-menu-item.is-active {
  color: #409EFF !important;
  background: rgba(64,158,255,0.1) !important;
}
.sidebar-menu .el-menu-item:hover,
.sidebar-menu .el-sub-menu__title:hover {
  background: rgba(255,255,255,0.05) !important;
}
.sidebar-menu .el-sub-menu .el-menu-item:hover {
  background: rgba(255,255,255,0.08) !important;
}
.app-main {
  background: #f0f2f5;
  padding: 20px;
  overflow-y: auto;
}

/* ===== 任务中心入口（侧栏底部） ===== */
.task-center-entry {
  display: flex;
  align-items: center;
  gap: 8px;
  padding: 12px 20px;
  color: #bfcbd9;
  cursor: pointer;
  border-top: 1px solid rgba(255,255,255,0.08);
  transition: background .2s;
}
.task-center-entry:hover {
  background: rgba(255,255,255,0.05);
  color: #fff;
}
.task-center-entry .task-entry-text {
  font-size: 14px;
}
.task-center-entry .task-entry-running {
  font-size: 12px;
  color: #e6a23c;
  margin-left: auto;
}
.task-badge {
  display: inline-flex;
}
.task-badge :deep(.el-badge__content) {
  background: #e6a23c;
  border: none;
  font-size: 11px;
}

/* ===== 任务中心抽屉 ===== */
.task-empty {
  color: #909399;
  text-align: center;
  padding: 48px 0;
  font-size: 14px;
}
.task-item-head {
  display: flex;
  align-items: center;
  justify-content: space-between;
  margin-bottom: 6px;
}
.task-item-name {
  font-weight: 600;
  font-size: 14px;
}
.task-item-meta {
  display: flex;
  justify-content: space-between;
  font-size: 12px;
  color: #909399;
  margin-bottom: 4px;
}
.task-item-file {
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  max-width: 60%;
}
.task-item-err {
  color: #f56c6c;
  font-size: 13px;
  margin: 4px 0;
  word-break: break-all;
}
.task-item-done {
  font-size: 13px;
  color: #67c23a;
  margin-top: 6px;
  display: flex;
  align-items: center;
  justify-content: space-between;
  word-break: break-all;
}
.task-item-done .el-button {
  margin-left: 12px;
  flex-shrink: 0;
}
</style>
