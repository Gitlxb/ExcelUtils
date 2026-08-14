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
        <el-menu-item index="/">
          <el-icon><HomeFilled /></el-icon><span>主页</span>
        </el-menu-item>
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
            <el-menu-item index="/hr/word-merge">Word文档合并</el-menu-item>
          </el-sub-menu>
          <el-menu-item index="/hr/gzb-pipei">工资表匹配</el-menu-item>
        </el-sub-menu>
      </el-menu>
    </el-aside>

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
import AiChatBot from './components/AiChatBot.vue'

export default {
  name: 'App',
  components: { AiChatBot },
  computed: {
    activeMenu() {
      return this.$route.path
    }
  },
  methods: {
    handleMenuSelect(index) {
      if (this.$route.path !== index) {
        this.$router.push(index)
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
</style>
