import { createRouter, createWebHistory } from 'vue-router'
import HomePage from '../components/HomePage.vue'

// ===== General =====
import CfBiaogeLi from '../components/general/CfBiaogeLi.vue'
import CfDuogeWenj from '../components/general/CfDuogeWenj.vue'
import NcengjiwenjjFzzt from '../components/general/NcengjiwenjjFzzt.vue'
import BiaogegesiZh from '../components/general/BiaogegesiZh.vue'
import ZhPdfExcel from '../components/general/ZhPdfExcel.vue'

// ===== Finance =====
import CwKgzSx from '../components/finance/CwKgzSx.vue'
import CwGysQzy from '../components/finance/CwGysQzy.vue'
import CwYhlsZyxg from '../components/finance/CwYhlsZyxg.vue'
import PlCfPizhu from '../components/finance/PlCfPizhu.vue'
import CwYzpzPmh from '../components/finance/CwYzpzPmh.vue'

// ===== Finance - cwgjx =====
import CwgjxYuzhi from '../components/finance/cwgjx/CwgjxYuzhi.vue'
import CwgjxBijiao from '../components/finance/cwgjx/CwgjxBijiao.vue'
import CwgjxBxcf from '../components/finance/cwgjx/CwgjxBxcf.vue'
import CwgjxDgfCf from '../components/finance/cwgjx/CwgjxDgfCf.vue'

// ===== Finance - cxm =====
import HsHuizongCxm from '../components/finance/cxm/HsHuizongCxm.vue'
import HsHuizongCxmChecker from '../components/finance/cxm/HsHuizongCxmChecker.vue'

// ===== Finance - jxkh =====
import JxkhSjpp from '../components/finance/jxkh/JxkhSjpp.vue'
import JxkhWgl from '../components/finance/jxkh/JxkhWgl.vue'
import JxkhHbbg from '../components/finance/jxkh/JxkhHbbg.vue'

// ===== HR =====
import CfDuogeWenjZz from '../components/hr/CfDuogeWenjZz.vue'
import CfZdnrfyqNrcsWsx from '../components/hr/CfZdnrfyqNrcsWsx.vue'
import CfGzbcfHyp from '../components/hr/CfGzbcfHyp.vue'
import CfGzbcfHypSimple from '../components/hr/CfGzbcfHypSimple.vue'
import HrGzbPipei from '../components/hr/HrGzbPipei.vue'

const routes = [
  { path: '/', name: 'home', component: HomePage },

  // ===== 通用功能 =====
  { path: '/general/split-sheets', name: 'general-split-sheets', component: CfBiaogeLi },
  { path: '/general/split-files', name: 'general-split-files', component: CfDuogeWenj },
  { path: '/general/extract', name: 'general-extract', component: NcengjiwenjjFzzt },
  { path: '/general/convert-format', name: 'general-convert-format', component: BiaogegesiZh },
  { path: '/general/pdf-excel', name: 'general-pdf-excel', component: ZhPdfExcel },

  // ===== 财务功能 =====
  { path: '/finance/date-filter', name: 'finance-date-filter', component: CwKgzSx },
  { path: '/finance/merge-supplier', name: 'finance-merge-supplier', component: CwGysQzy },
  { path: '/finance/bank-flow', name: 'finance-bank-flow', component: CwYhlsZyxg },
  { path: '/finance/comment', name: 'finance-comment', component: PlCfPizhu },
  { path: '/finance/yzpz', name: 'finance-yzpz', component: CwYzpzPmh },
  // 财务工具箱
  { path: '/finance/toolbox/fee-split', name: 'finance-toolbox-fee-split', component: CwgjxYuzhi },
  { path: '/finance/toolbox/data-compare', name: 'finance-toolbox-data-compare', component: CwgjxBijiao },
  { path: '/finance/toolbox/expense-split', name: 'finance-toolbox-expense-split', component: CwgjxBxcf },
  { path: '/finance/toolbox/labor-split', name: 'finance-toolbox-labor-split', component: CwgjxDgfCf },
  // 绩效考核
  { path: '/finance/jxkh/sjpp', name: 'finance-jxkh-sjpp', component: JxkhSjpp },
  { path: '/finance/jxkh/wgl', name: 'finance-jxkh-wgl', component: JxkhWgl },
  { path: '/finance/jxkh/hbbg', name: 'finance-jxkh-hbbg', component: JxkhHbbg },
  // 核算汇总
  { path: '/finance/cxm/summary', name: 'finance-cxm-summary', component: HsHuizongCxm },
  { path: '/finance/cxm/checker', name: 'finance-cxm-checker', component: HsHuizongCxmChecker },

  // ===== 人事功能 =====
  { path: '/hr/split-files', name: 'hr-split-files', component: CfDuogeWenjZz },
  { path: '/hr/split-config', name: 'hr-split-config', component: CfZdnrfyqNrcsWsx },
  { path: '/hr/split-salary', name: 'hr-split-salary', component: CfGzbcfHyp },
  { path: '/hr/split-salary-simple', name: 'hr-split-salary-simple', component: CfGzbcfHypSimple },
  { path: '/hr/gzb-pipei', name: 'hr-gzb-pipei', component: HrGzbPipei },
]

const router = createRouter({
  history: createWebHistory(),
  routes
})

export default router
