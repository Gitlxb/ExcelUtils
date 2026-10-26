import { createApp } from 'vue'
import ElementPlus from 'element-plus'
import 'element-plus/dist/index.css'
import zhCn from 'element-plus/dist/locale/zh-cn.mjs'
import * as ElementPlusIconsVue from '@element-plus/icons-vue'
import axios from 'axios'
import App from './App.vue'
import router from './router'
import { getDeviceId } from './utils/deviceId'

// 全局请求拦截器：附加设备号（任务中心多浏览器隔离；后端兼容缺失头）
axios.interceptors.request.use(config => {
  config.headers['X-Device-Id'] = getDeviceId()
  return config
})

const app = createApp(App)
app.use(ElementPlus, { locale: zhCn })
app.use(router)
// 注册所有图标
for (const [key, component] of Object.entries(ElementPlusIconsVue)) {
  app.component(key, component)
}
app.mount('#app')
