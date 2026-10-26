/**
 * 设备号工具：localStorage 持久生成一次，用于任务中心多浏览器隔离。
 * 所有 axios 请求经 main.js 拦截器自动附加 X-Device-Id 头。
 */
export function getDeviceId() {
  try {
    let d = localStorage.getItem('excel_device_id')
    if (!d) {
      d = (crypto.randomUUID ? crypto.randomUUID() : 'd' + Date.now() + Math.random().toString(16).slice(2))
      localStorage.setItem('excel_device_id', d)
    }
    return d
  } catch (e) {
    return 'unknown-device'
  }
}

/** 当前设备在跑的任务 taskId 列表（localStorage 持久，刷新/切页可恢复） */
const TASK_KEY = 'excel_active_tasks'

export function getActiveTasks() {
  try {
    const raw = localStorage.getItem(TASK_KEY)
    return raw ? JSON.parse(raw) : []
  } catch (e) {
    return []
  }
}

export function rememberTask(taskId) {
  try {
    const list = getActiveTasks()
    if (!list.includes(taskId)) {
      list.push(taskId)
      localStorage.setItem(TASK_KEY, JSON.stringify(list))
    }
  } catch (e) { /* ignore */ }
}

export function forgetTask(taskId) {
  try {
    const list = getActiveTasks().filter(id => id !== taskId)
    localStorage.setItem(TASK_KEY, JSON.stringify(list))
  } catch (e) { /* ignore */ }
}
