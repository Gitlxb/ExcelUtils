/**
 * 异步任务请求包装：处理端点已改造为"提交即返回 taskId + 后台执行"。
 *
 * postAsync(url, formData, { pollMs=1500 }) 与 axios.post 用法一致，
 * 区别：
 *  - 若响应含 taskId → 自动轮询 /api/tasks/{taskId}（模块级定时器，页面卸载不中断后台任务）
 *  - 轮询到 SUCCESS → resolve({ data: {...原响应, ...task.result} })（与改造前同步响应同构，
 *    页面既有 res.data.fileName / res.data.message / res.data.logs 等照常使用）
 *  - 轮询到 FAILED  → reject 带 message（兼容现有 catch 里读 err.response?.data?.message 的逻辑）
 *  - 若响应无 taskId（秒级辅助接口如读 sheets/detect 未改造）→ 原样直返，行为不变
 *
 * 组件卸载后轮询仍继续，但 resolve/reject 只在组件仍存活时有意义（Vue 对已卸载组件
 * 的 setState 静默忽略）；"切走回来仍可下载"由任务中心抽屉（App.vue）兜底——已完成
 * 任务随时可从任务中心下载，本工具也把 taskId 记入 localStorage 便于恢复。
 */
import axios from 'axios'
import { rememberTask, forgetTask } from './deviceId'

/** 模块级轮询表：key=taskId, value={done, resolve, reject, failMsg} */
const pollTable = {}
let globalTimer = null
let POLL_MS = 1500

function ensureGlobalTimer() {
  if (globalTimer) return
  globalTimer = setInterval(() => {
    for (const [taskId, h] of Object.entries(pollTable)) {
      if (h.done) { delete pollTable[taskId]; continue }
      axios.get('/api/tasks/' + taskId)
        .then(res => {
          const task = res.data && res.data.task
          if (!task) { h.done = true; h.reject(apiError('任务不存在或已过期')); forgetTask(taskId); return }
          if (task.status === 'SUCCESS') {
            h.done = true
            h.resolve(task)
            forgetTask(taskId)
          } else if (task.status === 'FAILED') {
            h.done = true
            h.reject(apiError(task.message || '处理失败', task.result || null))
            forgetTask(taskId)
          }
          // RUNNING/PENDING 继续等
        })
        .catch(() => { /* 网络瞬断，下轮重试 */ })
    }
    if (Object.keys(pollTable).length === 0 && globalTimer) {
      clearInterval(globalTimer)
      globalTimer = null
    }
  }, POLL_MS)
}

function apiError(message, extra) {
  const err = new Error(message)
  err.response = { data: { message, ...(extra || {}) } }
  return err
}

/**
 * 统一请求入口。
 * @returns Promise<{data: Object}> 与 axios.post 返回结构兼容
 */
export function postAsync(url, formData, { pollMs = 1500, timeout } = {}) {
  POLL_MS = pollMs   // 生效自定义轮询间隔（默认 1500ms）
  const axiosOpts = {}
  if (timeout) axiosOpts.timeout = timeout   // 透传上传/首响应超时（如 hbgs 大文件 10min）
  return axios.post(url, formData, axiosOpts).then(res => {
    const d = res.data || {}
    // 未改造端点（同步辅助）直返
    if (!d.taskId) return res
    const taskId = d.taskId
    rememberTask(taskId)
    // 已改造：轮询到结束
    return new Promise((resolve, reject) => {
      pollTable[taskId] = { done: false, resolve, reject }
      ensureGlobalTimer()
      // 立即先查一次，减少首屏等待
      axios.get('/api/tasks/' + taskId).then(r2 => {
        const h = pollTable[taskId]
        if (!h || h.done) return
        const task = r2.data && r2.data.task
        if (!task) { h.done = true; h.reject(apiError('任务不存在或已过期')); forgetTask(taskId); return }
        if (task.status === 'SUCCESS') {
          h.done = true; h.resolve(task); forgetTask(taskId)
        } else if (task.status === 'FAILED') {
          h.done = true; h.reject(apiError(task.message || '处理失败', task.result || null)); forgetTask(taskId)
        }
      }).catch(() => {})
    }).then(task => {
      // 把 task.result 合入 res.data：页面原有 res.data.fileName 等继续可用
      return { data: Object.assign({}, res.data, task.result || {}) }
    })
  })
}

export default postAsync
