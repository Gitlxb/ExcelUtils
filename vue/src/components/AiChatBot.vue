<template>
  <div class="ai-chat-bot">
    <!-- 悬浮按钮 -->
    <div class="chat-fab" @click="toggleChat" v-show="!chatVisible">
      <el-icon :size="28"><ChatDotRound /></el-icon>
    </div>

    <!-- 聊天面板 -->
    <transition name="chat-slide">
      <div class="chat-panel" v-show="chatVisible">
        <!-- 头部 -->
        <div class="chat-header">
          <div class="chat-header-left">
            <div class="chat-avatar">🤖</div>
            <div class="chat-header-info">
              <span class="chat-title">AI助手 · 小锦</span>
              <span class="chat-subtitle">DeepSeek 驱动</span>
            </div>
          </div>
          <div class="chat-header-actions">
            <el-button text circle @click="clearChat" title="清空对话">
              <el-icon><Delete /></el-icon>
            </el-button>
            <el-button text circle @click="toggleChat" title="最小化">
              <el-icon><Close /></el-icon>
            </el-button>
          </div>
        </div>

        <!-- 消息列表 -->
        <div class="chat-messages" ref="messagesContainer">
          <!-- 欢迎消息 -->
          <div v-if="messages.length === 0" class="chat-welcome">
            <div class="welcome-icon">🎉</div>
            <p>你好！我是「小锦」，你的Excel处理助手</p>
            <p class="welcome-hint">可以问我关于Excel操作、数据处理等方面的问题</p>
          </div>

          <div
            v-for="(msg, index) in messages"
            :key="index"
            :class="['chat-message', msg.role === 'user' ? 'message-user' : 'message-assistant']"
          >
            <div class="message-avatar">
              {{ msg.role === 'user' ? '👤' : '🤖' }}
            </div>
            <div class="message-body">
              <!-- 思考过程（可折叠） -->
              <div v-if="msg.reasoning" class="reasoning-box">
                <div class="reasoning-header" @click="toggleReasoning(index)">
                  <span class="reasoning-arrow">{{ msg.reasoningExpanded ? '▼' : '▶' }}</span>
                  <span>💭 {{ msg.reasoningDone ? '思考过程' : '正在思考...' }}</span>
                </div>
                <div v-show="msg.reasoningExpanded" class="reasoning-content">{{ msg.reasoning }}</div>
              </div>
              <!-- 正式回答 -->
              <div class="message-bubble" v-html="renderContent(msg.content)"></div>
            </div>
          </div>

          <!-- 正在输入动画 -->
          <div v-if="isLoading" class="chat-message message-assistant">
            <div class="message-avatar">🤖</div>
            <div class="message-bubble typing-bubble">
              <span class="typing-dot"></span>
              <span class="typing-dot"></span>
              <span class="typing-dot"></span>
            </div>
          </div>
        </div>

        <!-- 输入区 -->
        <div class="chat-input-area">
          <el-input
            v-model="inputText"
            type="textarea"
            :rows="1"
            :autosize="{ minRows: 1, maxRows: 4 }"
            placeholder="输入你的问题..."
            @keydown.enter.exact.prevent="sendMessage"
            :disabled="isLoading"
            resize="none"
          />
          <el-button
            type="primary"
            circle
            @click="sendMessage"
            :disabled="isLoading || !inputText.trim()"
            class="send-btn"
          >
            <el-icon><Promotion /></el-icon>
          </el-button>
        </div>
      </div>
    </transition>
  </div>
</template>

<script>
export default {
  name: 'AiChatBot',
  data() {
    return {
      chatVisible: false,
      inputText: '',
      messages: [],
      isLoading: false,
      conversationHistory: [],
    }
  },
  methods: {
    toggleChat() {
      this.chatVisible = !this.chatVisible
      if (this.chatVisible) {
        this.$nextTick(() => this.scrollToBottom())
      }
    },

    clearChat() {
      this.messages = []
      this.conversationHistory = []
    },

    toggleReasoning(index) {
      this.messages[index].reasoningExpanded = !this.messages[index].reasoningExpanded
    },

    async sendMessage() {
      const text = this.inputText.trim()
      if (!text || this.isLoading) return

      this.messages.push({ role: 'user', content: text })
      this.conversationHistory.push({ role: 'user', content: text })
      this.inputText = ''
      this.scrollToBottom()

      this.isLoading = true

      // 添加助手消息，包含 reasoning 和 content 两个独立字段
      this.messages.push({
        role: 'assistant',
        content: '',
        reasoning: '',
        reasoningDone: false,
        reasoningExpanded: true
      })
      const aiIndex = this.messages.length - 1

      try {
        const response = await fetch('/api/deepseek/chat', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ messages: this.conversationHistory })
        })

        if (!response.ok) {
          throw new Error('请求失败: ' + response.status)
        }

        // 读取 SSE 流
        const reader = response.body.getReader()
        const decoder = new TextDecoder()
        let buffer = ''
        let fullContent = ''

        // eslint-disable-next-line no-constant-condition
        while (true) {
          const { done, value } = await reader.read()
          if (done) break

          buffer += decoder.decode(value, { stream: true })
          const lines = buffer.split('\n')
          buffer = lines.pop() || ''

          for (let i = 0; i < lines.length; i++) {
            const line = lines[i]

            // 解析 event 行
            if (line.startsWith('event:')) {
              const eventType = line.substring(6).trim()
              const nextLine = lines[i + 1]
              if (nextLine && nextLine.startsWith('data:')) {
                const data = nextLine.substring(5).trim()
                i++ // 跳过下一行

                if (eventType === 'reasoning' && data) {
                  this.messages[aiIndex].reasoning += data
                  this.scrollToBottom()
                } else if (eventType === 'message' && data) {
                  fullContent += data
                  this.messages[aiIndex].content = fullContent
                  this.messages[aiIndex].reasoningDone = true
                  this.scrollToBottom()
                } else if (eventType === 'done') {
                  // 完成
                } else if (eventType === 'error') {
                  fullContent = data
                  this.messages[aiIndex].content = data
                }
              }
              continue
            }

            // 兼容旧格式：直接 data: 开头
            if (line.startsWith('data:')) {
              const data = line.substring(5).trim()
              if (data === '[DONE]') continue
              fullContent += data
              this.messages[aiIndex].content = fullContent
              this.messages[aiIndex].reasoningDone = true
              this.scrollToBottom()
            }
          }
        }

        const finalContent = this.messages[aiIndex].content
        if (finalContent) {
          this.conversationHistory.push({ role: 'assistant', content: finalContent })
        } else {
          this.messages[aiIndex].content = '抱歉，我没有获取到回复，请重试。'
        }
      } catch (err) {
        this.messages[aiIndex].content = '网络连接失败，请检查网络并重试。'
        console.error('DeepSeek API error:', err)
      } finally {
        this.isLoading = false
        this.scrollToBottom()
      }
    },

    renderContent(content) {
      if (!content) return ''
      let html = content
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/\n/g, '<br>')
        .replace(/```(\w*)\n?([\s\S]*?)```/g, '<pre><code>$2</code></pre>')
        .replace(/`([^`]+)`/g, '<code>$1</code>')
        .replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>')
      return html
    },

    scrollToBottom() {
      this.$nextTick(() => {
        const container = this.$refs.messagesContainer
        if (container) {
          container.scrollTop = container.scrollHeight
        }
      })
    }
  }
}
</script>

<style scoped>
.ai-chat-bot {
  position: fixed;
  right: 24px;
  bottom: 24px;
  z-index: 9999;
  font-family: 'Microsoft YaHei', sans-serif;
}

/* ===== 悬浮按钮 ===== */
.chat-fab {
  width: 56px;
  height: 56px;
  border-radius: 50%;
  background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
  box-shadow: 0 4px 16px rgba(102, 126, 234, 0.45);
  display: flex;
  align-items: center;
  justify-content: center;
  cursor: pointer;
  transition: all 0.3s ease;
  color: #fff;
  position: absolute;
  right: 0;
  bottom: 0;
}

.chat-fab:hover {
  transform: scale(1.08);
  box-shadow: 0 6px 24px rgba(102, 126, 234, 0.6);
}

/* ===== 聊天面板 ===== */
.chat-panel {
  position: absolute;
  right: 0;
  bottom: 0;
  width: 400px;
  height: 560px;
  background: #fff;
  border-radius: 16px;
  box-shadow: 0 8px 40px rgba(0, 0, 0, 0.15);
  display: flex;
  flex-direction: column;
  overflow: hidden;
}

/* ===== 头部 ===== */
.chat-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: 14px 16px;
  background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
  color: #fff;
}

.chat-header-left {
  display: flex;
  align-items: center;
  gap: 10px;
}

.chat-avatar {
  width: 40px;
  height: 40px;
  background: rgba(255,255,255,0.2);
  border-radius: 12px;
  display: flex;
  align-items: center;
  justify-content: center;
  font-size: 22px;
}

.chat-header-info {
  display: flex;
  flex-direction: column;
}

.chat-title {
  font-size: 15px;
  font-weight: 600;
}

.chat-subtitle {
  font-size: 11px;
  opacity: 0.8;
}

.chat-header-actions .el-button {
  color: #fff !important;
}

/* ===== 消息区域 ===== */
.chat-messages {
  flex: 1;
  overflow-y: auto;
  padding: 16px;
  background: #f7f8fa;
  display: flex;
  flex-direction: column;
  gap: 16px;
}

.chat-messages::-webkit-scrollbar {
  width: 4px;
}

.chat-messages::-webkit-scrollbar-thumb {
  background: #d0d5dd;
  border-radius: 2px;
}

/* 欢迎消息 */
.chat-welcome {
  text-align: center;
  padding: 40px 20px;
  color: #909399;
}

.welcome-icon {
  font-size: 48px;
  margin-bottom: 12px;
}

.welcome-hint {
  font-size: 12px;
  margin-top: 4px;
  opacity: 0.6;
}

/* ===== 消息气泡 ===== */
.chat-message {
  display: flex;
  gap: 10px;
  max-width: 100%;
}

.message-user {
  flex-direction: row-reverse;
}

.message-avatar {
  width: 34px;
  height: 34px;
  border-radius: 10px;
  display: flex;
  align-items: center;
  justify-content: center;
  font-size: 18px;
  flex-shrink: 0;
  background: #fff;
  box-shadow: 0 1px 3px rgba(0,0,0,0.08);
}

.message-body {
  max-width: 75%;
  display: flex;
  flex-direction: column;
  gap: 6px;
}

/* ===== 思考过程 ===== */
.reasoning-box {
  background: #fef7e0;
  border: 1px solid #f5d873;
  border-radius: 10px;
  overflow: hidden;
}

.reasoning-header {
  display: flex;
  align-items: center;
  gap: 6px;
  padding: 8px 12px;
  cursor: pointer;
  font-size: 12px;
  color: #8b6914;
  user-select: none;
  transition: background 0.2s;
}

.reasoning-arrow {
  font-size: 10px;
  display: inline-block;
  transition: transform 0.2s;
}

.reasoning-header:hover {
  background: rgba(245, 216, 115, 0.3);
}

.reasoning-content {
  padding: 8px 12px 12px;
  font-size: 12px;
  color: #6b5b1a;
  line-height: 1.5;
  white-space: pre-wrap;
  border-top: 1px solid rgba(245, 216, 115, 0.4);
  max-height: 200px;
  overflow-y: auto;
}

.message-bubble {
  max-width: 75%;
  padding: 10px 14px;
  border-radius: 14px;
  font-size: 14px;
  line-height: 1.6;
  word-break: break-word;
}

.message-user .message-bubble {
  background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
  color: #fff;
  border-bottom-right-radius: 4px;
}

.message-assistant .message-bubble {
  background: #fff;
  color: #303133;
  border-bottom-left-radius: 4px;
  box-shadow: 0 1px 4px rgba(0,0,0,0.06);
}

.message-bubble :deep(code) {
  background: rgba(0,0,0,0.06);
  padding: 2px 6px;
  border-radius: 4px;
  font-size: 13px;
}

.message-user .message-bubble :deep(code) {
  background: rgba(255,255,255,0.2);
}

.message-bubble :deep(pre) {
  background: #282c34;
  color: #abb2bf;
  padding: 12px;
  border-radius: 8px;
  overflow-x: auto;
  margin: 8px 0;
  font-size: 13px;
}

.message-bubble :deep(pre code) {
  background: none;
  padding: 0;
}

.message-bubble :deep(strong) {
  font-weight: 600;
}

/* ===== 打字动画 ===== */
.typing-bubble {
  display: flex;
  align-items: center;
  gap: 4px;
  padding: 14px 18px;
}

.typing-dot {
  width: 8px;
  height: 8px;
  background: #909399;
  border-radius: 50%;
  animation: typing-bounce 1.4s infinite ease-in-out;
}

.typing-dot:nth-child(1) { animation-delay: 0s; }
.typing-dot:nth-child(2) { animation-delay: 0.2s; }
.typing-dot:nth-child(3) { animation-delay: 0.4s; }

@keyframes typing-bounce {
  0%, 60%, 100% { transform: translateY(0); }
  30% { transform: translateY(-8px); }
}

/* ===== 输入区 ===== */
.chat-input-area {
  display: flex;
  align-items: flex-end;
  gap: 10px;
  padding: 12px 16px;
  background: #fff;
  border-top: 1px solid #ebeef5;
}

.chat-input-area :deep(.el-textarea__inner) {
  border-radius: 12px;
  background: #f5f7fa;
  border: 1px solid transparent;
  padding: 10px 14px;
  font-size: 14px;
  transition: all 0.2s;
}

.chat-input-area :deep(.el-textarea__inner:focus) {
  background: #fff;
  border-color: #667eea;
  box-shadow: 0 0 0 2px rgba(102,126,234,0.15);
}

.send-btn {
  flex-shrink: 0;
  width: 40px;
  height: 40px;
  background: linear-gradient(135deg, #667eea 0%, #764ba2 100%) !important;
  border: none !important;
}

.send-btn:hover {
  opacity: 0.9;
}

.send-btn.is-disabled {
  opacity: 0.4;
}

/* ===== 过渡动画 ===== */
.chat-slide-enter-active {
  transition: all 0.35s cubic-bezier(0.4, 0, 0.2, 1);
}

.chat-slide-leave-active {
  transition: all 0.25s cubic-bezier(0.4, 0, 0.2, 1);
}

.chat-slide-enter-from {
  opacity: 0;
  transform: translateY(20px) scale(0.95);
}

.chat-slide-leave-to {
  opacity: 0;
  transform: translateY(20px) scale(0.95);
}
</style>
