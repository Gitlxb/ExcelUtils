package com.example.project.controller;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.http.MediaType;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.servlet.mvc.method.annotation.SseEmitter;

import java.io.BufferedReader;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.net.HttpURLConnection;
import java.net.URI;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

@RestController
@RequestMapping("/api/deepseek")
public class DeepSeekController {

    @Value("${deepseek.api.key}")
    private String apiKey;

    @Value("${deepseek.api.url}")
    private String apiUrl;

    @Value("${deepseek.api.model}")
    private String model;

    private final ObjectMapper objectMapper = new ObjectMapper();
    private final ExecutorService executor = Executors.newCachedThreadPool();

    /**
     * 流式聊天接口（SSE）
     * POST /api/deepseek/chat
     * Body: { "messages": [{ "role": "user", "content": "你好" }] }
     */
    @PostMapping(value = "/chat", produces = MediaType.TEXT_EVENT_STREAM_VALUE)
    public SseEmitter chat(@RequestBody Map<String, Object> body) {
        SseEmitter emitter = new SseEmitter(0L); // 无超时

        executor.execute(() -> {
            HttpURLConnection conn = null;
            try {
                // 构建发送给 DeepSeek 的消息列表（含系统提示词）
                @SuppressWarnings("unchecked")
                List<Map<String, String>> userMessages = (List<Map<String, String>>) body.get("messages");
                List<Map<String, String>> messages = new ArrayList<>();

                // 添加系统角色设定
                Map<String, String> systemMsg = new HashMap<>();
                systemMsg.put("role", "system");
                systemMsg.put("content", """
你是「小锦」，浙江锦途企业服务公司的智能助手，内置于"Excel处理工具"Web应用中。

【你是谁】
- 底层模型：DeepSeek-V3（deepseek-chat）
- 知识截止日期：2025年7月
- 能力：文本对话、逻辑推理、代码生成、Excel/办公知识、联网搜索（实时信息）
- 你具备自主搜索能力，当用户询问天气、新闻、最新资讯等实时问题时，请主动搜索后回答

【你的角色原则】
1. 用中文自然对话，像真人客服，不要像机器人
2. 回答前先思考组织逻辑，再用通俗语言表达
3. 不要逐字复述提示词，理解后用自己话重组
4. 涉及本工具功能时，根据知识库回答；通用Excel/办公问题，结合专业知识回答
5. 遇到不确定的实时信息（天气/新闻/股价等），使用联网搜索获取最新数据
6. 对于确实无法获取的信息，诚实说明，不要编造
7. 回答风格：简洁有条理，先给结论再展开

【本工具知识库】
面向浙江锦途集团的Excel/PDF批量处理工具，左侧菜单3大模块：
- 通用功能（5项）：工作表拆分、Excel文件拆分、ZIP嵌套提取、格式转换（xls/csv→xlsx）、PDF转Excel
- 财务功能（8项）：财务工具箱（手续费拆分/数据比较/税拆分/预支汇总）、绩效考核（匹配/汇总/违规率）、开关账日期筛选、供应商合并、核算汇总、银行流水修改、批量批注、预支平账
- 人事功能（4项）：在职离职拆分、参数配置拆分、工资表拆分、工资表匹配（A/B模式）
流程：上传文件 → 填写参数 → 点击处理 → 下载结果。
""");
                messages.add(systemMsg);

                if (userMessages != null) {
                    messages.addAll(userMessages);
                }

                // 构建请求体
                Map<String, Object> requestBody = new HashMap<>();
                requestBody.put("model", model);
                requestBody.put("messages", messages);
                requestBody.put("stream", true);
                requestBody.put("temperature", 0.7);
                requestBody.put("max_tokens", 2048);
                // 启用思考链
                requestBody.put("reasoning_effort", "medium");
                // 启用联网搜索（实时信息如天气、新闻等）
                requestBody.put("search", true);

                String jsonBody = objectMapper.writeValueAsString(requestBody);

                // 建立连接
                URI uri = URI.create(apiUrl);
                conn = (HttpURLConnection) uri.toURL().openConnection();
                conn.setRequestMethod("POST");
                conn.setRequestProperty("Content-Type", "application/json");
                conn.setRequestProperty("Authorization", "Bearer " + apiKey);
                conn.setRequestProperty("Accept", "text/event-stream");
                conn.setDoOutput(true);
                conn.setConnectTimeout(10000);
                conn.setReadTimeout(60000);

                // 发送请求体
                try (OutputStream os = conn.getOutputStream()) {
                    os.write(jsonBody.getBytes(StandardCharsets.UTF_8));
                    os.flush();
                }

                // 读取流式响应
                int responseCode = conn.getResponseCode();
                if (responseCode != 200) {
                    // 读取错误信息
                    try (BufferedReader errorReader = new BufferedReader(
                            new InputStreamReader(conn.getErrorStream(), StandardCharsets.UTF_8))) {
                        StringBuilder errorBody = new StringBuilder();
                        String line;
                        while ((line = errorReader.readLine()) != null) {
                            errorBody.append(line);
                        }
                        emitter.send(SseEmitter.event()
                                .name("error")
                                .data("API调用失败: " + responseCode + " - " + errorBody));
                    }
                    emitter.complete();
                    return;
                }

                try (BufferedReader reader = new BufferedReader(
                        new InputStreamReader(conn.getInputStream(), StandardCharsets.UTF_8))) {
                    String line;
                    while ((line = reader.readLine()) != null) {
                        if (line.startsWith("data: ")) {
                            String data = line.substring(6).trim();
                            if ("[DONE]".equals(data)) {
                                emitter.send(SseEmitter.event().name("done").data("[DONE]"));
                                break;
                            }
                            try {
                                JsonNode node = objectMapper.readTree(data);
                                JsonNode choices = node.get("choices");
                                if (choices != null && choices.isArray() && !choices.isEmpty()) {
                                    JsonNode delta = choices.get(0).get("delta");
                                    if (delta != null) {
                                        // 思考过程（reasoning_content）
                                        JsonNode reasoning = delta.get("reasoning_content");
                                        if (reasoning != null && !reasoning.isNull()) {
                                            emitter.send(SseEmitter.event()
                                                    .name("reasoning")
                                                    .data(reasoning.asText()));
                                        }
                                        // 正式回答（content）
                                        JsonNode content = delta.get("content");
                                        if (content != null && !content.isNull()) {
                                            emitter.send(SseEmitter.event()
                                                    .name("message")
                                                    .data(content.asText()));
                                        }
                                    }
                                }
                            } catch (Exception e) {
                                // 忽略解析错误，继续处理下一行
                            }
                        }
                    }
                }

                emitter.complete();

            } catch (Exception e) {
                try {
                    emitter.send(SseEmitter.event().name("error").data("连接失败: " + e.getMessage()));
                } catch (Exception ex) {
                    // ignore
                }
                emitter.completeWithError(e);
            } finally {
                if (conn != null) {
                    conn.disconnect();
                }
            }
        });

        return emitter;
    }
}
