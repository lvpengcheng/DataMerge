# 规则整理独立页面实现计划

## 概述
创建独立的"规则整理"页面，带 AI 对话交互界面，用户上传文件后通过对话生成结构化 rules.md。

## 页面布局设计
```
┌──────────────────────────────────────────────────────┐
│ Header: [智训] [智算] [规则✓] [管理]     用户信息    │
├──────────────┬───────────────────────────────────────┤
│  文件上传区   │           AI 对话区                    │
│              │                                       │
│ 源文件:      │  ┌─────────────────────────────────┐  │
│ [选择文件]   │  │  AI: 已分析完成，以下是规则...    │  │
│              │  │  ...规则内容...                   │  │
│ 目标文件:    │  │                                  │  │
│ [选择文件]   │  │  用户: 请调整D列公式...           │  │
│              │  │                                  │  │
│ 设计文档:    │  │  AI: 好的，已修改...              │  │
│ [选择文件]   │  │                                  │  │
│              │  └─────────────────────────────────┘  │
│ AI模型:      │  ┌─────────────────────────────────┐  │
│ [Claude ▾]   │  │ 请输入补充说明或调整要求...       │  │
│              │  │                    [发送] [下载]  │  │
│ [开始整理]   │  └─────────────────────────────────┘  │
└──────────────┴───────────────────────────────────────┘
```

## 实现步骤

### Step 1: 后端 - AI Provider 增加流式 chat
**文件**: `backend/ai_engine/ai_provider.py`
- 检查各 Provider 是否已有 `chat_stream(messages, chunk_callback)` 方法
- 如没有，新增 `BaseAIProvider.chat_stream()` 基类方法
- 各子类（Claude/DeepSeek/OpenAI）实现流式聊天

### Step 2: 后端 - RuleOrganizer 增加流式 + 多轮对话
**文件**: `backend/ai_engine/rule_organizer.py`
- 新增 `organize_rules_stream(...)` — 使用 `chat_stream` 逐块回调
- 新增 `chat_followup(messages_history, chunk_callback)` — 多轮追问
- 保留原有 `organize_rules()` 不变

### Step 3: 后端 - 新增 API 端点
**文件**: `backend/app/main.py`
- `POST /api/rules/organize/stream` — SSE 流式规则整理
- `POST /api/rules/chat` — 多轮对话追问 (SSE)
- `GET /rules` — 页面路由

### Step 4: 前端 - 创建页面文件
- `frontend/templates/rules.html`
- `frontend/static/css/rules.css`
- `frontend/static/js/rules.js`

### Step 5: 导航更新
- 在 training.html、compute.html、admin.html 的导航栏中添加"规则"链接

## 新增文件
```
frontend/templates/rules.html
frontend/static/css/rules.css
frontend/static/js/rules.js
```

## 修改文件
```
backend/ai_engine/ai_provider.py      (流式chat)
backend/ai_engine/rule_organizer.py   (流式+多轮)
backend/app/main.py                   (API端点+页面路由)
frontend/templates/training.html      (导航)
frontend/templates/compute.html       (导航)
frontend/templates/admin.html         (导航)
```
