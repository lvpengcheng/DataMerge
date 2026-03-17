# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 项目概述

这是一个AI驱动的Excel数据整合SaaS系统，能够自动解析复杂的多表头Excel数据，根据用户提供的规则文件，生成并验证数据处理脚本，实现企业级数据自动化整合。

## 核心架构

### 1. 训练-计算双阶段架构

**训练阶段** (`/api/train`):
- 用户上传规则文档 + 示例Excel + 预期结果
- AI生成Python脚本（支持迭代优化，最多5次）
- 每次迭代都会执行代码并与预期结果对比，生成差异报告
- 将差异作为提示词反馈给AI进行代码修正
- 保存最佳匹配代码（即使未达到100%）和文档格式模版

**计算阶段** (`/api/calculate`):
- 验证上传文档与训练模版格式一致性
- 加载已训练脚本在沙箱中执行
- 返回处理结果

### 2. 多租户隔离存储

所有租户数据存储在 `tenants/{tenant_id}/` 目录下：
```
tenants/{tenant_id}/
├── training/           # 训练文件
│   ├── rules/         # 规则文档
│   ├── source/        # 示例数据
│   └── expected/      # 预期结果
├── scripts/           # 生成的脚本
├── calculations/      # 计算结果
├── best_history.json  # 历史最佳代码和分数
└── active_script.json # 当前激活脚本
```

### 3. AI提供者架构

支持多种AI提供者（`backend/ai_engine/ai_provider.py`）：
- OpenAI (GPT-4, GPT-3.5)
- Anthropic (Claude)
- DeepSeek
- Ollama (本地模型)

所有提供者继承 `BaseAIProvider` 基类，实现统一接口。支持主提供者失败时自动切换到备用提供者。

### 4. 代码生成模式

系统支持两种代码生成模式：

**公式模式** (默认，推荐):
- 生成Excel公式而非Python代码
- 使用 `openpyxl` 直接写入公式到单元格
- 性能更好，更易调试

**模块化模式**:
- 生成结构化的Python代码
- 分模块生成：数据加载、映射、计算、输出
- 适合复杂业务逻辑

## 关键组件

### Excel解析器 (`excel_parser.py`)
- 支持复杂多级表头的智能识别
- 支持手动表头规则 (`manual_headers` 参数)
- 格式: `{"文件名": {"Sheet名": [起始行, 结束行]}}`

### 训练引擎 (`backend/ai_engine/training_engine.py`)
- 管理AI训练迭代过程
- 支持历史最佳代码缓存（避免重复训练）
- 支持强制重新训练 (`force_retrain=True`)
- 训练成功阈值可配置（`.env` 中的 `TRAINING_SUCCESS_THRESHOLD`）

### 代码沙箱 (`backend/sandbox/code_sandbox.py`)
- 安全执行AI生成的代码
- 限制模块访问和资源使用
- 超时控制和内存限制

### 文档验证器 (`backend/document_validator.py`)
- 提取并保存训练时的文档格式模版
- 计算时验证上传文档与模版一致性
- 验证Sheet名称、表头行范围、列数、表头内容

### Excel对比器 (`backend/utils/excel_comparator.py`)
- 对比生成结果与预期结果
- 计算匹配分数（0.0-1.0）
- 生成详细差异报告

## 开发命令

### 启动服务
```bash
# 开发模式（自动重载）
python -m uvicorn backend.app.main:app --reload --port 8000

# 生产模式
python -m uvicorn backend.app.main:app --host 0.0.0.0 --port 8000
```

### 运行测试
```bash
# 运行所有测试
pytest

# 运行特定测试
pytest tests/test_excel_parser.py

# 生成覆盖率报告
pytest --cov=backend --cov-report=html
```

### 环境配置
```bash
# 复制环境变量示例
cp .env.example .env

# 编辑 .env 配置AI API密钥和其他参数
```

## 重要配置项

### 训练配置 (`.env`)
- `MAX_TRAINING_ITERATIONS`: 最大训练迭代次数（默认5）
- `TRAINING_SUCCESS_THRESHOLD`: 训练成功阈值（默认0.95）
- `TRAINING_PERFECT_THRESHOLD`: 完美匹配阈值（默认1.0）
- `USE_FORMULA_MODE`: 是否使用公式模式（默认true）
- `USE_MODULAR_GENERATION`: 是否使用模块化生成（默认auto）

### AI提供者配置
- `AI_PROVIDER`: 主AI提供者（openai/claude/deepseek/ollama）
- `AI_FALLBACK_PROVIDER`: 备用提供者
- `OPENAI_API_KEY`, `ANTHROPIC_API_KEY`, `DEEPSEEK_API_KEY`: API密钥

### 沙箱配置
- `CODE_SANDBOX_TIMEOUT`: 代码执行超时时间（秒，默认300）
- `CODE_SANDBOX_MAX_MEMORY`: 最大内存限制（MB，默认1024）

## 关键工作流程

### 训练流程
1. 接收上传文件并保存到租户目录
2. 使用 `IntelligentExcelParser` 解析示例数据和预期结果的结构
3. 读取规则文档内容
4. 构建提示词（包含数据结构、规则、预期格式）
5. AI生成代码
6. 在沙箱中执行代码
7. 使用 `compare_excel_files` 对比结果
8. 如果未达到完美匹配，将差异反馈给AI进行修正
9. 重复步骤5-8，最多迭代 `MAX_TRAINING_ITERATIONS` 次
10. 保存最佳代码和文档格式模版

### 计算流程
1. 接收上传的新数据文件
2. 使用 `DocumentValidator` 验证文档格式与训练模版一致
3. 如果格式不一致，返回详细错误信息并停止
4. 加载租户的最佳脚本
5. 在沙箱中执行脚本处理新数据
6. 返回处理结果文件

### 历史最佳代码机制
- 每次训练后，如果分数高于历史最佳，更新 `best_history.json`
- 下次训练时，如果历史最佳分数已达到 `TRAINING_PERFECT_THRESHOLD`，直接使用历史代码（除非 `force_retrain=True`）
- 这避免了重复训练，节省API调用成本

## 前端页面

- `/training`: 训练页面（智训）
- `/compute`: 计算页面（智算）
- 使用原生JavaScript + HTML，无框架依赖
- 支持实时日志流式显示

## 注意事项

1. **并发训练保护**: 使用 `_training_locks` 防止同一租户并发训练导致数据冲突
2. **手动表头规则**: 对于复杂表头，必须通过 `manual_headers` 参数指定表头行范围
3. **代码沙箱限制**: 生成的代码只能访问白名单模块（pandas, openpyxl等），不能执行系统命令
4. **文件格式验证**: 计算时严格验证文档格式，防止格式不一致导致错误
5. **训练日志**: 所有训练过程都有详细日志，保存在 `tenants/{tenant_id}/training_logs/`
6. **迭代结果保存**: 每次迭代的代码和结果都保存在 `tenants/{tenant_id}/training/script_{id}/iterations/`

## 常见问题

### 如何添加新的AI提供者？
在 `backend/ai_engine/ai_provider.py` 中继承 `BaseAIProvider` 类，实现 `generate_code` 和 `chat` 方法，然后在 `AIProviderFactory` 中注册。

### 如何调整训练成功标准？
修改 `.env` 中的 `TRAINING_SUCCESS_THRESHOLD` 值（0.0-1.0）。

### 如何强制重新训练？
在训练请求中设置 `force_retrain=True`，这会清除历史最佳代码，从头开始训练。

### 如何处理只有表头的Excel？
系统会自动检测并处理只有表头的情况，不会抛出 KeyError。
