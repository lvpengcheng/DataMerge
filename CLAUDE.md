# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## User Preferences

- 每次回答都称呼用户为"rex"
- **默认不创建使用说明**：除非用户特殊指定，否则不需要创建使用说明.md文件
- **默认不创建测试脚本**：除非用户特殊说明，否则不创建测试脚本来验证代码修改
- **分步骤解决**：每次解决问题时，评估是否可以分步骤进行，控制上下文长度

## Common Commands

```bash
# Start dev server (port 8000, hot reload)
python run.py --start

# Or interactive menu (start / test / exit)
python run.py

# Initialize database
python -m backend.database.init_db

# Run tests
pytest backend/ -v
pytest backend/test_rule_extractor.py -v   # single test file

# Docker
docker-compose up -d

# Lint / format
black backend/
flake8 backend/
```

## 代码导航（先查表，再局部读）

**规则：本仓库多个核心文件超过 3000 行，整读会直接撑爆上下文窗口。**

`backend/app/main.py` 7,680 行 · `excel_parser.py` 5,105 行 · `backend/api/training_chat.py` 4,206 行 · `backend/ai_engine/formula_code_generator.py` 3,546 行 · `frontend/static/js/tools.js` 3,366 行 · `backend/ai_engine/ai_provider.py` 2,639 行

> 任何超过 500 行的文件：先用 Grep 定位符号名 → 再用 Read 的 `offset`/`limit` 读目标 ±80 行。**禁止整读。**

| 要改什么 | 去哪 |
|---|---|
| Excel 解析 / 表头识别 / 区域边界 | `excel_parser.py:1187` `IntelligentExcelParser`、`:323` `HeaderRuleEngine`、`:882` `EnhancedRowAnalyzer`、`:470` `ColumnConsistencyValidator`、`:723` `BoundaryCandidateEvaluator` |
| 公式模式代码生成 | `backend/ai_engine/formula_code_generator.py:51` `FormulaCodeGenerator`、`:1856` `load_source_data`、`:2004` `write_source_sheets`、`:2063` `find_source_sheet` |
| 模板填充（结果表） | `backend/ai_engine/template_code_generator.py:419` `fill_template`、`:1372` `_resolve_target_sheets`、`:1228` `_restore_number_formats` |
| 输出后处理 / 日期格式还原 | `backend/utils/output_postprocess.py:363` `restore_formats_from_template`、`:458` `normalize_date_formatted_values`、`:612` `restore_template_region_format`、`:989` `normalize_source_sheet_formats` |
| Excel 比对 | `backend/utils/excel_comparator.py:911` `compare_excel_files`、`:1465` 多表版、`:1518` `_compare_dataframes_core` |
| 合并 / 整合对比 | `backend/utils/merge_engine.py:36` `compute_header_fingerprint`、`:256` `norm_compare`；`backend/utils/integrate_engine.py:61` `build_key_index`、`:131` `_normalize_excel_formula` |
| 表头 / 列匹配 | `backend/utils/fast_header_matcher.py:44`、`backend/utils/smart_matcher.py:15`、`backend/utils/ai_source_mapping.py:5` |
| 模板与目标表定位 | `backend/utils/template_resolver.py:107` `resolve_template_path`、`backend/utils/target_sheet_resolver.py:75` `resolve_target_sheets` |
| 模板行列规划 / 清行 | `backend/utils/template_row_planner.py:80` `build_row_plan`、`:202`、`:265` `clean_template_rows` |
| 沙箱执行 | `backend/sandbox/code_sandbox.py:19` `CodeSandbox`、`:1086` `_execute_script_in_proc` |
| 子进程隔离 / 超时 / 内存护栏 | `backend/utils/subprocess_runner.py`、`backend/utils/subprocess_worker.py:69` `_run_task` |
| 上传并发闸门 | `backend/utils/upload_stream.py:12` `ExcelWorkGate` |
| AI 提供方 / 流式 / 思考流 | `backend/ai_engine/ai_provider.py`、`backend/ai_engine/training_logger.py:538` `StreamAwareAIProvider` |
| 提示词构造 | `backend/ai_engine/prompt_generator.py:13` `PromptGenerator` |
| 训练主循环 | `backend/ai_engine/training_engine.py:20` `TrainingEngine` |
| 表格结构分析 | `backend/ai_engine/table_analyzer.py:97` `TableAnalyzer` |

**`main.py` 内联端点速查**：训练 `:688` / `:1050`(SSE) · 智算 `:5454` / `:5684`(SSE) / `:1489` / `:2349`(split) · 对比 `:2774` · 调代码 `:3340` · 重校验 `:2932` · 加密检查 `:141` · 规则整理 `:175`/`:268` · 邮件 `:3096`/`:3148` · 存储与历史 `:1817`/`:1828` · 对话训练见 `backend/api/training_chat.py:35`

> 排除规则只认 `.claude/settings.json` 的 `permissions.deny`。`.claudeignore` 文件和 `settings.json` 的 `ignorePatterns` 键**从未生效**（已于 2026-09 删除），不要重新添加。

## Architecture Overview

**DataMerge** is an AI-driven Excel data integration system for HR/payroll scenarios (salary, attendance, social insurance, tax). It uses AI to generate Python scripts from user-provided rules, validates them against expected results through iterative refinement, then executes those scripts on new data.

### Tech Stack

- **Backend**: FastAPI (Python 3.11) — serves both REST API and server-rendered HTML pages
- **Frontend**: Vanilla HTML/CSS/JS in `frontend/templates/` and `frontend/static/` — no framework
- **Database**: SQLAlchemy ORM (PostgreSQL in production, SQLite `data.db` for local dev)
- **Excel parsing**: Aspose.Cells for .NET via pythonnet bridge (`excel_parser.py`) — NOT openpyxl for core parsing. Requires .NET runtime loaded into the Python process. DLLs live in `libs/`
- **AI providers**: Multi-provider abstraction in `backend/ai_engine/ai_provider.py` — supports OpenAI, Claude, DeepSeek, Ollama. Active provider configured via `AI_PROVIDER` in `.env`

### Core Workflows

**Training (智训)**: Upload rule docs + source Excel + expected result → AI generates Python script → runs in `CodeSandbox` → output compared via `excel_comparator` → differences fed back to AI for refinement → best script saved to DB.

**Computation (智算)**: Upload new data files → validate against training template → match headers → execute saved script in sandbox → return generated Excel.

### Two Code Generation Modes

1. **Formula mode** (`FormulaCodeGenerator`): Generates Python that writes Excel formulas (VLOOKUP, IF, etc.) — preferred, more transparent
2. **Modular mode** (`ModularCodeGenerator`): Generates pure Python computation code

Controlled by `USE_FORMULA_MODE` and `USE_MODULAR_GENERATION` in `.env`.

### Key Directories

- `backend/app/main.py` — Main FastAPI app (**7,680 lines**). Contains 40+ inline endpoint handlers plus router registrations — see 代码导航 above before reading
- `backend/api/` — Factored-out API routers (assets, compute, training, training_chat, rules, dashboard)
- `backend/ai_engine/` — AI code generation, prompt building, training loop, rule extraction
- `backend/database/` — SQLAlchemy models (14 tables), connection setup, DB init/migrations
- `backend/sandbox/code_sandbox.py` — Sandboxed execution of AI-generated code
- `backend/auth/` — JWT authentication (login, token creation, password hashing)
- `backend/admin/` — User/role/org/tenant management
- `backend/utils/` — Excel comparison, header matching, data validation, Aspose helpers
- `excel_parser.py` — Core Excel parser (**5,105 lines / ~237KB**), wraps Aspose.Cells .NET via pythonnet
- `aspose_init.py` — .NET runtime initialization for Aspose.Cells
- `tenants/` — Per-tenant isolated file storage (gitignored)
- `global_assets/` — Global reference data files shared across tenants
- `libs/` — .NET assemblies (Aspose.Cells.dll, SkiaSharp.dll, license file)

### API Routes

| Prefix | Router | Purpose |
|--------|--------|---------|
| `/api/auth` | `backend/auth/router.py` | Login, logout, JWT tokens |
| `/api/admin` | `backend/admin/router.py` | User/role/org/tenant CRUD |
| `/api/assets` | `backend/api/assets.py` | Data asset management |
| `/api/compute2` | `backend/api/compute.py` | Two-step compute flow |
| `/api/training` | `backend/api/training.py` | Training session history |
| `/api/training/chat` | `backend/api/training_chat.py` | Interactive chat-based training |
| `/api/rules` | `backend/api/rules.py` | Rule session CRUD |
| `/api/dashboard` | `backend/api/dashboard.py` | Tenant status overview |

Plus ~40 inline endpoints in `main.py` covering original training, calculation, download, comparison, email, and frontend page routes.

### Multi-Tenancy

Each tenant gets isolated storage at `tenants/{tenant_id}/` with sub-directories for training files, scripts, and results. Database-level authorization ties tenants to organizations via `tenant_authorizations` table.

### Aspose.Cells .NET Bridge

The project loads .NET Core runtime into the Python process via pythonnet to use Aspose.Cells for Excel parsing. This is a critical non-standard dependency:
- DLLs in `libs/` (Aspose.Cells.dll, SkiaSharp.dll)
- License file: `libs/Aspose.Total.NET.lic`
- Initialization: `aspose_init.py`
- Docker requires .NET 9 runtime base image
- Environment vars: `LD_LIBRARY_PATH`, `DOTNET_SYSTEM_GLOBALIZATION_INVARIANT`, `DOTNET_ROLL_FORWARD`
