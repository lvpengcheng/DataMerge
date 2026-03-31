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

- `backend/app/main.py` — Main FastAPI app (~5200 lines). Contains 40+ inline endpoint handlers plus router registrations
- `backend/api/` — Factored-out API routers (assets, compute, training, training_chat, rules, dashboard)
- `backend/ai_engine/` — AI code generation, prompt building, training loop, rule extraction
- `backend/database/` — SQLAlchemy models (14 tables), connection setup, DB init/migrations
- `backend/sandbox/code_sandbox.py` — Sandboxed execution of AI-generated code
- `backend/auth/` — JWT authentication (login, token creation, password hashing)
- `backend/admin/` — User/role/org/tenant management
- `backend/utils/` — Excel comparison, header matching, data validation, Aspose helpers
- `excel_parser.py` — Core Excel parser (~158KB), wraps Aspose.Cells .NET via pythonnet
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
