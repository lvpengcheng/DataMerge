# 🤖 AI驱动的Excel数据整合SaaS系统

一个完全由AI驱动的SaaS系统，能够自动解析复杂的多表头Excel数据，根据用户提供的规则文件，生成并验证数据处理脚本，最终实现企业级数据自动化整合。

## 📋 项目结构

```
DataMerge/
├── backend/                    # 后端服务
│   ├── app/                   # FastAPI应用
│   │   ├── main.py           # 主应用入口
│   │   └── models.py         # 数据模型
│   ├── ai_engine/            # AI引擎组件
│   │   ├── ai_provider.py    # AI提供者（OpenAI/Claude/本地）
│   │   ├── training_engine.py # 训练引擎
│   │   └── prompt_generator.py # 提示词生成器
│   ├── storage/              # 文件存储管理
│   │   └── storage_manager.py
│   ├── sandbox/              # 代码沙箱执行环境
│   │   └── code_sandbox.py
│   ├── excel_parser.py       # Excel解析器
│   └── document_validator.py # 文档格式验证器
├── tenants/                   # 租户数据目录
├── tests/                    # 测试文件
│   ├── test_excel_parser.py
│   ├── test_ai_engine.py
│   └── test_integration.py
├── examples/                 # 示例文件
│   ├── README.md
│   └── demo.py
├── requirements.txt          # Python依赖
├── .env.example             # 环境变量示例
├── docker-compose.yml       # Docker编排
├── Dockerfile               # Docker镜像
├── run.py                   # 运行脚本
└── README.md                # 项目说明
```

## 🚀 快速开始

### 1. 环境准备

```bash
# 克隆项目
git clone <repository-url>
cd ai-excel-integration

# 创建虚拟环境
python -m venv venv

# 激活虚拟环境
# Windows:
venv\Scripts\activate
# Linux/Mac:
source venv/bin/activate

# 安装依赖
pip install -r requirements.txt
```

### 2. 配置环境变量

```bash
# 复制环境变量示例文件
cp .env.example .env

# 编辑.env文件，配置以下关键变量：
# - AI_API_KEY: 你的AI API密钥
# - DATABASE_URL: 数据库连接URL
# - REDIS_URL: Redis连接URL
```

### 3. 启动服务

#### 方式一：直接运行（开发环境）

```bash
# 启动FastAPI服务
python -m uvicorn backend.app.main:app --reload --port 8000

# 访问API文档
# http://localhost:8000/docs
```

#### 方式二：使用Docker（推荐）

```bash
# 构建并启动所有服务
docker-compose up -d

# 查看日志
docker-compose logs -f

# 停止服务
docker-compose down
```

## 📡 API接口

### 训练接口

**POST** `/api/{tenant_id}/train`

上传规则和数据文件，训练AI生成数据处理脚本

**请求参数：**
- `rule_files`: 规则文件列表（multipart/form-data）
- `source_files`: 源数据Excel文件列表
- `expected_result`: 预期结果Excel文件
- `ai_provider`: AI提供者（openai/claude/local）
- `ai_model`: AI模型名称
- `manual_headers`: 手动表头规则（JSON字符串，可选）

**响应示例：**
```json
{
  "task_id": "train_abc123",
  "status": "started",
  "message": "训练任务已启动",
  "files_uploaded": {
    "rules": 2,
    "source_data": 3,
    "expected_result": 1
  }
}
```

### 计算接口

**POST** `/api/{tenant_id}/calculate`

使用已训练脚本处理新数据。

**请求参数：**
- `data_files`: 新数据Excel文件列表（multipart/form-data）

**响应示例：**
```json
{
  "task_id": "calc_xyz789",
  "status": "started",
  "message": "计算任务已启动",
  "files_uploaded": 5
}
```

### 下载结果

**GET** `/api/{tenant_id}/download/{filename}`

下载处理结果文件。

### 存储统计

**GET** `/api/{tenant_id}/storage/stats`

获取租户存储使用统计。

## 🔧 核心功能

### 1. Excel解析器

[`excel_parser.py`](backend/excel_parser.py:1) - 支持复杂多级表头的Excel文件解析

```python
from backend.excel_parser import ExcelParser

# 初始化解析器
parser = ExcelParser(manual_headers={
    "备忘录": {
        "工资调整项": [3, 3],
        "补当月工时": [6, 6]
    }
})

# 解析Excel文件
data = parser.parse_excel_to_json("example.xlsx")
```

### 2. AI引擎

[`ai_engine/`](backend/ai_engine/__init__.py:1) - 支持多种AI提供者

```python
from backend.ai_engine import AIProviderFactory, TrainingEngine

# 创建AI提供者
provider = AIProviderFactory.create_provider("openai", {
    "api_key": "your-api-key",
    "model": "gpt-4"
})

# 创建训练引擎
engine = TrainingEngine(provider, max_iterations=5)
```

### 3. 代码沙箱

[`sandbox/code_sandbox.py`](backend/sandbox/code_sandbox.py:1) - 安全执行生成的代码

```python
from backend.sandbox.code_sandbox import CodeSandbox

sandbox = CodeSandbox(timeout=300, max_memory_mb=1024)
result = sandbox.execute_script(script_content, input_data)
```

### 4. 存储管理

[`storage/storage_manager.py`](backend/storage/storage_manager.py:1) - 租户文件管理

```python
from backend.storage.storage_manager import StorageManager

storage = StorageManager()
script_info = storage.save_script(tenant_id, script_content)
```

### 5. 文档格式验证 ⭐ 新功能

[`document_validator.py`](backend/document_validator.py:1) - 确保上传文档与训练模版格式一致

```python
from backend.document_validator import DocumentValidator

validator = DocumentValidator()

# 提取文档格式
schema = validator.extract_document_schema(parsed_data)

# 验证文档格式
is_valid, errors = validator.validate_document(document_data, template_schema)
```

**功能特点：**
- ✅ 自动提取并保存训练时的文档格式模版
- ✅ 计算时自动验证上传文档与模版的一致性
- ✅ 详细的错误提示，精确定位格式差异
- ✅ 验证Sheet名称、表头行范围、列数、表头内容等

**详细文档：** 查看 [`docs/DOCUMENT_VALIDATION.md`](docs/DOCUMENT_VALIDATION.md:1)

## 🛡️ 安全特性

- ✅ 代码沙箱隔离执行
- ✅ 租户数据完全隔离
- ✅ 文件类型和大小验证
- ✅ API密钥认证
- ✅ 资源使用限制
- ✅ 文档格式验证，防止格式不一致导致的错误

## 📊 监控和日志

- 所有操作都有详细日志记录
- 支持任务状态实时查询
- 存储使用统计
- 审计日志追踪

## 🧪 测试

```bash
# 运行所有测试
pytest

# 运行特定测试
pytest tests/test_excel_parser.py

# 生成覆盖率报告
pytest --cov=backend --cov-report=html
```

## 📝 开发指南

### 添加新的AI提供者

在 [`ai_engine/ai_provider.py`](backend/ai_engine/ai_provider.py:1) 中继承 [`BaseAIProvider`](backend/ai_engine/ai_provider.py:11) 类：

```python
class NewAIProvider(BaseAIProvider):
    def generate_code(self, prompt: str, **kwargs) -> str:
        # 实现代码生成逻辑
        pass
    
    def chat(self, messages: List[Dict[str, str]], **kwargs) -> str:
        # 实现对话逻辑
        pass
```

### 自定义表头解析规则

在训练请求中传入 `manual_headers` 参数：

```json
{
  "manual_headers": {
    "文件名": {
      "Sheet名": [起始行, 结束行]
    }
  }
}
```

## 🐳 Docker部署

### 构建镜像

```bash
docker build -t ai-excel-integration:latest .
```

### 使用Docker Compose

```bash
# 启动所有服务
docker-compose up -d

# 查看服务状态
docker-compose ps

# 查看日志
docker-compose logs -f backend

# 停止服务
docker-compose down
```

## 🔄 工作流程

1. **训练阶段**
   - 用户上传规则文档 + 示例Excel + 验证结果
   - 系统把示例excel、规则文档、验证结果分别保存到租户下的训练文件夹内
   - 使用excel_parse.py来提取示例excel的数据结构，主要是表头和表头所在列的letter以供后续计算时的数据的匹配关系，这个解析出来的结构要保存起来，以供后续计算时上传的文件和这些文件之间文件名和文件格式的对比
   - 使用excel_parse.py来提取验证结果excel的数据结构，后续生成的结果文件以这个文件结构来生成
   - 读取规则文件，把规则文件的所有内容解析成文本或者md格式，以供后续ai训练时使用
   - 把数据源解构、预期格式、数据生成规则喂给api，让api知道我们的目的是告诉他我有几个excel，其数据结构是什么，我想生成一个预期结果的数据结构，各个数据表，之间的引用或者计算关系如规则内容，请根据以上提示词，生成一个py语句，可以读取固定文件夹下的excel文件，生成最终结果，这个py语句要首先对文件名和文件格式进行验证，如果有错，直接抛出，如果都通过，则继续进行计算。计算的步骤可以先解析格式、设定mapping关系，填入主数据，根据数据规则内各列的引用关系，对这个excel进行逐列填充。注意，对excel进行数据读取时，也采用同样的excel_parse.py和接口传过来的同样的手动表头。
   - AI生成Python脚本（最多5次迭代）可调整，每次试算时，都要给出生成的excel和预期的excel之间的差异，把差异以提示词的方式传给api，让api进行修正代码
   - 验证通过后保存脚本和模版格式，如果迭代后还无法达到100%，则保留匹配率最高的那版代码作为这个租户的计算代码

2. **计算阶段**
   - 用户上传新数据Excel
   - **验证文档格式与模版一致性** ⭐
   - 如果格式不一致，返回详细错误信息并停止
   - 如果格式一致，加载已验证脚本
   - 在沙箱中执行脚本
   - 返回处理结果

## 🎯 实现完成总结

基于README.md的设计说明，我已经成功实现了完整的AI驱动的Excel数据整合SaaS系统。系统包含以下核心功能：

### ✅ 已实现功能

1. **智能Excel解析器** (`excel_parser.py`)
   - 支持复杂多级表头解析
   - 自动识别表头和数据区域
   - 支持手动表头规则指定

2. **多AI提供者支持** (`ai_engine/ai_provider.py`)
   - OpenAI GPT系列
   - Claude系列
   - 本地AI模拟（可扩展）

3. **训练引擎** (`ai_engine/training_engine.py`)
   - 自动生成数据处理代码
   - 最多5次迭代优化
   - 结果对比和自动修正

4. **代码沙箱** (`sandbox/code_sandbox.py`)
   - 安全执行生成的Python代码
   - 模块和函数访问控制
   - 资源使用限制

5. **多租户支持** (`storage/storage_manager.py`)
   - 租户数据完全隔离
   - 文件版本管理
   - 存储使用统计

6. **文档格式验证** (`document_validator.py`)
   - 自动提取文档格式模版
   - 上传文件格式一致性验证
   - 详细的错误提示

7. **完整的API接口** (`app/main.py`)
   - 训练接口: `/api/{tenant_id}/train`
   - 计算接口: `/api/{tenant_id}/calculate`
   - 下载接口: `/api/{tenant_id}/download/{filename}`
   - 统计接口: `/api/{tenant_id}/storage/stats`

### 🔧 使用方法

1. **安装依赖**:
   ```bash
   pip install -r requirements.txt
   ```

2. **配置环境**:
   ```bash
   cp .env.example .env
   # 编辑.env文件设置AI API密钥
   ```

3. **运行系统**:
   ```bash
   python run.py
   # 或直接运行: uvicorn backend.app.main:app --reload --port 8000
   ```

4. **访问API文档**:
   ```
   http://localhost:8000/docs
   ```

### 🧪 测试验证

系统包含完整的测试套件：
```bash
# 运行所有测试
pytest tests/ -v

# 运行演示
python examples/demo.py
```

### 🐳 Docker部署

```bash
# 使用Docker Compose
docker-compose up -d

# 查看服务状态
docker-compose ps

# 查看日志
docker-compose logs -f backend
```

### 📊 系统特点

- **完全自动化**: 从规则到可执行代码的完整流程
- **智能修正**: 基于结果对比的自动代码优化
- **安全可靠**: 代码沙箱执行，租户数据隔离
- **易于扩展**: 支持新的AI提供者和文件格式
- **企业级**: 支持多租户，完整的API接口

## 📞 技术支持

- 问题反馈：查看测试文件和示例
- API文档：启动服务后访问 `/docs`
- 演示脚本：`python examples/demo.py`

## 📄 许可证

MIT License

## 🙏 致谢

- FastAPI - 现代化的Web框架
- Pandas & Openpyxl - 强大的数据处理库
- OpenAI/Claude - AI能力支持
- 智能Excel解析器 - 基于C#版本的Python移植实现

---

**注意：** 这是一个完整的系统实现。在生产环境使用前，请确保：
1. 配置正确的AI API密钥
2. 设置适当的安全策略
3. 配置HTTPS和访问控制
4. 根据业务需求调整参数配置
