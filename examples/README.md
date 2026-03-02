# 使用示例

## 快速开始

### 1. 安装依赖
```bash
pip install -r requirements.txt
```

### 2. 配置环境变量
```bash
# 复制环境变量示例文件
cp .env.example .env

# 编辑.env文件，设置AI API密钥
# OPENAI_API_KEY=your_openai_api_key_here
# 或
# ANTHROPIC_API_KEY=your_anthropic_api_key_here
```

### 3. 运行示例
```bash
python run.py
```

## 示例文件说明

### 1. 规则文件 (rules.md)
```
# 工资计算规则

## 数据源
1. employees.xlsx - 员工基本信息和基本工资
2. performance.xlsx - 绩效工资、津贴和扣款

## 计算规则
总工资 = 基本工资 + 绩效工资 + 津贴 - 扣款

## 输出要求
1. 包含员工ID、姓名、部门
2. 计算总工资
3. 添加合计行
```

### 2. 源文件示例
- `employees.xlsx`: 员工基本信息
- `performance.xlsx`: 绩效数据

### 3. 预期结果文件
- `expected_salary.xlsx`: 预期的工资计算结果

## API使用示例

### 训练模型
```python
import requests

# 训练请求
url = "http://localhost:8000/api/test_tenant/train"

files = {
    'rule_files': [('rules.md', open('rules.md', 'rb'))],
    'source_files': [
        ('employees.xlsx', open('employees.xlsx', 'rb')),
        ('performance.xlsx', open('performance.xlsx', 'rb'))
    ],
    'expected_result': ('expected_salary.xlsx', open('expected_salary.xlsx', 'rb'))
}

data = {
    'ai_provider': 'openai',
    'ai_model': 'gpt-4'
}

response = requests.post(url, files=files, data=data)
print(response.json())
```

### 计算数据
```python
import requests

# 计算请求
url = "http://localhost:8000/api/test_tenant/calculate"

files = {
    'data_files': [
        ('new_employees.xlsx', open('new_employees.xlsx', 'rb')),
        ('new_performance.xlsx', open('new_performance.xlsx', 'rb'))
    ]
}

response = requests.post(url, files=files)
print(response.json())
```

### 下载结果
```bash
# 直接访问下载链接
curl -O http://localhost:8000/api/test_tenant/download/result.xlsx
```

## 手动表头规则

对于复杂的Excel文件，可以指定手动表头规则：

```json
{
  "备忘录": {
    "工资调整项": [3, 3],
    "补当月工时": [6, 6]
  }
}
```

在训练请求中传入：
```python
data = {
    'ai_provider': 'openai',
    'ai_model': 'gpt-4',
    'manual_headers': '{"备忘录": {"工资调整项": [3, 3], "补当月工时": [6, 6]}}'
}
```

## 多租户支持

系统支持多租户，每个租户的数据完全隔离：

```python
# 租户A的训练
response_a = requests.post(
    "http://localhost:8000/api/tenant_a/train",
    files=files_a,
    data=data_a
)

# 租户B的训练（使用不同的规则和文件）
response_b = requests.post(
    "http://localhost:8000/api/tenant_b/train",
    files=files_b,
    data=data_b
)
```

## 系统架构

### 核心组件
1. **Excel解析器** (`excel_parser.py`): 智能解析复杂Excel文件
2. **AI引擎** (`ai_engine/`): 支持多种AI服务提供者
3. **训练引擎** (`training_engine.py`): 管理AI训练和代码生成
4. **代码沙箱** (`code_sandbox.py`): 安全执行生成的代码
5. **存储管理器** (`storage_manager.py`): 管理租户文件和数据
6. **文档验证器** (`document_validator.py`): 确保文档格式一致性

### 工作流程
1. **训练阶段**:
   - 上传规则、示例数据和预期结果
   - AI生成数据处理脚本
   - 验证脚本准确性
   - 保存验证通过的脚本

2. **计算阶段**:
   - 上传新数据
   - 验证文档格式
   - 执行保存的脚本
   - 返回处理结果

## 测试

运行所有测试：
```bash
pytest tests/ -v
```

运行特定测试：
```bash
pytest tests/test_excel_parser.py -v
pytest tests/test_ai_engine.py -v
pytest tests/test_integration.py -v
```

## 故障排除

### 常见问题

1. **AI API密钥未设置**
   ```
   错误: 未设置AI API密钥
   解决: 在.env文件中设置OPENAI_API_KEY或ANTHROPIC_API_KEY
   ```

2. **文件格式验证失败**
   ```
   错误: 文档格式验证失败
   解决: 确保上传的文件与训练时的格式一致
   ```

3. **内存不足**
   ```
   错误: 内存不足
   解决: 减少训练迭代次数或使用更小的文件
   ```

### 日志查看
```bash
# 查看应用日志
tail -f logs/app.log

# 查看错误日志
tail -f logs/error.log
```

## 性能优化建议

1. **文件大小**: 建议单个Excel文件不超过50MB
2. **训练迭代**: 默认5次迭代，可根据需要调整
3. **AI模型**: 对于简单任务可使用较小模型以节省成本
4. **缓存**: 系统会自动缓存训练结果，避免重复训练

## 安全注意事项

1. **API密钥**: 不要将API密钥提交到版本控制系统
2. **文件上传**: 系统会验证文件类型和大小
3. **代码执行**: 所有生成的代码都在沙箱中执行
4. **数据隔离**: 租户数据完全隔离，确保隐私安全