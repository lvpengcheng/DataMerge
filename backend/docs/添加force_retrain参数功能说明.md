# 添加force_retrain参数 - 强制重新训练功能

## 功能概述

添加了 `force_retrain` 参数，用于控制训练行为，支持三种训练模式：

1. **智能训练模式**（`force_retrain=False`，默认）
   - 历史最佳分数 = 100% → 直接使用历史最佳代码，跳过训练
   - 历史最佳分数 < 100% → 重新训练，尝试提升分数

2. **强制重新训练模式**（`force_retrain=True`）
   - 清除所有历史训练数据
   - 清除历史最佳代码
   - 从头开始全新训练

## 修改的文件

### 1. 训练引擎 (`ai_engine/training_engine.py`)

#### 修改 `train` 方法签名

```python
def train(
    self,
    source_files: List[str],
    expected_file: str,
    rule_files: List[str],
    manual_headers: Optional[Dict[str, Any]] = None,
    tenant_id: str = "default",
    salary_year: Optional[int] = None,
    salary_month: Optional[int] = None,
    monthly_standard_hours: Optional[float] = None,
    force_retrain: bool = False  # ← 新增参数
) -> Dict[str, Any]:
    """训练AI生成数据处理脚本

    Args:
        ...
        force_retrain: 是否强制重新训练（默认False）
            - False: 如果历史最佳分数=100%，直接使用历史最佳代码；如果<100%，重新训练
            - True: 清除所有历史训练数据和最佳代码，从头开始全新训练
    """
```

#### 添加强制重新训练逻辑

```python
# 处理强制重新训练
if force_retrain:
    self.training_logger.log_info("=" * 60)
    self.training_logger.log_info("强制重新训练模式：清除所有历史数据")
    self.training_logger.log_info("=" * 60)

    # 清除历史最佳分数和代码
    historical_best_file = self._get_historical_best_file(tenant_id)
    if historical_best_file.exists():
        historical_best_file.unlink()
        self.training_logger.log_info(f"已删除历史最佳分数文件: {historical_best_file}")

    self.training_logger.log_info("历史数据清除完成，开始全新训练")
```

#### 修改 `_train_formula_mode` 方法

```python
def _train_formula_mode(
    self,
    source_files: List[str],
    expected_file: str,
    rules_content: str,
    source_structure: Dict[str, Any],
    expected_structure: Dict[str, Any],
    manual_headers: Optional[Dict[str, Any]] = None,
    tenant_id: str = "default",
    force_retrain: bool = False  # ← 新增参数
) -> Dict[str, Any]:
```

#### 修改100%跳过逻辑

```python
# 如果不是强制重新训练，且历史最佳分数已经是100%，直接返回
if not force_retrain and historical_best_score >= 1.0 and historical_best_code:
    self.training_logger.log_info("历史最佳分数已达到100%，跳过训练，直接使用历史最佳代码")
    self.training_logger.end_training(historical_best_score, 0)
    return {
        "success": True,
        "best_score": historical_best_score,
        ...
    }
```

### 2. API端点 (`app/main.py`)

#### 修改普通训练API

```python
@app.post("/api/train")
async def train_model(
    tenant_id: str = Form(...),
    rule_files: List[UploadFile] = File(...),
    source_files: List[UploadFile] = File(...),
    expected_result: UploadFile = File(...),
    manual_headers: Optional[str] = Form(None),
    salary_year: Optional[int] = Form(None),
    salary_month: Optional[int] = Form(None),
    monthly_standard_hours: Optional[float] = Form(None),
    force_retrain: bool = Form(False)  # ← 新增参数
):
    """训练AI生成数据处理脚本

    Args:
        ...
        force_retrain: 是否强制重新训练（默认False）
            - False: 如果历史最佳分数=100%，直接使用历史最佳代码；如果<100%，重新训练
            - True: 清除所有历史训练数据和最佳代码，从头开始全新训练
    """
    ...
    training_result = training_engine.train(
        source_files=saved_files["source"],
        expected_file=saved_files["expected"],
        rule_files=saved_files["rules"],
        manual_headers=manual_headers_dict,
        tenant_id=tenant_id,
        salary_year=salary_year,
        salary_month=salary_month,
        monthly_standard_hours=monthly_standard_hours,
        force_retrain=force_retrain  # ← 传递参数
    )
```

#### 修改流式训练API

```python
@app.post("/api/train/stream")
async def train_model_stream(
    tenant_id: str = Form(...),
    rule_files: List[UploadFile] = File(...),
    source_files: List[UploadFile] = File(...),
    expected_result: UploadFile = File(...),
    manual_headers: Optional[str] = Form(None),
    salary_year: Optional[int] = Form(None),
    salary_month: Optional[int] = Form(None),
    monthly_standard_hours: Optional[float] = Form(None),
    force_retrain: bool = Form(False)  # ← 新增参数
):
    """流式训练AI生成数据处理脚本（支持实时日志输出）

    Args:
        ...
        force_retrain: 是否强制重新训练（默认False）
    """
    ...
    training_result = await loop.run_in_executor(
        executor,
        lambda: training_engine.train(
            source_files=saved_files["source"],
            expected_file=saved_files["expected"],
            rule_files=saved_files["rules"],
            manual_headers=manual_headers_dict,
            tenant_id=tenant_id,
            salary_year=salary_year,
            salary_month=salary_month,
            monthly_standard_hours=monthly_standard_hours,
            force_retrain=force_retrain  # ← 传递参数
        )
    )
```

## 使用场景

### 场景1：日常训练（默认行为）

```bash
# 不传force_retrain参数，或传False
curl -X POST "http://localhost:8000/api/train" \
  -F "tenant_id=达美乐62" \
  -F "rule_files=@规则.docx" \
  -F "source_files=@源数据.xlsx" \
  -F "expected_result=@预期结果.xlsx"
```

**行为**：
- 如果历史最佳分数 = 100% → 直接使用历史最佳代码（节省时间和成本）
- 如果历史最佳分数 < 100% → 重新训练，尝试提升分数

### 场景2：规则或数据变化，需要重新训练

```bash
# 传force_retrain=true
curl -X POST "http://localhost:8000/api/train" \
  -F "tenant_id=达美乐62" \
  -F "rule_files=@新规则.docx" \
  -F "source_files=@新源数据.xlsx" \
  -F "expected_result=@新预期结果.xlsx" \
  -F "force_retrain=true"
```

**行为**：
- 清除历史最佳分数文件
- 清除历史最佳代码
- 从头开始全新训练
- 生成适应新规则和数据的代码

### 场景3：调试或测试新的提示词

```bash
# 修改了提示词模板，需要重新训练验证效果
curl -X POST "http://localhost:8000/api/train" \
  -F "tenant_id=测试租户" \
  -F "rule_files=@规则.docx" \
  -F "source_files=@源数据.xlsx" \
  -F "expected_result=@预期结果.xlsx" \
  -F "force_retrain=true"
```

**行为**：
- 忽略历史最佳代码
- 使用新的提示词模板重新生成代码
- 验证新提示词的效果

## 日志输出

### 智能训练模式（历史最佳=100%）

```
历史最佳分数: 100.00%
历史最佳分数已达到100%，跳过训练，直接使用历史最佳代码
训练结束，总迭代次数: 0
```

### 智能训练模式（历史最佳<100%）

```
历史最佳分数: 85.50%
开始第 1/2 次迭代 (训练)
--- 公式模式迭代 1/2 ---
=== 公式模式：开始生成代码 ===
...
```

### 强制重新训练模式

```
============================================================
强制重新训练模式：清除所有历史数据
============================================================
已删除历史最佳分数文件: tenants/达美乐62/training/historical_best.json
历史数据清除完成，开始全新训练
============================================================
公式模式训练开始
说明: 基础列填充数据，计算列使用Excel公式
============================================================
开始第 1/2 次迭代 (训练)
--- 公式模式迭代 1/2 ---
=== 公式模式：开始生成代码 ===
...
```

## 前端集成建议

### 1. 添加复选框

```html
<form>
  <input type="file" name="rule_files" />
  <input type="file" name="source_files" />
  <input type="file" name="expected_result" />

  <!-- 新增：强制重新训练选项 -->
  <label>
    <input type="checkbox" name="force_retrain" value="true" />
    强制重新训练（清除历史数据，从头开始）
  </label>

  <button type="submit">开始训练</button>
</form>
```

### 2. 添加提示信息

```html
<div class="info-box">
  <h4>训练模式说明：</h4>
  <ul>
    <li><strong>智能训练（默认）</strong>：如果历史最佳分数已达100%，直接使用历史最佳代码，节省时间和成本</li>
    <li><strong>强制重新训练</strong>：清除所有历史数据，从头开始全新训练。适用于：
      <ul>
        <li>规则文件发生变化</li>
        <li>源数据结构发生变化</li>
        <li>需要测试新的提示词模板</li>
      </ul>
    </li>
  </ul>
</div>
```

### 3. JavaScript示例

```javascript
async function trainModel(formData) {
  const forceRetrain = document.getElementById('force_retrain').checked;

  if (forceRetrain) {
    const confirmed = confirm(
      '确定要强制重新训练吗？\n\n' +
      '这将清除所有历史训练数据和最佳代码，从头开始全新训练。\n' +
      '如果历史最佳分数已经很高，建议使用智能训练模式。'
    );

    if (!confirmed) {
      return;
    }
  }

  formData.append('force_retrain', forceRetrain);

  const response = await fetch('/api/train', {
    method: 'POST',
    body: formData
  });

  const result = await response.json();
  console.log('训练结果:', result);
}
```

## 优势

1. **节省成本**：历史最佳分数=100%时，不会重复调用AI，节省API费用
2. **节省时间**：跳过不必要的训练，快速返回结果
3. **灵活控制**：用户可以根据需要选择是否强制重新训练
4. **向后兼容**：默认值为False，不影响现有调用

## 注意事项

1. **历史最佳代码的有效性**：
   - 如果规则文件变化，历史最佳代码可能不再适用
   - 如果源数据结构变化，历史最佳代码可能无法执行
   - 建议在规则或数据变化时使用 `force_retrain=True`

2. **训练日志保留**：
   - 强制重新训练只删除历史最佳分数文件
   - 不删除 `training_logs` 目录中的历史训练日志
   - 便于追溯和调试

3. **并发训练**：
   - 如果多个请求同时训练同一个租户，可能会产生竞争
   - 建议在前端添加训练状态检查，避免并发训练

## 总结

✅ 添加了 `force_retrain` 参数，支持三种训练模式
✅ 修改了训练引擎和API端点
✅ 向后兼容，默认行为不变
✅ 提供了灵活的训练控制
✅ 节省了不必要的AI调用和成本

现在用户可以根据实际需求选择合适的训练模式，既能享受智能训练的便利，又能在需要时强制重新训练。
