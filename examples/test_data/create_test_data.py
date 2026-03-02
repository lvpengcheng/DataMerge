"""
创建测试数据的脚本

测试场景：员工薪资计算
- 源文件1：员工基本信息（工号、姓名、部门、基本工资）
- 源文件2：考勤数据（工号、出勤天数、加班小时）
- 源文件3：绩效数据（工号、绩效等级、绩效系数）
- 预期结果：薪资汇总表

计算规则：
1. 应发工资 = 基本工资 × (出勤天数/21.75) + 加班费
2. 加班费 = 基本工资/21.75/8 × 加班小时 × 1.5
3. 绩效奖金 = 基本工资 × 绩效系数
4. 实发工资 = 应发工资 + 绩效奖金 - 社保扣款
5. 社保扣款 = 基本工资 × 10.5%
"""

import pandas as pd
from pathlib import Path

# 获取当前目录
current_dir = Path(__file__).parent

# ============ 源文件1：员工基本信息 ============
employee_data = {
    '工号': ['E001', 'E002', 'E003', 'E004', 'E005'],
    '姓名': ['张三', '李四', '王五', '赵六', '钱七'],
    '部门': ['技术部', '市场部', '技术部', '财务部', '市场部'],
    '基本工资': [10000, 9500, 12000, 8500, 11000]
}
df_employee = pd.DataFrame(employee_data)
df_employee.to_excel(current_dir / 'source' / '01_员工信息.xlsx', index=False, sheet_name='员工信息')
print("已创建: source/01_员工信息.xlsx")

# ============ 源文件2：考勤数据 ============
attendance_data = {
    '工号': ['E001', 'E002', 'E003', 'E004', 'E005'],
    '出勤天数': [21, 20, 22, 19, 21],
    '加班小时': [10, 5, 15, 0, 8]
}
df_attendance = pd.DataFrame(attendance_data)
df_attendance.to_excel(current_dir / 'source' / '02_考勤数据.xlsx', index=False, sheet_name='考勤记录')
print("已创建: source/02_考勤数据.xlsx")

# ============ 源文件3：绩效数据 ============
performance_data = {
    '工号': ['E001', 'E002', 'E003', 'E004', 'E005'],
    '绩效等级': ['A', 'B', 'S', 'C', 'A'],
    '绩效系数': [1.2, 1.0, 1.5, 0.8, 1.2]
}
df_performance = pd.DataFrame(performance_data)
df_performance.to_excel(current_dir / 'source' / '03_绩效数据.xlsx', index=False, sheet_name='绩效评定')
print("已创建: source/03_绩效数据.xlsx")

# ============ 计算预期结果 ============
# 合并数据
df_merged = df_employee.merge(df_attendance, on='工号').merge(df_performance, on='工号')

# 计算各项
df_merged['日薪'] = df_merged['基本工资'] / 21.75
df_merged['出勤工资'] = df_merged['日薪'] * df_merged['出勤天数']
df_merged['加班费'] = df_merged['日薪'] / 8 * df_merged['加班小时'] * 1.5
df_merged['应发工资'] = df_merged['出勤工资'] + df_merged['加班费']
df_merged['绩效奖金'] = df_merged['基本工资'] * df_merged['绩效系数']
df_merged['社保扣款'] = df_merged['基本工资'] * 0.105
df_merged['实发工资'] = df_merged['应发工资'] + df_merged['绩效奖金'] - df_merged['社保扣款']

# 保留两位小数
for col in ['日薪', '出勤工资', '加班费', '应发工资', '绩效奖金', '社保扣款', '实发工资']:
    df_merged[col] = df_merged[col].round(2)

# 选择输出列
output_columns = ['工号', '姓名', '部门', '基本工资', '出勤天数', '加班小时',
                  '绩效等级', '绩效系数', '应发工资', '绩效奖金', '社保扣款', '实发工资']
df_result = df_merged[output_columns]

df_result.to_excel(current_dir / 'expected' / '薪资汇总表.xlsx', index=False, sheet_name='薪资明细')
print("已创建: expected/薪资汇总表.xlsx")

# 打印预期结果
print("\n预期结果预览:")
print(df_result.to_string(index=False))
