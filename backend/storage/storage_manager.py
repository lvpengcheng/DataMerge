"""
存储管理器 - 管理租户文件和数据
"""

import json
import shutil
import hashlib
from pathlib import Path
from typing import Dict, List, Any, Optional, Tuple
import logging


class StorageManager:
    """存储管理器"""

    def __init__(self, base_dir: str = "tenants"):
        self.base_dir = Path(base_dir)
        self.base_dir.mkdir(exist_ok=True)
        self.logger = logging.getLogger(__name__)

    def get_tenant_dir(self, tenant_id: str) -> Path:
        """获取租户目录"""
        tenant_dir = self.base_dir / tenant_id
        tenant_dir.mkdir(exist_ok=True)
        return tenant_dir

    def save_training_files(
        self,
        tenant_id: str,
        rule_files: List[str],
        source_files: List[str],
        expected_file: str
    ) -> Dict[str, Any]:
        """保存训练文件"""
        tenant_dir = self.get_tenant_dir(tenant_id)
        training_dir = tenant_dir / "training"
        training_dir.mkdir(exist_ok=True)

        # 创建子目录
        rules_dir = training_dir / "rules"
        source_dir = training_dir / "source"
        expected_dir = training_dir / "expected"
        rules_dir.mkdir(exist_ok=True)
        source_dir.mkdir(exist_ok=True)
        expected_dir.mkdir(exist_ok=True)

        saved_files = {
            "rules": [],
            "source": [],
            "expected": None
        }

        # 保存规则文件
        for rule_file in rule_files:
            dest_path = rules_dir / Path(rule_file).name
            shutil.copy2(rule_file, dest_path)
            saved_files["rules"].append(str(dest_path))

        # 保存源文件
        for source_file in source_files:
            dest_path = source_dir / Path(source_file).name
            shutil.copy2(source_file, dest_path)
            saved_files["source"].append(str(dest_path))

        # 保存预期文件
        if expected_file:
            dest_path = expected_dir / Path(expected_file).name
            shutil.copy2(expected_file, dest_path)
            saved_files["expected"] = str(dest_path)

        # 保存文件信息
        file_info = {
            "tenant_id": tenant_id,
            "training_time": self._get_current_time(),
            "file_counts": {
                "rules": len(rule_files),
                "source": len(source_files),
                "expected": 1 if expected_file else 0
            },
            "file_paths": saved_files
        }

        info_file = training_dir / "training_info.json"
        with open(info_file, 'w', encoding='utf-8') as f:
            json.dump(file_info, f, ensure_ascii=False, indent=2)

        self.logger.info(f"保存训练文件完成，租户: {tenant_id}")
        return saved_files

    def save_script(
        self,
        tenant_id: str,
        script_content: str,
        training_result: Dict[str, Any],
        template_schema: Dict[str, Any]
    ) -> Dict[str, Any]:
        """保存生成的脚本"""
        tenant_dir = self.get_tenant_dir(tenant_id)
        scripts_dir = tenant_dir / "scripts"
        scripts_dir.mkdir(exist_ok=True)

        # 同时保存到训练文件夹
        training_dir = tenant_dir / "training"
        training_dir.mkdir(exist_ok=True)

        # 生成脚本ID（基于内容哈希）
        script_hash = hashlib.md5(script_content.encode('utf-8')).hexdigest()[:12]
        script_id = f"script_{script_hash}"

        # 保存脚本文件
        script_file = scripts_dir / f"{script_id}.py"
        with open(script_file, 'w', encoding='utf-8') as f:
            f.write(script_content)

        # 保存脚本信息
        script_info = {
            "script_id": script_id,
            "tenant_id": tenant_id,
            "created_time": self._get_current_time(),
            "score": training_result.get("best_score", 0),
            "iterations": training_result.get("total_iterations", 0),
            "success": training_result.get("success", False),
            "file_path": str(script_file),
            "template_schema": template_schema,
            "manual_headers": training_result.get("manual_headers"),
            "source_structure": training_result.get("source_structure", {}),
            "expected_structure": training_result.get("expected_structure", {}),
            "rules_content": training_result.get("rules_content", ""),
            "validation_rules": training_result.get("validation_rules", {}),
            "ai_provider": training_result.get("ai_provider"),
        }

        info_file = scripts_dir / f"{script_id}_info.json"
        with open(info_file, 'w', encoding='utf-8') as f:
            json.dump(script_info, f, ensure_ascii=False, indent=2)

        # 更新活跃脚本
        self._update_active_script(tenant_id, script_id, script_info)

        # 保存完整的训练结果到训练文件夹
        self._save_training_result(tenant_id, script_id, training_result, script_content)

        self.logger.info(f"保存脚本完成，租户: {tenant_id}, 脚本ID: {script_id}")
        return script_info

    def _save_training_result(
        self,
        tenant_id: str,
        script_id: str,
        training_result: Dict[str, Any],
        script_content: str
    ) -> None:
        """保存完整的训练结果到训练文件夹

        Args:
            tenant_id: 租户ID
            script_id: 脚本ID
            training_result: 训练结果
            script_content: 脚本内容
        """
        try:
            tenant_dir = self.get_tenant_dir(tenant_id)
            training_dir = tenant_dir / "training"
            training_dir.mkdir(exist_ok=True)

            # 创建训练结果目录
            result_dir = training_dir / script_id
            result_dir.mkdir(exist_ok=True)

            # 保存脚本文件
            script_file = result_dir / f"{script_id}.py"
            with open(script_file, 'w', encoding='utf-8') as f:
                f.write(script_content)

            # 准备训练结果数据
            result_data = {
                "script_id": script_id,
                "tenant_id": tenant_id,
                "saved_time": self._get_current_time(),
                "best_score": training_result.get("best_score", 0),
                "total_iterations": training_result.get("total_iterations", 0),
                "success": training_result.get("success", False),
                "log_file": training_result.get("log_file"),
                "training_summary": training_result.get("training_summary"),
                "manual_headers": training_result.get("manual_headers"),
                "source_structure": training_result.get("source_structure", {}),
                "expected_structure": training_result.get("expected_structure", {}),
                "rules_content": training_result.get("rules_content", ""),
                "iteration_results": training_result.get("iteration_results", []),
                "script_content_preview": script_content[:1000] + "..." if len(script_content) > 1000 else script_content
            }

            # 保存训练结果信息
            result_info_file = result_dir / "training_result.json"
            with open(result_info_file, 'w', encoding='utf-8') as f:
                json.dump(result_data, f, ensure_ascii=False, indent=2)

            # 保存每个迭代的详细信息
            iterations_dir = result_dir / "iterations"
            iterations_dir.mkdir(exist_ok=True)

            iteration_results = training_result.get("iteration_results", [])
            for i, iteration in enumerate(iteration_results):
                iteration_data = {
                    "iteration_number": iteration.get("iteration", i + 1),
                    "score": iteration.get("score", 0),
                    "error_description": iteration.get("error_description", ""),
                    "comparison_result": iteration.get("comparison_result", ""),
                    "execution_result": iteration.get("execution_result", {}),
                    "code_preview": iteration.get("raw_response", "")  # 保存原始的API响应
                }

                iteration_file = iterations_dir / f"iteration_{i+1}.json"
                with open(iteration_file, 'w', encoding='utf-8') as f:
                    json.dump(iteration_data, f, ensure_ascii=False, indent=2)

            self.logger.info(f"保存训练结果完成，租户: {tenant_id}, 脚本ID: {script_id}, 目录: {result_dir}")

        except Exception as e:
            self.logger.error(f"保存训练结果失败: {e}")

    def get_active_script(self, tenant_id: str) -> Optional[Dict[str, Any]]:
        """获取活跃脚本"""
        tenant_dir = self.get_tenant_dir(tenant_id)
        active_file = tenant_dir / "active_script.json"

        if not active_file.exists():
            return None

        try:
            with open(active_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            self.logger.error(f"读取活跃脚本失败: {e}")
            return None

    def get_script_content(self, tenant_id: str, script_id: str) -> Optional[str]:
        """获取脚本内容（磁盘优先，DB 兜底）"""
        tenant_dir = self.get_tenant_dir(tenant_id)
        script_file = tenant_dir / "scripts" / f"{script_id}.py"

        if script_file.exists():
            try:
                with open(script_file, 'r', encoding='utf-8') as f:
                    return f.read()
            except Exception as e:
                self.logger.error(f"读取脚本内容失败: {e}")

        # 磁盘没找到，尝试从 DB 加载
        try:
            from backend.database.connection import SessionLocal
            from backend.database.models import Script
            db = SessionLocal()
            try:
                db_script = (
                    db.query(Script)
                    .filter(
                        Script.tenant_id == tenant_id,
                        Script.is_active == True,
                        Script.name == script_id,
                    )
                    .order_by(Script.version.desc())
                    .first()
                )
                if db_script and db_script.code:
                    self.logger.info(f"从 DB 加载脚本: tenant={tenant_id}, name={script_id}, id={db_script.id}")
                    return db_script.code
            finally:
                db.close()
        except Exception as e:
            self.logger.warning(f"DB 脚本加载失败: {e}")

        return None

    def save_calculation_files(
        self, tenant_id: str, data_files: List[str]
    ) -> Dict[str, Any]:
        """保存计算文件"""
        tenant_dir = self.get_tenant_dir(tenant_id)
        calculation_dir = tenant_dir / "calculations"
        calculation_dir.mkdir(exist_ok=True)

        # 创建新的计算批次
        batch_id = self._generate_batch_id()
        batch_dir = calculation_dir / batch_id
        batch_dir.mkdir(exist_ok=True)

        input_dir = batch_dir / "input"
        output_dir = batch_dir / "output"
        input_dir.mkdir(exist_ok=True)
        output_dir.mkdir(exist_ok=True)

        saved_files = {
            "batch_id": batch_id,
            "input_files": [],
            "output_files": []
        }

        # 保存输入文件
        for data_file in data_files:
            dest_path = input_dir / Path(data_file).name
            shutil.copy2(data_file, dest_path)
            saved_files["input_files"].append(str(dest_path))

        # 保存批次信息
        batch_info = {
            "batch_id": batch_id,
            "tenant_id": tenant_id,
            "calculation_time": self._get_current_time(),
            "input_file_count": len(data_files),
            "status": "pending",
            "input_dir": str(input_dir),
            "output_dir": str(output_dir)
        }

        info_file = batch_dir / "batch_info.json"
        with open(info_file, 'w', encoding='utf-8') as f:
            json.dump(batch_info, f, ensure_ascii=False, indent=2)

        self.logger.info(f"保存计算文件完成，租户: {tenant_id}, 批次ID: {batch_id}")
        return saved_files

    def save_calculation_result(
        self, tenant_id: str, batch_id: str, result_file: str, execution_result: Dict[str, Any]
    ) -> Dict[str, Any]:
        """保存计算结果"""
        tenant_dir = self.get_tenant_dir(tenant_id)
        batch_dir = tenant_dir / "calculations" / batch_id

        if not batch_dir.exists():
            raise ValueError(f"批次目录不存在: {batch_id}")

        output_dir = batch_dir / "output"
        output_dir.mkdir(exist_ok=True)

        self.logger.info(f"保存计算结果，租户: {tenant_id}, 批次: {batch_id}")
        self.logger.info(f"源文件: {result_file}")
        self.logger.info(f"目标目录: {output_dir}")

        # 复制结果文件
        result_filename = Path(result_file).name
        dest_path = output_dir / result_filename

        # 规范化路径，避免Windows路径和POSIX路径的差异
        src_path = Path(result_file).resolve()
        dst_path = dest_path.resolve()

        # 检查是否是同一个文件
        if src_path == dst_path:
            self.logger.info(f"源文件和目标文件相同，跳过复制: {src_path}")
        else:
            # 确保目标目录存在
            dest_path.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(result_file, dest_path)
            self.logger.info(f"复制结果文件: {result_file} -> {dest_path}")

        # 更新批次信息
        info_file = batch_dir / "batch_info.json"
        if info_file.exists():
            with open(info_file, 'r', encoding='utf-8') as f:
                batch_info = json.load(f)

            batch_info.update({
                "status": "completed",
                "completion_time": self._get_current_time(),
                "result_file": str(dest_path),
                "execution_result": execution_result
            })

            with open(info_file, 'w', encoding='utf-8') as f:
                json.dump(batch_info, f, ensure_ascii=False, indent=2)

        self.logger.info(f"保存计算结果完成，租户: {tenant_id}, 批次ID: {batch_id}")
        return {
            "batch_id": batch_id,
            "result_file": str(dest_path),
            "status": "completed"
        }

    def get_storage_stats(self, tenant_id: str) -> Dict[str, Any]:
        """获取存储统计"""
        tenant_dir = self.get_tenant_dir(tenant_id)

        stats = {
            "tenant_id": tenant_id,
            "total_size": 0,
            "directory_stats": {},
            "file_counts": {
                "training": 0,
                "scripts": 0,
                "calculations": 0
            }
        }

        # 计算各目录大小
        for item in tenant_dir.iterdir():
            if item.is_dir():
                dir_size = self._get_directory_size(item)
                stats["directory_stats"][item.name] = {
                    "size_bytes": dir_size,
                    "size_human": self._format_size(dir_size)
                }
                stats["total_size"] += dir_size

                # 统计文件数量
                if item.name == "training":
                    stats["file_counts"]["training"] = self._count_files(item)
                elif item.name == "scripts":
                    stats["file_counts"]["scripts"] = self._count_files(item)
                elif item.name == "calculations":
                    stats["file_counts"]["calculations"] = self._count_files(item)

        stats["total_size_human"] = self._format_size(stats["total_size"])
        return stats

    def save_calculation_history(
        self,
        tenant_id: str,
        batch_id: str,
        salary_year: int,
        salary_month: int,
        output_file: str
    ) -> Dict[str, Any]:
        """将计算结果保存为历史数据

        Args:
            tenant_id: 租户ID
            batch_id: 计算批次ID
            salary_year: 薪资年份
            salary_month: 薪资月份
            output_file: 输出文件路径

        Returns:
            保存结果信息
        """
        import shutil

        tenant_dir = self.get_tenant_dir(tenant_id)
        history_dir = tenant_dir / "history" / f"{salary_year}_{salary_month:02d}"
        history_dir.mkdir(parents=True, exist_ok=True)

        # 复制输出文件到历史目录
        source_path = Path(output_file)
        dest_path = history_dir / "output.xlsx"
        if source_path.exists():
            shutil.copy2(source_path, dest_path)

        # 更新 calculation_history.json
        history_file = tenant_dir / "calculation_history.json"
        history_data = self.get_calculation_history(tenant_id)

        record = {
            "batch_id": batch_id,
            "salary_year": salary_year,
            "salary_month": salary_month,
            "output_file": str(dest_path),
            "calculation_time": self._get_current_time()
        }

        # 替换同年月的旧记录
        history_data["records"] = [
            r for r in history_data.get("records", [])
            if not (r["salary_year"] == salary_year and r["salary_month"] == salary_month)
        ]
        history_data["records"].append(record)

        # 更新最后计算信息
        history_data["last_batch_id"] = batch_id
        history_data["last_salary_year"] = salary_year
        history_data["last_salary_month"] = salary_month
        history_data["last_calculation_time"] = record["calculation_time"]

        with open(history_file, 'w', encoding='utf-8') as f:
            json.dump(history_data, f, ensure_ascii=False, indent=2)

        self.logger.info(f"历史数据已保存: 租户={tenant_id}, {salary_year}年{salary_month}月, 文件={dest_path}")
        return {"history_dir": str(history_dir), "output_file": str(dest_path)}

    def get_calculation_history(self, tenant_id: str) -> Dict[str, Any]:
        """读取租户的计算历史元数据

        Args:
            tenant_id: 租户ID

        Returns:
            计算历史数据，包含 records 列表和最后计算信息
        """
        tenant_dir = self.get_tenant_dir(tenant_id)
        history_file = tenant_dir / "calculation_history.json"

        if history_file.exists():
            try:
                with open(history_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception as e:
                self.logger.error(f"读取计算历史失败: {e}")

        return {"records": []}

    def _update_active_script(self, tenant_id: str, script_id: str, script_info: Dict[str, Any]):
        """更新活跃脚本"""
        tenant_dir = self.get_tenant_dir(tenant_id)
        active_file = tenant_dir / "active_script.json"

        active_info = {
            "script_id": script_id,
            "updated_time": self._get_current_time(),
            "script_info": script_info
        }

        with open(active_file, 'w', encoding='utf-8') as f:
            json.dump(active_info, f, ensure_ascii=False, indent=2)

    def _generate_batch_id(self) -> str:
        """生成批次ID"""
        import time
        import random
        timestamp = int(time.time())
        random_suffix = random.randint(1000, 9999)
        return f"batch_{timestamp}_{random_suffix}"

    def _get_current_time(self) -> str:
        """获取当前时间字符串"""
        from datetime import datetime
        return datetime.now().isoformat()

    def _get_directory_size(self, directory: Path) -> int:
        """计算目录大小"""
        total_size = 0
        for file_path in directory.rglob('*'):
            if file_path.is_file():
                total_size += file_path.stat().st_size
        return total_size

    def _format_size(self, size_bytes: int) -> str:
        """格式化文件大小"""
        for unit in ['B', 'KB', 'MB', 'GB']:
            if size_bytes < 1024.0:
                return f"{size_bytes:.2f} {unit}"
            size_bytes /= 1024.0
        return f"{size_bytes:.2f} TB"

    def _count_files(self, directory: Path) -> int:
        """统计文件数量"""
        count = 0
        for file_path in directory.rglob('*'):
            if file_path.is_file():
                count += 1
        return count