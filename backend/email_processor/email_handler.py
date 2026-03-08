"""
邮件处理模块 - 自动接收邮件、处理附件、触发计算
"""

import os
import re
import logging
import poplib
import smtplib
import email
import json
from email.header import decode_header
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from datetime import datetime, timezone
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Any
import shutil

logger = logging.getLogger(__name__)


class EmailHandler:
    """邮件处理器"""

    def __init__(self, config_file: str = "email_config.json"):
        """初始化邮件处理器

        Args:
            config_file: 配置文件路径
        """
        self.config_file = config_file
        self.config = self._load_config()

    def _load_config(self) -> Dict[str, Any]:
        """加载配置文件"""
        if os.path.exists(self.config_file):
            with open(self.config_file, 'r', encoding='utf-8') as f:
                config = json.load(f)
                # 确保processed_message_ids字段存在
                if "processed_message_ids" not in config:
                    config["processed_message_ids"] = []
                return config
        return {
            "last_check_time": datetime.now().isoformat(),
            "email_accounts": [],
            "processed_message_ids": []
        }

    def _save_config(self):
        """保存配置文件"""
        with open(self.config_file, 'w', encoding='utf-8') as f:
            json.dump(self.config, f, ensure_ascii=False, indent=2)

    def add_email_account(
        self,
        email_address: str,
        pop3_server: str,
        pop3_port: int,
        pop3_ssl: bool,
        pop3_password: str,
        smtp_server: str,
        smtp_port: int,
        smtp_ssl: bool,
        smtp_password: str,
        recipients: List[str] = None
    ) -> Dict[str, Any]:
        """添加邮件账户配置

        Args:
            email_address: 邮箱地址
            pop3_server: POP3服务器地址
            pop3_port: POP3端口号
            pop3_ssl: POP3是否使用SSL
            pop3_password: POP3密码或授权码
            smtp_server: SMTP服务器地址
            smtp_port: SMTP端口号
            smtp_ssl: SMTP是否使用SSL
            smtp_password: SMTP密码或授权码
            recipients: 收件人列表

        Returns:
            添加结果
        """
        account = {
            "email_address": email_address,
            "pop3_server": pop3_server,
            "pop3_port": pop3_port,
            "pop3_ssl": pop3_ssl,
            "pop3_password": pop3_password,
            "smtp_server": smtp_server,
            "smtp_port": smtp_port,
            "smtp_ssl": smtp_ssl,
            "smtp_password": smtp_password,
            "recipients": recipients or [],
            "last_check_time": self.config.get("last_check_time", datetime.now().isoformat())
        }

        # 检查是否已存在
        for i, acc in enumerate(self.config.get("email_accounts", [])):
            if acc["email_address"] == email_address:
                self.config["email_accounts"][i] = account
                self._save_config()
                return {"success": True, "message": "邮箱账户已更新"}

        if "email_accounts" not in self.config:
            self.config["email_accounts"] = []

        self.config["email_accounts"].append(account)
        self._save_config()

        return {"success": True, "message": "邮箱账户已添加"}

    def _connect_pop3(self, account: Dict[str, Any]) -> poplib.POP3:
        """连接POP3服务器"""
        if account["pop3_ssl"]:
            mail = poplib.POP3_SSL(account["pop3_server"], account["pop3_port"])
        else:
            mail = poplib.POP3(account["pop3_server"], account["pop3_port"])

        mail.user(account["email_address"])
        mail.pass_(account["pop3_password"])
        return mail

    def _decode_header_value(self, header_value: str) -> str:
        """解码邮件头"""
        if not header_value:
            return ""

        decoded_parts = decode_header(header_value)
        result = []

        for part, encoding in decoded_parts:
            if isinstance(part, bytes):
                if encoding:
                    try:
                        result.append(part.decode(encoding))
                    except:
                        result.append(part.decode('utf-8', errors='ignore'))
                else:
                    result.append(part.decode('utf-8', errors='ignore'))
            else:
                result.append(str(part))

        return ''.join(result)

    def _parse_subject(self, subject: str) -> Optional[Tuple[str, int, int]]:
        """解析邮件主题，提取租户名称、薪资年、薪资月

        Args:
            subject: 邮件主题

        Returns:
            (租户名称, 薪资年, 薪资月) 或 None
        """
        # 匹配格式: 租户名称_年份_月份 或 租户名称_年份-月份
        # 支持下划线和横杠分隔符，允许主题后面有额外文字
        pattern1 = r'^(.+?)_(\d{4})_(\d{1,2})'  # 租户_年_月 (后面可以有其他文字)
        pattern2 = r'^(.+?)_(\d{4})-(\d{1,2})'  # 租户_年-月 (后面可以有其他文字)

        match = re.match(pattern1, subject.strip())
        if not match:
            match = re.match(pattern2, subject.strip())

        if match:
            tenant_name = match.group(1)
            year = int(match.group(2))
            month = int(match.group(3))
            return (tenant_name, year, month)

        return None

    def _save_attachment(
        self,
        part: email.message.Message,
        save_dir: Path
    ) -> Optional[Path]:
        """保存邮件附件

        Args:
            part: 邮件部分
            save_dir: 保存目录

        Returns:
            保存的文件路径
        """
        filename = part.get_filename()
        if not filename:
            return None

        # 解码文件名
        filename = self._decode_header_value(filename)

        # 检查是否是Excel文件
        if not (filename.endswith('.xlsx') or filename.endswith('.xls')):
            logger.info(f"跳过非Excel文件: {filename}")
            return None

        save_dir.mkdir(parents=True, exist_ok=True)
        file_path = save_dir / filename

        # 保存文件
        with open(file_path, 'wb') as f:
            f.write(part.get_payload(decode=True))

        logger.info(f"附件已保存: {file_path}")
        return file_path

    def _fetch_emails_pop3(
        self,
        account: Dict[str, Any],
        since_date: datetime
    ) -> List[email.message.Message]:
        """从POP3服务器获取邮件

        Args:
            account: 邮箱账户配置
            since_date: 起始时间

        Returns:
            邮件列表
        """
        mail = self._connect_pop3(account)

        # 获取邮件列表
        response, messages, octets = mail.list()
        num_messages = len(messages)

        logger.info(f"邮箱中共有 {num_messages} 封邮件")

        # 获取已处理的Message-ID列表
        processed_ids = set(self.config.get("processed_message_ids", []))

        # 确保since_date有时区信息
        # last_check_time存的是本地时间，需要标记为本地时区（而非UTC）
        if since_date.tzinfo is None:
            local_tz = datetime.now(timezone.utc).astimezone().tzinfo
            since_date_aware = since_date.replace(tzinfo=local_tz)
        else:
            since_date_aware = since_date

        emails = []
        for i in range(num_messages):
            try:
                # 获取邮件内容
                response, lines, octets = mail.retr(i + 1)
                email_content = b'\n'.join(lines)
                msg = email.message_from_bytes(email_content)

                # 获取邮件主题和日期
                subject = self._decode_header_value(msg.get('Subject', ''))
                date_str = msg.get('Date')
                message_id = msg.get('Message-ID', '').strip()

                logger.info(f"邮件 {i+1}: 主题={subject}, 日期={date_str}, Message-ID={message_id}")

                # 检查是否已处理过（通过Message-ID去重）
                if message_id and message_id in processed_ids:
                    logger.info(f"邮件已处理过，跳过: Message-ID={message_id}")
                    continue

                # POP3需要手动过滤日期
                if date_str:
                    try:
                        email_date = email.utils.parsedate_to_datetime(date_str)
                        # 确保email_date有时区信息
                        if email_date.tzinfo is None:
                            email_date = email_date.replace(tzinfo=timezone.utc)

                        logger.info(f"邮件日期: {email_date}, 过滤时间: {since_date_aware}")

                        if email_date >= since_date_aware:
                            emails.append(msg)
                            logger.info(f"邮件符合时间条件，已添加")
                        else:
                            logger.info(f"邮件不符合时间条件，跳过")
                    except Exception as e:
                        logger.warning(f"解析邮件日期失败: {e}")
                        # 如果日期解析失败，也添加邮件
                        emails.append(msg)
                else:
                    # 如果没有日期，也添加邮件
                    logger.info(f"邮件没有日期信息，添加邮件")
                    emails.append(msg)

            except Exception as e:
                logger.error(f"获取邮件 {i+1} 失败: {e}", exc_info=True)

        mail.quit()
        logger.info(f"共获取到 {len(emails)} 封符合条件的邮件")
        return emails

    def fetch_new_emails(self, account: Dict[str, Any]) -> List[email.message.Message]:
        """获取新邮件

        Args:
            account: 邮箱账户配置

        Returns:
            新邮件列表
        """
        last_check_time = datetime.fromisoformat(account.get("last_check_time", datetime.now().isoformat()))

        logger.info(f"开始检查邮件，上次检查时间: {last_check_time}")

        try:
            emails = self._fetch_emails_pop3(account, last_check_time)
            logger.info(f"获取到 {len(emails)} 封新邮件")
            return emails

        except Exception as e:
            logger.error(f"获取邮件失败: {e}", exc_info=True)
            return []

    def process_email(
        self,
        msg: email.message.Message,
        storage_manager,
        excel_parser,
        ai_provider
    ) -> Dict[str, Any]:
        """处理单封邮件

        Args:
            msg: 邮件消息
            storage_manager: 存储管理器
            excel_parser: Excel解析器
            ai_provider: AI提供者

        Returns:
            处理结果
        """
        # 解析主题
        subject = self._decode_header_value(msg.get('Subject', ''))
        logger.info(f"处理邮件: {subject}")

        parsed = self._parse_subject(subject)
        if not parsed:
            logger.info(f"邮件主题格式不匹配: {subject}")
            return {"success": False, "message": "邮件主题格式不匹配"}

        tenant_name, salary_year, salary_month = parsed
        logger.info(f"解析结果 - 租户: {tenant_name}, 年份: {salary_year}, 月份: {salary_month}")

        # 检查租户是否有训练成功的脚本
        active_script = storage_manager.get_active_script(tenant_name)
        if not active_script:
            logger.warning(f"租户 {tenant_name} 没有活跃脚本")
            return {"success": False, "message": f"租户 {tenant_name} 没有训练成功的脚本"}

        # 检查脚本是否训练成功
        script_info = active_script.get("script_info", {})
        if not script_info.get("success"):
            logger.warning(f"租户 {tenant_name} 的脚本训练未成功")
            return {"success": False, "message": f"租户 {tenant_name} 没有训练成功的脚本"}

        # 创建临时目录
        tenant_dir = storage_manager.get_tenant_dir(tenant_name)
        calc_dir = tenant_dir / "calculations" / f"{salary_year}{salary_month:02d}"
        temp_dir = calc_dir / "temp"
        temp_dir.mkdir(parents=True, exist_ok=True)

        # 保存附件
        attachments = []
        for part in msg.walk():
            if part.get_content_maintype() == 'multipart':
                continue
            if part.get('Content-Disposition') is None:
                continue

            file_path = self._save_attachment(part, temp_dir)
            if file_path:
                attachments.append(file_path)

        if not attachments:
            logger.warning("邮件中没有Excel附件")
            return {"success": False, "message": "邮件中没有Excel附件"}

        logger.info(f"保存了 {len(attachments)} 个附件")

        # 匹配文件并重命名
        script_info = active_script.get("script_info", {})
        source_structure = script_info.get("source_structure", {})
        expected_file_count = len(source_structure.get("files", {}))

        matched_files = self._match_and_rename_files(
            attachments,
            source_structure,
            calc_dir,
            excel_parser
        )

        logger.info(f"匹配了 {len(matched_files)} 个文件，预期 {expected_file_count} 个")

        # 检查文件数量是否一致
        if len(matched_files) != expected_file_count:
            logger.warning(f"文件数量不一致: 实际 {len(matched_files)}, 预期 {expected_file_count}")
            return {
                "success": False,
                "message": f"文件数量不一致: 实际 {len(matched_files)}, 预期 {expected_file_count}",
                "matched_files": len(matched_files),
                "expected_files": expected_file_count
            }

        # 获取法定工作小时
        monthly_standard_hours = self._get_monthly_standard_hours(
            salary_year,
            salary_month,
            ai_provider
        )

        logger.info(f"{salary_year}年{salary_month}月法定工作小时: {monthly_standard_hours}")

        # 调用计算接口
        result = {
            "success": True,
            "tenant_name": tenant_name,
            "salary_year": salary_year,
            "salary_month": salary_month,
            "matched_files": len(matched_files),
            "monthly_standard_hours": monthly_standard_hours,
            "message": "文件已准备好，可以开始计算"
        }

        # 清理临时目录
        shutil.rmtree(temp_dir, ignore_errors=True)

        return result

    def _match_and_rename_files(
        self,
        attachments: List[Path],
        source_structure: Dict[str, Any],
        target_dir: Path,
        excel_parser
    ) -> List[Path]:
        """匹配文件并重命名

        Args:
            attachments: 附件文件列表
            source_structure: 训练时的源文件结构
            target_dir: 目标目录
            excel_parser: Excel解析器

        Returns:
            匹配并重命名后的文件列表
        """
        matched_files = []
        expected_files = source_structure.get("files", {})

        for attachment in attachments:
            try:
                # 解析附件结构
                parsed_sheets = excel_parser.parse_excel_file(str(attachment))

                # 尝试匹配
                best_match_file = None
                best_score = 0

                # 遍历训练时的文件
                for expected_filename, expected_file_info in expected_files.items():
                    expected_sheets = expected_file_info.get("sheets", {})

                    # 比较每个sheet的表头
                    for sheet_data in parsed_sheets:
                        for region in sheet_data.regions:
                            attach_headers = set(region.head_data.keys())

                            for expected_sheet_name, expected_sheet_info in expected_sheets.items():
                                expected_headers = set(expected_sheet_info.get("headers", {}).keys())

                                # 计算匹配度
                                if expected_headers and attach_headers:
                                    intersection = expected_headers & attach_headers
                                    score = len(intersection) / len(expected_headers)

                                    if score > best_score:
                                        best_score = score
                                        best_match_file = expected_filename

                # 如果匹配度足够高，重命名文件
                if best_match_file and best_score > 0.8:
                    new_path = target_dir / best_match_file

                    shutil.copy(attachment, new_path)
                    matched_files.append(new_path)
                    logger.info(f"文件匹配成功: {attachment.name} -> {best_match_file} (匹配度: {best_score:.2%})")
                else:
                    logger.warning(f"文件匹配失败: {attachment.name} (最高匹配度: {best_score:.2%})")

            except Exception as e:
                logger.error(f"处理附件失败 {attachment}: {e}", exc_info=True)

        return matched_files

    def _get_monthly_standard_hours(
        self,
        year: int,
        month: int,
        ai_provider=None
    ) -> float:
        """获取指定年月的法定工作小时（使用chinese_calendar计算）

        使用chinese_calendar库获取中国法定节假日和调休数据，
        准确计算该月的法定工作日数，乘以8小时。
        包含法定节假日扣除和调休补班。

        Args:
            year: 年份
            month: 月份
            ai_provider: 保留参数，不再使用

        Returns:
            法定工作小时
        """
        import calendar
        import datetime

        try:
            from chinese_calendar import is_workday

            # 获取该月的天数
            _, days_in_month = calendar.monthrange(year, month)

            # 使用chinese_calendar判断每天是否为工作日
            # 已包含法定节假日和调休补班的处理
            workdays = 0
            for day in range(1, days_in_month + 1):
                d = datetime.date(year, month, day)
                if is_workday(d):
                    workdays += 1

            hours = workdays * 8.0
            logger.info(f"{year}年{month}月: {workdays}个工作日(含调休), {hours}小时")
            return hours

        except ImportError:
            logger.warning("chinese_calendar未安装，使用基础日历计算（不含节假日调休）")
            # 降级：仅按周一到周五计算
            _, days_in_month = calendar.monthrange(year, month)
            workdays = 0
            for day in range(1, days_in_month + 1):
                weekday = calendar.weekday(year, month, day)
                if weekday < 5:
                    workdays += 1
            hours = workdays * 8.0
            logger.info(f"{year}年{month}月: {workdays}个工作日(不含调休), {hours}小时")
            return hours

        except NotImplementedError:
            logger.warning(f"chinese_calendar不支持{year}年数据，使用基础日历计算")
            _, days_in_month = calendar.monthrange(year, month)
            workdays = 0
            for day in range(1, days_in_month + 1):
                weekday = calendar.weekday(year, month, day)
                if weekday < 5:
                    workdays += 1
            hours = workdays * 8.0
            logger.info(f"{year}年{month}月: {workdays}个工作日(不含调休), {hours}小时")
            return hours

        except Exception as e:
            logger.error(f"计算工作小时失败: {e}")
            # 默认值：按照每月平均21.75个工作日计算
            return 174.0

    def mark_email_processed(self, message_id: str):
        """标记邮件为已处理

        Args:
            message_id: 邮件的Message-ID
        """
        if not message_id:
            return

        processed_ids = self.config.get("processed_message_ids", [])
        if message_id not in processed_ids:
            processed_ids.append(message_id)
            # 只保留最近500条记录，防止无限增长
            if len(processed_ids) > 500:
                processed_ids = processed_ids[-500:]
            self.config["processed_message_ids"] = processed_ids
            self._save_config()
            logger.info(f"已标记邮件为已处理: Message-ID={message_id}")

    def update_last_check_time(self, account_email: str):
        """更新上次检查时间

        Args:
            account_email: 邮箱地址
        """
        current_time = datetime.now().isoformat()

        for account in self.config.get("email_accounts", []):
            if account["email_address"] == account_email:
                account["last_check_time"] = current_time
                break

        self._save_config()
        logger.info(f"已更新上次检查时间: {current_time}")

    def send_result_email(
        self,
        account: Dict[str, Any],
        recipients: List[str],
        tenant_name: str,
        salary_year: int,
        salary_month: int,
        result_file_path: str,
        success: bool = True,
        error_message: str = ""
    ) -> Dict[str, Any]:
        """通过SMTP发送结果邮件

        Args:
            account: 邮箱账户配置
            recipients: 收件人列表
            tenant_name: 租户名称
            salary_year: 薪资年份
            salary_month: 薪资月份
            result_file_path: 结果文件路径
            success: 计算是否成功
            error_message: 错误信息

        Returns:
            发送结果
        """
        try:
            msg = MIMEMultipart()
            msg['From'] = account['email_address']
            msg['To'] = ', '.join(recipients)

            if success:
                msg['Subject'] = f"{tenant_name}_{salary_year}_{salary_month:02d}_计算结果"
                body = (
                    f"您好，\n\n"
                    f"{tenant_name} {salary_year}年{salary_month}月的薪资计算已完成。\n"
                    f"请查看附件中的计算结果。\n\n"
                    f"此邮件由系统自动发送，请勿回复。"
                )
            else:
                msg['Subject'] = f"{tenant_name}_{salary_year}_{salary_month:02d}_处理失败"
                body = (
                    f"您好，\n\n"
                    f"{tenant_name} {salary_year}年{salary_month}月的薪资计算处理失败。\n"
                    f"错误信息: {error_message}\n\n"
                    f"此邮件由系统自动发送，请勿回复。"
                )

            msg.attach(MIMEText(body, 'plain', 'utf-8'))

            # 添加附件
            if success and result_file_path and os.path.exists(result_file_path):
                with open(result_file_path, 'rb') as f:
                    attachment = MIMEApplication(f.read())
                    filename = os.path.basename(result_file_path)
                    attachment.add_header(
                        'Content-Disposition', 'attachment',
                        filename=filename
                    )
                    msg.attach(attachment)

            # 发送邮件
            if account.get('smtp_ssl', True):
                smtp = smtplib.SMTP_SSL(account['smtp_server'], account['smtp_port'])
            else:
                smtp = smtplib.SMTP(account['smtp_server'], account['smtp_port'])

            smtp.login(account['email_address'], account['smtp_password'])
            smtp.sendmail(account['email_address'], recipients, msg.as_string())
            smtp.quit()

            logger.info(f"结果邮件已发送至: {', '.join(recipients)}")
            return {"success": True, "message": "结果邮件已发送"}

        except Exception as e:
            logger.error(f"发送结果邮件失败: {e}", exc_info=True)
            return {"success": False, "message": f"发送结果邮件失败: {str(e)}"}
