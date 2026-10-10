"""SOP 通知邮件，使用现有 email_config.json 中的 SMTP 发件账户。"""

import logging
import smtplib
from email.message import EmailMessage

from ..email_processor.email_handler import EmailHandler

logger = logging.getLogger(__name__)


def send_sop_mail(recipient: str, subject: str, body: str) -> bool:
    if not recipient:
        logger.warning("[sop] 通知未发送：收件人邮箱为空")
        return False
    accounts = EmailHandler().config.get("email_accounts") or []
    account = next((item for item in accounts if item.get("smtp_server") and item.get("email_address")), None)
    if not account:
        logger.warning("[sop] 通知未发送：未配置 SMTP 发件账户")
        return False
    message = EmailMessage()
    message["From"] = account["email_address"]
    message["To"] = recipient
    message["Subject"] = subject
    message.set_content(body)
    try:
        server_type = smtplib.SMTP_SSL if account.get("smtp_ssl", True) else smtplib.SMTP
        with server_type(account["smtp_server"], int(account["smtp_port"]), timeout=20) as server:
            if account.get("smtp_starttls") and not account.get("smtp_ssl", True):
                server.starttls()
            server.login(account["email_address"], account["smtp_password"])
            server.send_message(message)
        return True
    except Exception:
        logger.exception("[sop] 通知发送失败: %s", recipient)
        return False
