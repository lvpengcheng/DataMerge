"""
测试POP3邮件连接和获取
"""
import poplib
import email
from email.header import decode_header
import sys

# 设置输出编码
if sys.platform == 'win32':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# 配置
EMAIL = "lupech@163.com"
POP3_SERVER = "pop.163.com"
POP3_PORT = 995
PASSWORD = "YRbEUpDrNDTpCtHw"

print(f"连接到 {POP3_SERVER}:{POP3_PORT}...")

try:
    # 连接POP3服务器
    mail = poplib.POP3_SSL(POP3_SERVER, POP3_PORT)
    print("[OK] 连接成功")

    # 登录
    mail.user(EMAIL)
    mail.pass_(PASSWORD)
    print("[OK] 登录成功")

    # 获取邮件列表
    response, messages, octets = mail.list()
    num_messages = len(messages)
    print(f"[OK] 邮箱中共有 {num_messages} 封邮件\n")

    # 获取所有邮件的主题
    for i in range(num_messages):
        print(f"--- 邮件 {i+1} ---")
        try:
            response, lines, octets = mail.retr(i + 1)
            email_content = b'\n'.join(lines)
            msg = email.message_from_bytes(email_content)

            # 解码主题
            subject_header = msg.get('Subject', '')
            decoded_parts = decode_header(subject_header)
            subject = ''
            for part, encoding in decoded_parts:
                if isinstance(part, bytes):
                    if encoding:
                        subject += part.decode(encoding)
                    else:
                        subject += part.decode('utf-8', errors='ignore')
                else:
                    subject += str(part)

            # 获取日期
            date = msg.get('Date', '')

            print(f"主题: {subject}")
            print(f"日期: {date}")
            print()

        except Exception as e:
            print(f"获取邮件失败: {e}\n")

    # 关闭连接
    mail.quit()
    print("[OK] 测试完成")

except Exception as e:
    print(f"[ERROR] 错误: {e}")
    import traceback
    traceback.print_exc()
