"""
测试邮件附件
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

    # 查找rex相关的邮件
    for i in range(num_messages):
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

            # 只处理rex相关的邮件
            if 'rex' not in subject.lower():
                continue

            print(f"=== 邮件 {i+1} ===")
            print(f"主题: {subject}")
            print(f"日期: {msg.get('Date', '')}")
            print(f"\n附件列表:")

            # 遍历邮件的所有部分
            attachment_count = 0
            for part in msg.walk():
                # 获取内容类型
                content_type = part.get_content_type()
                content_disposition = part.get('Content-Disposition', '')

                print(f"  - Content-Type: {content_type}")
                print(f"    Content-Disposition: {content_disposition}")

                # 检查是否是附件
                if part.get_content_maintype() == 'multipart':
                    continue

                filename = part.get_filename()
                if filename:
                    # 解码文件名
                    decoded_parts = decode_header(filename)
                    decoded_filename = ''
                    for part_data, encoding in decoded_parts:
                        if isinstance(part_data, bytes):
                            if encoding:
                                decoded_filename += part_data.decode(encoding)
                            else:
                                decoded_filename += part_data.decode('utf-8', errors='ignore')
                        else:
                            decoded_filename += str(part_data)

                    attachment_count += 1
                    print(f"    附件 {attachment_count}: {decoded_filename}")

            print(f"\n总共 {attachment_count} 个附件\n")

        except Exception as e:
            print(f"处理邮件 {i+1} 失败: {e}\n")

    # 关闭连接
    mail.quit()
    print("[OK] 测试完成")

except Exception as e:
    print(f"[ERROR] 错误: {e}")
    import traceback
    traceback.print_exc()
