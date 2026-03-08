async def _send_result_email(
    account: Dict[str, Any],
    tenant_name: str,
    salary_year: int,
    salary_month: int,
    calc_result: Dict[str, Any]
) -> bool:
    """发送结果邮件
    Args:
        account: 邮箱账户配置
        tenant_name: 租户名称
        salary_year: 薪资年份
        salary_month: 薪资月份
        calc_result: 计算结果

  Returns:
        是否发送成功
    """
    try:
        import smtplib
        from email.mime.text import MIMEText
      from email.mime.multipart import MIMEMultipart
        from email.mime.application import MIMEApplication

        recipients = account.get("recipients", [])
        if not recipients:
         logger.warning("没有配置收件人，跳过发送邮件")
            return False

        # 检查是否已发送过（根据结果文件路径）
      saved_output_file = calc_result.get("saved_output_file")
        if saved_output_file:
            sent_record_file = Path(saved_output_file).parent / ".email_sent"
            if sent_record_file.exists():
                logger.info(f"该计算结果已发送过邮件，跳过发送: {saved_output_file}")
                return False

        # 创建邮件
        msg = MIMEMultipart()
        msg['From'] = account["email_address"]
        msg['To'] = ', '.join(recipients)
        msg['Subject'] = f"{tenant_name}_{salary_year}_{salary_month}_计算结果"

      # 邮件正文
        body = f"""
{tenant_name} {salary_year}年{salary_month}月薪资计算已完成

计算状态: {'成功' if calc_result.get('success') else '失败'}
处理时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

"""
        if calc_result.get("success"):
            body += f"\n下载链接: {calc_result.get('download_url', '无')}\n"
        else:
            body += f"\n错误信息: {calc_result.get('error', '未知错误')}\n"

        msg.attach(MIMEText(body, 'plain', 'utf-8'))

        # 附加结果文件
        if saved_output_file:
            output_file = Path(saved_output_file)
            if output_file.exists():
                with open(output_file, 'rb') as f:
              attachment = MIMEApplication(f.read(), _subtype="xlsx")
                    attachment.add_header('Content-Disposition', 'attachment', filename=output_file.name)
                    msg.attach(attachment)

        # 发送邮件
        if account["smtp_ssl"]:
     server = smtplib.SMTP_SSL(account["smtp_server"], account["smtp_port"])
        else:
            server = smtplib.SMTP(account["smtp_server"], account["smtp_port"])

        server.login(account["email_address"], account["smtp_password"])
        server.send_message(msg)
        server.quit()

        logger.info(f"结果邮件已发送给: {', '.join(recipients)}")

        # 记录已发送
        if saved_output_file:
            sent_record_file = Path(saved_output_file).parent / ".email_sent"
            with open(sent_record_file, 'w', encoding='utf-8') as f:
              f.write(f"{datetime.now().isoformat()}\n")
                f.write(f"Recipients: {', '.join(recipients)}\n")

        return True

    except Exception as e:
        logger.error(f"发送结果邮件失败: {e}", exc_info=True)
        return False
