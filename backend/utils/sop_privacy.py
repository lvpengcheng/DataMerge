"""Per-review aliases for personal fields and a final prompt redaction pass."""
import re


class SopPrivacy:
    FIELDS = {
        "证件": ("身份证", "证件号", "护照", "idcard", "idnumber", "identity", "passport", "ssn"),
        "地址": ("地址", "住址", "居住地", "通讯地", "籍贯", "address"),
        "电话": ("手机", "电话", "phone", "mobile", "tel"),
        "邮箱": ("邮箱", "邮件", "email"),
        "账户": ("银行卡", "银行账号", "卡号", "开户账号", "accountnumber", "bankaccount"),
        "单位": ("客户名称", "公司名称", "单位名称", "企业名称", "companyname"),
        "姓名": ("姓名", "名字", "中文名", "英文名", "员工名称", "联系人", "收件人", "负责人", "employee_name", "fullname", "name"),
    }
    ID = re.compile(r"(?<!\d)(?:\d{17}[\dXx]|\d{15})(?!\d)")
    PHONE = re.compile(r"(?<!\d)1[3-9]\d{9}(?!\d)")
    EMAIL = re.compile(r"[\w.+-]+@[\w.-]+\.[A-Za-z]{2,}")
    ADDRESS = re.compile(r"[\u4e00-\u9fff]{2,}(?:省|自治区|市)[\u4e00-\u9fffA-Za-z0-9\-]{2,}(?:号|室|楼|弄|村)")
    LABEL = re.compile(
        r"(?P<label>姓名|员工姓名|联系人|身份证(?:号码|号)?|证件号码?|通讯地址|联系地址|家庭住址|地址|住址|手机号码?|联系电话|邮箱|银行账号|employee[ _]name|full[ _]name|\bname|id[ _]number|address|email)"
        r"(?P<sep>\s*[:：=]\s*)(?P<value>[^\n\r|,，;；<>]+)", re.IGNORECASE)

    def __init__(self):
        self.aliases = {}
        self.counts = {}

    @classmethod
    def field_kind(cls, label):
        normalized = re.sub(r"[\s_\-]+", "", str(label).lower())
        if len(normalized) > 45:
            return None
        for kind, words in cls.FIELDS.items():
            if any(re.sub(r"[\s_\-]+", "", word.lower()) in normalized for word in words):
                return kind
        return None

    def alias(self, value, kind):
        if value is None or not str(value).strip():
            return value
        raw = str(value).strip()
        if re.fullmatch(r"\[(?:姓名|证件|地址|电话|邮箱|账户|单位)\d+\]", raw):
            return raw
        if raw not in self.aliases:
            self.counts[kind] = self.counts.get(kind, 0) + 1
            self.aliases[raw] = f"[{kind}{self.counts[kind]:03d}]"
        return self.aliases[raw]

    def learn_text(self, text):
        fields = {}
        for line in str(text).splitlines():
            if '|' not in line:
                fields = {}
                continue
            cells = [cell.strip().strip('`*') for cell in line.strip().strip('|').split('|')]
            detected = {i: self.field_kind(cell) for i, cell in enumerate(cells) if self.field_kind(cell)}
            if detected:
                fields = detected
            elif fields and not all(re.fullmatch(r'[:\-\s]*', cell) for cell in cells):
                for i, kind in fields.items():
                    if i < len(cells) and cells[i] and not cells[i].startswith('['):
                        self.alias(cells[i], kind)
        for match in self.LABEL.finditer(str(text)):
            value = match["value"].strip().strip('"\'`*')
            if value and not value.startswith("["):
                self.alias(value, self.field_kind(match["label"]) or "姓名")
        for kind, pattern in (("证件", self.ID), ("电话", self.PHONE), ("邮箱", self.EMAIL), ("地址", self.ADDRESS)):
            for match in pattern.finditer(str(text)):
                self.alias(match[0], kind)

    def redact(self, text):
        text = str(text)
        self.learn_text(text)
        def redact_label(match):
            value = match['value'].strip().strip('"\'`*')
            alias = self.aliases.get(value)
            return match['label'] + match['sep'] + alias if alias else match[0]
        text = self.LABEL.sub(redact_label, text)
        if self.aliases:
            # Match once so replacement aliases are not processed as raw values again.
            tokens = r"\[(?:姓名|证件|地址|电话|邮箱|账户|单位)\d+\]"
            raw_patterns = []
            for raw in sorted(self.aliases, key=len, reverse=True):
                if raw.isdigit() and len(raw) < 6:
                    continue
                escaped = re.escape(raw)
                if re.fullmatch(r'[A-Za-z0-9]+', raw):
                    escaped = r'(?<![A-Za-z0-9])' + escaped + r'(?![A-Za-z0-9])'
                raw_patterns.append(escaped)
            if raw_patterns:
                pattern = re.compile(f"({tokens})|(?:" + "|".join(raw_patterns) + ")")
                text = pattern.sub(lambda match: match[0] if match[1] else self.aliases[match[0]], text)
        return text
