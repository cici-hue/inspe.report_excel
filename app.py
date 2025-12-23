
import re
from typing import Dict, List

def _find_first(pats: List[str], text: str, flags=re.DOTALL) -> str:
    for pat in pats:
        m = re.search(pat, text, flags)
        if m:
            return m.group(1).strip()
    return ""

def parse_fields(text: str) -> Dict[str, str]:
    fields: Dict[str, str] = {}

    # ---- 基础检验信息 ----
    fields["Inspection No."]  = _find_first([r"Inspection No\.\s*([A-Z0-9\-]+)"], text)
    fields["Inspection Seq."] = _find_first([r"Inspection Seq\.\s*(\d+)"], text)
    fields["Inspection Date"] = _find_first([r"Inspection Date\s*([A-Za-z]{3}\s\d{1,2},\s\d{2})"], text)

    # ---- PO / Split No. 与 PO Date：按表头定位后一行值 ----
    po_block = re.search(
        r"PO\s*/\s*Split No\.\s*PO Date\s*PO Type[^\n]*\n\s*([0-9]+)\s*([A-Za-z]{3}\s\d{1,2},\s\d{2})",
        text
    )
    if po_block:
        fields["PO / Split No."] = po_block.group(1).strip()
        fields["PO Date"]        = po_block.group(2).strip()
    else:
        # 兜底
        fields["PO / Split No."] = _find_first([r"PO\s*/\s*Split No\.\s*([0-9]+)"], text)
        fields["PO Date"]        = _find_first([r"PO Date\s*([A-Za-z]{3}\s\d{1,2},\s\d{2})"], text)

    # ---- Style No. 与 Item No.（重点加固）----
    # 1) 优先：从 Item Description 下一行抓“就近两串 6~8 位数字”，第一个→Style No.，第二个→Item No.
    item_desc_line = re.search(r"Item Description[\s\S]{0,200}?\n\s*(.+?)\n", text)
    two_nums = re.findall(r"\b(\d{6,8})\b", item_desc_line.group(1) if item_desc_line else "")
    style_from_desc = two_nums[0] if len(two_nums) >= 1 else ""
    item_from_desc  = two_nums[1] if len(two_nums) >= 2 else ""

    # 2) 兜底：各自标签直接抓（防止少数模板两数不在同一行）
    style_from_label = _find_first([r"Style No\.\s*([0-9A-Za-z/]+)"], text)
    item_from_label  = _find_first([r"Item No\.\s*([0-9A-Za-z/]+)"], text)

    # 3) 最终赋值（优先就近数字对；无则用标签）
    fields["Style No."] = style_from_desc or style_from_label
    fields["Item No."]  = item_from_desc  or item_from_label

    # ---- Delivered Quantity：优先总计行，再兜底头部 ----
    delivered = _find_first([
        r"Delivered Qty\.[\s\S]+?(\b\d{2,6}\b)\s*$",
        r"Delivered Quantity[\s\S]{0,60}?Item Quantity[\s\S]{0,30}?\n\s*[0-9]+\s*(\d{2,6})"
    ], text)
    fields["Delivered Quantity"] = delivered

    # ---- Customer / Dept（分拆）----
    m_cd = re.search(r"Customer\s*/\s*Dept\s*(.+?)\s*/\s*([0-9.]+)", text)
    if m_cd:
        fields["Customer"] = m_cd.group(1).strip()
        fields["Dept"]     = m_cd.group(2).strip()
    else:
        block = _find_first([r"Customer\s*/\s*Dept\s*([^\n]+)"], text)
        parts = [s.strip() for s in block.split("/") if block]
        fields["Customer"] = parts[0] if parts else ""
        fields["Dept"]     = parts[1] if len(parts) > 1 else ""

    # ---- Factory / FID Code（重点加固）----
    # 特定工厂强匹配（你这份是 Huangshan Yinghui ...）
    m_fac_specific = re.search(r"(Huangshan\s+Yinghui\s+Textile\s+Technology\s+Co\.,\s*Ltd\.)\s*/\s*([0-9]+)", text)
    if m_fac_specific:
        fields["Factory"]  = m_fac_specific.group(1).strip()
        fields["FID Code"] = m_fac_specific.group(2).strip()
    else:
        # 通用：先找到 “Factory / FID Code” 标题后的 “名称 / 数字”
        # 避免前面拼接了 “Inspection Location Customer / Dept …” 的噪声，先截取 Factory 段
        fac_block = re.search(r"Factory\s*/\s*FID Code\s*([\s\S]{0,200})", text)
        if fac_block:
            # 在该小块内部找 “名称 / 数字”
            m_fac = re.search(r"([^\n/][\s\S]*?)\s*/\s*([0-9.]+)", fac_block.group(1))
            if m_fac:
                name_clean = re.sub(r"Customer\s*/\s*Dept[\s\S]+", "", m_fac.group(1)).strip()
                fields["Factory"]  = name_clean
                fields["FID Code"] = m_fac.group(2).strip()
            else:
                fields["Factory"]  = ""
                fields["FID Code"] = ""
        else:
            fields["Factory"]  = ""
            fields["FID Code"] = ""

    # ---- Vendor（只要名称，不要编号）----
    m_vendor = re.search(r"Vendor\s*/\s*Vendor No\.\s*(.+?)\s*/\s*[0-9]+", text)
    fields["Vendor"] = m_vendor.group(1).strip() if m_vendor else ""

    return fields
