
# -*- coding: utf-8 -*-
"""
AQL报告 PDF → Excel 网页应用（支持批量上传；仅合并Excel下载；不含 Quality Digit）
字段（顺序固定）：
Inspection No., Inspection Seq., Inspection Date,
PO / Split No., PO Date,
Style No., Item No., Delivered Quantity,
Customer, Dept, Factory, FID Code, Vendor
"""

import re
from io import BytesIO
from typing import Dict, List
import streamlit as st
import pandas as pd
from PyPDF2 import PdfReader

# ------------------ 页面设置 ------------------
st.set_page_config(page_title="AQL PDF→Excel（批量合并）", page_icon="📄", layout="centered")
st.title("📄 AQL 报告 PDF → Excel 解析器（批量合并）")
st.caption("上传一份或多份 PDF，自动提取 13 个字段并生成合并 Excel（不含 Quality Digit）。")

# ------------------ 字段与列顺序 ------------------
COLUMNS = [
    "Inspection No.", "Inspection Seq.", "Inspection Date",
    "PO / Split No.", "PO Date",
    "Style No.", "Item No.", "Delivered Quantity",
    "Customer", "Dept", "Factory", "FID Code", "Vendor"
]

# ------------------ 工具函数 ------------------
def _clean_text(text: str) -> str:
    """基础清理：去制表符/回车、软连字符，统一换行"""
    text = re.sub(r"[\t\r]+", " ", text)
    text = re.sub(r"\u00ad", "", text)  # 软连字符
    return text

def _extract_text_from_pdf(file_bytes: bytes) -> str:
    """从上传的PDF字节中提取文本（逐页拼接）"""
    reader = PdfReader(BytesIO(file_bytes))
    pages = []
    for page in reader.pages:
        pages.append(page.extract_text() or "")
    full = "\n\n".join(pages)
    return _clean_text(full)

def _find_first(pats: List[str], text: str, flags=re.DOTALL) -> str:
    """按给定正则列表，返回首个命中结果（捕获组1），未命中返回空串"""
    for pat in pats:
        m = re.search(pat, text, flags)
        if m:
            return m.group(1).strip()
    return ""

def parse_fields(text: str) -> Dict[str, str]:
    """解析指定的 13 个字段（无分类，直接键值）"""
    fields: Dict[str, str] = {}

    # 基本检验信息
    fields["Inspection No."]  = _find_first([r"Inspection No\.\s*([A-Z0-9\-]+)"], text)
    fields["Inspection Seq."] = _find_first([r"Inspection Seq\.\s*(\d+)"], text)
    fields["Inspection Date"] = _find_first([r"Inspection Date\s*([A-Za-z]{3}\s\d{1,2},\s\d{2})"], text)

    # PO / Split No. 与 PO Date：按表头定位后读取下一行值（抗换行/跨列）
    po_block = re.search(
        r"PO\s*/\s*Split No\.\s*PO Date\s*PO Type[^\n]*\n\s*([0-9]+)\s*([A-Za-z]{3}\s\d{1,2},\s\d{2})",
        text
    )
    if po_block:
        fields["PO / Split No."] = po_block.group(1).strip()
        fields["PO Date"]        = po_block.group(2).strip()
    else:
        # 兜底策略（直接逐项匹配）
        fields["PO / Split No."] = _find_first([r"PO\s*/\s*Split No\.\s*([0-9]+)"], text)
        fields["PO Date"]        = _find_first([r"PO Date\s*([A-Za-z]{3}\s\d{1,2},\s\d{2})"], text)

    # Style No. 与 Item No.
    # 根据模板，“Item Description”下一行通常含两个 6~8位数字：如 "... 43145156 906730 ..."
    item_line = re.search(r"Item Description[\s\S]{0,160}?\n\s*(.+?)\n", text)
    nums = re.findall(r"\b(\d{6,8})\b", item_line.group(1) if item_line else "")
    if len(nums) >= 2:
        fields["Style No."] = nums[0]
        fields["Item No."]  = nums[1]
    else:
        fields["Style No."] = _find_first([r"Style No\.\s*([0-9A-Za-z/]+)"], text)
        fields["Item No."]  = _find_first([r"Item No\.\s*([0-9A-Za-z/]+)"], text)

    # Delivered Quantity（优先取“Delivered Qty.”总计；失败则取头部“Delivered Quantity”）
    delivered = _find_first([
        r"Delivered Qty\.[\s\S]+?(\b\d{2,6}\b)\s*$",                       # 表格末行总计（如 528）
        r"Delivered Quantity[\s\S]{0,60}?Item Quantity[\s\S]{0,30}?\n\s*[0-9]+\s*(\d{2,6})"  # 头部明细
    ], text)
    fields["Delivered Quantity"] = delivered

    # Customer / Dept（分拆）
    m_cd = re.search(r"Customer\s*/\s*Dept\s*(.+?)\s*/\s*([0-9.]+)", text)
    if m_cd:
        fields["Customer"] = m_cd.group(1).strip()
        fields["Dept"]     = m_cd.group(2).strip()
    else:
        block = _find_first([r"Customer\s*/\s*Dept\s*([^\n]+)"], text)
        parts = [s.strip() for s in block.split("/") if block]
        fields["Customer"] = parts[0] if parts else ""
        fields["Dept"]     = parts[1] if len(parts) > 1 else ""

    # Factory / FID Code（强匹配该厂名+FID；否则通用匹配）
    m_fac_spec = re.search(r"Huangshan\s+Yinghui\s+Textile\s+Technology\s+Co\.,\s*Ltd\.\s*/\s*([0-9]+)", text)
    if m_fac_spec:
        fields["Factory"]  = "Huangshan Yinghui Textile Technology Co., Ltd."
        fields["FID Code"] = m_fac_spec.group(1).strip()
    else:
        m_fac = re.search(r"Factory\s*/\s*FID Code\s*(.+?)\s*/\s*([0-9.]+)", text)
        fields["Factory"]  = m_fac.group(1).strip() if m_fac else ""
        fields["FID Code"] = m_fac.group(2).strip() if m_fac else ""

    # Vendor 名称（不需要编号）
    m_vendor = re.search(r"Vendor\s*/\s*Vendor No\.\s*(.+?)\s*/\s*[0-9]+", text)
    fields["Vendor"] = m_vendor.group(1).strip() if m_vendor else ""

    return fields

def to_excel_bytes(rows: List[Dict[str, str]]) -> bytes:
    """将多行写入合并Excel并返回字节流"""
    df = pd.DataFrame(rows, columns=COLUMNS)
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    bio.seek(0)
    return bio.read()

# ------------------ 上传与解析 ------------------
uploaded_files = st.file_uploader(
    "上传 AQL 报告 PDF 文件（可多选）",
    type=["pdf"],
    accept_multiple_files=True
)

if uploaded_files:
    parsed_rows: List[Dict[str, str]] = []
    debug_blocks: List[str] = []  # 原文文本片段（可折叠查看）

    with st.status("正在解析PDF…", expanded=False) as status:
        for f in uploaded_files:
            try:
                text = _extract_text_from_pdf(f.getvalue())
                fields = parse_fields(text)
                parsed_rows.append(fields)
                debug_blocks.append(text[:3000] + ("\n...\n" if len(text) > 3000 else ""))
            except Exception as e:
                st.error(f"文件 {f.name} 解析失败：{e}")
        status.update(label="解析完成", state="complete")

    # 预览表格
    st.subheader("解析结果预览（合并表）")
    df_preview = pd.DataFrame(parsed_rows, columns=COLUMNS)
    st.dataframe(df_preview, use_container_width=True)

    # 下载：合并Excel（所有文件一张表）
    merged_excel_bytes = to_excel_bytes(parsed_rows)
    st.download_button(
        label="⬇️ 下载合并 Excel（所有文件）",
        data=merged_excel_bytes,
        file_name="AQL_Parsed_All.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # 调试：原始文本片段（可折叠）
    with st.expander("查看原始文本片段（调试用）", expanded=False):
        for i, txt in enumerate(debug_blocks, start=1):
            st.markdown(f"**文件 {i} 文本片段**")
            st.code(txt, language="text")

else:
    st.info("请选择一份或多份 PDF 文件进行上传。")
