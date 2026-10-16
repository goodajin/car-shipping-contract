import streamlit as st
from docxtpl import DocxTemplate
import pandas as pd
import re
import os
import subprocess
import tempfile

# ===================== 密码配置区，在这里修改你的密码 =====================
PASSWORD = "770880"  # 修改这里改成你自己的登录密码
# ========================================================================

# -------------------------- 登录校验函数 --------------------------
def check_password():
    # 初始化session状态
    if "authenticated" not in st.session_state:
        st.session_state["authenticated"] = False

    if st.session_state["authenticated"]:
        return True

    st.title("🔐 车辆托运合同生成器 - 登录")
    input_pwd = st.text_input("请输入访问密码", type="password")
    if st.button("登录"):
        if input_pwd == PASSWORD:
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("密码错误，请重试")
    st.stop()


# -------------------------- 【复用你原来的解析函数】 --------------------------
def parse_raw_text(raw_text: str):
    lines = raw_text.strip().splitlines()
    raw_data = {}
    for line in lines:
        line = line.strip()
        if not line:
            continue
        if ":" in line:
            key_part, val_part = line.split(":", maxsplit=1)
            val = val_part.strip()
            key_clean = re.sub(r"\s+", "", key_part)
            raw_data[key_clean] = val

    # 费用计算容错
    try:
        board_fee = int(raw_data.get("板费", 0))
    except ValueError:
        board_fee = 0
    try:
        deposit = int(raw_data.get("定金", 0))
    except ValueError:
        deposit = 0
    balance = board_fee - deposit

    route_str = raw_data.get("路线", "")
    if "—" in route_str:
        start_place, end_place = route_str.split("—", maxsplit=1)
    else:
        start_place, end_place = route_str, ""

    data = {
        "start_date": raw_data.get("托运日期", ""),
        "start_place": start_place.strip(),
        "end_place": end_place.strip(),
        "sender_name": raw_data.get("发车姓名", ""),
        "sender_phone": raw_data.get("发车电话", ""),
        "receiver_name": raw_data.get("收车姓名", ""),
        "receiver_phone": raw_data.get("收车电话", ""),
        "car_model": raw_data.get("车型", ""),
        "car_no": raw_data.get("车牌号", ""),
        "car_count": raw_data.get("数量台", "1"),
        "car_value": raw_data.get("车辆价值", ""),
        "insurance_desc": raw_data.get("保险额度", "车辆等值"),
        "trans_mode": raw_data.get("运输方式", "大板车"),
        "total_freight": str(board_fee),
        "pick_method": "自送网点",
        "deliver_method": "网点自提",
        "pay_type": raw_data.get("付款方式", "定金提付"),
        "deposit": str(deposit),
        "balance": str(balance),
        "remark": raw_data.get("备注", "")
    }
    return data

# -------------------------- Word填充、PDF转换【云端Linux版本】 --------------------------
def render_contract(template_path, context):
    tpl = DocxTemplate(template_path)
    tpl.render(context)
    tmp_docx = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
    tpl.save(tmp_docx)
    return tmp_docx.name

def docx_to_pdf(docx_file):
    tmp_pdf = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
    tmp_pdf.close()
    # Linux云端：直接使用 libreoffice
    cmd = [
        "libreoffice",
        "--headless",
        "--convert-to", "pdf",
        "--outdir", os.path.dirname(tmp_pdf.name),
        docx_file
    ]
    subprocess.run(cmd, capture_output=True, timeout=30)
    pdf_path = os.path.splitext(docx_file)[0] + ".pdf"
    if os.path.exists(pdf_path):
        with open(pdf_path, "rb") as f:
            pdf_data = f.read()
        os.unlink(pdf_path)
        return pdf_data
    return None

# -------------------------- 记录保存到csv --------------------------
def save_record(ctx):
    csv_file = "records.csv"
    row = pd.DataFrame([{
        "托运日期": ctx["start_date"],
        "始发地": ctx["start_place"],
        "目的地": ctx["end_place"],
        "发车姓名": ctx["sender_name"],
        "发车电话": ctx["sender_phone"],
        "收车姓名": ctx["receiver_name"],
        "收车电话": ctx["receiver_phone"],
        "车型": ctx["car_model"],
        "车牌号": ctx["car_no"],
        "数量(台)": ctx["car_count"],
        "车辆价值": ctx["car_value"],
        "保险额度": ctx["insurance_desc"],
        "板费": ctx["total_freight"],
        "定金": ctx["deposit"],
        "余款": ctx["balance"],
        "提车方式": ctx["pick_method"],
        "送车方式": ctx["deliver_method"],
        "运输类型": ctx["trans_mode"],
        "付款方式": ctx["pay_type"],
        "备注": ctx["remark"]
    }])
    if os.path.exists(csv_file):
        df = pd.read_csv(csv_file)
        df = pd.concat([df, row], ignore_index=True)
    else:
        df = row
    df.to_csv(csv_file, index=False, encoding="utf-8-sig")

# -------------------------- 页面入口 --------------------------
st.set_page_config(page_title="托运合同生成器", layout="wide")

# 先执行密码校验，不通过直接停止后面所有代码
if not check_password():
    st.stop()

# 登录成功，进入主页面
st.title("🚗 车辆托运合同生成器（手机网页版）")

raw_text = st.text_area(
    "粘贴托运原始信息",
    height=280,
    value="""路       线: 版纳—昆明
托运日期: 10.15
发车姓名:王家芬
发车电话:13988190736
收车姓名:施丽芬
收车电话:13638741235
车       型:北京现代索纳塔
车牌号:云KK5385
数量台:1
车辆价值:10万
保险额度:
运输方式:大板车
付款方式:定金提付
板        费: 700
定        金: 100
备注：含20万保险
"""
)

parsed_ctx = None
if st.button("🔍 解析文本并预览", type="primary"):
    parsed_ctx = parse_raw_text(raw_text)
    st.session_state["ctx"] = parsed_ctx

# 预览&编辑区域
if "ctx" in st.session_state:
    st.subheader("📝 核对/修改信息")
    ctx = st.session_state["ctx"]
    col1, col2 = st.columns(2)
    with col1:
        ctx["start_date"] = st.text_input("托运日期", value=ctx["start_date"])
        ctx["start_place"] = st.text_input("始发地", value=ctx["start_place"])
        ctx["end_place"] = st.text_input("目的地", value=ctx["end_place"])
        ctx["sender_name"] = st.text_input("发车姓名", value=ctx["sender_name"])
        ctx["sender_phone"] = st.text_input("发车电话", value=ctx["sender_phone"])
        ctx["receiver_name"] = st.text_input("收车姓名", value=ctx["receiver_name"])
        ctx["receiver_phone"] = st.text_input("收车电话", value=ctx["receiver_phone"])
        ctx["car_model"] = st.text_input("车型", value=ctx["car_model"])
        ctx["car_no"] = st.text_input("车牌号", value=ctx["car_no"])
    with col2:
        ctx["car_count"] = st.text_input("数量(台)", value=ctx["car_count"])
        ctx["car_value"] = st.text_input("车辆价值", value=ctx["car_value"])
        ctx["insurance_desc"] = st.text_input("保险额度", value=ctx["insurance_desc"])
        ctx["trans_mode"] = st.selectbox("运输类型", ["大板车", "小板车"], index=0 if ctx["trans_mode"]=="大板车" else 1)
        ctx["pick_method"] = st.selectbox("提车方式", ["自送网点", "代驾提车", "小板提车"], index=0)
        ctx["deliver_method"] = st.selectbox("送车方式", ["网点自提", "代驾送车", "小板送车"], index=0)
        ctx["total_freight"] = st.text_input("板费(运费合计)", value=ctx["total_freight"])
        ctx["deposit"] = st.text_input("定金", value=ctx["deposit"])
        ctx["balance"] = st.text_input("余款", value=ctx["balance"])
        ctx["pay_type"] = st.selectbox("付款方式", ["定金提付", "全款到付", "全款支付"], index=0)
    ctx["remark"] = st.text_area("备注", value=ctx["remark"])

    st.session_state["ctx"] = ctx

    if st.button("✅ 生成合同 Word + PDF"):
        template_file = "车辆运输协议模版.docx"
        if not os.path.exists(template_file):
            st.error("❌ 找不到模板：车辆运输协议模版.docx，请放到同目录！")
        else:
            with st.spinner("正在生成文档..."):
                docx_path = render_contract(template_file, ctx)
                # 读取docx二进制
                with open(docx_path, "rb") as f:
                    docx_bytes = f.read()
                pdf_bytes = docx_to_pdf(docx_path)
                # 保存记录
                save_record(ctx)

                # 文件名清洗，适配iPhone
                def clean_filename(name):
                    return re.sub(r'[\\/*?:"<>|]', "", name)
                raw_name = f"{ctx['start_date']}_{ctx['start_place']}-{ctx['end_place']}_{ctx['car_no']}"
                safe_name = clean_filename(raw_name)

                st.success("✅ 生成完成，记录已存入表格！")
                # 下载按钮
                st.download_button("📄 下载Word(.docx)", data=docx_bytes, file_name=f"{safe_name}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                if pdf_bytes:
                    st.download_button("📕 下载PDF(.pdf)", data=pdf_bytes, file_name=f"{safe_name}.pdf", mime="application/pdf")
                else:
                    st.warning("PDF转换失败，可单独下载Word文档")
                os.unlink(docx_path)

# 托运记录查看&导出
st.divider()
st.subheader("📋 全部托运记录")
csv_name = "records.csv"
if os.path.exists(csv_name):
    df_all = pd.read_csv(csv_name, encoding="utf-8-sig")
    st.dataframe(df_all, use_container_width=True)
    csv_download = df_all.to_csv(index=False, encoding="utf-8-sig")
    st.download_button("💾 导出全部记录Excel(csv)", data=csv_download, file_name="托运记录.csv")
else:
    st.info("暂无托运记录，生成合同后自动新增")
