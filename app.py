import streamlit as st
from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER
import io
import os
import plotly.graph_objects as go
import pandas as pd
from openpyxl import load_workbook

# ==========================================
# パス
# ==========================================
current_dir = os.path.dirname(__file__)
logo_path = os.path.join(current_dir, "logo.png")
template_path = os.path.join(current_dir, "template.pptx")
quotation_template_path = os.path.join(current_dir, "quotation_template.xlsx")

# ==========================================
# 初期設定
# ==========================================
def init_session_state():
    defaults = {
        "my_company": "株式会社SonicAI",
        "my_zip": "108-0075",
        "my_address": "東京都港区港南2-16-1 7F Spaces品川",
        "my_name": "小林賢正",
        "my_mail": "ken-kobayashi@sonicai.jp",
        "my_tel": "080-8044-3236",
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value

init_session_state()

# ==========================================
# UI設定
# ==========================================
st.set_page_config(page_title="SonicAI AI Tools", layout="wide")

# ==========================================
# サイドバー
# ==========================================
with st.sidebar:
    if os.path.exists(logo_path):
        st.image(logo_path)

    current_tab = st.radio("Menu", ["Tools", "Settings"])

    tool_choice = None
    if current_tab == "Tools":
        tool_choice = st.selectbox("Select Tool", [
            "📄 Report Generator",
            "🔍 Optical Calc",
            "💰 Quotation Generator"
        ])

    elif current_tab == "Settings":
        st.text_input("Company", key="my_company")
        st.text_input("Zip", key="my_zip")
        st.text_input("Address", key="my_address")
        st.text_input("Name", key="my_name")
        st.text_input("Mail", key="my_mail")
        st.text_input("TEL", key="my_tel")

# ==========================================
# 💰 見積ツール
# ==========================================
if tool_choice == "💰 Quotation Generator":

    st.title("💰 Quotation Generator")

    # 基本情報
    col1, col2 = st.columns(2)

    with col1:
        customer = st.text_input("客先")
        subject = st.text_input("件名")
        quote_no = st.text_input("見積番号")

    with col2:
        date = st.date_input("日付")
        payment = st.text_input("支払条件")

    st.write("---")

    # 商品マスタ
    product_master = {
        "AI検査装置A": {"price": 500000, "unit": "式"},
        "カメラ": {"price": 120000, "unit": "台"},
    }

    rows = st.number_input("行数", 1, 20, 5)

    items = []

    for i in range(int(rows)):
        cols = st.columns(6)

        product = cols[0].selectbox(
            f"商品{i}",
            ["自由入力"] + list(product_master.keys()),
            key=f"prod_{i}"
        )

        if product != "自由入力":
            name = product
            price = product_master[product]["price"]
            unit = product_master[product]["unit"]
        else:
            name = cols[1].text_input("品名", key=f"name_{i}")
            price = cols[4].number_input("単価", key=f"price_{i}", value=0)
            unit = cols[3].text_input("単位", key=f"unit_{i}")

        qty = cols[2].number_input("数量", key=f"qty_{i}", value=1)
        discount = cols[5].number_input("値引き", key=f"disc_{i}", value=0)

        amount = qty * price - discount

        items.append({
            "name": name,
            "qty": qty,
            "unit": unit,
            "price": price,
            "amount": amount
        })

    # 集計
    subtotal = sum(i["amount"] for i in items)
    total_discount = st.number_input("全体値引き", value=0)
    taxable = subtotal - total_discount
    tax = int(taxable * 0.1)
    total = taxable + tax

    st.metric("小計", f"{subtotal:,}")
    st.metric("税抜", f"{taxable:,}")
    st.metric("税込", f"{total:,}")

    # Excel生成
    if st.button("Excel出力"):

        wb = load_workbook(quotation_template_path)
        ws = wb["見積書"]

        ws["C6"] = customer
        ws["C7"] = subject
        ws["H6"] = quote_no

        start = 19
        for idx, item in enumerate(items):
            row = start + idx
            ws[f"C{row}"] = item["name"]
            ws[f"F{row}"] = item["qty"]
            ws[f"G{row}"] = item["unit"]
            ws[f"H{row}"] = item["price"]
            ws[f"I{row}"] = item["amount"]

        ws["I28"] = subtotal
        ws["I29"] = tax
        ws["I30"] = total

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)

        st.download_button(
            "ダウンロード",
            data=buf,
            file_name="quotation.xlsx"
        )

# ==========================================
# 他ツールはそのまま（省略）
# ==========================================