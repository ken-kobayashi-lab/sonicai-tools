import streamlit as st
from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER
import io
import os
import plotly.graph_objects as go

# ==========================================
# 1. 共通ロジック
# ==========================================
current_dir = os.path.dirname(__file__)
logo_path = os.path.join(current_dir, "logo.png")
template_path = os.path.join(current_dir, "template.pptx")


def init_session_state():
    """初期値を session_state に設定"""
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


def fill_placeholder(slide, ph_type_idx, content, is_image=False, img_file=None):
    """
    スライド上のプレースホルダを上から下、左から右で並べて埋める
    is_image=True の場合は画像プレースホルダに画像を挿入
    """
    placeholders = sorted(
        [sh for sh in slide.shapes if sh.is_placeholder],
        key=lambda x: (x.top, x.left)
    )

    text_phs = [
        sh for sh in placeholders
        if sh.placeholder_format.type != PP_PLACEHOLDER.PICTURE
    ]
    img_phs = [
        sh for sh in placeholders
        if sh.placeholder_format.type == PP_PLACEHOLDER.PICTURE
    ]

    try:
        if is_image:
            if img_file and len(img_phs) > ph_type_idx:
                ph = img_phs[ph_type_idx]
                ph.insert_picture(img_file)
                return True
            return False
        else:
            if len(text_phs) > ph_type_idx:
                ph = text_phs[ph_type_idx]
                ph.text = str(content)
                return True
            return False
    except Exception:
        return False


def move_slide(prs, old_index, new_index):
    """スライド順を移動"""
    xml_slides = prs.slides._sldIdLst
    slide_id = xml_slides[old_index]
    xml_slides.remove(slide_id)
    xml_slides.insert(new_index, slide_id)


def validate_data(data):
    required_keys = [
        "company_name",
        "project_name",
        "my_company",
        "my_zip",
        "my_address",
        "my_name",
        "my_mail",
        "my_tel",
        "cond",
        "setup_img",
        "res",
        "uploaded_images",
        "image_comments",
        "summ",
    ]
    missing = [k for k in required_keys if k not in data]
    if missing:
        raise ValueError(f"必要なデータが不足しています: {missing}")


def generate_pptx(template_path_arg, data):
    validate_data(data)

    if not os.path.exists(template_path_arg):
        raise FileNotFoundError(f"テンプレートファイルが見つかりません: {template_path_arg}")

    prs = Presentation(template_path_arg)

    # 0枚目: 表紙
    fill_placeholder(prs.slides[0], 0, data["company_name"])
    fill_placeholder(prs.slides[0], 1, data["project_name"])

    creator_info = (
        f"{data['my_company']}\n"
        f"〒{data['my_zip']}\n"
        f"{data['my_address']}\n"
        f"{data['my_name']}\n"
        f"mail：{data['my_mail']}\n"
        f"TEL：{data['my_tel']}"
    )
    fill_placeholder(prs.slides[0], 2, creator_info)

    # 1枚目: 条件 / 構成図
    fill_placeholder(prs.slides[1], 1, data["cond"])
    if data["setup_img"] is not None:
        fill_placeholder(prs.slides[1], 0, "", is_image=True, img_file=data["setup_img"])

    # 2枚目: 結果
    fill_placeholder(prs.slides[2], 1, data["res"])

    # 3枚目以降: 画像一覧
    images = data["uploaded_images"]
    comments = data["image_comments"]

    if len(prs.slides) < 4:
        raise ValueError("template.pptx の4枚目（画像貼り付け用スライド）が不足しています。")

    img_layout = prs.slides[3].slide_layout

    for i in range(0, len(images), 2):
        slide = prs.slides[3] if i == 0 else prs.slides.add_slide(img_layout)

        if i > 0:
            move_slide(prs, len(prs.slides) - 1, 3 + (i // 2))

        # 左側
        fill_placeholder(slide, 0, "", is_image=True, img_file=images[i])
        if len(comments) > i:
            fill_placeholder(slide, 1, comments[i])

        # 右側
        if i + 1 < len(images):
            fill_placeholder(slide, 1, "", is_image=True, img_file=images[i + 1])
            if len(comments) > i + 1:
                fill_placeholder(slide, 2, comments[i + 1])
        else:
            fill_placeholder(slide, 2, "")

    # 最終ページ: まとめ
    fill_placeholder(prs.slides[-1], 1, data["summ"])

    ppt_io = io.BytesIO()
    prs.save(ppt_io)
    ppt_io.seek(0)
    return ppt_io


# ==========================================
# 2. 初期設定
# ==========================================
init_session_state()

st.set_page_config(page_title="SonicAI AI Tools", layout="wide")

st.markdown("""
<style>
    html, body, [class*="css"] {
        font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif;
    }
    .logo-container {
        text-align: center;
        padding-bottom: 20px;
    }
    div.stButton > button {
        background-color: #ff6600;
        color: white;
        height: 3em;
        width: 100%;
        font-weight: bold;
        border-radius: 10px;
        border: none;
        transition: 0.3s;
    }
    div.stButton > button:hover {
        background-color: #e65c00;
        box-shadow: 0 4px 15px rgba(255,102,0,0.3);
    }
    div.stDownloadButton > button {
        background-color: #28a745;
        color: white;
        width: 100%;
        height: 3em;
        font-weight: bold;
        border-radius: 10px;
        border: none;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background-color: #f1f3f6;
        padding: 5px;
        border-radius: 12px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 40px;
        border-radius: 8px;
        background-color: transparent;
        color: #555;
        font-weight: 500;
        border: none !important;
    }
    .stTabs [aria-selected="true"] {
        background-color: white !important;
        box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        color: #ff6600 !important;
    }
    .camera-info-card {
        background-color: #1e2630;
        color: #ffffff;
        padding: 20px;
        border-radius: 12px;
        border-left: 6px solid #007bff;
        margin: 10px 0;
        line-height: 1.6;
    }
    .camera-info-label {
        color: #8899ac;
        font-size: 0.85em;
        font-weight: bold;
        text-transform: uppercase;
    }
    .camera-info-value {
        color: #ffffff;
        font-size: 1.1em;
        margin-bottom: 8px;
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 3. サイドバー
# ==========================================
with st.sidebar:
    st.markdown('<div class="logo-container">', unsafe_allow_html=True)
    if os.path.exists(logo_path):
        st.image(logo_path, use_container_width=True)
    else:
        st.title("SonicAI")
    st.markdown('</div>', unsafe_allow_html=True)

    st.title("SonicAI Platform")
    current_tab = st.radio(
        "Navigation",
        ["Tools", "Settings", "Manual"],
        label_visibility="collapsed"
    )

    tool_choice = None
    if current_tab == "Tools":
        tool_choice = st.selectbox("Select Tool", ["📄 Report Generator", "🔍 Optical Calc"])

    elif current_tab == "Settings":
        st.subheader("👤 Creator Info")
        st.text_input("Company", key="my_company")
        st.text_input("Zip Code", key="my_zip")
        st.text_area("Address", key="my_address", height=80)
        st.text_input("Name", key="my_name")
        st.text_input("Email", key="my_mail")
        st.text_input("TEL", key="my_tel")
        st.caption("Settings are used for the report cover.")

    st.write("---")
    st.caption("SonicAI Inc. v1.7")

# ==========================================
# 4. メインコンテンツ
# ==========================================
if current_tab == "Manual":
    st.markdown("<h1 style='color: #6c757d;'>📖 User Manual</h1>", unsafe_allow_html=True)

    col_m1, col_m2 = st.columns(2)

    with col_m1:
        st.subheader("🔍 視野計算ツールの使い方")
        st.markdown("""
1. **カメラ・レンズ選択**: 使用するカメラ型式とレンズの焦点距離を選びます。  
2. **モード選択**: WD（距離）から視野を出すか、必要な視野からWDを出すかを選びます。  
3. **グラフの活用**: オレンジ色の線が選択中のレンズ特性です。赤い点線が現在の計算値を示します。  
        """)

    with col_m2:
        st.subheader("📄 レポート作成ツールの使い方")
        st.markdown("""
1. **基本情報入力**: 案件名や検証条件（照明設定など）を入力します。  
2. **画像アップロード**: 検証に使用した画像をドラッグ＆ドロップします。  
3. **生成**: 「パワーポイントを生成」ボタンを押し、完成したファイルをダウンロードします。  
        """)

    st.divider()
    st.warning("⚠️ **注意**: 本ツールはオフライン動作を前提としています。ライブラリの更新時以外はネット接続不要です。")

elif current_tab == "Settings":
    st.markdown("<h1 style='color: #6c757d;'>👤 Settings Preview</h1>", unsafe_allow_html=True)
    st.write("サイドバーで入力した情報は、以下の形式でレポートの表紙に反映されます。")
    st.code(
        f"{st.session_state.my_company}\n"
        f"〒{st.session_state.my_zip}\n"
        f"{st.session_state.my_address}\n"
        f"{st.session_state.my_name}\n"
        f"mail：{st.session_state.my_mail}\n"
        f"TEL：{st.session_state.my_tel}"
    )

elif tool_choice == "📄 Report Generator":
    st.markdown("<h1 style='color: #ff6600;'>📄 Report Generator</h1>", unsafe_allow_html=True)

    c_name = st.text_input("客先会社名", value="〇〇株式会社 御中")
    p_name = st.text_input("案件名", value="AI外観検査 導入可否検証")

    st.write("---")
    col1, col2 = st.columns(2)

    with col1:
        cond = st.text_area("検証条件", value="カメラ：\nレンズ：\n照明：", height=150)
        setup_img = st.file_uploader("構成図(任意)", type=["png", "jpg", "jpeg"])

    with col2:
        res = st.text_area("検証結果", value="良好な結果を確認しました。", height=150)
        summ = st.text_area("ご提案まとめ", value="本構成での導入を推奨します。", height=150)

    st.write("---")

    uploaded_images = st.file_uploader(
        "画像をアップロード",
        accept_multiple_files=True,
        type=["png", "jpg", "jpeg"]
    )

    image_comments = []
    if uploaded_images:
        cols = st.columns(2)
        for idx, img in enumerate(uploaded_images):
            with cols[idx % 2]:
                st.image(img, use_container_width=True)
                comment = st.text_input(
                    f"Comment {idx + 1}",
                    value=f"画像({idx + 1}): {img.name}",
                    key=f"c_{idx}"
                )
                image_comments.append(comment)

    if st.button("🚀 パワーポイントを生成", use_container_width=True):
        try:
            if not uploaded_images:
                st.error("画像を1枚以上選択してください。")
            else:
                data = {
                    "company_name": c_name,
                    "project_name": p_name,
                    "my_company": st.session_state.my_company,
                    "my_zip": st.session_state.my_zip,
                    "my_address": st.session_state.my_address,
                    "my_name": st.session_state.my_name,
                    "my_mail": st.session_state.my_mail,
                    "my_tel": st.session_state.my_tel,
                    "cond": cond,
                    "setup_img": setup_img,
                    "res": res,
                    "uploaded_images": uploaded_images,
                    "image_comments": image_comments,
                    "summ": summ,
                }

                final_ppt = generate_pptx(template_path, data)

                st.success("パワーポイントを生成しました。")
                st.balloons()
                st.download_button(
                    label="📥 ダウンロード",
                    data=final_ppt,
                    file_name=f"{p_name}.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width=True
                )
        except Exception as e:
            st.error(f"エラーが発生しました: {e}")

else:
    st.markdown("<h1 style='color: #007bff;'>🔍 Optical Calc Tool</h1>", unsafe_allow_html=True)

    camera_data = {
        "40万画素 (CS-41-C/B)": {"model": "CS-41-C/B", "sensor": "1/2.9", "w": 720, "h": 540, "px": 6.9},
        "160万画素 (CS-160-C/B)": {"model": "CS-160-C/B", "sensor": "1/2.9", "w": 1440, "h": 1080, "px": 3.45},
        "500万画素 (CS-500-C/B)": {"model": "CS-500-C/B", "sensor": "1/2", "w": 2600, "h": 2160, "px": 2.5}
    }
    focal_lengths = [8, 12, 16, 25, 35, 50, 75]

    c1, c2 = st.columns([1, 1])

    with c1:
        cam_label = st.selectbox("Select Camera", list(camera_data.keys()))
        cam = camera_data[cam_label]
        st.markdown(f"""
        <div class="camera-info-card">
            <div class="camera-info-label">Model</div>
            <div class="camera-info-value">{cam['model']}</div>
            <div class="camera-info-label">Sensor Size</div>
            <div class="camera-info-value">{cam['sensor']}</div>
            <div class="camera-info-label">Pixel Size</div>
            <div class="camera-info-value">{cam['px']}μm × {cam['px']}μm</div>
            <div class="camera-info-label">Resolution</div>
            <div class="camera-info-value">{cam['w']} × {cam['h']} px</div>
        </div>
        """, unsafe_allow_html=True)

    with c2:
        focal_len = st.selectbox("Focal Length (mm)", focal_lengths)
        f_value = st.number_input("F-Value", value=4.0, step=0.1)
        calc_mode = st.radio("Mode", ["WD -> FOV", "FOV -> WD"])

    sensor_w = cam["w"] * cam["px"] / 1000

    if calc_mode == "WD -> FOV":
        wd_val = st.number_input("Working Distance (mm)", value=300.0, step=10.0)
        fov_w = wd_val * sensor_w / focal_len
    else:
        fov_w = st.number_input("Target FOV (Width mm)", value=100.0, step=10.0)
        wd_val = focal_len * fov_w / sensor_w

    fig = go.Figure()
    wd_range = list(range(50, 2001, 10))

    for f in focal_lengths:
        f_line = [w * sensor_w / f for w in wd_range]
        fig.add_trace(
            go.Scatter(
                x=f_line,
                y=wd_range,
                name=f"{f}mm",
                line=dict(width=4 if f == focal_len else 1.5)
            )
        )

    fig.add_shape(
        type="line",
        x0=0, y0=wd_val, x1=fov_w, y1=wd_val,
        line=dict(color="Red", dash="dash")
    )
    fig.add_shape(
        type="line",
        x0=fov_w, y0=0, x1=fov_w, y1=wd_val,
        line=dict(color="Red", dash="dash")
    )

    fig.update_layout(
        xaxis_title="FOV Width (mm)",
        yaxis_title="WD (mm)",
        height=500,
        margin=dict(t=20, b=20)
    )

    st.plotly_chart(fig, use_container_width=True)

    m1, m2, m3 = st.columns(3)
    m1.metric("Current WD", f"{wd_val:.1f} mm")
    m2.metric("FOV Width", f"{fov_w:.1f} mm")
    m3.metric("Resolution", f"{(fov_w / cam['w'] * 1000):.2f} um/px")

    st.write("---")

    dof = (2 * f_value * (cam["px"] * 2 / 1000) * (wd_val ** 2)) / (focal_len ** 2)
    st.info(f"DOF: ±{dof / 2:.2f} mm (Total {dof:.2f} mm)")
