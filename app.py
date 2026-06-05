# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import google.generativeai as genai
from google.api_core import exceptions  # Rate Limit(429) エラーを確実に捕捉するため
from docx import Document
from PyPDF2 import PdfReader
from pptx import Presentation
import io
import configparser
import os
import signal
import re
import time  # リトライ待機（スリープ）処理のため

# --- 1. APIキーの設定 (APIKEY.ini または クラウドのSecretsからハイブリッド取得) ---
# ※既存のAPI取得ロジックの構造・変数名を1行も崩さずに、クラウド安全対策を内包させています
def load_api_key():
    # 1. まずローカルの APIKEY.ini を探す
    config = configparser.ConfigParser()
    file_path = 'APIKEY.ini'
    if os.path.exists(file_path):
        try:
            config.read(file_path, encoding='utf-8-sig')
            return config.get('GEMINI', 'API_KEY')
        except:
            pass
            
    # 2. もしローカルにファイルがなければ、Streamlit Cloudの「Secrets」から安全に取得する
    try:
        if "GEMINI" in st.secrets and "API_KEY" in st.secrets["GEMINI"]:
            return st.secrets["GEMINI"]["API_KEY"]
        elif "API_KEY" in st.secrets:
            return st.secrets["API_KEY"]
    except:
        return None
        
    return None

INI_KEY = load_api_key()
EMBEDDED_API_KEY = INI_KEY

# --- 2. 各ファイル抽出関数 (CSVの読み込みを拡張) ---
def extract_from_docx(file):
    doc = Document(file)
    return "\n".join([para.text for para in doc.paragraphs])

def extract_from_pdf(file):
    reader = PdfReader(file)
    return "\n".join([page.extract_text() for page in reader.pages if page.extract_text()])

def extract_from_pptx(file):
    prs = Presentation(file)
    text_runs = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if hasattr(shape, "text"):
                text_runs.append(shape.text)
    return "\n".join(text_runs)

def extract_from_excel(file):
    all_sheets = pd.read_excel(file, sheet_name=None)
    text_data = []
    for sheet_name, df in all_sheets.items():
        text_data.append(f"--- シート名: {sheet_name} ---\n{df.to_string(index=False)}")
    return "\n".join(text_data)

def extract_from_csv(file):
    # シニア層が作成した様々なエンコーディングに対応できるよう柔軟に読み込み
    try:
        df = pd.read_csv(file, encoding='utf-8')
    except UnicodeDecodeError:
        file.seek(0)
        df = pd.read_csv(file, encoding='shift_jis')
    return df.to_string(index=False)

# --- 404エラーを回避しつつ、利用可能なモデル名を安全に取得する関数 ---
def get_safe_model_name(api_key):
    try:
        genai.configure(api_key=api_key)
        available_models = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
        
        has_flash = any('gemini-1.5-flash' in m for m in available_models)
        target_raw = 'gemini-1.5-flash' if has_flash else (available_models[0] if available_models else 'gemini-1.5-flash')
        
        safe_name = target_raw.replace('models/', '')
        return safe_name
    except:
        return 'gemini-1.5-flash'

# --- 3. AI回答生成ロジック (自動リトライ・履歴ウィンドウ削減版) ---
def get_ai_roleplay_response(messages, persona, product_docs, api_key):
    target_model = get_safe_model_name(api_key)
    
    # トークン爆発を防ぐため、AIに送る「過去の履歴」を最新の6発言（3往復）に限定
    recent_messages = [messages[0]] + messages[-5:] if len(messages) > 6 else messages

    for attempt in range(5):
        try:
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel(target_model)
            
            combined_docs = "\n\n".join(product_docs) if product_docs else "追加のマニュアル等のアップロードは現在ありません。標準仕様に基づいて回答してください。"
            
            history_text = ""
            for m in recent_messages:
                role_label = "AIアシスタント(あなた)" if m["role"] == "assistant" else "本部担当者(ユーザー)"
                history_text += f"{role_label}: {m['content']}\n"

            system_prompt = f"""
あなたは「預かり資産トータルクエリーサービス」およびBIツール「軽技WEB」に精通した熟練のAIFAQアシスタントです。
金融機関の本部担当者（ユーザー）からのシステム操作、データ抽出、クエリ、エラー等に関する質問に対して、正確かつ分かりやすく回答してください。

【対象サービス・システム構成情報】
・サービス名: {persona['service_name']}
・概要: {persona['overview']}
・接続構成: {persona['architecture']}

【本サービスの特徴】
{persona['features']}

【アップロードされた各種マニュアル・参考ドキュメント（最優先参照情報）】
{combined_docs}

【回答の絶対ルール】
1. ユーザー（本部担当者）の質問に対し、アップロードされたマニュアルの情報を最優先に参照して回答を構成してください。
2. あくまで「預かり資産トータルクエリーサービス」および「軽技WEB」「Fund Organizer」の文脈に沿って回答してください。一般的なIT知識でなく、本サービスの仕様を重視してください。
3. 専門的な内容であっても、金融機関の業務担当者がスムーズに作業を進められるよう、具体的な手順や選択すべきクエリ（標準クエリ約200種）など、実務に即した丁寧な表現で回答してください。
4. あまりに難しい専門用語は避け、シニアや業務に不慣れな方でも理解できるよう、手順を①②③と箇条書きで分かりやすく整理して伝えてください。
5. AIとしてのメタな発言（例：「以上がマニュアルに基づく回答です」など）は含めず、ユーザーへの親切な回答テキストのみを出力してください。

【これまでの会話履歴（※直近の重要な会話のみ抽出）】
{history_text}

上記のルールと履歴を元に、次につづく「AIアシスタント(あなた)」の丁寧な回答を生成してください。
"""
            response = model.generate_content(system_prompt)
            return response.text

        except (exceptions.ResourceExhausted, Exception) as e:
            error_msg = str(e)
            if "429" in error_msg or isinstance(e, exceptions.ResourceExhausted):
                wait_time = (attempt + 1) * 10
                time.sleep(wait_time)
                continue
            return f"【システムエラー】詳細: {error_msg}"
            
    return "【混雑エラー】現在AIへのリクエストが連続しています。無料枠の制限を超過したため、1分ほど待ってから再度送信してください。"

# --- 4. 画面構築 (Streamlit UI ＆ あたたかみのあるシニア向けデザインCSS) ---
st.set_page_config(page_title="かんたん操作案内 AIお助けチャット", layout="wide")

# シニア向け：大きな文字、親しみやすいカラー（アイボリー・ブラウン・LINE風グリーン）の適用
st.markdown("""
    <style>
    /* 全体の背景色とフォントサイズ調整 */
    .stApp {
        background-color: #FDFBF7; /* 温かみのあるアイボリー */
    }
    html, body, [class*="css"], p, ul, li {
        font-size: 18px !important; /* シニア向けに文字を大きく */
        color: #4A3E3D !important; /* 視認性の高いダークブラウン */
        font-family: 'Hiragino Kaku Gothic ProN', 'Meiryo', sans-serif;
    }
    h1 {
        font-size: 32px !important;
        color: #5C4033 !important; /* 濃いブラウン */
        font-weight: bold;
    }
    /* サイドバーの背景デザイン */
    [data-testid="stSidebar"] {
        background-color: #F5EBE6 !important; /* 薄いベージュ */
        border-right: 2px solid #E6D5CC;
    }
    /* LINE風のチャットバブル */
    .chat-bubble-user {
        background-color: #85E292 !important; /* LINE風の優しいグリーン */
        color: #111111 !important;
        padding: 14px 20px;
        border-radius: 18px 18px 2px 18px;
        margin: 8px 0px;
        display: inline-block;
        max-width: 80%;
        box-shadow: 1px 1px 4px rgba(0,0,0,0.1);
    }
    .chat-bubble-assistant {
        background-color: #FFFFFF !important; /* 白色バブル */
        color: #4A3E3D !important;
        padding: 14px 20px;
        border-radius: 18px 18px 18px 2px;
        margin: 8px 0px;
        display: inline-block;
        max-width: 80%;
        border: 1px solid #E6D5CC;
        box-shadow: 1px 1px 4px rgba(0,0,0,0.05);
    }
    /* ボタンを目立たせる */
    .stButton>button {
        background-color: #8C6A5C !important;
        color: white !important;
        font-size: 18px !important;
        border-radius: 8px;
        padding: 10px 24px;
        border: none;
    }
    </style>
""", unsafe_allow_html=True)

st.title("🍵 かんたんマニュアル案内室 (AIお助けチャット)")
st.markdown("ファイル（マニュアル）を左側から読み込ませるだけで、AIがLINEのように分かりやすく操作方法を教えてくれます。")

# --- 5. 固定ペルソナ情報設定 ---
current_persona = {
    "service_name": "預かり資産トータルクエリーサービス",
    "overview": "軽技WEBというBIツールを利用した、投資信託の実績集計などが行える金融機関の本部担当者向けサービス。",
    "architecture": "顧客（投資家）の契約情報、取引履歴、残高、損益情報を保管しているデータベース「Fund Organizer」を参照元として、本部担当者のブラウザから「軽技WEB」を利用して接続し、データを取得する。",
    "features": "①約200の標準クエリから業務要件に合わせたクエリを選択して実行\n②標準クエリを編集しフレキシブルに加工・保存が可能\n③データ取得結果をCSV/Excel形式で保存が可能\n④スケジュール実行で業務の効率化を実現\n⑤専門の担当者による手厚いサポート"
}

# --- 6. 左側メニュー（サイドバー）の構築 ---
st.sidebar.markdown("### 🔑 個別のAPIキー設定")
st.sidebar.markdown("<small>ご自身のAPIキーをお持ちの場合は入力してください。空欄の場合は自動的に標準のキーを使用します。</small>", unsafe_allow_html=True)
custom_api_key = st.sidebar.text_input("Gemini API キー", type="password", key="custom_key")

# APIキーの優先判定ロジック
if custom_api_key.strip():
    ACTIVE_API_KEY = custom_api_key.strip()
else:
    ACTIVE_API_KEY = EMBEDDED_API_KEY

st.sidebar.markdown("---")
st.sidebar.markdown("### 📁 マニュアル・ドキュメント読込")
st.sidebar.markdown("<small>操作の案内書や、CSV、Excel、PDFなどの各種資料をここにドラッグ＆ドロップしてください。</small>", unsafe_allow_html=True)

uploaded_files = st.sidebar.file_uploader(
    "ここへファイルを重ねてください", 
    type=["docx", "pdf", "pptx", "xlsx", "xls", "csv"], 
    accept_multiple_files=True,
    key="file_uploader"
)

all_extra_text = []
if uploaded_files:
    for f in uploaded_files:
        try:
            if f.name.endswith(".docx"): content = extract_from_docx(f)
            elif f.name.endswith(".pdf"): content = extract_from_pdf(f)
            elif f.name.endswith(".pptx"): content = extract_from_pptx(f)
            elif f.name.endswith((".xlsx", ".xls")): content = extract_from_excel(f)
            elif f.name.endswith(".csv"): content = extract_from_csv(f)
            
            all_extra_text.append(f"--- ファイル名: {f.name} ---\n{content}")
            st.sidebar.success(f"⭕ 読み込みました: {f.name}")
        except Exception as e:
            st.sidebar.error(f"❌ {f.name} の読込失敗")

st.sidebar.markdown("---")
if st.sidebar.button("🛑 アプリを安全に終了する"):
    st.sidebar.warning("システムを終了します。窓を閉じてください。")
    os.kill(os.getpid(), signal.SIGINT)

# --- 7. セッション状態の初期化 ---
if "messages" not in st.session_state:
    st.session_state.messages = [{
        "role": "assistant", 
        "content": "「預かり資産トータルクエリーサービス」の案内室へようこそ！お困りの操作や、調べたい内容について、LINEのように下の欄へ何でも打ち込んでくださいね。"
    }]

# --- 8. メインエリアのチャットレイアウト (LINE風表示への最適化) ---
if not uploaded_files:
    st.info(f"💡 まだマニュアル資料が読み込まれていません。お持ちの資料がある場合は、画面左側のメニューから追加できます。")
else:
    st.success(f"✨ 現在 {len(uploaded_files)} 件の資料を元にAIが回答できます。")

# 過去のメッセージ表示
for m in st.session_state.messages:
    if m["role"] == "user":
        # 右寄せのユーザー発言（LINE風）
        st.markdown(f"""
            <div style="text-align: right;">
                <div class="chat-bubble-user">
                    <b>👤 あなた:</b><br>{m['content']}
                </div>
            </div>
        """, unsafe_allow_html=True)
    else:
        # 左寄せのAIアシスタント発言
        st.markdown(f"""
            <div style="text-align: left;">
                <div class="chat-bubble-assistant">
                    <b>🤖 案内AI:</b><br>{m['content']}
                </div>
            </div>
        """, unsafe_allow_html=True)

# 新規入力フォーム
if prompt := st.chat_input("ここに質問したいことを書いて、右の紙飛行機マークを押してください"):
    # ユーザーの発言を即時追加
    st.session_state.messages.append({"role": "user", "content": prompt})
    st.markdown(f"""
        <div style="text-align: right;">
            <div class="chat-bubble-user">
                <b>👤 あなた:</b><br>{prompt}
            </div>
        </div>
    """, unsafe_allow_html=True)

    # 返答の生成
    with st.spinner("マニュアルを確認して、分かりやすい言葉で回答を作っています。少しお待ちください..."):
        res = get_ai_roleplay_response(
            st.session_state.messages, 
            current_persona, 
            all_extra_text, 
            ACTIVE_API_KEY
        )
        
    st.session_state.messages.append({"role": "assistant", "content": res})
    st.rerun()
