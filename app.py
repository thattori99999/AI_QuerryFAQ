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

# --- 2. 各ファイル抽出関数 ---
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

# 追加：CSVファイルの抽出関数（Shift-JISとUTF-8の両方に対応するセーフモード）
def extract_from_csv(file):
    try:
        # まずはUTF-8で試みる
        df = pd.read_csv(file, encoding='utf-8')
    except Exception:
        try:
            # 失敗した場合はShift-JISで再試行
            file.seek(0)
            df = pd.read_csv(file, encoding='shift_jis')
        except Exception as e:
            return f"CSVファイルの読み込みに失敗しました（文字コードエラー）: {str(e)}"
    return df.to_string(index=False)

def get_text_from_file(file):
    """ファイル形式に応じてテキストを抽出する共通関数"""
    if file.name.endswith(".docx"): return extract_from_docx(file)
    elif file.name.endswith(".pdf"): return extract_from_pdf(file)
    elif file.name.endswith(".pptx"): return extract_from_pptx(file)
    elif file.name.endswith((".xlsx", ".xls")): return extract_from_excel(file)
    elif file.name.endswith(".csv"): return extract_from_csv(file)
    return ""

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
def get_ai_roleplay_response(messages, persona, product_docs, api_key, format_sample=None):
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

            # フォーマットサンプルの指示ブロックを作成
            format_instruction = ""
            if format_sample:
                format_instruction = f"""
【最優先：出力フォーマットのルール】
ユーザーから出力フォーマットのサンプルが指定されています。
回答を作成する際は、以下のサンプル内容、項目名、構成、トンマナを「必ず」参考にして、同様の出力形式で結果を出力してください：
--- 出力フォーマットサンプル 開始 ---
{format_sample}
--- 出力フォーマットサンプル 終了 ---
"""

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

{format_instruction}

【回答の絶対ルール】
1. ユーザー（本部担当者）の質問に対し、アップロードされたマニュアルの情報を最優先に参照して回答を構成してください。
2. あくまで「預かり資産トータルクエリーサービス」および「軽技WEB」「Fund Organizer」の文脈に沿って回答してください。一般的なIT知識でなく、本サービスの仕様を重視してください。
3. 専門的な内容であっても、金融機関の業務担当者がスムーズに作業を進められるよう、具体的な手順や選択すべきクエリ（標準クエリ約200種）など、実務に即した丁寧な表現で回答してください。
4. アップロードされた情報だけで判断がつかない不確実な事項については、知ったかぶりをせず、「専門のサポート窓口」へ案内するなどの対応を含めてください。
5. AIとしてのメタな発言（例：「以上がマニュアルに基づく回答です」など）は含めず、ユーザーへの親切な回答テキストのみを出力してください。
6. シニア層や不慣れな担当者でも一目で操作がわかるよう、手順は「①、②、③」などの箇条書きを使い、専門用語には分かりやすい言葉で補足を添えてください。

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

# --- 4. 画面構築 (Streamlit UI) ---
# ページ設定
st.set_page_config(
    page_title="預かり資産トータルクエリーサービス らくらくAIFAQ", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# セッション状態の初期化
if "messages" not in st.session_state:
    st.session_state.messages = [
        {
            "role": "assistant", 
            "content": "「預かり資産トータルクエリーサービス」らくらく相談窓口へようこそ！\n\n画面の左側で「マニュアル」を読み込ませることで、あなた専用 of ガイドブックになります。\n操作方法や画面の使い方のほか、「こういう風にデータを出したい」といったご質問にいつでもお答えしますよ。お気軽に何でも聞いてくださいね。"
        }
    ]

if "format_file_key" not in st.session_state:
    st.session_state.format_file_key = 0

if "format_sample" not in st.session_state:
    st.session_state.format_sample = ""

# --- 5. シニア向け・あたたかみのある緑ベースのカスタムCSS ---
# ラジオボタンで文字サイズを選択できるようにする
st.sidebar.markdown("### 🔤 画面の文字の大きさ")
font_size_choice = st.sidebar.radio(
    "文字のサイズを選んでください：",
    ["ふつう", "大きく", "とても大きく"],
    index=1, # デフォルトを「大きく」にして見やすさを最優先に
    horizontal=True,
    label_visibility="collapsed"
)

# 選択に応じたCSSフォントサイズの設定
if font_size_choice == "ふつう":
    text_size = "16px"
    title_size = "26px"
    bubble_padding = "12px"
elif font_size_choice == "大きく":
    text_size = "20px"
    title_size = "30px"
    bubble_padding = "16px"
else: # とても大きく
    text_size = "24px"
    title_size = "36px"
    bubble_padding = "20px"

# カスタムスタイルの注入 (緑ベースのやさしい3色デザイン)
st.markdown(f"""
<style>
    /* 全体フォント設定 */
    html, body, [class*="css"], .stMarkdown p, .stButton button {{
        font-size: {text_size} !important;
        line-height: 1.6 !important;
        color: #2D312E !important; /* 目が疲れないソフトな黒 */
    }}
    
    /* アプリ全体の背景色 */
    .stApp {{
        background-color: #F4F8F4 !important;
    }}
    
    /* サイドバーの背景色と境界線 */
    section[data-testid="stSidebar"] {{
        background-color: #E8F0E8 !important;
        border-right: 2px solid #D2E2D2 !important;
    }}
    
    /* ヘッダータイトル */
    .app-title {{
        font-size: {title_size} !important;
        color: #1B4332 !important;
        font-weight: bold;
        margin-bottom: 5px;
        display: flex;
        align-items: center;
        gap: 10px;
    }}
    
    /* サブ説明文 */
    .app-subtitle {{
        font-size: calc({text_size} - 2px) !important;
        color: #406040 !important;
        margin-bottom: 20px;
        background-color: #EAF4EA;
        padding: 10px 15px;
        border-radius: 8px;
        border-left: 5px solid #2D6A4F;
    }}

    /* LINE風チャット吹き出しのカスタマイズ */
    div[data-testid="stChatMessage"] {{
        border-radius: 20px !important;
        padding: {bubble_padding} !important;
        margin-bottom: 15px !important;
        box-shadow: 0 2px 5px rgba(0,0,0,0.05) !important;
    }}
    
    /* ユーザー発言（右側・黄緑色の吹き出し） */
    div[data-testid="stChatMessage"]:has(span[data-testid="chat-avatar-user"]) {{
        background-color: #A7E0A6 !important; /* LINE風の優しい緑 */
        margin-left: 15% !important;
        border: 1px solid #90C88F !important;
    }}
    
    /* アシスタント発言（左側・白色の吹き出し） */
    div[data-testid="stChatMessage"]:has(span[data-testid="chat-avatar-assistant"]) {{
        background-color: #FFFFFF !important;
        margin-right: 15% !important;
        border: 1px solid #E0ECE0 !important;
    }}
    
    /* ボタンのスタイルカスタマイズ */
    .stButton button {{
        background-color: #2D6A4F !important;
        color: #FFFFFF !important;
        border-radius: 12px !important;
        padding: 10px 24px !important;
        font-weight: bold !important;
        border: none !important;
        box-shadow: 0 3px 6px rgba(0,0,0,0.1) !important;
        transition: all 0.2s ease;
    }}
    .stButton button:hover {{
        background-color: #1B4332 !important;
        transform: translateY(-1px);
    }}
    
    /* 入力エリアの強調表示 */
    div[data-testid="stChatInput"] textarea {{
        font-size: {text_size} !important;
        border: 2px solid #2D6A4F !important;
        border-radius: 12px !important;
        background-color: #FFFFFF !important;
    }}

    /* サイドバーの見出し調整 */
    .sidebar-heading {{
        color: #1B4332 !important;
        font-weight: bold !important;
        font-size: calc({text_size} + 2px) !important;
        margin-top: 15px !important;
        margin-bottom: 5px !important;
        border-bottom: 2px solid #2D6A4F;
        padding-bottom: 3px;
    }}
    
    /* 読み込み完了バッジ風表示 */
    .load-success {{
        background-color: #D8F3DC !important;
        color: #1B4332 !important;
        padding: 5px 10px;
        border-radius: 5px;
        font-size: calc({text_size} - 2px);
        margin-bottom: 5px;
        border: 1px solid #B7E4C7;
    }}
</style>
""", unsafe_allow_html=True)

# メインエリアヘッダー
st.markdown('<div class="app-title">🍵 らくらく操作ガイド AIチャット窓口</div>', unsafe_allow_html=True)
st.markdown('<div class="app-subtitle">投資信託の集計システムや「軽技WEB」に関する疑問を、お手元のマニュアルを参考にしてAIがやさしく解決します。</div>', unsafe_allow_html=True)

# --- 6. 固定ペルソナ情報設定 ---
current_persona = {
    "service_name": "預かり資産トータルクエリーサービス",
    "overview": "軽技WEBというBIツールを利用した、投資信託の実績集計などが行える金融機関の本部担当者向けサービス。",
    "architecture": "顧客（投資家）の契約情報、取引履歴、残高、損益情報を保管しているデータベース「Fund Organizer」を参照元として、本部担当者のブラウザから「軽技WEB」を利用して接続し、データを取得する。",
    "features": "①約200の標準クエリから業務要件に合わせたクエリを選択して実行\n②標準クエリを編集しフレキシブルに加工・保存が可能\n③データ取得結果をCSV/Excel形式で保存が可能\n④スケジュール実行で業務の効率化を実現\n⑤専門の担当者による手厚いサポート"
}

# --- 7. サイドバー機能 (情報入力メニュー) ---
st.sidebar.markdown('<div class="sidebar-heading">🔑 APIキーの設定</div>', unsafe_allow_html=True)
st.sidebar.markdown("<small style='color:#555;'>※未入力の場合は自動的にシステムのデフォルトキーを使用します。</small>", unsafe_allow_html=True)
user_api_key = st.sidebar.text_input(
    "Gemini APIキー",
    type="password",
    placeholder="AIの鍵をお持ちなら入力してください",
    help="入力されたキーを最優先で使用します。空欄の場合は標準の共有キーが使われます。",
    label_visibility="collapsed"
)

# APIキー決定ロジック
active_api_key = user_api_key if user_api_key else EMBEDDED_API_KEY

st.sidebar.markdown('<div class="sidebar-heading">📁 操作マニュアルの読み込み</div>', unsafe_allow_html=True)
st.sidebar.markdown("<small style='color:#555;'>クエリーマニュアルや軽技WEB操作手順書などをここから取り込めます。</small>", unsafe_allow_html=True)

uploaded_files = st.sidebar.file_uploader(
    "マニュアル資料 (Word, PDF, PPT, Excel, CSV)", 
    type=["docx", "pdf", "pptx", "xlsx", "xls", "csv"], 
    accept_multiple_files=True,
    key="file_uploader",
    label_visibility="collapsed"
)

all_extra_text = []
if uploaded_files:
    st.sidebar.markdown("<div style='margin-top: 10px; font-weight: bold;'>📖 読み込み中のマニュアル:</div>", unsafe_allow_html=True)
    for f in uploaded_files:
        try:
            content = get_text_from_file(f)
            all_extra_text.append(f"--- ファイル名: {f.name} ---\n{content}")
            st.sidebar.markdown(f'<div class="load-success">✔️ 読込完了: {f.name}</div>', unsafe_allow_html=True)
        except Exception as e:
            st.sidebar.error(f"❌ {f.name} の読み込みに失敗しました。")

# --- 出力フォーマットサンプルの取り込み・確認・クリア機能 ---
st.sidebar.markdown('<div class="sidebar-heading">📝 出力フォーマットの指定</div>', unsafe_allow_html=True)
st.sidebar.markdown("<small style='color:#555;'>AIに決まった形式で回答させたい場合は、参考となるサンプルファイル（Excel, PDF, Word, CSV, テキスト）を取り込めます。</small>", unsafe_allow_html=True)

format_file = st.sidebar.file_uploader(
    "出力サンプルの取り込み (テキスト, CSV, Word, Excel, PDF)",
    type=["txt", "csv", "docx", "xlsx", "xls", "pdf"],
    key=f"format_uploader_{st.session_state.format_file_key}",
    label_visibility="collapsed"
)

# アップロードされたファイルをセッション状態に保存
if format_file:
    try:
        st.session_state.format_sample = get_text_from_file(format_file)
        st.sidebar.success("✔️ 出力サンプルを取り込みました！")
    except Exception as e:
        st.sidebar.error(f"サンプルファイルの読み込みに失敗しました。理由: {str(e)}")

# 読み込んだサンプルの内容確認
if st.session_state.format_sample:
    with st.sidebar.expander("🔍 取り込み中のフォーマットを確認"):
        st.text_area(
            "現在のサンプル内容",
            value=st.session_state.format_sample,
            height=150,
            disabled=True,
            label_visibility="collapsed"
        )
    
    # クリアボタン
    if st.sidebar.button("🧹 出力サンプルを消去する", use_container_width=True):
        st.session_state.format_sample = ""
        st.session_state.format_file_key += 1 # キーを変更してファイルアップローダーを強制リセット
        st.rerun()

# 区切り線とアプリ終了ボタン (実行環境ポリシーに配慮し安全に対策)
st.sidebar.markdown("<br><hr>", unsafe_allow_html=True)
if st.sidebar.button("🛑 アプリを終了する", use_container_width=True):
    st.sidebar.warning("システムを停止します。このブラウザタブを閉じてください。")
    try:
        os.kill(os.getpid(), signal.SIGINT)
    except Exception:
        pass # クラウド環境などで強制終了プロセスが無効化されている場合のクラッシュを防ぐ

# --- 8. メインエリアのチャットレイアウト ---
if not uploaded_files:
    st.info("💡 左メニューの『操作マニュアルの読み込み』から資料をアップロードすると、さらに詳しい手順まで正確に案内できるようになります。")
else:
    st.success(f"📂 現在 {len(uploaded_files)} 個のマニュアル資料を基に回答を作成します。")

# チャットメッセージの履歴描画
for m in st.session_state.messages:
    role = "assistant" if m["role"] == "assistant" else "user"
    avatar = "🤖" if role == "assistant" else "👤"
    with st.chat_message(role, avatar=avatar):
        st.markdown(m["content"])

# 質問の入力欄
if prompt := st.chat_input("ここに質問したいことを書いて「送信」を押してください（例：データのダウンロード方法は？）"):
    st.session_state.messages.append({"role": "user", "content": prompt})
    with st.chat_message("user", avatar="👤"):
        st.markdown(prompt)

    # アシスタントの返答作成
    with st.chat_message("assistant", avatar="🤖"):
        with st.spinner("マニュアルから正しい操作方法を調べています。少々お待ちください..."):
            res = get_ai_roleplay_response(
                st.session_state.messages, 
                current_persona, 
                all_extra_text, 
                active_api_key, 
                format_sample=st.session_state.format_sample if st.session_state.format_sample else None
            )
        st.markdown(res)
        
    st.session_state.messages.append({"role": "assistant", "content": res})
    st.rerun()
