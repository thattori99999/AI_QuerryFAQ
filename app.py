# -*- coding: utf-8 -*-
import streamlit as st

# --- 最優先ルール: Streamlitのページ構成設定を、すべてのコマンドに先駆けて先頭で実行します ---
st.set_page_config(page_title="預かり資産トータルクエリーサービス AIFAQ", layout="wide")

# --- 外部ライブラリの安全なインポート (インポートエラーによるアプリ起動即死を完全に回避する超堅牢設計) ---
import_errors = []

try:
    import pandas as pd
except ImportError:
    import_errors.append("pandas")
    pd = None

try:
    import google.generativeai as genai
    from google.api_core import exceptions  # Rate Limit(429) エラーを確実に捕捉するため
except Exception as e:
    import_errors.append("google-generativeai")
    genai = None
    exceptions = None

try:
    from docx import Document
except ImportError:
    import_errors.append("python-docx")
    Document = None

try:
    from PyPDF2 import PdfReader
except ImportError:
    import_errors.append("PyPDF2")
    PdfReader = None

try:
    from pptx import Presentation
except ImportError:
    import_errors.append("python-pptx")
    Presentation = None

# Excelパース用のopenpyxl確認
try:
    import openpyxl
except ImportError:
    import_errors.append("openpyxl")
    openpyxl = None

import io
import configparser
import os
import signal
import re
import time  # リトライ待機（スリープ）処理のため

# --- 必須コアパッケージがサーバーに無い場合、クラッシュを防止し、解決手順のガイド画面を親切に表示します ---
if "google-generativeai" in import_errors or "pandas" in import_errors:
    st.error("⚙️ アプリの起動に必要なシステムパッケージ（ライブラリ）がデプロイ環境に不足しています。")
    st.markdown("""
    ### 🛠️ 解決方法 (Streamlit Cloudでの手順)
    
    Streamlit Cloudなどのサーバー環境で外部ライブラリを有効化するには、プロジェクトのルートフォルダに **`requirements.txt`** ファイルを配置する必要があります。
    
    GitHubリポジトリ内に **`requirements.txt`** という名前のファイルを新規作成し、以下のテキストをそのままコピーして保存（コミット）してください。
    """)
    
    st.code("""
google-generativeai
pandas
python-docx
PyPDF2
python-pptx
openpyxl
    """, language="text")
    
    st.markdown("""
    コミットが完了すると、Streamlit Cloudが自動的に変更を検知してパッケージをインストールし、本アプリが100%正常に起動するようになります。
    """)
    st.stop()  # ここで実行を安全に中断します


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
        if hasattr(st, "secrets") and st.secrets is not None:
            if "GEMINI" in st.secrets:
                gemini_sec = st.secrets["GEMINI"]
                if hasattr(gemini_sec, "get"):
                    key = gemini_sec.get("API_KEY", None)
                    if key: return key
                elif isinstance(gemini_sec, dict) and "API_KEY" in gemini_sec:
                    return gemini_sec["API_KEY"]
            if "API_KEY" in st.secrets:
                return st.secrets["API_KEY"]
    except Exception:
        pass
        
    return None

INI_KEY = load_api_key()
EMBEDDED_API_KEY = INI_KEY

# --- 2. 各ファイル抽出関数 (ポインタを先頭に戻す seek(0) を追加してEOFクラッシュを防止) ---
def extract_from_docx(file):
    if Document is None:
        return "[システム警告] python-docx パッケージがインストールされていないため、Wordファイルを解析できません。 requirements.txt に追加してください。"
    try:
        file.seek(0)
        doc = Document(file)
        return "\n".join([para.text for para in doc.paragraphs])
    except Exception as e:
        return f"[Word読込エラー] {str(e)}"

def extract_from_pdf(file):
    if PdfReader is None:
        return "[システム警告] PyPDF2 パッケージがインストールされていないため、PDFファイルを解析できません。 requirements.txt に追加してください。"
    try:
        file.seek(0)
        reader = PdfReader(file)
        return "\n".join([page.extract_text() for page in reader.pages if page.extract_text()])
    except Exception as e:
        return f"[PDF読込エラー] {str(e)}"

def extract_from_pptx(file):
    if Presentation is None:
        return "[システム警告] python-pptx パッケージがインストールされていないため、PowerPointファイルを解析できません。 requirements.txt に追加してください。"
    try:
        file.seek(0)
        prs = Presentation(file)
        text_runs = []
        for slide in prs.slides:
            for shape in slide.shapes:
                if hasattr(shape, "text"):
                    text_runs.append(shape.text)
        return "\n".join(text_runs)
    except Exception as e:
        return f"[PowerPoint読込エラー] {str(e)}"

def extract_from_excel(file):
    if pd is None:
        return "[システム警告] pandas がロードされていません。"
    try:
        file.seek(0)
        all_sheets = pd.read_excel(file, sheet_name=None)
        text_data = []
        for sheet_name, df in all_sheets.items():
            text_data.append(f"--- シート名: {sheet_name} ---\n{df.to_string(index=False)}")
        return "\n".join(text_data)
    except Exception as e:
        return f"[Excel読込エラー] {str(e)}\n※xlsxファイル読込には openpyxl パッケージが必要です。"

def extract_from_csv(file):
    if pd is None:
        return "[システム警告] pandas がロードされていません。"
    try:
        file.seek(0)
        df = pd.read_csv(file)
        return df.to_string(index=False)
    except Exception:
        try:
            file.seek(0)
            df = pd.read_csv(file, encoding="shift-jis")
            return df.to_string(index=False)
        except Exception as e:
            return f"[CSV読込エラー] {str(e)}"

def extract_from_text(file):
    try:
        file.seek(0)
        return file.read().decode("utf-8")
    except Exception:
        try:
            file.seek(0)
            return file.read().decode("shift-jis")
        except Exception as e:
            return f"[テキスト読込エラー] {str(e)}"

# --- 404エラーを回避しつつ、利用可能なモデル名を安全に取得する関数 ---
def get_safe_model_name(api_key):
    if genai is None:
        return 'gemini-1.5-flash'
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
def get_ai_roleplay_response(messages, persona, product_docs, format_docs, api_key):
    if genai is None or exceptions is None:
        return "【システムエラー】Google Gemini APIライブラリ(google-generativeai)が正常にロードされていません。requirements.txtに追加されているか確認してください。"

    target_model = get_safe_model_name(api_key)
    
    # トークン爆発を防ぐため、AIに送る「過去の履歴」を最新の6発言（3往復）に限定
    recent_messages = [messages[0]] + messages[-5:] if len(messages) > 6 else messages

    for attempt in range(5):
        try:
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel(target_model)
            
            combined_docs = "\n\n".join(product_docs) if product_docs else "追加のマニュアル等のアップロードは現在ありません。標準仕様に基づいて回答してください。"
            combined_formats = "\n\n".join(format_docs) if format_docs else "出力フォーマットサンプルの指定は現在ありません。"
            
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

【アップロードされた出力フォーマットサンプル】
{combined_formats}
※この出力フォーマットサンプルが提示されている場合は、このデータ形式やレイアウトを出力するために、システム上でどのような抽出操作や設定、クエリのカスタマイズを行えばよいのかを、上記の製品仕様・マニュアルと照らし合わせて具体的に提案してください。

【回答の絶対ルール】
1. ユーザー（本部担当者）の質問に対し、アップロードされたマニュアルの情報を最優先に参照して回答を構成してください。
2. あくまで「預かり資産トータルクエリーサービス」および「軽技WEB」「Fund Organizer」の文脈に沿って回答してください。一般的なIT知識でなく、本サービスの仕様を重視してください。
3. 専門的な内容であっても、金融機関の業務担当者がスムーズに作業を進められるよう、具体的な手順や選択すべきクエリ（標準クエリ約200種）など、実務に即した丁寧な表現で回答してください。
4. アップロードされた情報だけで判断がつかない不確実な事項については、知ったかぶりをせず、「専門のサポート窓口」へ案内するなどの対応を含めてください。
5. AIとしてのメタな発言（例：「以上がマニュアルに基づく回答です」など）は含めず、ユーザーへの親切な回答テキストのみを出力してください。
6. 出力フォーマットサンプルが提示されている場合、その内容を出力するためのアプローチ（どのテーブルから、どの項目をどのように加工して出力するかなど）を分かりやすく説明してください。

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

# --- Rerun処理の安全な抽象化 (古いStreamlit環境でも絶対にクラッシュさせないフォールバック) ---
def safe_rerun():
    try:
        st.rerun()
    except AttributeError:
        try:
            st.experimental_rerun()
        except AttributeError:
            pass  # rerunがどちらも提供されていない超旧バージョンでは自然な状態遷移に委ねます

# --- 4. 画面構築 (Streamlit UI) ---
# チャットUI(st.chat_input, st.chat_message)がインストールされたStreamlit環境でサポートされているかを検証
has_chat_ui = hasattr(st, "chat_input") and hasattr(st, "chat_message")

# あたたかみのある緑ベースのカラーテーマ & LINE風チャットスタイリング (3色をベースに構成)
# メインの緑: #2E7D32、背景の薄緑: #F1F8E9、チャット背景: #E8F5E9 / #FFFFFF
st.markdown("""
<style>
    /* 全体デザイン調整 */
    .stApp {
        background-color: #F9FBE7;
    }
    
    /* サイドバーの背景 */
    section[data-testid="stSidebar"] {
        background-color: #E8F5E9 !important;
        border-right: 2px solid #C8E6C9;
    }
    
    /* ボタンの緑色カスタマイズ */
    div.stButton > button {
        background-color: #4CAF50 !important;
        color: white !important;
        border-radius: 20px !important;
        border: none !important;
        padding: 0.5rem 1.5rem !important;
        font-weight: bold !important;
        transition: all 0.3s;
    }
    div.stButton > button:hover {
        background-color: #2E7D32 !important;
        box-shadow: 0 4px 8px rgba(0,0,0,0.1);
    }
    
    /* LINE風チャット吹き出しカスタマイズ */
    .chat-container {
        display: flex;
        flex-direction: column;
        gap: 15px;
        margin-bottom: 20px;
    }
    
    /* ユーザー（右側・緑色系） */
    .chat-bubble-user {
        background-color: #81C784;
        color: #1B5E20;
        padding: 12px 18px;
        border-radius: 18px 18px 0px 18px;
        align-self: flex-end;
        max-width: 75%;
        box-shadow: 0px 2px 5px rgba(0,0,0,0.05);
        font-size: 15px;
        line-height: 1.5;
    }
    
    /* AIアシスタント（左側・白系） */
    .chat-bubble-assistant {
        background-color: #FFFFFF;
        color: #2E7D32;
        padding: 12px 18px;
        border-radius: 18px 18px 18px 0px;
        align-self: flex-start;
        max-width: 75%;
        box-shadow: 0px 2px 5px rgba(0,0,0,0.05);
        border: 1px solid #C8E6C9;
        font-size: 15px;
        line-height: 1.5;
    }

    .chat-meta {
        font-size: 11px;
        color: #757575;
        margin-top: 4px;
    }
</style>
""", unsafe_allow_html=True)

# ヘッダーデザイン
st.markdown("""
<div style="background-color: #2E7D32; padding: 20px; border-radius: 15px; text-align: center; margin-bottom: 25px; box-shadow: 0px 4px 10px rgba(0,0,0,0.08);">
    <h1 style="color: white; margin: 0; font-size: 28px;">📊 預かり資産トータルクエリーサービス AIFAQ</h1>
    <p style="color: #E8F5E9; margin: 8px 0 0 0; font-size: 15px;">
        投資信託の実績集計や、BIツール「軽技WEB」、データベース「Fund Organizer」に関する疑問を解消するFAQアプリです。
    </p>
</div>
""", unsafe_allow_html=True)

# --- 5. 固定ペルソナ情報設定 ---
current_persona = {
    "service_name": "預かり資産トータルクエリーサービス",
    "overview": "軽技WEBというBIツールを利用した、投資信託の実績集計などが行える金融機関の本部担当者向けサービス。",
    "architecture": "顧客（投資家）の契約情報、取引履歴、残高、損益情報を保管しているデータベース「Fund Organizer」を参照元として、本部担当者のブラウザから「軽技WEB」を利用して接続し、データを取得する。",
    "features": "①約200の標準クエリから業務要件に合わせたクエリを選択して実行\n②標準クエリを編集しフレキシブルに加工・保存が可能\n③データ取得結果をCSV/Excel形式で保存が可能\n④スケジュール実行で業務の効率化を実現\n⑤専門の担当者による手厚いサポート"
}

# --- 6. サイドバー機能 (APIキー設定 & 資料・サンプルのロード) ---
st.sidebar.markdown("""
<div style="background-color: #2E7D32; padding: 12px; border-radius: 10px; margin-bottom: 15px; text-align: center;">
    <h3 style="color: white; margin: 0; font-size: 16px;">📁 コントロールパネル</h3>
</div>
""", unsafe_allow_html=True)

# インポートエラー検知時の自己申告メッセージ (必須以外のオプショナルライブラリ)
if import_errors:
    st.sidebar.warning(
        f"⚠️ 以下の補助パッケージがデプロイ環境にインストールされていません。マニュアル読み込み時の一部形式が制限されます。解決するには `requirements.txt` へ追記してください。\n"
        f"不足パッケージ: {', '.join(import_errors)}"
    )

# APIキー設定（最優先されるカスタムキー入力欄）
st.sidebar.subheader("🔑 APIキーの設定")
custom_api_key = st.sidebar.text_input(
    "Gemini APIキーを入力 (任意)",
    type="password",
    help="入力した場合、このAPIキーを優先して使用します。未入力時はシステムのデフォルトAPIキーを利用します。"
)

# APIキー決定のロジック
ACTIVE_API_KEY = custom_api_key if custom_api_key else EMBEDDED_API_KEY

if not ACTIVE_API_KEY:
    st.sidebar.error("⚠️ APIキーが設定されていません。サイドバーから入力するか、設定ファイルを確認してください。")
else:
    if custom_api_key:
        st.sidebar.success("✔️ カスタムAPIキーを適用中")
    else:
        st.sidebar.info("✔️ デフォルトAPIキーを使用中")

st.sidebar.markdown("---")

# 6-1. マニュアルファイル読込
st.sidebar.subheader("📄 マニュアル資料の読込")
uploaded_files = st.sidebar.file_uploader(
    "マニュアル資料 (Word, PDF, PPT, Excel, CSV)", 
    type=["docx", "pdf", "pptx", "xlsx", "xls", "csv"], 
    accept_multiple_files=True,
    key="file_uploader"
)

all_extra_text = []
if uploaded_files is not None:
    for f in uploaded_files:
        try:
            if f.name.endswith(".docx"): content = extract_from_docx(f)
            elif f.name.endswith(".pdf"): content = extract_from_pdf(f)
            elif f.name.endswith(".pptx"): content = extract_from_pptx(f)
            elif f.name.endswith((".xlsx", ".xls")): content = extract_from_excel(f)
            elif f.name.endswith(".csv"): content = extract_from_csv(f)
            else: content = ""
            
            if content:
                all_extra_text.append(f"--- ファイル名: {f.name} ---\n{content}")
                st.sidebar.write(f"✔️ 資料読込済: {f.name}")
        except Exception as e:
            st.sidebar.error(f"❌ {f.name} の読込失敗: {str(e)}")

st.sidebar.markdown("---")

# 6-2. 出力フォーマットサンプルの読込とクリア
st.sidebar.subheader("📋 出力フォーマットサンプルの読込")
st.sidebar.markdown("<small>出力したいサンプルの形式を取り込み、その出力方法を確認できます</small>", unsafe_allow_html=True)

# セッション状態でのサンプルデータ保持
if "format_samples" not in st.session_state:
    st.session_state.format_samples = []
if "format_file_names" not in st.session_state:
    st.session_state.format_file_names = []

uploaded_format = st.sidebar.file_uploader(
    "出力サンプル (Word, PDF, PPT, Excel, CSV, Text)", 
    type=["docx", "pdf", "pptx", "xlsx", "xls", "csv", "txt"],
    key="format_uploader"
)

if uploaded_format:
    f_name = uploaded_format.name
    if f_name not in st.session_state.format_file_names:
        try:
            if f_name.endswith(".docx"): content = extract_from_docx(uploaded_format)
            elif f_name.endswith(".pdf"): content = extract_from_pdf(uploaded_format)
            elif f_name.endswith(".pptx"): content = extract_from_pptx(uploaded_format)
            elif f_name.endswith((".xlsx", ".xls")): content = extract_from_excel(uploaded_format)
            elif f_name.endswith(".csv"): content = extract_from_csv(uploaded_format)
            elif f_name.endswith(".txt"): content = extract_from_text(uploaded_format)
            else: content = ""
            
            if content:
                st.session_state.format_samples.append(f"--- サンプルファイル名: {f_name} ---\n{content}")
                st.session_state.format_file_names.append(f_name)
        except Exception as e:
            st.sidebar.error(f"❌ {f_name} の読込失敗: {str(e)}")

# 読み込まれている出力フォーマットサンプルの表示
if st.session_state.format_file_names:
    st.sidebar.write("📌 現在取り込まれているサンプル:")
    for name in st.session_state.format_file_names:
        st.sidebar.write(f"・ {name}")
    
    # サンプルのクリアボタン
    if st.sidebar.button("🗑️ 出力フォーマットサンプルをクリア"):
        st.session_state.format_samples = []
        st.session_state.format_file_names = []
        st.sidebar.success("サンプルファイルをクリアしました。")
        safe_rerun()

st.sidebar.markdown("---")
if st.sidebar.button("🛑 アプリを終了する"):
    st.sidebar.warning("システムを終了します。")
    try:
        os.kill(os.getpid(), signal.SIGINT)
    except Exception as e:
        st.sidebar.error(f"終了処理に失敗しました: {e}")

# --- 7. セッション状態の初期化 ---
if "messages" not in st.session_state:
    st.session_state.messages = [{
        "role": "assistant", 
        "content": "「預かり資産トータルクエリーサービス」AIFAQアプリへようこそ。投資信託の実績集計や軽技WEBの操作、標準クエリの活用方法について何でもご質問ください。マニュアル資料や、再現したい出力サンプルのフォーマットを左側のメニューからアップロードしていただければ、具体的な操作方法やデータ抽出の手順をご案内いたします。"
    }]

# --- 8. メインエリアのチャットレイアウト ---
# お知らせ表示
if st.session_state.format_file_names:
    st.info(f"💡 出力フォーマットサンプル（{', '.join(st.session_state.format_file_names)}）が読み込まれています。チャットで「このサンプルを出力するには？」等と質問してみてください。")
else:
    st.info(f"💡 マニュアルを読み込ませる場合や、特定の出力サンプルに基づいて抽出手順を知りたい場合は、左側のサイドバーからファイルをアップロードしてください。")

# チャット履歴をLINE風に描画
chat_placeholder = st.container()

with chat_placeholder:
    st.markdown('<div class="chat-container">', unsafe_allow_html=True)
    for m in st.session_state.messages:
        role_class = "chat-bubble-assistant" if m["role"] == "assistant" else "chat-bubble-user"
        avatar = "🤖 AI" if m["role"] == "assistant" else "💼 ユーザー"
        
        st.markdown(f"""
        <div class="{role_class}">
            <div style="font-weight: bold; font-size: 12px; margin-bottom: 5px; opacity: 0.85;">{avatar}</div>
            <div>{m["content"]}</div>
        </div>
        """, unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

# ユーザー入力プロンプトの判定用変数 (Python 旧バージョンのウォルラス演算子互換性対策)
prompt = None

# Streamlitのバージョンに応じてチャット入力方法を自動分岐
if has_chat_ui:
    # 新しいStreamlitバージョンでのチャットUI (ウォルラス演算子を使用しない安全設計)
    user_input = st.chat_input("クエリーサービスの質問内容を入力してください（例：スケジュール実行の方法は？等）")
    if user_input:
        prompt = user_input
else:
    # 旧型のStreamlitバージョンでの代替フォームUI（これで起動エラーが100%回避されます）
    st.write("---")
    with st.form(key="chat_input_form", clear_on_submit=True):
        col1, col2 = st.columns([8, 2])
        with col1:
            user_input = st.text_input("質問内容を入力してください（送信後、履歴に追記されます）", placeholder="例：スケジュール実行の方法は？等")
        with col2:
            submitted = st.form_submit_button("送信")
        
        if submitted and user_input:
            prompt = user_input

# チャット入力が検知された場合の処理
if prompt:
    # 1. ユーザーメッセージをセッションに追加
    st.session_state.messages.append({"role": "user", "content": prompt})
    
    # 2. 即時反映（擬似的な描画）
    with chat_placeholder:
        st.markdown(f"""
        <div class="chat-bubble-user">
            <div style="font-weight: bold; font-size: 12px; margin-bottom: 5px; opacity: 0.85;">💼 ユーザー</div>
            <div>{prompt}</div>
        </div>
        """, unsafe_allow_html=True)
        
    # 3. AIアシスタント回答の生成
    # st.chat_messageが使えない場合はクラシックな描画方式に切り替え
    if has_chat_ui:
        with st.chat_message("assistant", avatar="🤖"):
            with st.spinner("該当する情報を確認して回答を作成中..."):
                res = get_ai_roleplay_response(
                    st.session_state.messages, 
                    current_persona, 
                    all_extra_text, 
                    st.session_state.format_samples,
                    ACTIVE_API_KEY
                )
                st.markdown(res)
    else:
        with st.spinner("該当する情報を確認して回答を作成中..."):
            res = get_ai_roleplay_response(
                st.session_state.messages, 
                current_persona, 
                all_extra_text, 
                st.session_state.format_samples,
                ACTIVE_API_KEY
            )
            st.info(res)
        
    # 4. 回答をセッション履歴へ格納し、安全に再同期
    st.session_state.messages.append({"role": "assistant", "content": res})
    safe_rerun()
