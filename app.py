# -*- coding: utf-8 -*-
import streamlit as st

# --- 【最優先ルール】Streamlitのページ構成設定は、他のあらゆるコマンドより先に最上部で実行します ---
st.set_page_config(page_title="預かり資産トータルクエリーサービス AIFAQ", layout="wide")

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
# ※既存 of API取得ロジックの構造・変数名を1行も崩さずに、クラウド安全対策を内包させています
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


# --- 2. 各ファイル抽出関数 (ポインタを先頭に戻す seek(0) を追加して読み込みエラーを完全に防止) ---
def extract_from_docx(file):
    file.seek(0)
    doc = Document(file)
    return "\n".join([para.text for para in doc.paragraphs])

def extract_from_pdf(file):
    file.seek(0)
    reader = PdfReader(file)
    return "\n".join([page.extract_text() for page in reader.pages if page.extract_text()])

def extract_from_pptx(file):
    file.seek(0)
    prs = Presentation(file)
    text_runs = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if hasattr(shape, "text"):
                text_runs.append(shape.text)
    return "\n".join(text_runs)

def extract_from_excel(file):
    file.seek(0)
    all_sheets = pd.read_excel(file, sheet_name=None)
    text_data = []
    for sheet_name, df in all_sheets.items():
        text_data.append(f"--- シート名: {sheet_name} ---\n{df.to_string(index=False)}")
    return "\n".join(text_data)

def extract_from_csv(file):
    file.seek(0)
    try:
        df = pd.read_csv(file)
        return df.to_string(index=False)
    except:
        try:
            file.seek(0)
            df = pd.read_csv(file, encoding="shift-jis")
            return df.to_string(index=False)
        except Exception as e:
            return f"[CSV読込エラー] {str(e)}"

def extract_from_text(file):
    file.seek(0)
    try:
        return file.read().decode("utf-8")
    except:
        try:
            file.seek(0)
            return file.read().decode("shift-jis")
        except Exception as e:
            return f"[テキスト読込エラー] {str(e)}"


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
def get_ai_roleplay_response(messages, persona, product_docs, format_docs, api_key):
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


# --- 4. アプリケーションのメインロジック ---
def main_app():
    if st.session_state.get("app_terminated", False):
        st.warning("🛑 システムは終了しました。再度ご利用になる場合は、ブラウザをリロード（再読み込み）してください。")
        st.stop()

    # あたたかみのある緑ベースのカラーテーマ & LINE風チャットデザイン
    # メインカラー3色: 深緑（#2E7D32）、薄緑（#F1F8E9）、チャット背景白（#FFFFFF）
    st.markdown("""
    <style>
        .stApp {
            background-color: #F9FBE7;
        }
        section[data-testid="stSidebar"] {
            background-color: #E8F5E9 !important;
            border-right: 2px solid #C8E6C9;
        }
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
        
        /* LINE風チャットコンテナと吹き出し */
        .chat-container {
            display: flex;
            flex-direction: column;
            gap: 15px;
            margin-bottom: 20px;
        }
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
    </style>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div style="background-color: #2E7D32; padding: 20px; border-radius: 15px; text-align: center; margin-bottom: 25px; box-shadow: 0px 4px 10px rgba(0,0,0,0.08);">
        <h1 style="color: white; margin: 0; font-size: 28px;">📊 預かり資産トータルクエリーサービス AIFAQ</h1>
        <p style="color: #E8F5E9; margin: 8px 0 0 0; font-size: 15px;">
            投資信託の実績集計や、BIツール「軽技WEB」、データベース「Fund Organizer」に関する疑問を解消するFAQアプリです。
        </p>
    </div>
    """, unsafe_allow_html=True)

    current_persona = {
        "service_name": "預かり資産トータルクエリーサービス",
        "overview": "軽技WEBというBIツールを利用した、投資信託の実績集計などが行える金融機関の本部担当者向け service。",
        "architecture": "顧客（投資家）の契約情報、取引履歴、残高、損益情報を保管しているデータベース「Fund Organizer」を参照元として、本部担当者のブラウザから「軽技WEB」を利用して接続し、データを取得する。",
        "features": "①約200の標準クエリから業務要件に合わせたクエリを選択して実行\n②標準クエリを編集しフレキシブルに加工・保存が可能\n③データ取得結果をCSV/Excel形式で保存が可能\n④スケジュール実行で業務の効率化を実現\n⑤専門の担当者による手厚いサポート"
    }

    # --- 左側サイドバー情報入力メニュー ---
    st.sidebar.markdown("""
    <div style="background-color: #2E7D32; padding: 12px; border-radius: 10px; margin-bottom: 15px; text-align: center;">
        <h3 style="color: white; margin: 0; font-size: 16px;">📁 コントロールパネル</h3>
    </div>
    """, unsafe_allow_html=True)

    st.sidebar.subheader("🔑 APIキーの設定")
    custom_api_key = st.sidebar.text_input(
        "Gemini APIキーを入力 (任意)",
        type="password",
        help="入力した場合、このAPIキーを優先して使用します。未入力時はシステムのデフォルトAPIキーを利用します。"
    )

    # APIキーの優先度選択
    ACTIVE_API_KEY = custom_api_key if custom_api_key else EMBEDDED_API_KEY

    if not ACTIVE_API_KEY:
        st.sidebar.error("⚠️ APIキーが設定されていません。サイドバーから入力するか、設定ファイルを確認してください。")
    else:
        if custom_api_key:
            st.sidebar.success("✔️ カスタムAPIキーを適用中")
        else:
            st.sidebar.info("✔️ デフォルトAPIキーを使用中")

    st.sidebar.markdown("---")

    # 1. 情報入力用マニュアルファイル
    st.sidebar.subheader("📄 マニュアル資料の読込")
    uploaded_files = st.sidebar.file_uploader(
        "資料 (Excel, Word, PDF, CSV, PPT)", 
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
                else: content = ""
                
                if content:
                    all_extra_text.append(f"--- ファイル名: {f.name} ---\n{content}")
                    st.sidebar.write(f"✔️ 資料読込済: {f.name}")
            except Exception as e:
                st.sidebar.error(f"❌ {f.name} の読込失敗: {str(e)}")

    st.sidebar.markdown("---")

    # 2. 出力フォーマットサンプル読込
    st.sidebar.subheader("📋 出力フォーマットサンプルの読込")
    st.sidebar.markdown("<small>出力したいサンプルの形式を取り込み、出力手順を確認できます</small>", unsafe_allow_html=True)

    if "format_samples" not in st.session_state:
        st.session_state.format_samples = []
    if "format_file_names" not in st.session_state:
        st.session_state.format_file_names = []

    uploaded_format = st.sidebar.file_uploader(
        "出力サンプル (Excel, Word, PDF, CSV, PPT, TXT)", 
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

    if st.session_state.format_file_names:
        st.sidebar.write("📌 現在取り込まれているサンプル:")
        for name in st.session_state.format_file_names:
            st.sidebar.write(f"・ {name}")
        
        # クリア機能
        if st.sidebar.button("🗑️ 出力フォーマットサンプルをクリア"):
            st.session_state.format_samples = []
            st.session_state.format_file_names = []
            st.sidebar.success("サンプルファイルをクリアしました。")
            st.rerun()

    st.sidebar.markdown("---")
    
    if st.sidebar.button("🛑 アプリを終了する"):
        st.session_state.app_terminated = True
        st.sidebar.warning("システムを終了しました。")
        st.rerun()

    # --- 5. セッション状態の初期化 ---
    if "messages" not in st.session_state:
        st.session_state.messages = [{
            "role": "assistant", 
            "content": "「預かり資産トータルクエリーサービス」AIFAQアプリへようこそ。投資信託の実績集計や軽技WEBの操作、標準クエリの活用方法について何でもご質問ください。マニュアル資料や再現したい出力サンプルのフォーマットを左側のメニューからアップロードしていただければ、具体的な操作方法やデータ抽出の手順をご案内いたします。"
        }]

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

    # ユーザー入力を受付
    if prompt := st.chat_input("クエリーサービスの質問内容を入力してください（例：スケジュール実行の方法は？等）"):
        st.session_state.messages.append({"role": "user", "content": prompt})
        
        with chat_placeholder:
            st.markdown(f"""
            <div class="chat-bubble-user">
                <div style="font-weight: bold; font-size: 12px; margin-bottom: 5px; opacity: 0.85;">💼 ユーザー</div>
                <div>{prompt}</div>
            </div>
            """, unsafe_allow_html=True)
            
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
            
        st.session_state.messages.append({"role": "assistant", "content": res})
        st.rerun()

# --- メインロジックを実行 ---
main_app()
