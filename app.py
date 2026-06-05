# -*- coding: utf-8 -*-
import streamlit as st

# --- 1. ページタイトルを動的に変更するためのセッション状態の初期化 ---
if "app_title" not in st.session_state:
    st.session_state.app_title = "🤖 汎用 AIFAQチャットシステム"

# --- 【最優先ルール】Streamlitのページ構成設定は、他のあらゆるコマンドより先に最上部で実行します ---
st.set_page_config(page_title=st.session_state.app_title, layout="wide")

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

# --- 2. APIキーの設定 (APIKEY.ini または クラウドのSecretsからハイブリッド取得) ---
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


# --- 3. 各ファイル抽出関数 (ポインタを先頭に戻す seek(0) を追加して読み込みエラーを完全に防止) ---
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


# --- 3-2. ファイル名からメインのシステム・サービス名を抽出するクレンジング関数（予備処理） ---
def clean_service_name(filename):
    base = os.path.splitext(filename)[0]
    # 一般的なドキュメント名称やノイズを正規表現で排除
    patterns = [
        r"(操作)?マニュアル", r"取扱説明書", r"手順書", r"仕様書", r"概要書", r"説明書",
        r"【.*】", r"\[.*\]", r"（.*）", r"\(.*\)",
        r"[vV]er\.?\d+(\.\d+)*", r"\d{8}", r"\d{6}",
        r"[-_]"
    ]
    cleaned = base
    for pat in patterns:
        cleaned = re.sub(pat, " ", cleaned)
    cleaned = cleaned.strip()
    return cleaned if cleaned else base


# --- 3-3. マニュアルテキストを解析して、メインの製品・システム名をGeminiからスマートに特定する関数 ---
def extract_service_name_via_ai(text, default_name, api_key):
    try:
        genai.configure(api_key=api_key)
        target_model = get_safe_model_name(api_key)
        model = genai.GenerativeModel(target_model)
        
        # テキストの先頭1500文字程度を利用して推測
        prompt = f"""
以下は、ユーザーからアップロードされたマニュアルまたは資料テキストの冒頭部分です。
この資料が「何のツール」「何のサービス」または「どのシステム」について説明しているものか、最もメインとなる固有名称を日本語で1つだけ見つけ出してください。
余計な説明、前置き、記号、拡張子などは絶対に含めず、純粋な名称のみを返してください。（例：「軽技WEB」「預かり資産トータルクエリーサービス」など）
最大でも20文字以内とします。特定が難しい場合は「{default_name}」を返してください。

【資料テキストの一部】
{text[:1500]}
"""
        response = model.generate_content(prompt)
        res_text = response.text.strip()
        res_text = re.sub(r"[`'\"]", "", res_text)  # 引用符の除去
        res_text = res_text.split("\n")[0].strip()
        return res_text if res_text else default_name
    except:
        return default_name


# --- 4. AI回答生成ロジック (自動リトライ・履歴ウィンドウ削減版 / 汎用化プロンプト) ---
def get_ai_roleplay_response(messages, persona, product_docs, format_docs, api_key):
    target_model = get_safe_model_name(api_key)
    
    # トークン爆発を防ぐため、AIに送る「過去の履歴」を最新の6発言（3往復）に限定
    recent_messages = [messages[0]] + messages[-5:] if len(messages) > 6 else messages

    for attempt in range(5):
        try:
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel(target_model)
            
            combined_docs = "\n\n".join(product_docs) if product_docs else "追加のマニュアル等のアップロードは現在ありません。一般的な知識に基づいてユーザーの課題解決を支援してください。"
            combined_formats = "\n\n".join(format_docs) if format_docs else "出力フォーマットサンプルの指定は現在ありません。"
            
            history_text = ""
            for m in recent_messages:
                role_label = "AIアシスタント(あなた)" if m["role"] == "assistant" else "ユーザー"
                history_text += f"{role_label}: {m['content']}\n"

            # 汎用FAQ用のシステムプロンプト構成
            system_prompt = f"""
あなたはユーザーから提供されたマニュアルや資料に基づいて、操作手順や記載内容を正確に説明する熟練のAIFAQアシスタントです。
操作方法、機能説明、データ構造、エラー解決等に関する質問に対して、丁寧かつ極めて分かりやすく回答してください。

【対象マニュアル・参照資料情報】
{persona['description']}

【アップロードされた各種マニュアル・参考ドキュメント（最優先参照情報）】
{combined_docs}

【アップロードされた出力フォーマットサンプル】
{combined_formats}
※この出力フォーマットサンプルが提示されている場合は、ユーザーが「これと同じデータ形式やレイアウトを出力したい」と希望しています。
現在アップロードされている各種マニュアル・参考ドキュメントを参照し、このフォーマットを出力するにはどのような操作、設定、データの選択や加工手順を行えばいいのかを、手順を追って具体的に説明してください。

【預かり資産トータルクエリーサービスに関する絶対判定ルール】
もしアップロードされた資料の内容や質問の文脈が「預かり資産トータルクエリーサービス」に関連する場合、ユーザーのやりたいデータ抽出や操作要望に対して、以下の思考プロセスを厳格に適用して回答を構成してください。
1. **「標準クエリ（約200種類）の確認」**: まず第一に、ユーザーのやりたい要件にそのまま合致する既存の標準クエリがすでに提供されているかを判断して案内してください。
2. **「既存クエリの修正・加工方法」**: もし完全に合致する既存の標準クエリがそのままでは見つからない場合、どの標準クエリをベース（ひな形）に選択し、それをどのように修正（項目追加、結合、フィルター条件の編集など）すれば目的の結果が得られるかを、具体的かつ分かりやすい手順として説明してください。

【回答の絶対ルール】
1. ユーザーの質問に対し、アップロードされたマニュアルの情報を最も信頼できる「絶対の基準（最優先情報）」として参照し、正確に回答を構成してください。
2. アップロードされた情報だけで判断がつかない不確実な事項やマニュアルに記載がない操作については、知ったかぶりをせず、「マニュアル等に記載がありませんでした」と明示したうえで、一般的な推奨方法を補足するか、専門の窓口や管理者への確認を案内してください。
3. 専門用語が使われている場合でも、操作担当者がスムーズに迷わず作業を進められるよう、ステップ・バイ・ステップの具体的な手順や丁寧な表現で回答してください。
4. AIとしてのメタな発言（例：「以上がアップロードされたマニュアルに基づく回答です」など）は含めず、ユーザーへの親切な回答テキストのみを親身なLINE風の対話形式で出力してください。

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


# --- 5. アプリケーションのメインロジック ---
def main_app():
    if st.session_state.get("app_terminated", False):
        st.warning("🛑 システムは終了しました。再度ご利用になる場合は、ブラウザをリロード（再読み込み）してください。")
        st.stop()

    # あたたかみのある緑ベースのカラーテーマ & LINE風チャットデザイン
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
            width: 100%;
        }
        /* 発言者を左右に寄せるための行ラッパー */
        .chat-row-user {
            display: flex;
            justify-content: flex-end;
            width: 100%;
        }
        .chat-row-assistant {
            display: flex;
            justify-content: flex-start;
            width: 100%;
        }
        .chat-bubble-user {
            background-color: #81C784;
            color: #1B5E20;
            padding: 12px 18px;
            border-radius: 18px 18px 0px 18px;
            max-width: 75%;
            box-shadow: 0px 2px 5px rgba(0,0,0,0.05);
            font-size: 15px;
            line-height: 1.5;
            text-align: left;
        }
        .chat-bubble-assistant {
            background-color: #FFFFFF;
            color: #2E7D32;
            padding: 12px 18px;
            border-radius: 18px 18px 18px 0px;
            max-width: 75%;
            box-shadow: 0px 2px 5px rgba(0,0,0,0.05);
            border: 1px solid #C8E6C9;
            font-size: 15px;
            line-height: 1.5;
            text-align: left;
        }
    </style>
    """, unsafe_allow_html=True)

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

    # --- タイトルとキャラクター情報の動的書き換えロジック ---
    all_extra_text = []
    file_names = []
    
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
                    file_names.append(f.name)
                    st.sidebar.write(f"✔️ 資料読込済: {f.name}")
            except Exception as e:
                st.sidebar.error(f"❌ {f.name} の読込失敗: {str(e)}")

    # 読み込まれたマニュアルに基づいてタイトルを賢く自動抽出・変更
    if file_names:
        # 新しくファイルを読み込んだ際、セッションキャッシュを更新
        if "last_processed_files" not in st.session_state or st.session_state.last_processed_files != file_names:
            # 1. まずは正規表現によるクレンジングをベースラインにする
            default_service_name = clean_service_name(file_names[0])
            
            # 2. APIキーが有効な場合は、AIを活用してドキュメントの本当のシステム・サービス名を正確に特定
            if ACTIVE_API_KEY and all_extra_text:
                joined_samples = "\n".join(all_extra_text)
                detected_name = extract_service_name_via_ai(joined_samples, default_service_name, ACTIVE_API_KEY)
                st.session_state.detected_service_name = detected_name
            else:
                st.session_state.detected_service_name = default_service_name
                
            st.session_state.last_processed_files = file_names
            
        # アプリタイトルの設定
        st.session_state.app_title = f"📖 {st.session_state.detected_service_name} 操作説明 AIFAQ"
        current_persona = {
            "description": f"提供されたマニュアル「{', '.join(file_names)}」（対象システム/ツール: {st.session_state.detected_service_name}）に精通した、専属の優秀なAIFAQ操作説明アシスタントです。"
        }
    else:
        st.session_state.app_title = "🤖 汎用 AIFAQチャットシステム"
        current_persona = {
            "description": "現在は特定のマニュアルはロードされていません。ロードされる多様なシステム・ツール資料や操作手順、一般的な疑問に対して柔軟に回答する汎用AIFAQアシスタントです。"
        }

    # アプリヘッダー表示
    st.markdown(f"""
    <div style="background-color: #2E7D32; padding: 20px; border-radius: 15px; text-align: center; margin-bottom: 25px; box-shadow: 0px 4px 10px rgba(0,0,0,0.08);">
        <h1 style="color: white; margin: 0; font-size: 28px;">{st.session_state.app_title}</h1>
        <p style="color: #E8F5E9; margin: 8px 0 0 0; font-size: 15px;">
            マニュアル資料や操作手順書を自動でインテリジェントに学習し、チャット形式で分かりやすく回答します。
        </p>
    </div>
    """, unsafe_allow_html=True)

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

    # --- 6. セッション状態の初期化 ---
    if "messages" not in st.session_state:
        st.session_state.messages = [{
            "role": "assistant", 
            "content": "汎用 AIFAQチャットシステムへようこそ！お手元の操作マニュアルや資料（Word, PDF, Excel, CSV, PPT等）を左側のメニューからアップロードしていただければ、即座にその内容を学習した専用の回答アシスタントとしてお答えいたします。\nまた、「再現したい成果物の出力サンプル」もお持ちの場合は、そちらをアップロードしていただくことで、マニュアルに沿ったデータ作成手順をお調べします。"
        }]

    if st.session_state.format_file_names:
        st.info(f"💡 出力フォーマットサンプル（{', '.join(st.session_state.format_file_names)}）が読み込まれています。チャットで「このサンプルを出力するには？」等と質問してみてください。")
    elif file_names:
        st.info(f"💡 現在、マニュアル資料（{', '.join(file_names)}）が読み込まれています。学習したシステム・ツール情報に基づいて的確に案内いたします！")
    else:
        st.info(f"💡 マニュアル資料を学習させたい場合は、左側のサイドバーからファイルをアップロードしてください。現在アップロードされたマニュアルはありません。")

    # チャット履歴をLINE風に描画
    chat_placeholder = st.container()

    with chat_placeholder:
        st.markdown('<div class="chat-container">', unsafe_allow_html=True)
        for m in st.session_state.messages:
            row_class = "chat-row-assistant" if m["role"] == "assistant" else "chat-row-user"
            bubble_class = "chat-bubble-assistant" if m["role"] == "assistant" else "chat-bubble-user"
            avatar = "🤖 AI" if m["role"] == "assistant" else "💼 ユーザー"
            
            st.markdown(f"""
            <div class="{row_class}">
                <div class="{bubble_class}">
                    <div style="font-weight: bold; font-size: 12px; margin-bottom: 5px; opacity: 0.85;">{avatar}</div>
                    <div>{m["content"]}</div>
                </div>
            </div>
            """, unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

    # ユーザー入力を受付
    if prompt := st.chat_input("マニュアルに関する質問や操作方法の疑問を入力してください..."):
        st.session_state.messages.append({"role": "user", "content": prompt})
        
        with chat_placeholder:
            st.markdown(f"""
            <div class="chat-row-user">
                <div class="chat-bubble-user">
                    <div style="font-weight: bold; font-size: 12px; margin-bottom: 5px; opacity: 0.85;">💼 ユーザー</div>
                    <div>{prompt}</div>
                </div>
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
