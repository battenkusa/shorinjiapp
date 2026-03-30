import streamlit as st
from pathlib import Path
from langchain_ollama import ChatOllama
from langchain_core.messages import HumanMessage
import json
import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime, timedelta
import pdfplumber
from docx import Document
import platform

st.set_page_config(page_title="葛飾区少林寺拳法連盟AIチーム", layout="wide")
st.title("🤝 葛飾区少林寺拳法連盟AIチーム")

llm = ChatOllama(model="llama3.2")

# ファイル監視・メール送信機能
def get_submission_folder():
    """提出資料監視フォルダのパスを取得"""
    if platform.system() == "Windows":
        return Path(r"C:\Users\batte\OneDrive\少林寺\葛飾区連盟\AI作業フォルダ\提出関連資料")
    else:
        return Path("input")  # フォールバック

def get_last_scan_time():
    """最終スキャン時刻を取得"""
    scan_file = Path("output/last_scan.json")
    if scan_file.exists():
        try:
            with open(scan_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
                return datetime.fromisoformat(data.get('last_scan', '1970-01-01T00:00:00'))
        except Exception:
            return datetime(1970, 1, 1)
    return datetime(1970, 1, 1)

def update_last_scan_time():
    """最終スキャン時刻を更新"""
    scan_file = Path("output/last_scan.json")
    scan_file.parent.mkdir(parents=True, exist_ok=True)
    with open(scan_file, 'w', encoding='utf-8') as f:
        json.dump({'last_scan': datetime.now().isoformat()}, f, ensure_ascii=False)

def extract_text_from_file(file_path):
    """ファイルからテキストを抽出"""
    file_path = Path(file_path)
    text = ""
    
    try:
        if file_path.suffix.lower() == '.pdf':
            with pdfplumber.open(file_path) as pdf:
                text = "\n".join(page.extract_text() or "" for page in pdf.pages)
        elif file_path.suffix.lower() == '.docx':
            doc = Document(file_path)
            text = "\n".join(paragraph.text for paragraph in doc.paragraphs)
        elif file_path.suffix.lower() == '.txt':
            with open(file_path, 'r', encoding='utf-8') as f:
                text = f.read()
    except Exception as e:
        st.error(f"ファイル読み込みエラー: {file_path.name} - {e}")
    
    return text

def analyze_document(text, filename):
    """文書内容をOllamaで分析"""
    prompt = f"""以下の文書を分析し、下記の形式で日本語で回答してください。

文書名: {filename}
内容:
{text[:2000]}

【分析結果】
概要: （この文書の概要を1-2行で）
実施内容: （具体的に何をするのか）
提出期限: （文書に記載の期限、不明な場合は「不明」）
依頼内容締切: （提出期限の10日前、期限が不明な場合は「不明」）

注意: 日付は元の文書に記載された通りに抽出してください。推測しないでください。"""
    
    try:
        res = llm.invoke([HumanMessage(content=prompt)])
        return res.content
    except Exception as e:
        return f"分析エラー: {e}"

def send_notification_email(analysis_result, filename):
    """Gmail経由で通知メールを送信"""
    gmail_address = os.getenv('GMAIL_ADDRESS')
    gmail_password = os.getenv('GMAIL_APP_PASSWORD')
    
    if not gmail_address or not gmail_password:
        st.error("Gmail設定が不正です。環境変数 GMAIL_ADDRESS と GMAIL_APP_PASSWORD を設定してください。")
        return False
    
    try:
        msg = MIMEMultipart()
        msg['From'] = gmail_address
        msg['To'] = "battenkusayokayoka@gmail.com"
        msg['Subject'] = f"【資料提出依頼】{filename}"
        
        body = f"""新しい提出資料が検出されました。

ファイル名: {filename}
検出日時: {datetime.now().strftime('%Y年%m月%d日 %H:%M:%S')}

{analysis_result}

このメールは自動送信されました。
葛飾区少林寺拳法連盟AIチーム"""
        
        msg.attach(MIMEText(body, 'plain', 'utf-8'))
        
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(gmail_address, gmail_password)
        text = msg.as_string()
        server.sendmail(gmail_address, "battenkusayokayoka@gmail.com", text)
        server.quit()
        return True
    except Exception as e:
        st.error(f"メール送信エラー: {e}")
        return False

def scan_for_new_files():
    """新しいファイルをスキャンして処理"""
    folder_path = get_submission_folder()
    last_scan = get_last_scan_time()
    new_files = []
    
    if not folder_path.exists():
        st.warning(f"監視フォルダが見つかりません: {folder_path}")
        return new_files
    
    for file_path in folder_path.rglob("*"):
        if file_path.is_file() and file_path.suffix.lower() in ['.pdf', '.docx', '.txt']:
            try:
                file_mtime = datetime.fromtimestamp(file_path.stat().st_mtime)
                if file_mtime > last_scan:
                    new_files.append(file_path)
            except Exception:
                continue
    
    return new_files

menu = st.sidebar.radio("機能メニュー", [
    "📋 議事録作成",
    "✉️ 連絡文作成",
    "📢 お知らせ作成",
    "📅 活動計画作成",
    "📄 大会資料作成",
    "📬 提出資料監視",
])

# 起動時自動スキャン（初回のみ）
if 'auto_scan_done' not in st.session_state:
    st.session_state.auto_scan_done = True
    with st.spinner("起動時スキャンを実行中..."):
        new_files = scan_for_new_files()
        if new_files:
            st.info(f"起動時スキャンで新しいファイルを{len(new_files)}件検出しました")
            for file_path in new_files:
                text = extract_text_from_file(file_path)
                if text.strip():
                    analysis = analyze_document(text, file_path.name)
                    send_notification_email(analysis, file_path.name)
            update_last_scan_time()
        else:
            st.info("起動時スキャン完了: 新しいファイルはありませんでした")

# ── 議事録作成 ──────────────────────────
if menu == "📋 議事録作成":
    st.subheader("議事録作成")
    col1, col2 = st.columns(2)
    title   = col1.text_input("会議名", placeholder="例：3月定例ミーティング")
    date    = col2.text_input("日時",   placeholder="例：2025年3月28日 19:00〜")
    members = st.text_input("参加者",   placeholder="例：田中・鈴木・佐藤（計3名）")
    memo    = st.text_area("会議メモ", height=150,
                placeholder="・文化祭の出し物をカフェに決定\n・担当は田中・鈴木・佐藤\n・次回は4月5日集合")

    if st.button("議事録を生成する"):
        with st.spinner("AIが生成中..."):
            prompt = f"""以下の会議メモから、見やすい議事録を日本語で作成してください。
会議名：{title} / 日時：{date} / 参加者：{members}
メモ：{memo}"""
            res = llm.invoke([HumanMessage(content=prompt)])
        st.success("完成しました！")
        st.markdown(res.content)
        st.download_button("テキストで保存", res.content, "議事録.txt")

# ── 連絡文作成 ──────────────────────────
elif menu == "✉️ 連絡文作成":
    st.subheader("連絡文作成")
    col1, col2 = st.columns(2)
    to      = col1.text_input("宛先",   placeholder="例：部員の皆さん")
    sender  = col2.text_input("差出人", placeholder="例：部長 田中")
    tone    = st.selectbox("トーン", ["丁寧・敬語", "やや砕けた", "フレンドリー"])
    content = st.text_area("伝えたい内容", height=120,
                placeholder="例：次の練習は4月5日に変更。場所は第2体育館。")

    if st.button("連絡文を生成する"):
        with st.spinner("AIが生成中..."):
            prompt = f"""以下の情報をもとに、{tone}なトーンでメール文を日本語で作成してください。
宛先：{to} / 差出人：{sender}
内容：{content}"""
            res = llm.invoke([HumanMessage(content=prompt)])
        st.success("完成しました！")
        st.markdown(res.content)
        st.download_button("テキストで保存", res.content, "連絡文.txt")

# ── お知らせ作成 ────────────────────────
elif menu == "📢 お知らせ作成":
    st.subheader("お知らせ作成")
    usage   = st.radio("用途", ["掲示板・印刷", "LINE / SNS", "チラシ"], horizontal=True)
    content = st.text_area("お知らせ内容", height=120,
                placeholder="例：5月の大会でカフェを出します。メンバー募集中です。")

    if st.button("お知らせを生成する"):
        with st.spinner("AIが生成中..."):
            prompt = f"""{usage}向けのお知らせ文を日本語で作成してください。
内容：{content}"""
            res = llm.invoke([HumanMessage(content=prompt)])
        st.success("完成しました！")
        st.markdown(res.content)
        st.download_button("テキストで保存", res.content, "お知らせ.txt")

# ── 活動計画作成 ────────────────────────
elif menu == "📅 活動計画作成":
    st.subheader("活動計画作成")
    goal    = st.text_input("目標・イベント名", placeholder="例：春季区民体育大会")
    col1, col2 = st.columns(2)
    start   = col1.text_input("開始日", placeholder="例：4月1日")
    end     = col2.text_input("締切日", placeholder="例：5月31日")
    tasks   = st.text_area("やること（わかっているもの）",
                placeholder="例：参加者募集、会場準備、役割分担、リハーサル")

    if st.button("計画を生成する"):
        with st.spinner("AIが生成中..."):
            prompt = f"""以下の情報をもとに、週単位の活動スケジュールを日本語で作成してください。
目標：{goal} / 期間：{start}〜{end}
やること：{tasks}"""
            res = llm.invoke([HumanMessage(content=prompt)])
        st.success("完成しました！")
        st.markdown(res.content)
        st.download_button("テキストで保存", res.content, "活動計画.txt")

# ── 大会資料作成 ────────────────────────
elif menu == "📄 大会資料作成":
    st.subheader("大会資料作成（PDF → Word）")
    st.caption("昨年のPDFを読み込んで、今年版のWordファイルを自動生成します")

    pdf_path = st.text_input(
        "元になるPDFのパス",
        value=r"C:\Users\batte\OneDrive\少林寺\葛飾区連盟\2025年度\春季区大会\2025年度春季区民体育大会案rev２.pdf"
    )

    col1, col2, col3 = st.columns(3)
    year    = col1.text_input("今年度", value="2026")
    date    = col2.text_input("実施日", value="5/31")
    members = col3.text_input("人数",   value="80")

    out_path = st.text_input(
        "保存先（Wordファイル）",
        value=r"C:\Users\batte\myagent\output\2026年度春季区大会資料.docx"
    )

    if st.button("Word資料を生成する"):
        import pdfplumber
        from docx import Document as DocxDocument

        # ① PDFを読み込む
        with st.spinner("PDFを読み込み中..."):
            try:
                with pdfplumber.open(pdf_path) as pdf:
                    pdf_text = "\n".join(
                        page.extract_text() for page in pdf.pages
                        if page.extract_text()
                    )
                st.info(f"PDF読み込み完了（{len(pdf_text)}文字）")
            except Exception as e:
                st.error(f"PDF読み込みエラー: {e}")
                st.stop()

        # ② AIに内容を更新させる
        with st.spinner("AIが内容を更新中..."):
            prompt = f"""以下は2025年度の大会資料のテキストです。
これをもとに、下記の変更点を反映した{year}年度版の資料文章を作成してください。

【変更点】
- 年度：2025年度 → {year}年度
- 実施日：{date}
- 参加人数：{members}人
- その他の日付・年度の記載も{year}年度に合わせて更新

【元の資料】
{pdf_text[:3000]}

【出力形式】
タイトル、目的、実施日、会場、参加者、スケジュール、注意事項などの項目を
整理した資料文章を日本語で出力してください。"""

            res = llm.invoke([HumanMessage(content=prompt)])
            updated_text = res.content

        # ③ Wordファイルに書き出す
        with st.spinner("Wordファイルを作成中..."):
            doc = DocxDocument()
            doc.add_heading(f"{year}年度 春季区民体育大会資料", level=1)

            for line in updated_text.split("\n"):
                line = line.strip()
                if not line:
                    continue
                if line.startswith("【") or line.startswith("■") or line.endswith("】"):
                    doc.add_heading(line, level=2)
                else:
                    doc.add_paragraph(line)

            Path(out_path).parent.mkdir(parents=True, exist_ok=True)
            doc.save(out_path)

        st.success("完成しました！")
        st.code(out_path)
        st.subheader("生成された内容（プレビュー）")
        st.markdown(updated_text)

# ── 提出資料監視 ────────────────────────
elif menu == "📬 提出資料監視":
    st.subheader("提出資料監視・メール送信")
    
    # 設定情報表示
    folder_path = get_submission_folder()
    last_scan = get_last_scan_time()
    
    st.info(f"📁 監視フォルダ: `{folder_path}`")
    st.info(f"⏰ 最終スキャン: {last_scan.strftime('%Y年%m月%d日 %H:%M:%S')}")
    
    # 手動スキャンボタン
    if st.button("🔍 手動スキャン実行"):
        with st.spinner("スキャンを実行中..."):
            new_files = scan_for_new_files()
            
            if new_files:
                st.success(f"新しいファイルを{len(new_files)}件検出しました！")
                
                for i, file_path in enumerate(new_files):
                    st.write(f"**{i+1}. {file_path.name}**")
                    
                    with st.spinner(f"{file_path.name}を分析中..."):
                        text = extract_text_from_file(file_path)
                        
                        if text.strip():
                            analysis = analyze_document(text, file_path.name)
                            st.markdown(f"```\n{analysis}\n```")
                            
                            if send_notification_email(analysis, file_path.name):
                                st.success(f"✅ {file_path.name} の通知メールを送信しました")
                            else:
                                st.error(f"❌ {file_path.name} のメール送信に失敗しました")
                        else:
                            st.warning(f"⚠️ {file_path.name} からテキストを抽出できませんでした")
                
                update_last_scan_time()
                st.rerun()
                
            else:
                st.info("新しいファイルは見つかりませんでした")
    
    # 環境変数確認
    st.subheader("📧 メール設定確認")
    gmail_address = os.getenv('GMAIL_ADDRESS')
    gmail_password = os.getenv('GMAIL_APP_PASSWORD')
    
    if gmail_address and gmail_password:
        st.success(f"✅ Gmail設定済み: {gmail_address}")
    else:
        st.error("❌ Gmail設定が不完全です")
        st.code("""
環境変数を設定してください:
GMAIL_ADDRESS=your-email@gmail.com
GMAIL_APP_PASSWORD=your-app-password
        """)
    
    # 監視フォルダ確認
    st.subheader("📁 監視フォルダ確認")
    if folder_path.exists():
        st.success(f"✅ 監視フォルダが存在します: {folder_path}")
        try:
            files = list(folder_path.rglob("*"))
            pdf_files = [f for f in files if f.suffix.lower() == '.pdf']
            docx_files = [f for f in files if f.suffix.lower() == '.docx']
            txt_files = [f for f in files if f.suffix.lower() == '.txt']
            
            st.write(f"📄 PDFファイル: {len(pdf_files)}件")
            st.write(f"📝 Wordファイル: {len(docx_files)}件") 
            st.write(f"📃 テキストファイル: {len(txt_files)}件")
        except Exception as e:
            st.error(f"フォルダアクセスエラー: {e}")
    else:
        st.warning(f"⚠️ 監視フォルダが見つかりません: {folder_path}")
        if platform.system() != "Windows":
            st.info("Windows以外の環境では input/ フォルダを使用します")