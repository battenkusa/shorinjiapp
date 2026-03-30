import streamlit as st
from pathlib import Path
from langchain_ollama import ChatOllama
from langchain_core.messages import HumanMessage
import os
import json
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime, timedelta
import pdfplumber
from docx import Document as DocxDocument

st.set_page_config(page_title="葛飾区少林寺拳法連盟AIチーム", layout="wide")
st.title("🤝 葛飾区少林寺拳法連盟AIチーム")

llm = ChatOllama(model="llama3.2")

# ═══ 提出資料監視機能の補助関数 ═══════════════════════

def get_submission_folder():
    """監視対象フォルダのパスを取得"""
    primary_path = Path(r"C:\Users\batte\OneDrive\少林寺\葛飾区連盟\AI作業フォルダ\提出関連資料")
    if primary_path.exists():
        return primary_path
    else:
        fallback_path = Path("input")
        fallback_path.mkdir(exist_ok=True)
        return fallback_path

def load_last_scan_time():
    """最終スキャン時刻を読み込み"""
    scan_file = Path("output/last_scan.json")
    if scan_file.exists():
        try:
            with open(scan_file, "r", encoding="utf-8") as f:
                data = json.load(f)
                return datetime.fromisoformat(data["last_scan"])
        except:
            return datetime.min
    return datetime.min

def save_last_scan_time():
    """現在時刻を最終スキャン時刻として保存"""
    scan_file = Path("output/last_scan.json")
    scan_file.parent.mkdir(exist_ok=True)
    with open(scan_file, "w", encoding="utf-8") as f:
        json.dump({"last_scan": datetime.now().isoformat()}, f, ensure_ascii=False)

def read_document_content(file_path):
    """ファイルの内容を読み込み"""
    try:
        file_path = Path(file_path)
        
        if file_path.suffix.lower() == ".pdf":
            with pdfplumber.open(file_path) as pdf:
                return "\n".join(
                    page.extract_text() for page in pdf.pages
                    if page.extract_text()
                )
        elif file_path.suffix.lower() == ".docx":
            doc = DocxDocument(file_path)
            return "\n".join(paragraph.text for paragraph in doc.paragraphs)
        elif file_path.suffix.lower() == ".txt":
            with open(file_path, "r", encoding="utf-8") as f:
                return f.read()
        else:
            return ""
    except Exception as e:
        return f"ファイル読み込みエラー: {e}"

def analyze_document_with_ollama(content):
    """Ollama を使って文書内容を分析"""
    prompt = f"""以下の文書を分析し、下記の情報を日本語で抽出してください。

【文書内容】
{content[:2000]}

【抽出する情報】
1. 概要: この文書の内容を簡潔に（1-2行）
2. 実施内容: 何をする必要があるか（具体的な作業内容）
3. 提出期限: 文書に記載されている期限日（YYYY-MM-DD形式で）
4. 依頼内容締切: 提出期限の10日前の日付（YYYY-MM-DD形式で）

【出力形式】
概要: [ここに概要]
実施内容: [ここに実施内容]
提出期限: [YYYY-MM-DD]
依頼内容締切: [YYYY-MM-DD]
"""
    
    try:
        res = llm.invoke([HumanMessage(content=prompt)])
        return res.content
    except Exception as e:
        return f"分析エラー: {e}"

def send_notification_email(file_name, analysis_result):
    """Gmail経由でメール送信"""
    try:
        gmail_address = os.getenv("GMAIL_ADDRESS")
        gmail_password = os.getenv("GMAIL_APP_PASSWORD")
        
        if not gmail_address or not gmail_password:
            return "環境変数 GMAIL_ADDRESS または GMAIL_APP_PASSWORD が設定されていません"
        
        msg = MIMEMultipart()
        msg['From'] = gmail_address
        msg['To'] = "battenkusayokayoka@gmail.com"
        msg['Subject'] = f"【提出資料監視】新しいファイルが検出されました: {file_name}"
        
        body = f"""
新しい提出資料が検出されました。

ファイル名: {file_name}
検出日時: {datetime.now().strftime('%Y年%m月%d日 %H:%M:%S')}

【AI分析結果】
{analysis_result}

--
葛飾区少林寺拳法連盟AIチーム
自動監視システム
        """
        
        msg.attach(MIMEText(body, 'plain', 'utf-8'))
        
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(gmail_address, gmail_password)
        server.send_message(msg)
        server.quit()
        
        return "メール送信完了"
        
    except Exception as e:
        return f"メール送信エラー: {e}"

def scan_for_new_files():
    """新しいファイルをスキャン"""
    folder = get_submission_folder()
    last_scan = load_last_scan_time()
    new_files = []
    
    for ext in ['*.pdf', '*.docx', '*.txt']:
        for file_path in folder.glob(ext):
            if file_path.is_file() and file_path.stat().st_mtime > last_scan.timestamp():
                new_files.append(file_path)
    
    return new_files

# ═══ 自動スキャン処理（起動時実行） ═══════════════════════

if "auto_scan_done" not in st.session_state:
    st.session_state.auto_scan_done = True
    
    try:
        new_files = scan_for_new_files()
        if new_files:
            for file_path in new_files:
                content = read_document_content(file_path)
                if content and not content.startswith("ファイル読み込みエラー"):
                    analysis = analyze_document_with_ollama(content)
                    send_notification_email(file_path.name, analysis)
            save_last_scan_time()
    except Exception as e:
        pass

menu = st.sidebar.radio("機能メニュー", [
    "📋 議事録作成",
    "✉️ 連絡文作成",
    "📢 お知らせ作成",
    "📅 活動計画作成",
    "📄 大会資料作成",
    "📬 提出資料監視",
])

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
    st.caption("指定フォルダの新しいファイルを監視し、AI分析結果をメールで通知します")
    
    folder = get_submission_folder()
    st.info(f"監視フォルダ: {folder}")
    
    # 最終スキャン時刻表示
    last_scan = load_last_scan_time()
    if last_scan != datetime.min:
        st.write(f"最終スキャン: {last_scan.strftime('%Y年%m月%d日 %H:%M:%S')}")
    else:
        st.write("最終スキャン: 未実行")
    
    col1, col2 = st.columns(2)
    
    with col1:
        if st.button("📂 手動スキャン実行"):
            with st.spinner("フォルダをスキャン中..."):
                new_files = scan_for_new_files()
                
                if not new_files:
                    st.success("新しいファイルは見つかりませんでした")
                else:
                    st.success(f"{len(new_files)}個の新しいファイルを検出しました")
                    
                    for file_path in new_files:
                        with st.expander(f"📄 {file_path.name}"):
                            st.write(f"ファイルパス: {file_path}")
                            st.write(f"更新日時: {datetime.fromtimestamp(file_path.stat().st_mtime)}")
                            
                            with st.spinner("内容を分析中..."):
                                content = read_document_content(file_path)
                                
                                if content.startswith("ファイル読み込みエラー"):
                                    st.error(content)
                                    continue
                                
                                analysis = analyze_document_with_ollama(content)
                                st.subheader("AI分析結果")
                                st.markdown(analysis)
                                
                                # メール送信
                                with st.spinner("メール送信中..."):
                                    result = send_notification_email(file_path.name, analysis)
                                    if "エラー" in result:
                                        st.error(result)
                                    else:
                                        st.success(result)
                    
                    save_last_scan_time()
    
    with col2:
        if st.button("🔄 スキャン履歴リセット"):
            try:
                scan_file = Path("output/last_scan.json")
                if scan_file.exists():
                    scan_file.unlink()
                st.success("スキャン履歴をリセットしました")
            except Exception as e:
                st.error(f"リセットエラー: {e}")
    
    st.subheader("設定情報")
    st.write("**対象ファイル形式:** PDF (.pdf), Word (.docx), テキスト (.txt)")
    st.write("**メール送信先:** battenkusayokayoka@gmail.com")
    st.write("**必要な環境変数:**")
    st.code("GMAIL_ADDRESS=your_email@gmail.com\nGMAIL_APP_PASSWORD=your_app_password")
    
    # 環境変数の状態確認
    gmail_address = os.getenv("GMAIL_ADDRESS")
    gmail_password = os.getenv("GMAIL_APP_PASSWORD")
    
    if gmail_address and gmail_password:
        st.success("✅ 環境変数が正しく設定されています")
    else:
        st.warning("⚠️ 環境変数 GMAIL_ADDRESS または GMAIL_APP_PASSWORD が設定されていません")