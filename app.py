import streamlit as st
import os
import json
import smtplib
from pathlib import Path
from datetime import datetime, timedelta
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from langchain_ollama import ChatOllama
from langchain_core.messages import HumanMessage

st.set_page_config(page_title="葛飾区少林寺拳法連盟AIチーム", layout="wide")
st.title("🤝 葛飾区少林寺拳法連盟AIチーム")

llm = ChatOllama(model="llama3.2")

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
    st.caption("指定フォルダの新しい提出資料を監視し、内容を分析してメール通知します")
    
    # 監視フォルダの設定（Windows環境の場合は指定パス、それ以外はinput/フォルダ）
    if os.name == 'nt':  # Windows
        monitor_folder = Path(r"C:\Users\batte\OneDrive\少林寺\葛飾区連盟\AI作業フォルダ\提出関連資料")
    else:
        monitor_folder = Path("input")
    
    # output フォルダの作成
    output_folder = Path("output")
    output_folder.mkdir(exist_ok=True)
    last_scan_file = output_folder / "last_scan.json"
    
    st.info(f"監視対象フォルダ: {monitor_folder}")
    
    # 自動スキャン（初回起動時のみ）
    if "auto_scan_done" not in st.session_state:
        with st.spinner("起動時スキャンを実行中..."):
            auto_scan_results = scan_for_new_files(monitor_folder, last_scan_file)
            if auto_scan_results:
                st.success(f"自動スキャンで{len(auto_scan_results)}件の新しいファイルを発見しました")
                for file_path in auto_scan_results:
                    process_document(file_path, llm)
            else:
                st.info("自動スキャン完了：新しいファイルはありませんでした")
        st.session_state.auto_scan_done = True
    
    col1, col2 = st.columns(2)
    
    with col1:
        if st.button("手動スキャン実行"):
            with st.spinner("新しいファイルをスキャン中..."):
                new_files = scan_for_new_files(monitor_folder, last_scan_file)
                if new_files:
                    st.success(f"{len(new_files)}件の新しいファイルを発見しました")
                    for file_path in new_files:
                        process_document(file_path, llm)
                else:
                    st.info("新しいファイルは見つかりませんでした")
    
    with col2:
        if st.button("スキャン履歴をリセット"):
            if last_scan_file.exists():
                last_scan_file.unlink()
            st.success("スキャン履歴をリセットしました")
    
    # 最終スキャン時刻の表示
    if last_scan_file.exists():
        with open(last_scan_file, 'r', encoding='utf-8') as f:
            scan_data = json.load(f)
            st.caption(f"最終スキャン: {scan_data.get('last_scan_time', '未記録')}")


def scan_for_new_files(monitor_folder, last_scan_file):
    """新しいファイルをスキャンする"""
    new_files = []
    
    # 前回のスキャン時刻を取得
    last_scan_time = None
    if last_scan_file.exists():
        try:
            with open(last_scan_file, 'r', encoding='utf-8') as f:
                scan_data = json.load(f)
                last_scan_time = datetime.fromisoformat(scan_data['last_scan_time'])
        except:
            last_scan_time = None
    
    # フォルダが存在しない場合は作成
    if not monitor_folder.exists():
        monitor_folder.mkdir(parents=True, exist_ok=True)
        st.warning(f"監視フォルダを作成しました: {monitor_folder}")
        return new_files
    
    # サポートされるファイル形式
    supported_extensions = {'.pdf', '.docx', '.txt'}
    
    # ファイルをスキャン
    for file_path in monitor_folder.rglob('*'):
        if file_path.is_file() and file_path.suffix.lower() in supported_extensions:
            file_mod_time = datetime.fromtimestamp(file_path.stat().st_mtime)
            
            if last_scan_time is None or file_mod_time > last_scan_time:
                new_files.append(file_path)
    
    # スキャン時刻を更新
    current_time = datetime.now()
    scan_data = {
        'last_scan_time': current_time.isoformat(),
        'scanned_files': len(new_files)
    }
    
    with open(last_scan_file, 'w', encoding='utf-8') as f:
        json.dump(scan_data, f, ensure_ascii=False, indent=2)
    
    return new_files


def process_document(file_path, llm):
    """文書を処理してメール送信する"""
    st.subheader(f"📄 {file_path.name}")
    
    # ファイル内容を読み込み
    try:
        content = read_document_content(file_path)
        if not content:
            st.error("ファイルの内容を読み取れませんでした")
            return
        
        st.info(f"文書を読み込みました（{len(content)}文字）")
        
    except Exception as e:
        st.error(f"ファイル読み込みエラー: {e}")
        return
    
    # AI分析
    with st.spinner("AI分析中..."):
        analysis_result = analyze_document_with_ai(content, llm)
        if not analysis_result:
            st.error("AI分析に失敗しました")
            return
    
    # 分析結果を表示
    st.json(analysis_result)
    
    # メール送信
    with st.spinner("メール送信中..."):
        email_sent = send_notification_email(file_path.name, analysis_result)
        if email_sent:
            st.success("メール送信完了")
        else:
            st.error("メール送信に失敗しました（環境変数を確認してください）")


def read_document_content(file_path):
    """ファイル内容を読み取る"""
    suffix = file_path.suffix.lower()
    
    try:
        if suffix == '.txt':
            with open(file_path, 'r', encoding='utf-8') as f:
                return f.read()
        
        elif suffix == '.pdf':
            import pdfplumber
            with pdfplumber.open(file_path) as pdf:
                return "\n".join(
                    page.extract_text() for page in pdf.pages
                    if page.extract_text()
                )
        
        elif suffix == '.docx':
            from docx import Document
            doc = Document(file_path)
            return "\n".join(paragraph.text for paragraph in doc.paragraphs)
        
        else:
            return None
            
    except Exception as e:
        st.error(f"ファイル読み込みエラー ({suffix}): {e}")
        return None


def analyze_document_with_ai(content, llm):
    """AI で文書内容を分析する"""
    prompt = f"""以下の提出資料を分析し、JSON形式で結果を返してください。

【分析対象文書】
{content[:2000]}

【求める情報】
- 概要: 文書の概要を1-2文で
- 実施内容: 実施すべき内容・業務を箇条書きで
- 期限: 文書に記載されている期限・締切日（YYYY-MM-DD形式）
- 依頼締切: 期限の10日前の日付（YYYY-MM-DD形式）

【出力形式】
{{
  "概要": "文書の概要",
  "実施内容": ["実施内容1", "実施内容2"],
  "期限": "YYYY-MM-DD",
  "依頼締切": "YYYY-MM-DD"
}}

期限が不明な場合は null を設定してください。"""

    try:
        response = llm.invoke([HumanMessage(content=prompt)])
        result_text = response.content.strip()
        
        # JSONの抽出を試みる
        if '{' in result_text and '}' in result_text:
            json_start = result_text.find('{')
            json_end = result_text.rfind('}') + 1
            json_str = result_text[json_start:json_end]
            return json.loads(json_str)
        else:
            return None
            
    except Exception as e:
        st.error(f"AI分析エラー: {e}")
        return None


def send_notification_email(filename, analysis_result):
    """分析結果をメールで送信する"""
    try:
        # 環境変数からメール設定を取得
        gmail_address = os.getenv('GMAIL_ADDRESS')
        gmail_password = os.getenv('GMAIL_APP_PASSWORD')
        
        if not gmail_address or not gmail_password:
            st.warning("環境変数 GMAIL_ADDRESS または GMAIL_APP_PASSWORD が設定されていません")
            return False
        
        # メール内容を作成
        subject = f"【提出資料検出】{filename}"
        
        body = f"""新しい提出資料が検出されました。

【ファイル名】
{filename}

【概要】
{analysis_result.get('概要', '不明')}

【実施内容】
"""
        
        for item in analysis_result.get('実施内容', []):
            body += f"• {item}\n"
        
        body += f"""
【提出期限】
{analysis_result.get('期限', '不明')}

【依頼内容締切】
{analysis_result.get('依頼締切', '不明')}

---
葛飾区少林寺拳法連盟AIチームより自動送信
"""
        
        # メール送信
        msg = MIMEMultipart()
        msg['From'] = gmail_address
        msg['To'] = "battenkusayokayoka@gmail.com"
        msg['Subject'] = subject
        
        msg.attach(MIMEText(body, 'plain', 'utf-8'))
        
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(gmail_address, gmail_password)
            server.send_message(msg)
        
        return True
        
    except Exception as e:
        st.error(f"メール送信エラー: {e}")
        return False