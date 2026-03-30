import streamlit as st
from pathlib import Path
from langchain_ollama import ChatOllama
from langchain_core.messages import HumanMessage
import os
import json
from datetime import datetime, timedelta
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import pdfplumber
from docx import Document as DocxDocument

st.set_page_config(page_title="葛飾区少林寺拳法連盟AIチーム", layout="wide")
st.title("🤝 葛飾区少林寺拳法連盟AIチーム")

llm = ChatOllama(model="llama3.2")

# =====================================
# 提出資料監視関連の関数定義
# =====================================

def scan_and_process_documents(target_folder, manual=True):
    """フォルダをスキャンして新しいドキュメントを処理する"""
    if manual:
        with st.spinner("フォルダをスキャン中..."):
            result = _scan_documents(target_folder)
    else:
        result = _scan_documents(target_folder)
    
    if result['new_files']:
        if manual:
            st.success(f"新しいファイルが{len(result['new_files'])}件見つかりました")
            for file_info in result['new_files']:
                st.write(f"📄 {file_info['name']} ({file_info['modified']})")
        
        # 各ファイルを分析してメール送信
        for file_info in result['new_files']:
            if manual:
                with st.spinner(f"'{file_info['name']}'を分析中..."):
                    analysis = analyze_document(file_info['path'])
                    if analysis:
                        send_notification_email(file_info, analysis)
                        st.success(f"'{file_info['name']}'の分析完了・メール送信済み")
            else:
                analysis = analyze_document(file_info['path'])
                if analysis:
                    send_notification_email(file_info, analysis)
    elif manual:
        st.info("新しいファイルは見つかりませんでした")


def _scan_documents(target_folder):
    """ドキュメントフォルダをスキャンして新しいファイルを検出"""
    folder_path = Path(target_folder)
    if not folder_path.exists():
        return {'new_files': []}
    
    # 前回のスキャン時刻を取得
    last_scan_path = Path("output/last_scan.json")
    last_scan_time = None
    
    if last_scan_path.exists():
        try:
            with open(last_scan_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                last_scan_time = datetime.fromisoformat(data['timestamp'])
        except:
            pass
    
    # 対象ファイル形式
    target_extensions = ['.pdf', '.docx', '.txt']
    new_files = []
    
    for ext in target_extensions:
        for file_path in folder_path.glob(f"*{ext}"):
            if file_path.is_file():
                modified_time = datetime.fromtimestamp(file_path.stat().st_mtime)
                
                # 前回スキャン後に更新されたファイルのみ対象
                if last_scan_time is None or modified_time > last_scan_time:
                    new_files.append({
                        'name': file_path.name,
                        'path': str(file_path),
                        'modified': modified_time.strftime('%Y-%m-%d %H:%M')
                    })
    
    # 今回のスキャン時刻を記録
    current_time = datetime.now()
    last_scan_path.parent.mkdir(exist_ok=True)
    with open(last_scan_path, 'w', encoding='utf-8') as f:
        json.dump({
            'timestamp': current_time.isoformat(),
            'scanned_files': len(new_files)
        }, f, ensure_ascii=False, indent=2)
    
    return {'new_files': new_files}


def analyze_document(file_path):
    """ドキュメントをAIで分析する"""
    try:
        # ファイル内容を読み込み
        file_path = Path(file_path)
        content = ""
        
        if file_path.suffix.lower() == '.pdf':
            with pdfplumber.open(file_path) as pdf:
                content = "\n".join(
                    page.extract_text() for page in pdf.pages
                    if page.extract_text()
                )
        elif file_path.suffix.lower() == '.docx':
            doc = DocxDocument(file_path)
            content = "\n".join(paragraph.text for paragraph in doc.paragraphs)
        elif file_path.suffix.lower() == '.txt':
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
        
        if not content.strip():
            return None
        
        # AIで分析
        prompt = f"""以下の提出資料を分析して、下記の項目を抽出してください：

1. 概要（何についての資料か）
2. 実施内容（具体的に何をするのか）
3. 提出期限（資料に記載されている期限日）
4. 依頼内容締切（提出期限の10日前の日付）

資料内容：
{content[:2000]}

以下のJSON形式で出力してください：
{{
  "概要": "...",
  "実施内容": "...",
  "提出期限": "YYYY-MM-DD",
  "依頼内容締切": "YYYY-MM-DD"
}}"""
        
        response = llm.invoke([HumanMessage(content=prompt)])
        
        # JSON部分を抽出
        response_text = response.content.strip()
        start = response_text.find('{')
        end = response_text.rfind('}') + 1
        
        if start >= 0 and end > start:
            json_str = response_text[start:end]
            return json.loads(json_str)
        
        return None
        
    except Exception as e:
        return None


def send_notification_email(file_info, analysis):
    """分析結果をもとにメール通知を送信"""
    try:
        # 環境変数から認証情報を取得
        gmail_address = os.getenv('GMAIL_ADDRESS')
        gmail_app_password = os.getenv('GMAIL_APP_PASSWORD')
        
        if not gmail_address or not gmail_app_password:
            return False
        
        # メール内容作成
        subject = f"📬 新しい提出資料検出: {file_info['name']}"
        
        body = f"""新しい提出資料が検出されました。

────────────────────────────────────────
📄 ファイル情報
────────────────────────────────────────
ファイル名: {file_info['name']}
更新日時: {file_info['modified']}
ファイルパス: {file_info['path']}

────────────────────────────────────────
🤖 AI分析結果
────────────────────────────────────────
概要: {analysis.get('概要', '分析できませんでした')}

実施内容: {analysis.get('実施内容', '分析できませんでした')}

提出期限: {analysis.get('提出期限', '期限不明')}
依頼内容締切: {analysis.get('依頼内容締切', '期限不明')}

────────────────────────────────────────

自動生成メール by 葛飾区少林寺拳法連盟AIチーム
"""
        
        # メール送信
        msg = MIMEMultipart()
        msg['From'] = gmail_address
        msg['To'] = 'battenkusayokayoka@gmail.com'
        msg['Subject'] = subject
        
        msg.attach(MIMEText(body, 'plain', 'utf-8'))
        
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(gmail_address, gmail_app_password)
        text = msg.as_string()
        server.sendmail(gmail_address, 'battenkusayokayoka@gmail.com', text)
        server.quit()
        
        return True
        
    except Exception as e:
        return False

# =====================================
# メインUI処理
# =====================================

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
    st.caption("新しい提出資料を自動検出して、AIで分析後にメール通知を送信します")
    
    # 監視対象フォルダの設定
    target_folder = r"C:\Users\batte\OneDrive\少林寺\葛飾区連盟\AI作業フォルダ\提出関連資料"
    
    # Windows環境でない場合のフォールバック
    if not Path(target_folder).exists():
        target_folder = "input"
        st.info(f"メインフォルダが見つからないため、{target_folder}フォルダを監視します")
    
    st.info(f"監視対象: {target_folder}")
    
    # 手動スキャン実行
    if st.button("今すぐスキャン実行"):
        scan_and_process_documents(target_folder, manual=True)
    
    # 最終スキャン時刻表示
    last_scan_path = Path("output/last_scan.json")
    if last_scan_path.exists():
        with open(last_scan_path, 'r', encoding='utf-8') as f:
            last_scan_data = json.load(f)
        st.text(f"最終スキャン: {last_scan_data.get('timestamp', '未記録')}")
    else:
        st.text("最終スキャン: 未実行")

# 自動スキャン機能（セッション開始時に1回実行）
if 'auto_scan_done' not in st.session_state:
    st.session_state.auto_scan_done = True
    # 監視対象フォルダの設定
    target_folder = r"C:\Users\batte\OneDrive\少林寺\葛飾区連盟\AI作業フォルダ\提出関連資料"
    if not Path(target_folder).exists():
        target_folder = "input"
    
    # バックグラウンドで自動スキャン実行
    try:
        scan_and_process_documents(target_folder, manual=False)
    except Exception as e:
        pass  # エラーは無視してUI表示を妨げない