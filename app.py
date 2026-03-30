import streamlit as st
from pathlib import Path
from langchain_ollama import ChatOllama
from langchain_core.messages import HumanMessage
import os
import json
import smtplib
import platform
from datetime import datetime, timedelta
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

st.set_page_config(page_title="葛飾区少林寺拳法連盟AIチーム", layout="wide")
st.title("🤝 葛飾区少林寺拳法連盟AIチーム")

llm = ChatOllama(model="llama3.2")

# ── ヘルパー関数（提出資料監視用） ──────────────────────────
def get_monitor_folder():
    """監視対象フォルダのパスを取得"""
    if platform.system() == "Windows":
        return Path(r"C:\Users\batte\OneDrive\少林寺\葛飾区連盟\AI作業フォルダ\提出関連資料")
    else:
        # Windows以外の環境ではinputフォルダを使用
        return Path("input")

def get_last_scan_time():
    """最後のスキャン時刻を取得"""
    scan_file = Path("output/last_scan.json")
    if scan_file.exists():
        try:
            with open(scan_file, "r", encoding="utf-8") as f:
                data = json.load(f)
                return datetime.fromisoformat(data["last_scan"])
        except:
            pass
    return datetime.min

def save_last_scan_time():
    """最後のスキャン時刻を保存"""
    scan_file = Path("output/last_scan.json")
    scan_file.parent.mkdir(parents=True, exist_ok=True)
    data = {"last_scan": datetime.now().isoformat()}
    with open(scan_file, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def read_file_content(file_path):
    """ファイル内容を読み込む"""
    file_path = Path(file_path)
    
    try:
        if file_path.suffix.lower() == '.pdf':
            import pdfplumber
            with pdfplumber.open(file_path) as pdf:
                text = "\n".join([
                    page.extract_text() or "" 
                    for page in pdf.pages
                ])
                return text
        
        elif file_path.suffix.lower() == '.docx':
            from docx import Document
            doc = Document(file_path)
            text = "\n".join([paragraph.text for paragraph in doc.paragraphs])
            return text
        
        elif file_path.suffix.lower() == '.txt':
            with open(file_path, 'r', encoding='utf-8') as f:
                return f.read()
        
        else:
            return None
    except Exception as e:
        st.error(f"ファイル読み込みエラー ({file_path.name}): {e}")
        return None

def analyze_document(content, filename):
    """文書をOllamaで分析"""
    prompt = f"""以下の文書を分析して、以下の項目を抽出してください。

文書名: {filename}
内容:
{content[:2000]}

以下の形式で回答してください：

【概要】
（この文書の概要を簡潔に）

【実施内容】
（行うべき内容・作業を具体的に）

【提出期限】
（文書に記載の期限日時。見つからない場合は「記載なし」）

【依頼内容締切】
（提出期限の10日前の日付。提出期限が記載ありの場合のみ計算）

【緊急度】
（高・中・低で評価）"""

    try:
        response = llm.invoke([HumanMessage(content=prompt)])
        return response.content
    except Exception as e:
        return f"分析エラー: {e}"

def send_email(subject, body):
    """Gmail経由でメール送信"""
    gmail_address = os.environ.get("GMAIL_ADDRESS")
    gmail_password = os.environ.get("GMAIL_APP_PASSWORD")
    
    if not gmail_address or not gmail_password:
        raise ValueError("環境変数 GMAIL_ADDRESS または GMAIL_APP_PASSWORD が設定されていません")
    
    msg = MIMEMultipart()
    msg['From'] = gmail_address
    msg['To'] = "battenkusayokayoka@gmail.com"
    msg['Subject'] = subject
    
    msg.attach(MIMEText(body, 'plain', 'utf-8'))
    
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
        server.login(gmail_address, gmail_password)
        server.send_message(msg)

def scan_for_new_files():
    """新しいファイルをスキャンして分析・メール送信"""
    monitor_folder = get_monitor_folder()
    
    if not monitor_folder.exists():
        return f"監視フォルダが見つかりません: {monitor_folder}"
    
    last_scan = get_last_scan_time()
    new_files = []
    
    # 対象ファイル拡張子
    target_extensions = {'.pdf', '.docx', '.txt'}
    
    for file_path in monitor_folder.rglob('*'):
        if (file_path.is_file() and 
            file_path.suffix.lower() in target_extensions and
            datetime.fromtimestamp(file_path.stat().st_mtime) > last_scan):
            new_files.append(file_path)
    
    if not new_files:
        save_last_scan_time()
        return "新しいファイルはありませんでした"
    
    results = []
    for file_path in new_files:
        st.info(f"分析中: {file_path.name}")
        
        # ファイル内容を読み込み
        content = read_file_content(file_path)
        if not content:
            continue
        
        # AI分析
        analysis = analyze_document(content, file_path.name)
        
        # メール送信
        try:
            subject = f"【提出資料監視】新しい資料が検出されました: {file_path.name}"
            body = f"""葛飾区少林寺拳法連盟AIチーム

新しい提出関連資料が検出されました。

ファイル名: {file_path.name}
検出時刻: {datetime.now().strftime('%Y年%m月%d日 %H:%M:%S')}
ファイルパス: {file_path}

=== AI分析結果 ===
{analysis}

=== システム情報 ===
この通知は提出資料監視システムによって自動送信されました。
"""
            
            send_email(subject, body)
            results.append(f"✅ {file_path.name} - 分析・メール送信完了")
            
        except Exception as e:
            results.append(f"❌ {file_path.name} - メール送信エラー: {e}")
    
    save_last_scan_time()
    return "\n".join(results)

# ── 起動時の自動スキャン ──────────────────────────
if 'auto_scan_done' not in st.session_state:
    st.session_state.auto_scan_done = True
    
    # バックグラウンドで自動スキャン実行
    try:
        result = scan_for_new_files()
        if "新しいファイルはありませんでした" not in result:
            st.info("🔍 起動時スキャン完了: " + result.split('\n')[0])
    except Exception as e:
        st.warning(f"⚠️ 起動時スキャンでエラーが発生: {e}")

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
    st.caption("指定フォルダの新しいファイルを監視し、AI分析結果をメールで自動通知します")
    
    # 監視設定表示
    monitor_folder = get_monitor_folder()
    col1, col2 = st.columns(2)
    
    with col1:
        st.info("📁 監視フォルダ")
        st.code(str(monitor_folder))
        
        st.info("📧 通知先メール")
        st.code("battenkusayokayoka@gmail.com")
    
    with col2:
        st.info("📋 対象ファイル")
        st.write("• PDF (.pdf)")
        st.write("• Word (.docx)")
        st.write("• テキスト (.txt)")
        
        st.info("⚙️ 環境変数設定")
        gmail_ok = bool(os.environ.get("GMAIL_ADDRESS"))
        password_ok = bool(os.environ.get("GMAIL_APP_PASSWORD"))
        st.write(f"• GMAIL_ADDRESS: {'✅' if gmail_ok else '❌'}")
        st.write(f"• GMAIL_APP_PASSWORD: {'✅' if password_ok else '❌'}")
    
    # スキャン履歴表示
    last_scan = get_last_scan_time()
    if last_scan != datetime.min:
        st.success(f"前回スキャン: {last_scan.strftime('%Y年%m月%d日 %H:%M:%S')}")
    else:
        st.warning("初回スキャンです")
    
    # 手動スキャンボタン
    if st.button("🔍 手動スキャン実行", type="primary"):
        if not gmail_ok or not password_ok:
            st.error("❌ Gmail送信の環境変数が設定されていません")
            st.info("環境変数 GMAIL_ADDRESS と GMAIL_APP_PASSWORD を設定してください")
        else:
            with st.spinner("フォルダをスキャン中..."):
                try:
                    result = scan_for_new_files()
                    st.success("スキャン完了!")
                    st.text_area("結果", value=result, height=200)
                except Exception as e:
                    st.error(f"スキャンエラー: {e}")
    
    # 使用方法の説明
    st.subheader("📖 使用方法")
    st.markdown("""
    **自動監視:**
    - Streamlit起動時に自動的にスキャンが実行されます
    - 新しいファイルが見つかった場合、自動的にメール通知されます
    
    **手動スキャン:**
    - 上記のボタンでいつでも手動スキャンできます
    - 前回スキャン時刻以降に更新されたファイルが対象です
    
    **分析内容:**
    - 概要、実施内容、提出期限、依頼内容締切、緊急度
    - Ollama (llama3.2) によるAI分析
    
    **セキュリティ:**
    - Gmail送信には環境変数を使用（APIキー直書き禁止）
    - ファイルパスは Path() で安全に処理
    """)
    
    # 環境変数設定ガイド
    if not gmail_ok or not password_ok:
        st.subheader("⚙️ 環境変数設定ガイド")
        st.code("""
# .envファイルまたはシステム環境変数に設定
GMAIL_ADDRESS=your-email@gmail.com
GMAIL_APP_PASSWORD=your-app-password

# Googleアプリパスワードの取得方法：
# 1. Googleアカウント設定 > セキュリティ
# 2. 2段階認証を有効化
# 3. アプリパスワードを生成
        """)
    
    st.subheader("📂 監視フォルダの状態")
    if monitor_folder.exists():
        st.success(f"✅ フォルダが存在します: {monitor_folder}")
        
        # フォルダ内のファイル一覧
        target_files = []
        target_extensions = {'.pdf', '.docx', '.txt'}
        
        try:
            for file_path in monitor_folder.rglob('*'):
                if file_path.is_file() and file_path.suffix.lower() in target_extensions:
                    target_files.append({
                        'ファイル名': file_path.name,
                        '更新日時': datetime.fromtimestamp(file_path.stat().st_mtime).strftime('%Y-%m-%d %H:%M:%S'),
                        '拡張子': file_path.suffix,
                        'サイズ': f"{file_path.stat().st_size / 1024:.1f} KB"
                    })
            
            if target_files:
                st.write(f"対象ファイル: {len(target_files)}件")
                st.dataframe(target_files[:10])  # 最初の10件のみ表示
                if len(target_files) > 10:
                    st.info(f"（他 {len(target_files) - 10}件のファイルがあります）")
            else:
                st.info("対象ファイル（.pdf/.docx/.txt）は見つかりませんでした")
                
        except Exception as e:
            st.error(f"フォルダスキャンエラー: {e}")
    else:
        st.error(f"❌ 監視フォルダが見つかりません: {monitor_folder}")
        st.info("Windows環境以外では 'input/' フォルダを監視します")