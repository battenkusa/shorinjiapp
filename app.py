import streamlit as st
from pathlib import Path
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
elif menu == "📬 提出資料監視":
    import os, json, smtplib, datetime, re
    from email.mime.text import MIMEText
    from email.mime.multipart import MIMEMultipart
    from pathlib import Path

    st.subheader("提出資料監視・メール自動送信")
    st.caption("フォルダに新しいファイルが追加されると自動で分析してメールします")

    WATCH_FOLDER = r"C:\Users\batte\OneDrive\少林寺\葛飾区連盟\AI作業フォルダ\提出関連資料"
    TO_EMAIL     = "battenkusayokayoka@gmail.com"
    HISTORY_FILE = r"C:\Users\batte\myagent\checked_files.json"

    gmail_address  = st.text_input("送信元Gmailアドレス")
    gmail_password = st.text_input("Gmailアプリパスワード（16桁）", type="password")

    def load_history():
        if Path(HISTORY_FILE).exists():
            return json.loads(Path(HISTORY_FILE).read_text(encoding="utf-8"))
        return []

    def save_history(files):
        Path(HISTORY_FILE).write_text(
            json.dumps(files, ensure_ascii=False, indent=2), encoding="utf-8"
        )

    def read_file(path: Path) -> str:
        if path.suffix.lower() == ".pdf":
            import pdfplumber
            with pdfplumber.open(str(path)) as pdf:
                return "\n".join(p.extract_text() for p in pdf.pages if p.extract_text())
        elif path.suffix.lower() == ".docx":
            from docx import Document as DocxDocument
            doc = DocxDocument(str(path))
            return "\n".join(p.text for p in doc.paragraphs if p.text)
        elif path.suffix.lower() == ".txt":
            return path.read_text(encoding="utf-8")
        return ""

    def analyze_document(text: str) -> dict:
        prompt = f"""以下の資料を分析してJSON形式のみで返してください。
資料内容：{text[:3000]}
{{
  "概要": "資料の概要を2〜3文で",
  "実施内容": "やるべきことを箇条書きで",
  "提出期限": "YYYY年MM月DD日形式、不明なら空文字"
}}"""
        res = llm.invoke([HumanMessage(content=prompt)])
        try:
            json_str = re.search(r'\{.*\}', res.content, re.DOTALL).group()
            return json.loads(json_str)
        except Exception:
            return {"概要": res.content[:200], "実施内容": "", "提出期限": ""}

    def is_expired(deadline_str: str) -> bool:
        try:
            nums = re.findall(r'\d+', deadline_str)
            if len(nums) >= 3:
                deadline = datetime.date(int(nums[0]), int(nums[1]), int(nums[2]))
                return deadline < datetime.date.today()
        except Exception:
            pass
        return False

    def calc_deadline(deadline_str: str):
        try:
            nums = re.findall(r'\d+', deadline_str)
            if len(nums) >= 3:
                dt     = datetime.date(int(nums[0]), int(nums[1]), int(nums[2]))
                before = dt - datetime.timedelta(days=10)
                return dt.strftime("%Y年%m月%d日"), before.strftime("%Y年%m月%d日")
        except Exception:
            pass
        return deadline_str, "（計算できませんでした）"

    def create_email_body(filename: str, analysis: dict) -> str:
        deadline, before_deadline = calc_deadline(analysis.get("提出期限", ""))
        expired = is_expired(analysis.get("提出期限", ""))
        expired_header = "⚠️ この資料は期限切れです ⚠️\n\n" if expired else ""
        expired_mark   = "（期限切れ）" if expired else ""
        return f"""葛飾区少林寺拳法連盟 関係者各位

{expired_header}【ファイル名】
{filename}

【概要】
{analysis.get("概要", "")}

【実施内容】
{analysis.get("実施内容", "")}

【期限】
・依頼内容締切：{before_deadline}{expired_mark}
・提出期限　　：{deadline}{expired_mark}

よろしくお願いいたします。
※このメールはAIが自動生成しました。
葛飾区少林寺拳法連盟 AIチーム"""

    def send_email(subject: str, body: str, gmail_addr: str, gmail_pass: str):
        msg = MIMEMultipart()
        msg["From"]    = gmail_addr
        msg["To"]      = TO_EMAIL
        msg["Subject"] = subject
        msg.attach(MIMEText(body, "plain", "utf-8"))
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(gmail_addr, gmail_pass)
            server.send_message(msg)

    if "auto_checked" not in st.session_state:
        st.session_state.auto_checked = False

    if not st.session_state.auto_checked:
        st.session_state.auto_checked = True
        watch_path = Path(WATCH_FOLDER)
        if watch_path.exists():
            history   = load_history()
            all_files = [
                f.name for f in watch_path.iterdir()
                if f.suffix.lower() in [".pdf", ".docx", ".txt"]
            ]
            new_files = [f for f in all_files if f not in history]
            if new_files:
                st.info(f"新しいファイルが {len(new_files)} 件見つかりました：{', '.join(new_files)}")
                st.session_state.new_files = new_files
            else:
                st.success("新しいファイルはありません")
                st.session_state.new_files = []
        else:
            st.warning(f"フォルダが見つかりません：{WATCH_FOLDER}")
            st.session_state.new_files = []

    if st.button("今すぐフォルダを確認する"):
        st.session_state.auto_checked = False
        st.rerun()

    new_files = st.session_state.get("new_files", [])
    if new_files and gmail_address and gmail_password:
        if st.button(f"新ファイル {len(new_files)} 件を分析してメール送信する"):
            watch_path = Path(WATCH_FOLDER)
            history    = load_history()
            for filename in new_files:
                file_path = watch_path / filename
                with st.spinner(f"{filename} を分析中..."):
                    text     = read_file(file_path)
                    analysis = analyze_document(text)
                    expired  = is_expired(analysis.get("提出期限", ""))
                    subject  = f"【期限切れ】{filename}" if expired else f"【提出資料のご連絡】{filename}"
                    body     = create_email_body(filename, analysis)
                try:
                    send_email(subject, body, gmail_address, gmail_password)
                    st.success(f"✅ {filename} → メール送信完了")
                except Exception as e:
                    st.error(f"❌ {filename} → 送信失敗：{e}")
                history.append(filename)
            save_history(history)
    elif new_files and not (gmail_address and gmail_password):
        st.warning("GmailアドレスとアプリパスワードをUI上に入力してください")