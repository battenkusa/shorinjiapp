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

# ── 提出資料監視 ────────────────────────
elif menu == "📬 提出資料監視":
    st.subheader("提出資料監視・メール送信")
    st.caption("共有フォルダの新しい資料を自動監視し、提出依頼文を生成します")
    
    import json
    import pdfplumber
    from docx import Document as DocxDocument
    from datetime import datetime, timedelta
    
    # 監視対象フォルダと出力フォルダのパス
    MONITOR_FOLDER = Path(r"C:\Users\batte\OneDrive\共有用")
    OUTPUT_FOLDER = Path(r"C:\Users\batte\OneDrive\少林寺\葛飾区連盟\AI作業フォルダ\提出関連資料")
    PROCESSED_FILES_JSON = Path("processed_files.json")
    
    # 処理済みファイルの記録を読み込み
    def load_processed_files():
        try:
            if PROCESSED_FILES_JSON.exists():
                with open(PROCESSED_FILES_JSON, 'r', encoding='utf-8') as f:
                    return json.load(f)
            return {}
        except Exception:
            return {}
    
    # 処理済みファイルの記録を保存
    def save_processed_files(processed_files):
        try:
            with open(PROCESSED_FILES_JSON, 'w', encoding='utf-8') as f:
                json.dump(processed_files, f, ensure_ascii=False, indent=2)
        except Exception as e:
            st.error(f"処理済みファイル記録の保存に失敗しました: {e}")
    
    # ファイル読み込み関数
    def read_document(file_path):
        try:
            file_path = Path(file_path)
            if file_path.suffix.lower() == '.pdf':
                with pdfplumber.open(file_path) as pdf:
                    text = "\n".join(
                        page.extract_text() for page in pdf.pages
                        if page.extract_text()
                    )
                return text
            elif file_path.suffix.lower() in ['.docx', '.doc']:
                doc = DocxDocument(file_path)
                text = "\n".join(paragraph.text for paragraph in doc.paragraphs)
                return text
            else:
                return None
        except Exception as e:
            st.error(f"ファイル読み込みエラー ({file_path.name}): {e}")
            return None
    
    # ファイル名にプレフィックスを追加して保存
    def rename_processed_file(file_path):
        try:
            file_path = Path(file_path)
            new_name = f"済_{file_path.name}"
            new_path = file_path.parent / new_name
            file_path.rename(new_path)
            return str(new_path)
        except Exception as e:
            st.error(f"ファイル名変更エラー: {e}")
            return None
    
    if st.button("📁 新しい資料をスキャン"):
        with st.spinner("フォルダをスキャン中..."):
            try:
                if not MONITOR_FOLDER.exists():
                    st.error(f"監視フォルダが見つかりません: {MONITOR_FOLDER}")
                    st.stop()
                
                # 処理済みファイルの記録を読み込み
                processed_files = load_processed_files()
                
                # フォルダ内のファイルをスキャン
                supported_extensions = {'.pdf', '.docx', '.doc'}
                new_files = []
                
                for file_path in MONITOR_FOLDER.rglob('*'):
                    if (file_path.is_file() and 
                        file_path.suffix.lower() in supported_extensions and
                        not file_path.name.startswith('済_')):
                        
                        file_key = str(file_path)
                        file_mtime = file_path.stat().st_mtime
                        
                        # 新しいファイルまたは更新されたファイルを検出
                        if (file_key not in processed_files or 
                            processed_files[file_key] != file_mtime):
                            new_files.append(file_path)
                
                if not new_files:
                    st.info("新しい資料は見つかりませんでした")
                else:
                    st.success(f"{len(new_files)}個の新しい資料を発見しました")
                    
                    # 各ファイルを処理
                    for file_path in new_files:
                        st.subheader(f"📄 {file_path.name}")
                        
                        # ファイル内容を読み込み
                        with st.spinner("ファイルを読み込み中..."):
                            document_text = read_document(file_path)
                            
                        if document_text:
                            st.info(f"読み込み完了（{len(document_text)}文字）")
                            
                            # AIで内容を分析
                            with st.spinner("AIが内容を分析中..."):
                                analysis_prompt = f"""以下の資料内容を分析して、以下の情報を抽出してください：

1. 概要（資料の主な内容）
2. 実施内容（具体的に何をするか）
3. 提出期限（いつまでに提出が必要か）

【資料内容】
{document_text[:2000]}

【出力形式】
概要: （概要を簡潔に）
実施内容: （実施内容を具体的に）
提出期限: （期限を明確に、日付がある場合は「YYYY年MM月DD日」形式で）
"""
                                
                                analysis_result = llm.invoke([HumanMessage(content=analysis_prompt)])
                                analysis_text = analysis_result.content
                            
                            st.markdown("### 📊 分析結果")
                            st.markdown(analysis_text)
                            
                            # 依頼文を生成
                            with st.spinner("依頼文を作成中..."):
                                request_prompt = f"""以下の分析結果をもとに、葛飾区少林寺拳法連盟のメンバーへの提出依頼文を作成してください。

【分析結果】
{analysis_text}

【依頼文の要件】
- 丁寧で分かりやすい文章
- 概要、実施内容、提出期限を含める
- 連盟内の提出期限は、資料上の提出期限の10日前に設定する
- 件名も含める

【出力形式】
件名: （メールの件名）

（依頼文の本文）
"""
                                
                                request_result = llm.invoke([HumanMessage(content=request_prompt)])
                                request_text = request_result.content
                            
                            st.markdown("### ✉️ 生成された依頼文")
                            st.markdown(request_text)
                            
                            # 依頼文を保存
                            try:
                                OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)
                                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                                output_filename = f"提出依頼_{file_path.stem}_{timestamp}.txt"
                                output_path = OUTPUT_FOLDER / output_filename
                                
                                with open(output_path, 'w', encoding='utf-8') as f:
                                    f.write(request_text)
                                
                                st.success(f"依頼文を保存しました: {output_filename}")
                                
                                # 元ファイルに「済_」プレフィックスを追加
                                renamed_path = rename_processed_file(file_path)
                                if renamed_path:
                                    st.info(f"元ファイルを「済_」付きで保存しました")
                                
                                # 処理済みファイルとして記録
                                processed_files[str(file_path)] = file_path.stat().st_mtime
                                
                            except Exception as e:
                                st.error(f"ファイル保存エラー: {e}")
                        
                        else:
                            st.warning(f"{file_path.name}の読み込みに失敗しました")
                    
                    # 処理済みファイル記録を保存
                    save_processed_files(processed_files)
                    
            except Exception as e:
                st.error(f"スキャン中にエラーが発生しました: {e}")
    
    st.markdown("---")
    st.markdown("### 📝 設定情報")
    st.info(f"**監視フォルダ:** {MONITOR_FOLDER}")
    st.info(f"**出力フォルダ:** {OUTPUT_FOLDER}")
    st.caption("対応形式: PDF (.pdf), Word (.docx, .doc)")
