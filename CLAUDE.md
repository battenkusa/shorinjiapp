\# CLAUDE.md



このファイルはClaude Codeがshorinjiappリポジトリで作業する際の指示です。



\## プロジェクト概要



葛飾区少林寺拳法連盟の団体活動支援AIツール。

Streamlit + LangChain + Ollama（ローカルLLM）で構成。



\## セットアップ

```bash

cd C:\\Users\\batte\\myagent

myagent\\Scripts\\activate

pip install -r requirements.txt

streamlit run app.py

```



\## アーキテクチャ



\- `app.py`: Streamlit UI（議事録・連絡文・お知らせ・計画・大会資料）

\- `input/`: 自動処理対象フォルダ（.txtを置くと自動判断）

\- `output/`: 生成ファイルの保存先



<important if="コードを書くとき">



\## セキュリティルール（必ず守ること）



\- APIキー・パスワードは絶対にコードに直書きしない。必ず環境変数から読む

\- ユーザー入力をそのまま subprocess やファイルパスに使わない

\- ファイルパスは Path() を使い、input/ output/ フォルダ外へのアクセスを禁止する

\- エラーメッセージに内部パス・APIキー・スタックトレースを含めない

\- .env ファイルは .gitignore に含まれているため絶対にコミットしない



</important>



<important if="新機能を追加するとき">



\## コーディング規則



\- 既存のメニュー構造（sidebar.radio）に従い elif で追加する

\- 日本語の文字列は UTF-8 で扱い、ファイル保存時も encoding="utf-8" を明示する

\- Streamlit の st.spinner() で処理中表示、st.success() で完了表示を統一する

\- PDF読み込みは pdfplumber、Word出力は python-docx を使う（他ライブラリ不可）

\- LLM呼び出しは必ず ChatOllama(model="llama3.2") を使う（外部APIは使わない）



</important>



\## 禁止事項



\- OpenAI・Anthropic等の外部LLM APIをapp.pyで直接呼び出すこと（Ollama限定）

\- requirements.txt にないライブラリを勝手に追加すること

\- input/ output/ 以外のフォルダへのファイル書き込み



\## タスク進め方



1\. 複雑な変更は必ずplanモードで設計してから実装する

2\. 既存機能を壊さないよう、変更前に該当箇所を確認する

3\. 日本語のUIテキストは変更しない



\## デバッグ



問題が発生した場合は以下を確認：

\- Ollamaが起動しているか（`ollama run llama3.2`）

\- 仮想環境が有効か（`myagent\\Scripts\\activate`）

\- 依存ライブラリが入っているか（`pip install -r requirements.txt`）

