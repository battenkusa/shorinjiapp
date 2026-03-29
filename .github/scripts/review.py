import openai
import subprocess
import sys

diff = subprocess.check_output(["git", "diff", "HEAD~1"]).decode("utf-8")

if not diff.strip():
    print("差分なし。スキップします。")
    sys.exit(0)

client = openai.OpenAI()

res = client.chat.completions.create(
    model="gpt-4o",
    messages=[
        {
            "role": "system",
            "content": """あなたはセキュアコーディングの専門家です。
以下の観点でコードを厳しくレビューしてください。

【セキュリティ確認項目】
1. APIキー・パスワードがコードに直書きされていないか
2. ユーザー入力をそのまま使っていないか（インジェクション）
3. ファイルパスに任意の値を使っていないか（パストラバーサル）
4. 外部からの入力を検証しているか
5. エラーメッセージに内部情報が含まれていないか
6. 不要な権限を持っていないか

【コード品質確認項目】
7. 明らかなバグやエラーがないか
8. 日本語文字列が正しく扱われているか
9. 例外処理が適切にされているか

問題があれば必ず「問題あり：」で始めて、
該当箇所と修正方法を具体的に指摘してください。
問題がなければ「問題なし」とだけ答えてください。"""
        },
        {
            "role": "user",
            "content": f"以下のコード差分をレビューしてください。\n\n{diff[:4000]}"
        }
    ]
)

review = res.choices[0].message.content
print("===== Codex セキュリティレビュー結果 =====")
print(review)
print("==========================================")

if "問題あり" in review:
    print("レビューNG：セキュリティまたは品質の問題が検出されました")
    sys.exit(1)

print("レビューOK：問題なし")
sys.exit(0)