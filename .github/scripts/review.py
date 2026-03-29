import openai
import subprocess
import sys

try:
    diff = subprocess.check_output(
        ["git", "diff", "HEAD~1", "HEAD"],
        stderr=subprocess.DEVNULL
    ).decode("utf-8")
except subprocess.CalledProcessError:
    diff = subprocess.check_output(
        ["git", "show", "HEAD"]
    ).decode("utf-8")

if not diff.strip():
    print("No diff. Skipping.")
    sys.exit(0)

client = openai.OpenAI()

res = client.chat.completions.create(
    model="gpt-4o",
    messages=[
        {
            "role": "system",
            "content": "You are a secure coding expert. Review the code diff for: 1.Hardcoded API keys or passwords 2.Injection vulnerabilities 3.Path traversal risks 4.Missing input validation 5.Sensitive info in error messages 6.Excessive permissions 7.Obvious bugs 8.Missing exception handling. If any problem found, start with 'PROBLEM:' and describe it. If no problem, just say 'OK'."
        },
        {
            "role": "user",
            "content": f"Review this diff:\n\n{diff[:4000]}"
        }
    ]
)

review = res.choices[0].message.content
print("===== Security Review =====")
print(review)
print("===========================")

if "PROBLEM:" in review:
    print("Review NG")
    sys.exit(1)

print("Review OK")
sys.exit(0)