import subprocess
import sys
import anthropic

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

client = anthropic.Anthropic()

system_prompt = "You are a secure coding expert. Review the code diff for CRITICAL issues only. Flag ONLY these: 1. Hardcoded API keys or passwords (NOT file paths or emails). 2. SQL injection or command injection from external user input. Do NOT flag hardcoded file paths or email addresses. If critical issues found start with CRITICAL: otherwise just say OK."

res = client.messages.create(
    model="claude-haiku-4-5-20251001",
    max_tokens=1024,
    system=system_prompt,
    messages=[
        {
            "role": "user",
            "content": "Review this diff:\n\n" + diff[:4000]
        }
    ]
)

review = res.content[0].text
print("===== Security Review =====")
print(review)
print("===========================")

if "CRITICAL:" in review:
    print("Review NG")
    sys.exit(1)

print("Review OK")
sys.exit(0)