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

res = client.messages.create(
    model="claude-haiku-4-5-20251001",
    max_tokens=1024,
    system="""You are a secure coding expert. Review the code diff for CRITICAL issues only:
1. Hardcoded API keys, passwords, or secrets (NOT file paths or email addresses)
2. SQL injection or command injection vulnerabilities  
3. Path traversal risks from EXTERNAL user input (NOT hardcoded internal paths)

Do NOT flag:
- Hardcoded file paths (these are configuration values)
- Hardcoded email addresses (these are configuration values)
- Windows-specific paths

Only output 'CRITICAL:' if you find API keys or passwords hardcoded, or injection vulnerabilities.
If no critical issues found, just output 'OK'.""",

Only output 'CRITICAL:' if you find one of the above serious issues.
For minor style issues or suggestions, ignore them.
If no critical issues found, just output 'OK'.""",
    messages=[
        {
            "role": "user",
            "content": f"Review this diff for critical security issues only:\n\n{diff[:4000]}"
        }
    ]
)

review = res.content[0].text
print("===== Security Review =====")
print(review)
print("===========================")

if "CRITICAL:" in review:
    print("Review NG - Critical security issue found")
    sys.exit(1)

print("Review OK")
sys.exit(0)