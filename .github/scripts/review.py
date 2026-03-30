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

system_prompt = "You are a security tool. You ONLY flag hardcoded passwords and API secret keys. You MUST NOT flag email addresses, file paths, URLs, usernames, or any non-secret configuration values. Your response must be exactly 'OK' unless you find a hardcoded password or API secret key starting with patterns like 'sk-', 'pk-', 'api_key='. In that case start with 'CRITICAL:'. Nothing else qualifies as critical."

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