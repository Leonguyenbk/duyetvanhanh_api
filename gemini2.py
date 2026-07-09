import os
import time
import random
from dotenv import load_dotenv
from google import genai
from google.genai import errors

load_dotenv()
client = genai.Client(api_key=os.getenv("GEMINI_API_KEY"))

PROMPT = """
Đây là văn bản tạm dừng, hãy kiểm tra lý do và tóm tắt, nếu không có lý do hoặc lý do không phù hợp hãy trả về không phù hợp và tóm tắt lý do không phù hợp
"""

RETRYABLE = {429, 500, 502, 503, 504}


def generate_with_retry(model, contents, max_retries=5):
    for attempt in range(max_retries):
        try:
            return client.models.generate_content(model=model, contents=contents)
        except errors.ServerError as e:
            code = getattr(e, "code", None)
            if code not in RETRYABLE or attempt == max_retries - 1:
                raise
            wait = (2 ** attempt) + random.uniform(0, 1)
            print(f"[{code}] Quá tải, thử lại sau {wait:.1f}s ({attempt + 1}/{max_retries})")
            time.sleep(wait)


# Upload 1 lần, dùng lại cho mọi lần retry
pdf = client.files.upload(file="test2.pdf")

response = generate_with_retry(
    model="gemini-3.5-flash",
    contents=[pdf, PROMPT],
)

print(response.text)