from openai import OpenAI
from dotenv import load_dotenv
import os

load_dotenv()

client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

# Upload PDF
uploaded = client.files.create(
    file=open("test2.pdf", "rb"),
    purpose="user_data"
)

# Đọc PDF
response = client.responses.create(
    model="gpt-5.5",
    input=[
        {
            "role": "user",
            "content": [
                {
                    "type": "input_file",
                    "file_id": uploaded.id,
                },
                {
                    "type": "input_text",
                    "text": """
Đây là văn bản tạm dừng, hãy kiểm tra lý do và tóm tắt, nếu không có lý do hoặc lý do không phù hợp hãy trả về không phù hợp và tóm tắt lý do không phù hợp"""

                }
            ]
        }
    ]
)

print(response.output_text)