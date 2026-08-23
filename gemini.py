from dotenv import load_dotenv
from google import genai
import os

load_dotenv()

client = genai.Client(
    api_key=os.getenv("GEMINI_API_KEY")
)

pdf = client.files.upload(file="VBTD.pdf")

response = client.models.generate_content(
    model="gemini-3-flash",
    contents=[
        pdf,
        """
Đây là văn bản tạm dừng, hãy kiểm tra lý do và tóm tắt, nếu không có lý do hoặc lý do không phù hợp hãy trả về không phù hợp và tóm tắt lý do không phù hợp, trường
hợp không có chữ ký và con dấu thì coi như không phù hợp"""
    ]
)

print(response.text)