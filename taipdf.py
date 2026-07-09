import requests

url = "https://dla.mplis.gov.vn/dc/Handlers/FileHandler.ashx"

params = {
    "DocId": "954585fb-a2bb-47c5-935d-54b2661c3a8e",
    "MimeType": "application/pdf"
}

headers = {
    "User-Agent": "Mozilla/5.0",
    "Referer": "https://dla.mplis.gov.vn/dc/",
}

# ===== DÁN COOKIE VÀO ĐÂY =====
cookie = """
ASP.NET_SessionId=kxfufebqhjutc5fz0b03qfy1; _Vbdlis.DC.Cookie=ULe5baUiiUP_sKZT-o7jT5Sp9kykHhThc28vDxxMhqIPqFzw7RQ2NVad2TBGJmSPmVrk9nCuk4rebTcYGP6MhnbqopYObGipS-ljR16fZ1eZbf-Aa6v15erqtjKa5KT-0tsCjZ3Ed7CaHz7Gjc4ksRyeZdpjiGlMvXOVIQLQ-4sE_jvUImVXgbl04pjF9qZ8iVUjlOeqyUxuvRduH1dgplAW8ld9kTU8wgUCO0jjKqDilK-OgXXFtcVHsOZcxfUClD7bjABNoYrD0yLyrg1cqtKMU8AuVxyzCP2-8i2yEA3h9RciPwESpmkyPIOrT0XuR097_X0Ky1Vu07iJaC0BT911Q7w63grTm6OJYwLJTQZ2zPUVmUVwmqYqzzIET2eOTSfRh4CWHCohhHSFflPkdafLswjL4g0ChKlzSup75pnHWRWfPLb6K4DoTtw0iQ2HqSxYSg; __RequestVerificationTokenDC=4zEaWdtYGj8kUiRHkHPCFKSdeP26hptk6kXhm7N5K-MJ4LyBeaSyh1_lFg0nxS3zUWSuF86Ah4igzFh6shbydTDRoGQ1
""".strip()

session = requests.Session()

# Chuyển chuỗi cookie thành requests cookies
for item in cookie.split(";"):
    if "=" in item:
        k, v = item.strip().split("=", 1)
        session.cookies.set(k, v)

r = session.get(
    url,
    params=params,
    headers=headers,
    allow_redirects=False
)

print("Status:", r.status_code)
print("Content-Type:", r.headers.get("Content-Type"))
print("Location:", r.headers.get("Location"))
print(r.text[:300] if "text" in r.headers.get("Content-Type", "") else "Binary")

if r.status_code == 200 and "application/pdf" in r.headers.get("Content-Type", ""):
    with open("HĐ.pdf", "wb") as f:
        f.write(r.content)
    print("HĐ.pdf")