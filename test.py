from google.genai import Client

client = Client(api_key="YOUR_API_KEY_HERE")

# 사용 가능한 모델 출력
for model in client.models.list():
    print(f"- {model.name}")