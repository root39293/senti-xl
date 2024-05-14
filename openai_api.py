from openai import OpenAI
import os

OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
SYSTEM_PROMPT = os.getenv("SYSTEM_PROMPT")

client = OpenAI(api_key=OPENAI_API_KEY)

def analyze_sentiment(text):
    try:
        completion = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": f'사용자의 입력 : {text}'}
            ],
            temperature=0.3
        )
        sentiment = completion.choices[0].message.content.strip()
        return sentiment
    except Exception as e:
        print(f"Error: {e}")
        return "분석 오류"
