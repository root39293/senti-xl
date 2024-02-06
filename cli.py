import pandas as pd
from openai import OpenAI

OPENAI_API_KEY="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"

client = OpenAI(
  api_key=OPENAI_API_KEY,
)

df = pd.read_excel("rawdata2.xlsx")
df['번호'] = df['번호'].astype(str)
df.dropna(subset=['주관식응답내용'], inplace=True)

def analyze_sentiment(text):
    try:
        completion = client.chat.completions.create(
        model="gpt-3.5-turbo",
        messages=[
            {"role": "system", "content": """ 

            #지시문
            너는 지금부터 감정분석 기계이다. 감정분석기계의 역할로서 #제약조건, #출력형태 를 준수하여 사용자가 입력하는 내용에 대해 감정분석을 하는것이 너의 역할이다. #예시를 참고하여 답하여라.

            #제약조건
            - 가급적 짧게 답한다.
            - 부연설명은 하지않는다.

            #출력형태
            [ 사용자의 입력내용에 따른 감정 긍정, 부정, 평이 출력 ]

            #예시
            친절하게 잘 알려주셨습니다.
            긍정

            검사관이 늦게오셨어요
            부정

            나쁘지않았습니다
            평이 """
             
             },
            {"role": "user", "content": f'#사용자의 입력 {text}'}
        ]
        )
        sentiment = completion.choices[0].message.content
        return sentiment
    except Exception as e:
        print(f"Error: {e}")
        return "분석 오류"

sentiment_results = []

for text in df['주관식응답내용']:
    sentiment = analyze_sentiment(text)
    sentiment_results.append(sentiment)

df['응답내용분석결과'] = sentiment_results

df.to_excel("감정분석결과.xlsx", index=False)
