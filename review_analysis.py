import re
from collections import Counter
import pandas as pd
import matplotlib.pyplot as plt


# =========================
# 1. 설정
# =========================
EXCEL_FILE = "reviews.xlsx"   # 분석할 엑셀 파일명
SHEET_NAME = 0                # 첫 번째 시트 사용
TOP_N = 10                    # 상위 키워드 개수

# 불용어(분석에서 제외할 단어)
STOPWORDS = {
    "진짜", "정말", "너무", "아주", "그냥", "약간", "조금", "완전",
    "그리고", "근데", "그런데", "그래서", "하지만", "또한",
    "사용", "사용했어요", "사용합니다", "사용했는데", "느낌",
    "제품", "상품", "구매", "재구매", "선물", "배송", "포장",
    "좋아요", "좋습니다", "좋았어요", "만족", "추천", "괜찮아요",
    "같아요", "있어요", "없어요", "했어요", "되었어요", "입니다",
    "이거", "저거", "그거", "이건", "저는", "제가", "하나",
    "정도", "부분", "때문", "이번", "계속", "바로", "처음",
    "무난", "보통", "살짝", "확실히", "요즘", "하루", "매우"
}


# =========================
# 2. 데이터 불러오기
# =========================
def load_reviews(file_path, sheet_name=0):
    df = pd.read_excel(file_path, sheet_name=sheet_name)

    # 컬럼명 공백 제거
    df.columns = [str(col).strip() for col in df.columns]

    required_cols = ["날짜", "리뷰", "주제"]
    for col in required_cols:
        if col not in df.columns:
            raise ValueError(f"엑셀에 '{col}' 열이 없습니다. 현재 열: {list(df.columns)}")

    # 리뷰 결측 제거
    df = df.dropna(subset=["리뷰"]).copy()
    df["리뷰"] = df["리뷰"].astype(str).str.strip()
    df = df[df["리뷰"] != ""]

    # 날짜 변환
    df["날짜"] = pd.to_datetime(df["날짜"], errors="coerce")

    return df


# =========================
# 3. 텍스트 전처리
# =========================
def clean_text(text):
    text = str(text)

    # 이모지 / 특수기호 / 숫자 일부 제거
    text = re.sub(r"[^\w\s가-힣]", " ", text)
    text = re.sub(r"\d+", " ", text)
    text = re.sub(r"_", " ", text)

    # 공백 정리
    text = re.sub(r"\s+", " ", text).strip()

    return text


def extract_keywords_from_reviews(reviews, stopwords=None, min_len=2):
    if stopwords is None:
        stopwords = set()

    words = []

    for review in reviews:
        cleaned = clean_text(review)
        tokens = cleaned.split()

        for token in tokens:
            token = token.strip()
            if len(token) < min_len:
                continue
            if token in stopwords:
                continue
            words.append(token)

    return Counter(words)


# =========================
# 4. 분석 함수
# =========================
def analyze_keywords(df, top_n=10):
    counter = extract_keywords_from_reviews(df["리뷰"], stopwords=STOPWORDS)
    return counter.most_common(top_n)


def analyze_topics(df):
    topic_series = df["주제"].fillna("미분류").astype(str).str.strip()
    topic_series = topic_series.replace("", "미분류")
    return topic_series.value_counts()


def analyze_monthly_count(df):
    monthly = df.dropna(subset=["날짜"]).copy()
    monthly["연월"] = monthly["날짜"].dt.to_period("M").astype(str)
    return monthly["연월"].value_counts().sort_index()


# =========================
# 5. 결과 저장
# =========================
def save_results(keyword_result, topic_result, monthly_result):
    keyword_df = pd.DataFrame(keyword_result, columns=["키워드", "빈도"])
    topic_df = topic_result.reset_index()
    topic_df.columns = ["주제", "리뷰 수"]

    monthly_df = monthly_result.reset_index()
    monthly_df.columns = ["연월", "리뷰 수"]

    with pd.ExcelWriter("review_analysis_result.xlsx", engine="openpyxl") as writer:
        keyword_df.to_excel(writer, sheet_name="키워드_TOP10", index=False)
        topic_df.to_excel(writer, sheet_name="주제별_리뷰수", index=False)
        monthly_df.to_excel(writer, sheet_name="월별_리뷰수", index=False)

    print("\n분석 결과가 'review_analysis_result.xlsx' 파일로 저장되었습니다.")


# =========================
# 6. 시각화
# =========================
def plot_keyword_chart(keyword_result):
    if not keyword_result:
        print("시각화할 키워드 결과가 없습니다.")
        return

    keywords = [k for k, v in keyword_result]
    counts = [v for k, v in keyword_result]

    plt.figure(figsize=(10, 6))
    plt.bar(keywords, counts)
    plt.title("리뷰 키워드 TOP 10")
    plt.xlabel("키워드")
    plt.ylabel("빈도")
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.show()


def plot_topic_chart(topic_result):
    if topic_result.empty:
        print("시각화할 주제 결과가 없습니다.")
        return

    plt.figure(figsize=(10, 6))
    plt.bar(topic_result.index.astype(str), topic_result.values)
    plt.title("주제별 리뷰 수")
    plt.xlabel("주제")
    plt.ylabel("리뷰 수")
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.show()


# =========================
# 7. 메인 실행
# =========================
def main():
    try:
        df = load_reviews(EXCEL_FILE, SHEET_NAME)

        print(f"총 리뷰 수: {len(df)}")

        keyword_result = analyze_keywords(df, top_n=TOP_N)
        topic_result = analyze_topics(df)
        monthly_result = analyze_monthly_count(df)

        print("\n[키워드 TOP 10]")
        for word, count in keyword_result:
            print(f"- {word}: {count}")

        print("\n[주제별 리뷰 수]")
        for topic, count in topic_result.items():
            print(f"- {topic}: {count}")

        print("\n[월별 리뷰 수]")
        for month, count in monthly_result.items():
            print(f"- {month}: {count}")

        save_results(keyword_result, topic_result, monthly_result)

        plot_keyword_chart(keyword_result)
        plot_topic_chart(topic_result)

    except Exception as e:
        print(f"오류 발생: {e}")


if __name__ == "__main__":
    main()
