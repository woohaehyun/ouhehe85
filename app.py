import streamlit as st
import pandas as pd
import io
import zipfile
from datetime import datetime, timedelta
import xlsxwriter

# 페이지 설정
st.set_page_config(page_title="신명약품 발주서 생성 시스템", layout="wide")

# 상단 로고 & 타이틀
col1, col2 = st.columns([1, 5])
with col1:
    st.image("로고리뉴얼.png", width=100)
with col2:
    st.title("💊 신명약품 발주서 생성 시스템")
st.markdown("매입처별 발주서를 자동 생성하고, 조건별 필터링 후 Excel 파일로 다운로드하세요.")

# 파일 업로드
st.sidebar.header("📂 파일 업로드")
sales_file = st.sidebar.file_uploader("매출자료 업로드", type=["xlsx"])
purchase_file = st.sidebar.file_uploader("매입자료 업로드", type=["xlsx"])
stock_file = st.sidebar.file_uploader("현재고 업로드", type=["xlsx"])

# 분석 모드 선택
mode = st.sidebar.radio("📅 분석 모드 선택", ["자동 모드 (최근 3개월)", "수동 모드"])

# 표준 컬럼명 매핑
def normalize_columns(df, mapping):
    df.rename(columns={k: v for k, v in mapping.items() if k in df.columns}, inplace=True)
    return df

# 필수 컬럼 체크
def check_required_columns(df, required, name):
    missing = [col for col in required if col not in df.columns]
    if missing:
        st.error(f"{name}에 다음 컬럼이 없습니다: {', '.join(missing)}")
        st.stop()

if sales_file and purchase_file and stock_file:
    # 파일 읽기
    sales_df = pd.read_excel(sales_file)
    purchase_df = pd.read_excel(purchase_file)
    stock_df = pd.read_excel(stock_file)

    # 컬럼명 표준화
    sales_df = normalize_columns(sales_df, {
        "거래일자": "명세일자", "일자": "명세일자",
        "매출처": "매 출 처", "상품명": "상 품 명",
        "포장 단위": "포장단위"
    })
    purchase_df = normalize_columns(purchase_df, {
        "입고일": "입고일자", "거래처": "매 입 처",
        "상품명": "상 품 명", "포장 단위": "포장단위",
        "매입처": "매 입 처", "제조사": "제 조 사",
        "단가": "매입단가", "매입 단가": "매입단가"
    })
    stock_df = normalize_columns(stock_df, {
        "거래처": "매 입 처", "상품명": "상 품 명",
        "포장 단위": "포장단위", "제조사": "제 조 사",
        "재고": "재고수량", "단가": "매입단가", "매입 단가": "매입단가"
    })

    # 필수 컬럼 체크
    check_required_columns(sales_df, ["명세일자", "매 출 처", "상 품 명", "포장단위", "수량", "매출단가"], "매출자료")
    check_required_columns(purchase_df, ["입고일자", "매 입 처", "상 품 명", "제 조 사", "수량", "매입단가"], "매입자료")
    check_required_columns(stock_df, ["매 입 처", "제 조 사", "상 품 명", "포장단위", "재고수량", "매입단가"], "현재고")

    # 병합 키 표준화
    for df in [sales_df, purchase_df, stock_df]:
        df["상 품 명"] = df["상 품 명"].astype(str).str.strip().str.upper()
        df["포장단위"] = df["포장단위"].astype(str).str.strip().str.upper()

    # 날짜 변환
    sales_df["명세일자"] = pd.to_datetime(sales_df["명세일자"], errors="coerce")
    purchase_df["입고일자"] = pd.to_datetime(purchase_df["입고일자"], errors="coerce")

    # 기간 필터
    if mode == "자동 모드 (최근 3개월)":
        end_date = sales_df["명세일자"].max()
        start_date = end_date - pd.DateOffset(months=3)
        filtered_sales = sales_df[(sales_df["명세일자"] >= start_date) & (sales_df["명세일자"] <= end_date)]
    else:
        start_date = st.sidebar.date_input("시작일", value=sales_df["명세일자"].min().date())
        end_date = st.sidebar.date_input("종료일", value=sales_df["명세일자"].max().date())
        filtered_sales = sales_df[(sales_df["명세일자"] >= pd.to_datetime(start_date)) &
                                  (sales_df["명세일자"] <= pd.to_datetime(end_date))]

    # 매입처, 거래처, 마진율 필터
    suppliers = sorted(purchase_df["매 입 처"].dropna().unique())
    customers = sorted(sales_df["매 출 처"].dropna().unique())
    selected_suppliers = st.sidebar.multiselect("매입처 선택", suppliers)
    selected_customers = st.sidebar.multiselect("거래처 선택", customers)
    margin_options = list(range(1, 101))
    selected_margins = st.sidebar.multiselect("마진율% 선택", margin_options)

    if selected_suppliers:
        purchase_df = purchase_df[purchase_df["매 입 처"].isin(selected_suppliers)]
        stock_df = stock_df[stock_df["매 입 처"].isin(selected_suppliers)]
    if selected_customers:
        filtered_sales = filtered_sales[filtered_sales["매 출 처"].isin(selected_customers)]

    # 전월 판매량 계산
    last_month_end = sales_df["명세일자"].max().replace(day=1) - timedelta(days=1)
    last_month_start = last_month_end.replace(day=1)
    last_month_sales = sales_df[(sales_df["명세일자"] >= last_month_start) &
                                (sales_df["명세일자"] <= last_month_end)]
    last_month_qty = last_month_sales.groupby(["상 품 명", "포장단위"], as_index=False)["수량"].sum()
    last_month_qty.rename(columns={"수량": "전월판매량"}, inplace=True)

    # 데이터 병합
    merged = pd.merge(last_month_qty, stock_df, on=["상 품 명", "포장단위"], how="left")
    merged["과재고"] = (merged["재고수량"] - merged["전월판매량"]).apply(lambda x: x if x > 0 else 0)
    merged["부족수량"] = (merged["전월판매량"] - merged["재고수량"]).apply(lambda x: x if x > 0 else 0)
    merged["발주수량"] = merged["부족수량"]

    # 매입자료 병합
    merged = pd.merge(merged,
                      purchase_df[["상 품 명", "포장단위", "매 입 처", "제 조 사", "매입단가"]].drop_duplicates(),
                      on=["상 품 명", "포장단위"], how="left")

    # 매입단가 안전 처리
    if "매입단가" in merged.columns:
        merged["매입단가"] = merged["매입단가"].fillna(0)
    else:
        merged["매입단가"] = 0

    # 매출단가 병합
    merged = pd.merge(merged,
                      sales_df[["상 품 명", "포장단위", "매출단가"]].drop_duplicates(),
                      on=["상 품 명", "포장단위"], how="left")

    # 매입처 안전 처리
    if "매 입 처" in merged.columns:
        merged["매 입 처"] = merged["매 입 처"].fillna("미지정")
    else:
        merged["매 입 처"] = "미지정"

    # 금액·마진율 계산
    merged["합계금액"] = merged["발주수량"] * merged["매입단가"]
    merged["마진율"] = ((merged["매출단가"] - merged["매입단가"]) / merged["매출단가"] * 100).round(1)

    # 마진율 필터 적용
    if selected_margins:
        merged = merged[merged["마진율"].isin(selected_margins)]

    # 미리보기
    st.subheader("📊 발주서 데이터 미리보기")
    st.dataframe(merged)

    # ZIP 다운로드
    if st.button("📦 발주서 ZIP 다운로드"):
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zipf:
            for supplier, group in merged.groupby("매 입 처"):
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                    group.to_excel(writer, index=False, sheet_name="발주서")
                    workbook = writer.book
                    worksheet = writer.sheets["발주서"]
                    worksheet.merge_range("A1:K1", "신명약품 발주서",
                                          workbook.add_format({"bold": True, "align": "center", "valign": "vcenter", "font_size": 16}))
                    worksheet.write("A2", "담당자: __________")
                    worksheet.write("E2", f"발주일: {datetime.today().strftime('%Y-%m-%d')}")
                    worksheet.write("I2", "대표이사 결재 [          ]")
                zipf.writestr(f"{supplier}_발주서.xlsx", output.getvalue())
        zip_buffer.seek(0)
        st.download_button("📥 ZIP 파일 다운로드", data=zip_buffer,
                           file_name="발주서_엑셀.zip", mime="application/zip")

else:
    st.warning("📂 사이드바에서 매출, 매입, 현재고 파일을 업로드하세요.")
