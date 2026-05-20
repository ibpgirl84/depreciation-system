import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from io import BytesIO
import calendar

# =========================================
# 페이지 설정
# =========================================
st.set_page_config(
    page_title="감가상각 자동계산 시스템",
    page_icon="📊",
    layout="wide"
)

# =========================================
# 제목
# =========================================
st.title("📊 감가상각 자동계산 시스템")
st.write("실무용 유무형자산 감가상각 자동계산")

# =========================================
# 기준월 선택
# =========================================
st.subheader("기준월 선택")

month_list = [
    "1월", "2월", "3월", "4월", "5월", "6월",
    "7월", "8월", "9월", "10월", "11월", "12월"
]

today = datetime.today()

col1, col2 = st.columns(2)

with col1:

    selected_year = st.selectbox(
        "년도",
        list(range(today.year - 5, today.year + 6)),
        index=5
    )

with col2:

    selected_month_text = st.selectbox(
        "월",
        month_list,
        index=today.month - 1
    )

selected_month = month_list.index(selected_month_text) + 1

month_end_day = calendar.monthrange(
    selected_year,
    selected_month
)[1]

base_date = datetime(
    selected_year,
    selected_month,
    month_end_day
)

st.write(f"기준월 말일 : {base_date.strftime('%Y-%m-%d')}")

# =========================================
# 업로드
# =========================================
uploaded_file = st.file_uploader(
    "유무형자산 엑셀 업로드",
    type=["xlsx", "xls"]
)

# =========================================
# 숫자 포맷
# =========================================
def format_number(x):

    try:
        return f"{int(round(float(x),0)):,}"
    except:
        return x

# =========================================
# 감가상각 계산
# =========================================
if uploaded_file:

    try:

        excel_data = pd.read_excel(
            uploaded_file,
            sheet_name=None
        )

        total_result = []
        total_journal = []

        for sheet_name, df in excel_data.items():

            st.subheader(f"📁 시트 처리 : {sheet_name}")

            # 컬럼명 정리
            df.columns = [str(col).strip() for col in df.columns]

            required_cols = ["취득일", "품명", "취득가액"]

            missing_cols = []

            for col in required_cols:
                if col not in df.columns:
                    missing_cols.append(col)

            if missing_cols:
                st.warning(f"{sheet_name} 시트 컬럼 누락 : {missing_cols}")
                continue

            # 숫자 처리
            df["취득가액"] = pd.to_numeric(
                df["취득가액"],
                errors="coerce"
            ).fillna(0)

            # 내용연수 기본값
            if "내용연수" not in df.columns:
                df["내용연수"] = 5

            df["내용연수"] = pd.to_numeric(
                df["내용연수"],
                errors="coerce"
            ).fillna(5)

            # 잔존가액 기본값
            if "잔존가액" not in df.columns:
                df["잔존가액"] = 0

            df["잔존가액"] = pd.to_numeric(
                df["잔존가액"],
                errors="coerce"
            ).fillna(0)

            result_rows = []

            # =========================================
            # 자산별 계산
            # =========================================
            for idx, row in df.iterrows():

                try:

                    acquire_date = pd.to_datetime(row["취득일"])

                    amount = float(row["취득가액"])

                    life_years = float(row["내용연수"])

                    remain_value = float(row["잔존가액"])

                    monthly_dep = (
                        amount - remain_value
                    ) / (life_years * 12)

                    used_months = (
                        (base_date.year - acquire_date.year) * 12
                        + (base_date.month - acquire_date.month)
                    )

                    used_months = max(0, used_months)

                    accumulated = monthly_dep * used_months

                    if accumulated > amount:
                        accumulated = amount

                    book_value = amount - accumulated

                    # 미상각잔액 없으면 당월상각비 0
                    if book_value <= 0:
                        monthly_dep_display = 0
                        book_value = 0
                    else:
                        monthly_dep_display = monthly_dep

                    result_rows.append({
                        "계정명": sheet_name,
                        "품명": row["품명"],
                        "취득일": acquire_date.strftime("%Y-%m-%d"),
                        "취득가액": format_number(amount),
                        "당월감가상각비": format_number(monthly_dep_display),
                        "감가상각누계액": format_number(accumulated),
                        "미상각잔액": format_number(book_value)
                    })

                    # 전표
                    if monthly_dep_display > 0:

                        total_journal.append({
                            "계정명": sheet_name,
                            "차변계정": "감가상각비",
                            "대변계정": "감가상각누계액",
                            "금액": format_number(monthly_dep_display)
                        })

                except:
                    pass

            result_df = pd.DataFrame(result_rows)

            if len(result_df) > 0:

                st.dataframe(
                    result_df,
                    use_container_width=True
                )

                total_result.append(result_df)

        # =========================================
        # 전표
        # =========================================
        if len(total_journal) > 0:

            st.subheader("🧾 전표 생성")

            journal_df = pd.DataFrame(total_journal)

            st.dataframe(
                journal_df,
                use_container_width=True
            )

        # =========================================
        # 다운로드
        # =========================================
        output = BytesIO()

        with pd.ExcelWriter(
            output,
            engine="openpyxl"
        ) as writer:

            for i, df in enumerate(total_result):

                sheet = f"결과_{i+1}"

                df.to_excel(
                    writer,
                    sheet_name=sheet,
                    index=False
                )

            if len(total_journal) > 0:

                journal_df.to_excel(
                    writer,
                    sheet_name="전표",
                    index=False
                )

        st.download_button(
            label="📥 결과 엑셀 다운로드",
            data=output.getvalue(),
            file_name="감가상각_결과.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:

        st.error(str(e))