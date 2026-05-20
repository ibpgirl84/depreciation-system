import streamlit as st
import pandas as pd
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
base_date = st.date_input(
    "기준월 선택",
    datetime.today()
)

# 기준월 말일 계산
last_day = calendar.monthrange(
    base_date.year,
    base_date.month
)[1]

base_month_end = datetime(
    base_date.year,
    base_date.month,
    last_day
)

st.write(
    f"기준월 말일 : {base_month_end.strftime('%Y-%m-%d')}"
)

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
        return f"{int(round(float(x), 0)):,}"
    except:
        return x

# =========================================
# 회계정책 내용연수
# =========================================
life_mapping = {

    "차량": 4,              # 48개월
    "비품": 5,
    "시설장치": 5,
    "소프트웨어": 5,
    "상표권": 5,
    "기타무형자산": 5
}

# =========================================
# 전표 계정 매핑
# =========================================
account_mapping = {

    "차량": {
        "차변": "유형자산상각비/차량운반구",
        "대변": "차량운반구감가상각누계액"
    },

    "비품": {
        "차변": "유형자산상각비/비품",
        "대변": "비품감가상각누계액"
    },

    "시설장치": {
        "차변": "유형자산상각비/시설장치",
        "대변": "시설장치감가상각누계액"
    },

    "소프트웨어": {
        "차변": "무형자산상각비/소프트웨어",
        "대변": "소프트웨어"
    },

    "상표권": {
        "차변": "무형자산상각비/상표권",
        "대변": "상표권"
    },

    "기타무형자산": {
        "차변": "무형자산상각비/기타무형자산",
        "대변": "기타무형자산"
    }
}

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

        # 전표 합산용
        journal_summary = {}

        # =========================================
        # 시트별 처리
        # =========================================
        for sheet_name, df in excel_data.items():

            st.subheader(f"📁 시트 처리 : {sheet_name}")

            # 컬럼명 공백 제거
            df.columns = [
                str(col).strip()
                for col in df.columns
            ]

            required_cols = [
                "취득일",
                "품명",
                "취득가액"
            ]

            missing_cols = []

            for col in required_cols:

                if col not in df.columns:
                    missing_cols.append(col)

            # 필수 컬럼 누락
            if missing_cols:

                st.warning(
                    f"{sheet_name} 시트 컬럼 누락 : {missing_cols}"
                )

                continue

            # =========================================
            # 숫자 처리
            # =========================================
            df["취득가액"] = pd.to_numeric(
                df["취득가액"],
                errors="coerce"
            ).fillna(0)

            # 잔존가액
            if "잔존가액" not in df.columns:
                df["잔존가액"] = 0

            df["잔존가액"] = pd.to_numeric(
                df["잔존가액"],
                errors="coerce"
            ).fillna(0)

            # =========================================
            # 회계정책 내용연수 적용
            # =========================================
            if "내용연수" not in df.columns:

                df["내용연수"] = life_mapping.get(
                    sheet_name,
                    5
                )

            df["내용연수"] = pd.to_numeric(
                df["내용연수"],
                errors="coerce"
            ).fillna(
                life_mapping.get(sheet_name, 5)
            )

            result_rows = []

            # =========================================
            # 자산별 계산
            # =========================================
            for idx, row in df.iterrows():

                try:

                    acquire_date = pd.to_datetime(
                        row["취득일"]
                    )

                    amount = float(
                        row["취득가액"]
                    )

                    life_years = float(
                        row["내용연수"]
                    )

                    remain_value = float(
                        row["잔존가액"]
                    )

                    # 월 감가상각비
                    monthly_dep = (
                        amount - remain_value
                    ) / (life_years * 12)

                    # 사용개월수
                    used_months = (
                        (base_month_end.year - acquire_date.year) * 12
                        + (
                            base_month_end.month
                            - acquire_date.month
                        )
                        + 1
                    )

                    used_months = max(
                        0,
                        used_months
                    )

                    # 최대개월수
                    max_months = int(
                        life_years * 12
                    )

                    used_months = min(
                        used_months,
                        max_months
                    )

                    # 감가상각누계액
                    accumulated = (
                        monthly_dep * used_months
                    )

                    # 미상각잔액
                    book_value = (
                        amount - accumulated
                    )

                    # 1,000원 이하 제거
                    if abs(book_value) <= 1000:

                        book_value = 0

                    # 감가 완료 자산 제외
                    if book_value <= 0:

                        continue

                    # =========================================
                    # 결과 저장
                    # =========================================
                    result_rows.append({

                        "구분명": sheet_name,

                        "품명": row["품명"],

                        "취득일":
                        acquire_date.strftime(
                            "%Y-%m-%d"
                        ),

                        "취득가액":
                        format_number(amount),

                        "당월감가상각비":
                        format_number(monthly_dep),

                        "감가상각누계액":
                        format_number(accumulated),

                        "미상각잔액":
                        format_number(book_value)
                    })

                    # =========================================
                    # 전표 합산
                    # =========================================
                    if sheet_name in account_mapping:

                        debit_account = account_mapping[
                            sheet_name
                        ]["차변"]

                        credit_account = account_mapping[
                            sheet_name
                        ]["대변"]

                    else:

                        debit_account = "감가상각비"
                        credit_account = "감가상각누계액"

                    key = (
                        sheet_name,
                        debit_account,
                        credit_account
                    )

                    if key not in journal_summary:

                        journal_summary[key] = 0

                    journal_summary[key] += monthly_dep

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
        # 전표 생성
        # =========================================
        journal_rows = []

        for key, amount in journal_summary.items():

            sheet_name = key[0]
            debit = key[1]
            credit = key[2]

            journal_rows.append({

                "구분명": sheet_name,

                "차변계정": debit,

                "대변계정": credit,

                "합산금액":
                format_number(amount)
            })

        journal_df = pd.DataFrame(
            journal_rows
        )

        if len(journal_df) > 0:

            st.subheader("🧾 전표 생성")

            st.dataframe(
                journal_df,
                use_container_width=True
            )

        # =========================================
        # 엑셀 다운로드
        # =========================================
        output = BytesIO()

        with pd.ExcelWriter(
            output,
            engine="openpyxl"
        ) as writer:

            # 결과 저장
            for i, df in enumerate(total_result):

                sheet = f"결과_{i+1}"

                df.to_excel(
                    writer,
                    sheet_name=sheet,
                    index=False
                )

            # 전표 저장
            if len(journal_df) > 0:

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