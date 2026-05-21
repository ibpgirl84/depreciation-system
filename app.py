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
    value=datetime.today()
)

# 말일 계산
last_day = calendar.monthrange(
    base_date.year,
    base_date.month
)[1]

base_date = datetime(
    base_date.year,
    base_date.month,
    last_day
)

st.write(
    f"기준월 말일 : {base_date.strftime('%Y-%m-%d')}"
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
# 회계정책
# =========================================
account_mapping = {

    "차량": {
        "debit": "유형자산상각비/차량운반구",
        "credit": "차량운반구감가상각누계액",
        "life_months": 48
    },

    "비품": {
        "debit": "유형자산상각비/비품",
        "credit": "비품감가상각누계액",
        "life_months": 48
    },

    "시설장치": {
        "debit": "유형자산상각비/시설장치",
        "credit": "시설장치감가상각누계액",
        "life_months": 48
    },

    "소프트웨어": {
        "debit": "무형자산상각비/소프트웨어",
        "credit": "소프트웨어",
        "life_months": 60
    },

    "상표권": {
        "debit": "무형자산상각비/상표권",
        "credit": "상표권",
        "life_months": 60
    },

    "기타무형자산": {
        "debit": "무형자산상각비/기타무형자산",
        "credit": "기타무형자산",
        "life_months": 60
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

        # =====================================
        # 시트 반복
        # =====================================
        for sheet_name, df in excel_data.items():

            st.subheader(
                f"📁 시트 처리 : {sheet_name}"
            )

            # 컬럼명 정리
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

            if missing_cols:

                st.warning(
                    f"{sheet_name} 시트 컬럼 누락 : {missing_cols}"
                )

                continue

            # 숫자형 변환
            df["취득가액"] = pd.to_numeric(
                df["취득가액"],
                errors="coerce"
            ).fillna(0)

            if "잔존가액" not in df.columns:

                df["잔존가액"] = 0

            df["잔존가액"] = pd.to_numeric(
                df["잔존가액"],
                errors="coerce"
            ).fillna(0)

            result_rows = []

            # =====================================
            # 자산 반복
            # =====================================
            for idx, row in df.iterrows():

                try:

                    category = sheet_name.strip()

                    acquire_date = pd.to_datetime(
                        row["취득일"]
                    )

                    amount = float(
                        row["취득가액"]
                    )

                    remain_value = float(
                        row["잔존가액"]
                    )

                    # ================================
                    # 내용연수
                    # ================================
                    if category in account_mapping:

                        useful_months = (
                            account_mapping[
                                category
                            ]["life_months"]
                        )

                    else:

                        useful_months = 60

                    # ================================
                    # 상각대상금액
                    # ================================
                    depreciable_amount = (
                        amount - remain_value
                    )

                    # ================================
                    # 월상각비
                    # 소수 유지
                    # ================================
                    monthly_dep = (
                        depreciable_amount
                        / useful_months
                    )

                    # ================================
                    # 사용개월수
                    # 취득월 제외
                    # ================================
                    used_months = (
                        (base_date.year - acquire_date.year)
                        * 12
                        + (
                            base_date.month
                            - acquire_date.month
                        )
                    )

                    used_months = max(
                        0,
                        used_months
                    )

                    used_months = min(
                        used_months,
                        useful_months
                    )

                    # ================================
                    # 감가상각누계액
                    # ================================
                    accumulated = (
                        monthly_dep
                        * used_months
                    )

                    # ================================
                    # 미상각잔액
                    # ================================
                    book_value = (
                        amount - accumulated
                    )

                    # 1000원 이하 제거
                    if abs(book_value) <= 1000:

                        book_value = 0

                    # 감가상각 종료
                    if used_months >= useful_months:

                        monthly_dep_display = 0
                        book_value = 0

                    else:

                        monthly_dep_display = (
                            monthly_dep
                        )

                    # ================================
                    # 결과 저장
                    # ================================
                    result_rows.append({

                        "구분명": category,

                        "품명": row["품명"],

                        "취득일": acquire_date.strftime(
                            "%Y-%m-%d"
                        ),

                        "취득가액": format_number(
                            amount
                        ),

                        "당월감가상각비": format_number(
                            round(monthly_dep_display)
                        ),

                        "감가상각누계액": format_number(
                            round(accumulated)
                        ),

                        "미상각잔액": format_number(
                            round(book_value)
                        )
                    })

                    # ================================
                    # 전표 합산
                    # 소수 유지 상태로 합산
                    # ================================
                    if monthly_dep_display > 0:

                        if (
                            category
                            not in journal_summary
                        ):

                            journal_summary[
                                category
                            ] = 0

                        journal_summary[
                            category
                        ] += monthly_dep

                except:
                    pass

            result_df = pd.DataFrame(
                result_rows
            )

            if len(result_df) > 0:

                st.dataframe(
                    result_df,
                    use_container_width=True
                )

                total_result.append(
                    result_df
                )

        # =====================================
        # 전표 생성
        # =====================================
        journal_rows = []

        for category, amount in (
            journal_summary.items()
        ):

            if category in account_mapping:

                debit_account = (
                    account_mapping[
                        category
                    ]["debit"]
                )

                credit_account = (
                    account_mapping[
                        category
                    ]["credit"]
                )

            else:

                debit_account = "감가상각비"

                credit_account = "감가상각누계액"

            journal_rows.append({

                "구분명": category,

                "차변계정": debit_account,

                "대변계정": credit_account,

                # 마지막에만 반올림
                "합산금액": format_number(
                    int(round(amount, 0))
                )
            })

        # =====================================
        # 전표 출력
        # =====================================
        if len(journal_rows) > 0:

            st.subheader(
                "🧾 전표 생성"
            )

            journal_df = pd.DataFrame(
                journal_rows
            )

            st.dataframe(
                journal_df,
                use_container_width=True
            )

        # =====================================
        # 엑셀 저장
        # =====================================
        output = BytesIO()

        with pd.ExcelWriter(
            output,
            engine="openpyxl"
        ) as writer:

            for i, df in enumerate(
                total_result
            ):

                sheet = f"결과_{i+1}"

                df.to_excel(
                    writer,
                    sheet_name=sheet,
                    index=False
                )

            if len(journal_rows) > 0:

                journal_df.to_excel(
                    writer,
                    sheet_name="전표",
                    index=False
                )

        # =====================================
        # 다운로드 버튼
        # =====================================
        st.download_button(
            label="📥 결과 엑셀 다운로드",
            data=output.getvalue(),
            file_name="감가상각_결과.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:

        st.error(str(e))