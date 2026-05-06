# -*- coding: utf-8 -*-
import pandas as pd
import openpyxl
import json
import os
import sys
import math

_script_dir = os.path.dirname(os.path.abspath(__file__))
_default_xlsx = "raw.xlsx"

if len(sys.argv) > 1:
    arg = sys.argv[1]
    XLSX_PATH = arg if os.path.isabs(arg) else os.path.join(_script_dir, arg)
else:
    XLSX_PATH = os.path.join(_script_dir, _default_xlsx)

OUTPUT_DIR = os.path.join(_script_dir, "data")

# 출력 디렉터리 생성
os.makedirs(OUTPUT_DIR, exist_ok=True)

# 시트 → 파일명 매핑
SHEET_MAP = {
    "대학별 전형별 인원 및 전형방법": "jeonhyeong.json",
    "2028 특수학과": "teuksu.json",
    "수능최저학력기준": "suneung_choejeo.json",
    "교과성적 반영방법": "gyogwa.json",
    "수능 반영 방법": "suneung_bangyeong.json",
    "대학별 변동사항": "byeondong.json",
}


def is_nan(val):
    if val is None:
        return True
    if isinstance(val, float) and math.isnan(val):
        return True
    return False


def clean_value(val):
    """값을 문자열로 변환. NaN → "", float 정수 → 정수 문자열"""
    if is_nan(val):
        return ""
    if isinstance(val, float):
        if val == int(val):
            return str(int(val))
        return str(val)
    if isinstance(val, int):
        return str(val)
    s = str(val).strip()
    # 줄바꿈을 공백으로
    s = s.replace("\n", " ").replace("\r", "")
    return s


def row_is_empty(row):
    """행의 모든 값이 비어있으면 True"""
    return all(str(v).strip() == "" for v in row)


def looks_like_data_row(row):
    """
    데이터 행 판단 기준:
    - 첫 번째 컬럼(row[0])이 숫자이거나
    - 두 번째 컬럼(row[1])이 숫자이면서 첫 번째도 숫자이면 데이터 행
    - 첫 번째 비어있지 않은 셀이 숫자인 경우에도 해당 셀이 row의 앞쪽(인덱스 0~2)에 있을 때만
    """
    # row[0]이 숫자이면 데이터 행
    v0 = str(row[0]).strip() if row else ""
    if v0 != "":
        try:
            int(float(v0))
            return True
        except (ValueError, OverflowError):
            pass

    # row[0]이 비어있고 row[1]이 숫자이면 데이터 행
    if len(row) > 1:
        v1 = str(row[1]).strip()
        if v1 != "":
            try:
                int(float(v1))
                return True
            except (ValueError, OverflowError):
                pass

    return False


def fix_unnamed_columns(columns):
    """
    - 명명된 컬럼이 처음 등장하면 그대로 사용
    - 같은 이름이 반복되면 이름_2, 이름_3 형태로 suffix 부여
    - "Unnamed:" 또는 빈 컬럼명은 이전 명명 컬럼 기반 suffix 부여
    """
    result = []
    name_count = {}  # 각 이름이 몇 번 나왔는지 추적
    last_named = None

    for col in columns:
        col_str = str(col).strip()

        if col_str.startswith("Unnamed:") or col_str == "":
            # 빈 컬럼: 이전 명명된 컬럼 기반 suffix
            if last_named is not None:
                base = last_named
                count = name_count.get(base, 1) + 1
                name_count[base] = count
                result.append(f"{base}_{count}")
            else:
                result.append(f"col_{len(result)}")
        else:
            # 명명된 컬럼: 중복 체크
            if col_str in name_count:
                count = name_count[col_str] + 1
                name_count[col_str] = count
                result.append(f"{col_str}_{count}")
            else:
                name_count[col_str] = 1
                result.append(col_str)
            last_named = col_str

    return result


def build_columns_from_header_rows(header_rows, ncols):
    """
    1~2개의 헤더 행을 받아 컬럼명 리스트 생성.
    - 1행 헤더: 그대로 사용
    - 2행 헤더: 각 열별로 두 행 값을 합침 (빈 셀은 왼쪽 값으로 forward-fill)
    """
    if not header_rows:
        return [f"col_{i}" for i in range(ncols)]

    # 각 행에서 빈 셀을 왼쪽 값으로 채우기 (forward fill)
    def forward_fill(row):
        filled = []
        last = ""
        for v in row:
            v = str(v).strip().replace("\n", " ")
            if v == "":
                filled.append(last)
            else:
                last = v
                filled.append(v)
        return filled

    filled_rows = [forward_fill(r) for r in header_rows]

    # 컬럼 수 맞추기
    for i, row in enumerate(filled_rows):
        if len(row) < ncols:
            filled_rows[i] = row + [""] * (ncols - len(row))
        else:
            filled_rows[i] = row[:ncols]

    if len(filled_rows) == 1:
        columns = filled_rows[0]
    else:
        # 2행 합치기: 같으면 첫 행만, 다르면 "_" 연결
        columns = []
        for col_idx in range(ncols):
            parts = []
            seen = set()
            for row in filled_rows:
                v = row[col_idx]
                if v and v not in seen:
                    parts.append(v)
                    seen.add(v)
            columns.append("_".join(parts) if parts else "")

    return columns


def is_summary_row(row):
    """
    합계/총계 행 판단:
    - 대부분 비어있고, 소수(1~2개)의 숫자 값만 있는 행
    - 텍스트 셀이 하나도 없음
    """
    non_empty = [str(v).strip() for v in row if str(v).strip() != ""]
    if not non_empty:
        return True  # 완전 빈 행
    # 모두 숫자이면 합계 행
    all_numeric = all(True if _try_float(v) else False for v in non_empty)
    return all_numeric


def _try_float(v):
    try:
        float(v)
        return True
    except (ValueError, TypeError):
        return False


def process_sheet(xl, sheet_name):
    """시트를 읽어 (columns, data_rows) 반환"""

    df_raw = pd.read_excel(xl, sheet_name=sheet_name, header=None, dtype=str)
    df_raw = df_raw.fillna("")

    nrows, ncols = df_raw.shape
    rows = [list(df_raw.iloc[i]) for i in range(nrows)]

    # 구조 파악:
    # - 0번 행: 빈 행이거나 합계 행 (건너뜀)
    # - 다음에 헤더 행(들): 텍스트 위주
    # - 중간 빈 행 가능
    # - 그 다음 데이터 행: 첫 셀이 숫자

    # 단계 1: 선행 빈 행 및 합계(숫자만 있는) 행 건너뜀
    pos = 0
    while pos < nrows and (row_is_empty(rows[pos]) or is_summary_row(rows[pos])):
        pos += 1

    if pos >= nrows:
        return [f"col_{i}" for i in range(ncols)], []

    # 단계 2: 헤더 행 수집 (데이터 행이 나올 때까지, 최대 3행)
    header_rows = []
    max_header = 3
    checked = 0
    while pos < nrows and checked < max_header:
        row = rows[pos]
        if row_is_empty(row):
            # 헤더 중간 빈 행 건너뜀
            pos += 1
            checked += 1
            continue
        if looks_like_data_row(row):
            # 데이터 시작
            break
        header_rows.append(row)
        pos += 1
        checked += 1

    # 단계 3: 데이터 시작 전 빈 행 건너뜀
    while pos < nrows and row_is_empty(rows[pos]):
        pos += 1

    data_start = pos

    # 컬럼명 생성
    columns = build_columns_from_header_rows(header_rows, ncols)

    # Unnamed 처리
    columns = fix_unnamed_columns(columns)

    # 컬럼 수 맞추기
    if len(columns) < ncols:
        columns.extend([f"col_{i}" for i in range(len(columns), ncols)])
    elif len(columns) > ncols:
        columns = columns[:ncols]

    # 데이터 행 처리
    data_rows = []
    for row_idx in range(data_start, nrows):
        row = rows[row_idx]
        row_dict = {}
        for col_idx, col_name in enumerate(columns):
            if col_idx < len(row):
                row_dict[col_name] = clean_value(row[col_idx])
            else:
                row_dict[col_name] = ""

        # 완전히 빈 행 제거
        if all(v == "" for v in row_dict.values()):
            continue

        # 일련번호(No, 연번)만 채워진 빈 트레일링 행 제거
        # ID 성격 컬럼 제외 후 모든 값이 비어있으면 스킵
        ID_LIKE = {"No", "연번", "연 번"}
        non_id_values = [v for k, v in row_dict.items() if k not in ID_LIKE]
        if non_id_values and all(v == "" for v in non_id_values):
            continue

        data_rows.append(row_dict)

    return columns, data_rows


def main():
    xl = pd.ExcelFile(XLSX_PATH, engine="openpyxl")
    available_sheets = xl.sheet_names

    print(f"엑셀 파일 내 시트 목록: {available_sheets}\n")

    for sheet_name, filename in SHEET_MAP.items():
        # 시트명 매칭
        matched = None
        for s in available_sheets:
            if s.strip() == sheet_name.strip():
                matched = s
                break

        if matched is None:
            print(f"[경고] 시트를 찾을 수 없음: '{sheet_name}'")
            continue

        try:
            columns, data = process_sheet(xl, matched)

            output = {
                "columns": columns,
                "data": data
            }

            out_path = os.path.join(OUTPUT_DIR, filename)
            with open(out_path, "w", encoding="utf-8") as f:
                json.dump(output, f, ensure_ascii=False, indent=2)

            print(f"저장 완료: {filename} ({len(data)}행)")

        except Exception as e:
            print(f"[오류] {sheet_name} 처리 중 오류: {e}")
            import traceback
            traceback.print_exc()

    print("\n모든 처리 완료!")


if __name__ == "__main__":
    main()
