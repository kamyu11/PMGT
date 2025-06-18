import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import xlrd
import re
import io

st.set_page_config(page_title='부모-자식 관계 추출기', layout='wide')

st.title('부모-자식 관계 추출기')

st.markdown(
    '''
    1. **파일 선택** 버튼을 눌러 Excel(.xlsx 또는 .xls) 파일을 업로드합니다.  
    2. 업로드 후, **변환** 버튼을 누르면 아래 표에 부모-자식 관계가 나타나고,  
       동시에 Excel파일로 내려받을 수 있는 링크가 생성됩니다.
    '''
)

uploaded_file = st.file_uploader(
    '분석할 Excel 파일 선택 (.xlsx 또는 .xls)',
    type=['xlsx','xls']
)

if uploaded_file is not None:
    # 파일을 한 번 읽어서 바이너리로 저장
    content = uploaded_file.read()
    file_ext = uploaded_file.name.lower().split('.')[-1]

    # 변환 버튼
    if st.button('변환'):
        try:
            # ────────────────────────────────────────────
            # 1) pandas로 원본 전체 읽기 (5행부터 데이터가 있다고 가정)
            # ────────────────────────────────────────────
            if file_ext == 'xlsx':
                df_full = pd.read_excel(
                    io.BytesIO(content),
                    sheet_name=1,
                    header=None,
                    engine='openpyxl'
                )
                # 병합 정보 가져오기
                wb_temp = load_workbook(io.BytesIO(content), data_only=True)
                ws_temp = wb_temp.worksheets[1]
                merged_ranges = list(ws_temp.merged_cells.ranges)
            elif file_ext == 'xls':
                book_temp = xlrd.open_workbook(file_contents=content, formatting_info=True)
                sheet_temp = book_temp.sheet_by_index(1)
                merged_ranges = sheet_temp.merged_cells
                df_full = pd.read_excel(
                    io.BytesIO(content),
                    sheet_name=1,
                    header=None,
                    engine='xlrd'
                )
            else:
                st.error('지원되지 않는 파일 형식입니다.')
                st.stop()

            # ────────────────────────────────────────────
            # 2) 5행(엑셀 기준) 이후 데이터만 사용
            # ────────────────────────────────────────────
            START_ROW_IDX = 4  # 엑셀 1-based 5행 → 0-based 4
            df_data = df_full.iloc[START_ROW_IDX:].reset_index(drop=True)
            df_data.columns = [f'COLUMN{i+1}' for i in range(df_data.shape[1])]

            # ────────────────────────────────────────────
            # 3) 병합된 셀 영역만 골라서 위쪽 값으로 채우기
            # ────────────────────────────────────────────
            def fill_merged_cells(df, merged_ranges, is_xlsx=True):
                for mr in merged_ranges:
                    if is_xlsx:
                        # openpyxl CellRange: mr.bounds → (min_col, min_row, max_col, max_row) (1-based)
                        min_col, min_row, max_col, max_row = mr.bounds
                        start_r = min_row - 1 - START_ROW_IDX
                        end_r   = max_row - 1 - START_ROW_IDX
                        start_c = min_col - 1
                        if not (0 <= start_r < len(df) and 0 <= start_c < df.shape[1]):
                            continue
                        top_val = df.iat[start_r, start_c]
                        for rr in range(start_r, end_r+1):
                            for cc in range(start_c, max_col):
                                if 0 <= rr < len(df) and 0 <= cc < df.shape[1]:
                                    df.iat[rr, cc] = top_val
                    else:
                        # xlrd merged_cells: (rlo, rhi, clo, chi)  (0-based)
                        rlo, rhi, clo, chi = mr
                        start_r = rlo - START_ROW_IDX
                        end_r   = (rhi - 1) - START_ROW_IDX
                        start_c = clo
                        if not (0 <= start_r < len(df) and 0 <= start_c < df.shape[1]):
                            continue
                        top_val = df.iat[start_r, start_c]
                        for rr in range(start_r, end_r+1):
                            for cc in range(start_c, chi):
                                if 0 <= rr < len(df) and 0 <= cc < df.shape[1]:
                                    df.iat[rr, cc] = top_val

            if file_ext == 'xlsx':
                fill_merged_cells(df_data, merged_ranges, is_xlsx=True)
            else:
                fill_merged_cells(df_data, merged_ranges, is_xlsx=False)

            # ────────────────────────────────────────────
            # 4) 부모-자식 관계 추출
            # ────────────────────────────────────────────
            rows = []

            def clean_name(val):
                if pd.isna(val):
                    return None
                txt = re.sub(r'[()\s\n]', '', str(val))
                return txt[:5]

            # G열(COLUMN7) 부모 찾기 순서: F→E→D→C→B→A
            parent_columns = ['COLUMN6','COLUMN5','COLUMN4','COLUMN3','COLUMN2','COLUMN1']

            # G열부터 차례로 눕히면서, 값이 끊어져도 계속 처리
            for _, row in df_data.iterrows():
                raw_child = row.get('COLUMN7')  # G열
                if pd.isna(raw_child):
                    # G열이 비어도 넘어가며 다음 행 처리
                    continue

                # 부모를 순서대로 찾기
                parent = None
                for pc in parent_columns:
                    v = row.get(pc)
                    if not pd.isna(v):
                        parent = clean_name(v)
                        break
                if parent is None:
                    continue

                child_base = re.sub(r'[\n\s]', '', str(raw_child)).strip()

                # 접미사(P/C/S) 판단: 8~10열 숫자 여부
                suffix_list = []
                if df_data.shape[1] >= 8:
                    num8 = pd.to_numeric(row.get('COLUMN8'), errors='coerce')
                    if pd.notna(num8):
                        suffix_list.append('P')
                if df_data.shape[1] >= 9:
                    num9 = pd.to_numeric(row.get('COLUMN9'), errors='coerce')
                    if pd.notna(num9):
                        suffix_list.append('C')
                if df_data.shape[1] >= 10:
                    num10 = pd.to_numeric(row.get('COLUMN10'), errors='coerce')
                    if pd.notna(num10):
                        suffix_list.append('S')

                if suffix_list:
                    for s in suffix_list:
                        rows.append({'부모': parent, '자식': f'{child_base}{s}'})
                else:
                    rows.append({'부모': parent, '자식': child_base})

            # (2) 나머지 열(F~B)을 자식으로 보고, 왼쪽 열을 부모로
            def extract_parent_child(df, child_col, parent_cols):
                for _, row in df.iterrows():
                    raw_child = row.get(child_col)
                    if pd.isna(raw_child):
                        continue
                    child = clean_name(raw_child)
                    parent = None
                    for pc in parent_cols:
                        v = row.get(pc)
                        if not pd.isna(v):
                            parent = clean_name(v)
                            break
                    if parent:
                        rows.append({'부모': parent, '자식': child})

            extract_parent_child(
                df_data,
                'COLUMN6',
                ['COLUMN5','COLUMN4','COLUMN3','COLUMN2','COLUMN1']
            )  # F열
            extract_parent_child(
                df_data,
                'COLUMN5',
                ['COLUMN4','COLUMN3','COLUMN2','COLUMN1']
            )  # E열
            extract_parent_child(
                df_data,
                'COLUMN4',
                ['COLUMN3','COLUMN2','COLUMN1']
            )  # D열
            extract_parent_child(
                df_data,
                'COLUMN3',
                ['COLUMN2','COLUMN1']
            )  # C열
            extract_parent_child(
                df_data,
                'COLUMN2',
                ['COLUMN1']
            )  # B열

            # ────────────────────────────────────────────
            # 5) 결과 DataFrame 생성 및 화면에 표시
            # ────────────────────────────────────────────
            result_df = pd.DataFrame(rows).drop_duplicates().reset_index(drop=True)
            st.subheader('추출된 부모-자식 관계')
            st.dataframe(result_df)

            # ────────────────────────────────────────────
            # 6) Excel 파일로 다운로드할 수 있는 버튼
            # ────────────────────────────────────────────
            towrite = io.BytesIO()
            result_df.to_excel(towrite, index=False, engine='openpyxl')
            towrite.seek(0)
            st.download_button(
                label='엑셀 파일로 내려받기',
                data=towrite,
                file_name='parent_child_result.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )

        except Exception as e:
            st.error(f'처리 중 오류 발생: {e}')
