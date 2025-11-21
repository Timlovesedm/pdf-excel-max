import streamlit as st
import pandas as pd
import pdfplumber
import io
import re
from collections import defaultdict

# ==========================================
# --- ツール①：PDFからデータを抽出する関数 ---
# ==========================================
def extract_tables_from_multiple_pdfs(pdf_files, keywords, global_start, global_end, file_specific_ranges=None):
    """
    pdf_files: アップロードされたファイルリスト
    keywords: 検索キーワードリスト
    global_start: 共通開始ページ (Noneなら最初から)
    global_end: 共通終了ページ (Noneなら最後まで)
    file_specific_ranges: { "ファイル名": {"start": int, "end": int} } 形式の辞書
    """
    all_rows = []
    if not keywords:
        st.error("❗ キーワードが入力されていません。", icon="🚨")
        return None

    for pdf_file in pdf_files:
        all_rows.append([f"ファイル名: {pdf_file.name}"])
        all_rows.append([])
        
        # --- ページ範囲の決定ロジック ---
        # 個別設定があるか確認
        current_start = global_start
        current_end = global_end
        
        if file_specific_ranges and pdf_file.name in file_specific_ranges:
            spec = file_specific_ranges[pdf_file.name]
            # 個別設定で値が入っていればそれを採用、なければNone(全範囲)
            current_start = spec.get("start") 
            current_end = spec.get("end")

        found_in_file = False
        try:
            with pdfplumber.open(pdf_file) as pdf:
                # ページインデックスの計算 (1始まりを0始まりに変換)
                s_idx = (current_start - 1) if current_start else 0
                e_idx = current_end if current_end else len(pdf.pages)
                
                # 範囲外エラー回避
                s_idx = max(0, s_idx)
                e_idx = min(len(pdf.pages), e_idx)
                
                if s_idx >= e_idx:
                    st.warning(f"ファイル「{pdf_file.name}」: ページ範囲指定が無効です（開始 {current_start} ～ 終了 {current_end}）。スキップします。", icon="⚠️")
                    continue

                target_pages = pdf.pages[s_idx:e_idx]
                
                for page in target_pages:
                    text = page.extract_text() or ""
                    if any(kw in text for kw in keywords):
                        found_in_file = True
                        tables = page.extract_tables()
                        for table_index, table in enumerate(tables):
                            if not table:
                                continue
                            all_rows.append([f"--- ページ {page.page_number} / テーブル {table_index + 1} ---"])
                            for row in table:
                                cleaned_row = ["" if item is None else str(item).replace("\n", " ") for item in row]
                                all_rows.append(cleaned_row)
                            all_rows.append([])
        except Exception as e:
            st.error(f"ファイル「{pdf_file.name}」処理中にエラー: {e}", icon="🔥")
            continue
        
        if not found_in_file:
            st.warning(f"ファイル「{pdf_file.name}」では指定範囲内にキーワードを含む表が見つかりませんでした。", icon="⚠️")

    if not any(r for r in all_rows if r):
        return None
    return pd.DataFrame(all_rows)


# ==========================================
# --- ツール②：共通ユーティリティ（強化版） ---
# ==========================================

def detect_year_header(cell_value):
    """セル内の文字列から年次ヘッダーを検出する"""
    cell_value = str(cell_value).strip()
    
    patterns = [
        # YYYYQ1~4
        (re.compile(r"(20\d{2}Q[1-4])", re.IGNORECASE), lambda m: m.group(1).upper()),
        # (自 2024年4月...
        (re.compile(r"\(?自\s*(\d{4})年(\d{1,2})月"), lambda m: f"{m.group(1)}/{m.group(2)}"),
        # 2024年3月期, 2024年3月 等
        (re.compile(r"(\d{4})年(\d{1,2})月"), lambda m: f"{m.group(1)}/{m.group(2)}"),
        # 2024年度
        (re.compile(r"(\d{4})年度"), lambda m: f"{m.group(1)}年度"),
        # 24/3 (YY/M) 形式
        (re.compile(r"^\'?(\d{2})/(\d{1,2})$"), lambda m: f"20{m.group(1)}/{m.group(2)}"),
        # 2024/3 (YYYY/M) 形式
        (re.compile(r"(\d{4})/(\d{1,2})"), lambda m: f"{m.group(1)}/{m.group(2)}"),
        # シンプルな数値 2024 or 202403
        (re.compile(r"^20\d{2}(\d{2})?$"), lambda m: m.group(0))
    ]

    for pat, formatter in patterns:
        match = pat.search(cell_value)
        if match:
            return formatter(match)
            
    return None

# ==========================================
# --- ツール②：【縦方向】統合ロジック ---
# ==========================================
def tool2_extract_data_vertical(df_chunk):
    if df_chunk.empty:
        return None, []
    
    year_cells = []
    for r in range(df_chunk.shape[0]):
        for c in range(df_chunk.shape[1]):
            cell_value = str(df_chunk.iat[r, c])
            year_header = detect_year_header(cell_value)
            if year_header:
                year_cells.append({"row": r, "col": c, "year_header": year_header})

    if not year_cells:
        return None, []

    year_cells.sort(key=lambda x: (x["row"], x["col"]))
    processed_years = set()
    
    initial_items = df_chunk[0].astype(str).str.strip().dropna()
    initial_items = initial_items[initial_items != ""]
    is_sonota = initial_items == "その他"
    if is_sonota.any():
        sonota_counts = initial_items.groupby(initial_items).cumcount()
        initial_items.loc[is_sonota] = "その他_temp_" + sonota_counts[is_sonota].astype(str)
    
    all_items_ordered = initial_items.drop_duplicates(keep="first").tolist()
    df_result = pd.DataFrame({"共通項目": all_items_ordered})

    for cell in year_cells:
        year_header = cell["year_header"]
        if year_header in processed_years:
            continue
        processed_years.add(year_header)
        val_col = cell["col"]
        
        temp_df = df_chunk.iloc[cell["row"] + 1 :, [0, val_col]].copy()
        temp_df.columns = ["共通項目", year_header]
        temp_df["共通項目"] = temp_df["共通項目"].astype(str).str.strip()
        temp_df = temp_df[temp_df["共通項目"] != ""].dropna(subset=["共通項目"])
        
        is_sonota = temp_df["共通項目"] == "その他"
        if is_sonota.any():
            sonota_counts = temp_df.groupby("共通項目").cumcount()
            temp_df.loc[is_sonota, "共通項目"] = "その他_temp_" + sonota_counts[is_sonota].astype(str)
            
        temp_df[year_header] = (
            pd.to_numeric(temp_df[year_header].astype(str).str.replace(",", ""), errors='coerce').fillna(0)
        )
        temp_df = temp_df.drop_duplicates(subset=["共通項目"], keep="first")
        df_result = pd.merge(df_result, temp_df, on="共通項目", how="left")

    return df_result, all_items_ordered

# ==========================================
# --- ツール②：【横方向】統合ロジック ---
# ==========================================
def tool2_extract_data_horizontal(df_chunk):
    if df_chunk.empty:
        return None, []

    # 1. 空列の削除
    df_clean = df_chunk.replace(r'^\s*$', pd.NA, regex=True).dropna(axis=1, how='all')
    
    if df_clean.shape[1] < 2:
        return None, []

    df_target = df_clean.fillna("") 

    # 2. 年次ヘッダーを探す
    detected_header = None
    header_row_idx = -1
    
    for r in range(min(10, df_target.shape[0])): 
        for c in range(df_target.shape[1]):
            val = df_target.iat[r, c]
            header_cand = detect_year_header(val)
            if header_cand:
                detected_header = header_cand
                header_row_idx = r
                break
        if detected_header:
            break
    
    if not detected_header:
        detected_header = str(df_target.iloc[0, -1]).strip()
        if not detected_header:
            detected_header = "Unknown_Period"

    # 3. データ抽出（一番左の列 と 一番右の列）
    temp_df = df_target.iloc[:, [0, -1]].copy()
    temp_df.columns = ["共通項目", detected_header]
    
    start_row = header_row_idx + 1 if header_row_idx != -1 else 0
    temp_df = temp_df.iloc[start_row:]
    
    # クレンジング
    temp_df["共通項目"] = temp_df["共通項目"].astype(str).str.strip()
    temp_df = temp_df[temp_df["共通項目"] != ""].dropna(subset=["共通項目"])
    
    temp_df[detected_header] = (
        pd.to_numeric(temp_df[detected_header].astype(str).str.replace(",", ""), errors='coerce')
    )
    temp_df = temp_df.dropna(subset=[detected_header])

    is_sonota = temp_df["共通項目"] == "その他"
    if is_sonota.any():
        sonota_counts = temp_df.groupby("共通項目").cumcount()
        temp_df.loc[is_sonota, "共通項目"] = "その他_temp_" + sonota_counts[is_sonota].astype(str)

    temp_df = temp_df.groupby("共通項目", as_index=False).sum()
    item_list = temp_df["共通項目"].tolist()

    return temp_df, item_list


# ==========================================
# --- ツール②：ファイル処理メイン関数 ---
# ==========================================
def process_files_and_tables(excel_file, integration_mode):
    try:
        xls = pd.ExcelFile(excel_file)
        sheet_name_to_read = "抽出結果" if "抽出結果" in xls.sheet_names else xls.sheet_names[0]
        df_full = pd.read_excel(xls, sheet_name=sheet_name_to_read, header=None)
    except Exception as e:
        st.error(f"Excelファイル読み込み失敗: {e}")
        return None

    df_full[0] = df_full[0].astype(str)
    file_indices = df_full[df_full[0].str.contains(r"ファイル名:", na=False)].index.tolist()
    file_chunks = []
    
    if not file_indices:
        file_chunks.append(df_full)
    else:
        for i in range(len(file_indices)):
            start_idx = file_indices[i]
            end_idx = file_indices[i + 1] if i + 1 < len(file_indices) else len(df_full)
            file_chunks.append(df_full.iloc[start_idx:end_idx].reset_index(drop=True))

    grouped_tables = defaultdict(list)
    master_item_order = defaultdict(list)

    for file_chunk in file_chunks:
        page_indices = file_chunk[file_chunk[0].str.contains(r"--- ページ", na=False)].index.tolist()
        table_chunks = []
        last_idx = 0
        
        if not page_indices:
            clean_chunk = file_chunk[
                ~file_chunk[0].str.contains(r"ファイル名:|---|^\s*$", na=False, regex=True)
            ].dropna(how="all")
            if not clean_chunk.empty:
                table_chunks.append(clean_chunk)
        else:
            for idx in page_indices:
                chunk = file_chunk.iloc[last_idx:idx]
                if not chunk.empty:
                    table_chunks.append(chunk)
                last_idx = idx
            final_chunk = file_chunk.iloc[last_idx:]
            if not final_chunk.empty:
                table_chunks.append(final_chunk)

        for i, table_chunk in enumerate(table_chunks):
            clean_table_chunk = table_chunk[
                ~table_chunk[0].str.contains(r"ファイル名:|---", na=False, regex=True)
            ].dropna(how="all")
            
            if clean_table_chunk.empty:
                continue
            
            if integration_mode == "vertical":
                processed_df, item_order = tool2_extract_data_vertical(clean_table_chunk.reset_index(drop=True))
            else: # horizontal
                processed_df, item_order = tool2_extract_data_horizontal(clean_table_chunk.reset_index(drop=True))

            if processed_df is not None and not processed_df.empty:
                grouped_tables[i].append(processed_df)
                
                current_master_order = master_item_order[i]
                if not current_master_order:
                    master_item_order[i].extend(item_order)
                else:
                    last_known_index = -1
                    for item in item_order:
                        if item in current_master_order:
                            last_known_index = current_master_order.index(item)
                        else:
                            current_master_order.insert(last_known_index + 1, item)
                            last_known_index += 1

    final_summaries = []
    for table_index in sorted(grouped_tables.keys()):
        list_of_dfs = grouped_tables[table_index]
        ordered_items = master_item_order[table_index]
        
        if not list_of_dfs:
            continue
            
        result_df = pd.DataFrame({"共通項目": ordered_items})
        
        for df_to_merge in list_of_dfs:
            cols_to_drop = [
                col for col in df_to_merge.columns if col in result_df.columns and col != "共通項目"
            ]
            result_df = pd.merge(
                result_df, df_to_merge.drop(columns=cols_to_drop), on="共通項目", how="left"
            )
            
        result_df.fillna(0, inplace=True)
        
        def sort_key(col_name):
            s = str(col_name).upper().replace('/', '').replace('Q', '0').replace('年度', '').replace('年', '').replace('月', '')
            digits = "".join(filter(str.isdigit, s))
            if digits:
                return int(digits.ljust(6, '0'))
            return 99999999

        year_cols = sorted(
            [col for col in result_df.columns if col != "共通項目"],
            key=sort_key
        )
        final_cols = ["共通項目"] + year_cols
        result_df = result_df[final_cols]
        
        for col in year_cols:
            result_df[col] = pd.to_numeric(result_df[col], errors='coerce').fillna(0).astype(int)
            
        result_df["共通項目"] = result_df["共通項目"].str.replace(r"_temp_\d+$", "", regex=True)
        
        final_summaries.append(result_df)
        
    return final_summaries


# ==========================================
# --- Streamlit UI ---
# ==========================================
st.set_page_config(page_title="多機能ツール", layout="wide")
st.title("📄📊 多機能ツール")

# --- ツール① ---
with st.container(border=True):
    st.header("ツール①：PDF表データ抽出")
    pdf_files = st.file_uploader(
        "PDFファイルをアップロード（複数可）", type="pdf", accept_multiple_files=True
    )
    keyword_input_str = st.text_input("検索キーワード（カンマ区切り）")
    
    st.subheader("ページ範囲設定")
    # 設定モードの選択
    range_mode = st.radio(
        "範囲設定モード", 
        ("全てのファイルで同じ範囲にする", "ファイルごとに範囲を指定する"),
        index=0
    )
    
    global_start = None
    global_end = None
    file_specific_ranges = {}

    if range_mode == "全てのファイルで同じ範囲にする":
        col1, col2 = st.columns(2)
        s_in = col1.text_input("開始ページ (共通)", placeholder="例: 5")
        e_in = col2.text_input("終了ページ (共通)", placeholder="例: 10")
        if s_in.isdigit(): global_start = int(s_in)
        if e_in.isdigit(): global_end = int(e_in)
        
    else:
        st.info("各ファイルの開始・終了ページを入力してください（空欄の場合は全ページが対象になります）")
        if pdf_files:
            for i, f in enumerate(pdf_files):
                c1, c2, c3 = st.columns([4, 1, 1])
                c1.write(f"📄 **{f.name}**")
                s_in = c2.text_input("開始", key=f"start_{i}_{f.name}", placeholder="1")
                e_in = c3.text_input("終了", key=f"end_{i}_{f.name}", placeholder="Last")
                
                s_val = int(s_in) if s_in.isdigit() else None
                e_val = int(e_in) if e_in.isdigit() else None
                
                # 辞書に保存
                file_specific_ranges[f.name] = {"start": s_val, "end": e_val}
        else:
            st.warning("まずはファイルをアップロードしてください。")

    if st.button("抽出開始 ▶️"):
        if pdf_files:
            keywords = [kw.strip() for kw in keyword_input_str.split(",") if kw.strip()]
            
            with st.spinner("PDF解析中..."):
                # 以前の引数 start_page, end_page の代わりに global_start, global_end と specific_ranges を渡す
                df_result = extract_tables_from_multiple_pdfs(
                    pdf_files, keywords, 
                    global_start=global_start, 
                    global_end=global_end,
                    file_specific_ranges=file_specific_ranges
                )
                
                if df_result is not None and not df_result.empty:
                    st.success("抽出完了！", icon="✅")
                    st.dataframe(df_result)
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                        df_result.to_excel(writer, index=False, header=False, sheet_name="抽出結果")
                        workbook = writer.book
                        worksheet = writer.sheets["抽出結果"]
                        bold_format = workbook.add_format({"bold": True, "font_size": 20})
                        for idx, val in enumerate(df_result[0]):
                            if isinstance(val, str) and val.startswith("ファイル名:"):
                                worksheet.set_row(idx, None, bold_format)
                    
                    if keywords:
                        base_name = '_'.join(keywords)
                        download_filename = f"{base_name}_まとめ.xlsx"
                    else:
                        download_filename = "抽出結果_まとめ.xlsx"

                    st.download_button(
                        label="📥 Excelファイルをダウンロード",
                        data=output.getvalue(),
                        file_name=download_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
        else:
            st.error("PDFファイルをアップロードしてください。", icon="🚨")

st.divider()

# --- ツール② ---
with st.container(border=True):
    st.header("ツール②：統合データ作成")
    
    st.info("📝 データの並び方を選択してください")
    integration_mode_label = st.radio(
        "統合モード選択",
        ("縦方向統合 (従来の形式)", "横方向統合 (項目:左 / 数値:右)"),
        help="データが縦に積み上がっている場合は「縦方向」、横並びの年次データを結合する場合は「横方向」を選択してください"
    )
    integration_mode = "vertical" if "縦方向" in integration_mode_label else "horizontal"
    
    excel_file = st.file_uploader("Excelファイルをアップロード", type=["xlsx"])
    
    if st.button("統合まとめ表を作成 ▶️", disabled=(excel_file is None)):
        with st.spinner("データ整理中..."):
            all_summaries = process_files_and_tables(excel_file, integration_mode)
            
            if all_summaries:
                st.success(f"{len(all_summaries)}個のまとめ表を作成！", icon="✅")
                output_excel = io.BytesIO()
                with pd.ExcelWriter(output_excel, engine="xlsxwriter") as writer:
                    for i, summary_df in enumerate(all_summaries):
                        sheet_name = f"統合まとめ表_{i+1}"
                        summary_df.to_excel(writer, sheet_name=sheet_name, index=False)
                        worksheet = writer.sheets[sheet_name]
                        worksheet.set_column(0, 0, 30)

                base_name_input = excel_file.name.rsplit('.xlsx', 1)[0]
                mode_suffix = "_縦統合" if integration_mode == "vertical" else "_横統合"
                if base_name_input.endswith('_まとめ'):
                    base_name_output = base_name_input.removesuffix('_まとめ') + mode_suffix
                else:
                    base_name_output = base_name_input + mode_suffix
                download_filename = f"{base_name_output}.xlsx"

                st.download_button(
                    label="📥 統合まとめ表をダウンロード",
                    data=output_excel.getvalue(),
                    file_name=download_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            else:
                st.warning("有効なデータが見つかりませんでした。（ヘッダー未検出、または空列の問題の可能性があります）", icon="⚠️")
