from pathlib import Path
import re, py_compile

src = Path("/mnt/data/app_updated_requested.py")
code = src.read_text(encoding="utf-8")

# 1. Replace make_blank_template_csv with workbook xlsx + keep csv? Need add xlsx function.
old = r'''def make_blank_template_csv() -> bytes:
    """產出批次上傳用空白制式表單。"""
    return pd.DataFrame(columns=REQUIRED_COLUMNS).to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
'''
new = r'''def make_blank_template_csv() -> bytes:
    """產出批次上傳用空白制式表單 CSV。"""
    return pd.DataFrame(columns=REQUIRED_COLUMNS).to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")


def make_blank_template_excel():
    """產出含欄位說明的空白制式 Excel 表單。"""
    if find_spec("openpyxl") is None:
        return None

    template_df = pd.DataFrame(columns=REQUIRED_COLUMNS)

    field_notes = pd.DataFrame(
        [
            {"欄位名稱": "代碼", "填寫說明": "場址唯一識別碼，可自行編號。", "建議填寫值或選項": "例如：SITE-001"},
            {"欄位名稱": "縣市", "填寫說明": "場址所在縣市，需與 town_lookup.csv 中縣市名稱一致，才能自動帶入人口密度。", "建議填寫值或選項": "例如：臺北市、新北市、桃園市"},
            {"欄位名稱": "鄉鎮市區", "填寫說明": "場址所在鄉鎮市區，需與 town_lookup.csv 中鄉鎮市區名稱一致。", "建議填寫值或選項": "例如：中正區、板橋區、草屯鎮"},
            {"欄位名稱": "基地名稱", "填寫說明": "場址名稱。若該基地已存在於綠化單元資料，會用於排除自己。", "建議填寫值或選項": "文字"},
            {"欄位名稱": "TWD97X", "填寫說明": "TWD97 X 座標，單點模式要求 6 碼數字。", "建議填寫值或選項": "例如：250000"},
            {"欄位名稱": "TWD97Y", "填寫說明": "TWD97 Y 座標，單點模式要求 7 碼數字。", "建議填寫值或選項": "例如：2650000"},
            {"欄位名稱": "距主要道路距離(公尺)", "填寫說明": "目前保留人工填寫，用於空品壓力與道路線源情境判斷。", "建議填寫值或選項": "數值；若未知可先填 999"},
            {"欄位名稱": "土地權屬", "填寫說明": "用於判斷土地權屬單純度。", "建議填寫值或選項": "公有或單一權屬／公私混合或需協調／私有或權屬複雜／其他／待確認"},
            {"欄位名稱": "管理機關", "填寫說明": "用於判斷管理權責明確性。", "建議填寫值或選項": "環境部／縣市環保局／鄉鎮市區公所／學校／醫療院所／其他機關／待確認"},
            {"欄位名稱": "基地面積(公頃)", "填寫說明": "用於可植栽空間評分；若無面積但有長度，系統會用長度估算。", "建議填寫值或選項": "數值，例如 0.8"},
            {"欄位名稱": "基地長度(公里)", "填寫說明": "基地面積缺漏時，以長度 × 0.1 估算面積。", "建議填寫值或選項": "數值，例如 1.2"},
            {"欄位名稱": "短期推動性", "填寫說明": "用於判斷近期是否容易推動。", "建議填寫值或選項": "可短期推動／需部分協調／短期難以推動／待評估"},
            {"欄位名稱": "開放可及性", "填寫說明": "用於判斷公共服務性。", "建議填寫值或選項": "完全開放／部分開放／不開放／待確認"},
            {"欄位名稱": "基地內部是否有停留活動空間", "填寫說明": "是否有人會在基地內休憩、運動、活動或等候。", "建議填寫值或選項": "是／否／未知"},
            {"欄位名稱": "是否有步行通學騎行路徑", "填寫說明": "基地內或周邊是否有人會步行、通學、騎車或穿越。", "建議填寫值或選項": "是／否／未知"},
            {"欄位名稱": "是否有邊界防護需求", "填寫說明": "基地是否位於道路、工廠、空地、住宅或學校等交界，需要緩衝防護。", "建議填寫值或選項": "是／否／未知"},
            {"欄位名稱": "道路周邊是否有敏感受體或社區活動空間", "填寫說明": "道路旁是否有住宅、學校、醫療院所或社區活動空間。", "建議填寫值或選項": "是／否／未知"},
            {"欄位名稱": "是否有短期事件", "填寫說明": "基地或周邊是否有施工、整地、土方、堆置、臨時活動或短期作業。", "建議填寫值或選項": "是／否／未知"},
            {"欄位名稱": "短期事件是否影響內部空間", "填寫說明": "短期事件是否影響人們停留或活動的空間。", "建議填寫值或選項": "是／否／未知"},
            {"欄位名稱": "短期事件是否影響通行路徑", "填寫說明": "短期事件是否影響步行、通學、騎車或通行路線。", "建議填寫值或選項": "是／否／未知"},
            {"欄位名稱": "短期事件是否位於邊界或鄰近受體", "填寫說明": "短期事件是否靠近基地邊界、周界、住宅、學校或醫療院所。", "建議填寫值或選項": "是／否／未知"},
        ]
    )

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        template_df.to_excel(writer, index=False, sheet_name="批次上傳空白表單")
        field_notes.to_excel(writer, index=False, sheet_name="欄位填寫說明")
    return output.getvalue()
'''
code = code.replace(old, new)

# 2. Add scoring standards columns in single score table. replace build_single_score_table records adding 計分標準 maybe simpler after DataFrame created.
old = '''    return pd.DataFrame(records)
'''
new = '''    standards = {
        "空品壓力": "2分：工廠≥3或道路≤100m；1分：工廠1–2或道路101–300m；0分：其他",
        "敏感受體密度": "2分：≥3處；1分：1–2處；0分：0處",
        "鄉鎮人口暴露": "2分：≥1000人/km²；1分：500–999人/km²；0分：<500人/km²",
        "土地權屬單純度": "2分：公有或單一權屬；1分：公私混合或需協調；0分：私有或權屬複雜",
        "管理權責明確性": "2分：管理機關明確；1分：需確認或協調；0分：無明確管理單位",
        "可植栽空間": "2分：>1ha；1分：>0.5–1ha；0分：≤0.5ha",
        "短期推動性": "2分：可短期推動；1分：需部分協調；0分：短期難以推動",
        "開放可及性": "2分：完全開放；1分：部分開放；0分：不開放",
        "鄰近生活節點": "2分：≥4處；1分：2–3處；0分：0–1處",
        "節點潛力總分": "高潛力：14–18；中潛力：7–13；基礎潛力：0–6",
        "最近綠化單元距離": "2分：≤100m；1分：>100–300m；0分：>300m",
        "500m內其他綠化單元數": "2分：≥4處；1分：2–3處；0分：0–1處",
        "1000m內其他綠化單元數": "2分：≥8處；1分：4–7處；0分：0–3處",
        "1000m內其他綠化單元總面積": "2分：≥10ha；1分：3–9.99ha；0分：<3ha",
        "串聯潛力總分": "高潛力：6–8；中潛力：3–5；基礎潛力：0–2",
    }
    out = pd.DataFrame(records)
    out["計分標準"] = out["評分項目"].map(standards).fillna("")
    return out
'''
# replace first occurrence after build_single_score_table only; this string repeats? Use replace count=1 after function area
idx = code.index("def build_single_score_table")
idx2 = code.index("def render_nearby_object_tables")
section = code[idx:idx2].replace(old, new, 1)
code = code[:idx] + section + code[idx2:]

# 3. Add push order function before render_batch_summary
insert = r'''
def classify_push_order(recommendation: str) -> str:
    """將優先推動建議簡化為批次統計用的推動順序分類。"""
    if not isinstance(recommendation, str):
        return "未分類"
    if "建議優先推動" in recommendation:
        return "建議優先推動"
    if "建議優先強化基地功能" in recommendation:
        return "建議優先強化基地功能"
    if "建議作為綠網連接補點" in recommendation:
        return "建議作為綠網連接補點"
    if "建議納入第二階段評估" in recommendation:
        return "建議納入第二階段評估"
    if "建議作為局部連接或補強場址" in recommendation:
        return "建議作為局部連接或補強場址"
    if "建議視管理可行性再評估" in recommendation:
        return "建議視管理可行性再評估"
    if "建議暫列低優先序" in recommendation:
        return "建議暫列低優先序"
    return "其他"

'''
code = code.replace("def render_batch_summary", insert + "def render_batch_summary")

# 4. Replace priority stats code in render_batch_summary
old = '''    with chart_col3:
        st.write("優先推動建議統計")
        priority_counts = result_df["優先推動建議"].value_counts()
        st.bar_chart(priority_counts)
        st.caption("此統計整合節點潛力、串聯潛力與功能情境，可作為批次場址排序與後續分群檢討依據。")
'''
new = '''    with chart_col3:
        st.write("推動順序統計")
        push_order = result_df["優先推動建議"].apply(classify_push_order)
        order = [
            "建議優先推動",
            "建議優先強化基地功能",
            "建議作為綠網連接補點",
            "建議納入第二階段評估",
            "建議作為局部連接或補強場址",
            "建議視管理可行性再評估",
            "建議暫列低優先序",
            "其他",
        ]
        st.bar_chart(push_order.value_counts().reindex(order).dropna())
        st.caption("依優先推動建議整理為推動順序分類，方便批次場址進行排序與分群。")
'''
code = code.replace(old, new)

# 5. Single input manager selectbox
old = '''        with n1:
            road_distance = st.number_input("距主要道路距離(公尺)", min_value=0.0, value=999.0, step=10.0)
            land_ownership = st.selectbox("土地權屬", ["公有或單一權屬", "公私混合或需協調", "私有或權屬複雜", "其他／待確認"])

        with n2:
            management_agency = st.text_input("管理機關", value="")
            area = st.number_input("基地面積(公頃)", min_value=0.0, value=0.0, step=0.1)
'''
new = '''        with n1:
            road_distance = st.number_input("距主要道路距離(公尺)", min_value=0.0, value=999.0, step=10.0)
            land_ownership = st.selectbox("土地權屬", ["公有或單一權屬", "公私混合或需協調", "私有或權屬複雜", "其他／待確認"])

        with n2:
            management_agency = st.selectbox(
                "管理機關",
                ["環境部", "縣市環保局", "鄉鎮市區公所", "學校", "醫療院所", "其他機關", "待確認"],
            )
            area = st.number_input("基地面積(公頃)", min_value=0.0, value=0.0, step=0.1)
'''
code = code.replace(old, new)

# 6. Sidebar add data sources and excel download
old = '''with st.sidebar:
    st.header("空白制式表單")
    st.write("可下載空白表單後，填寫批次場址資料再上傳分析。")
    st.download_button(
        label="下載批次上傳空白表單 CSV",
        data=make_blank_template_csv(),
        file_name="城市空品植生綠網淨化單元_批次上傳空白表單.csv",
        mime="text/csv",
    )

    st.divider()
    st.header("外部資料")
    st.write("系統會自動補算對應指標。")
'''
new = '''with st.sidebar:
    st.header("空白制式表單")
    st.write("可下載空白表單後，填寫批次場址資料再上傳分析。")

    template_excel = make_blank_template_excel()
    if template_excel is not None:
        st.download_button(
            label="下載批次上傳空白表單 Excel（含欄位說明）",
            data=template_excel,
            file_name="城市空品植生綠網淨化單元_批次上傳空白表單.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    st.download_button(
        label="下載批次上傳空白表單 CSV",
        data=make_blank_template_csv(),
        file_name="城市空品植生綠網淨化單元_批次上傳空白表單.csv",
        mime="text/csv",
    )

    st.divider()
    st.header("外部資料")
    st.write("系統會自動補算對應指標。")
    with st.expander("自動帶入資料來源"):
        st.markdown(
            """
            - **人口密度**：114年人口密度資料，1150427下載  
            - **工廠資料**：環境保護許可管理系統（暨解除列管）對象基本資料－曾有空列管，1150427下載  
            - **空品淨化區**：114年第四季季報  
            - **綠牆**：113年第四季季報，面積由平方公尺換算為公頃  
            - **各級學校**：111學年度各級學校名錄  
            - **醫療院所**：112年12月醫療院所分布圖  
            """
        )
'''
code = code.replace(old, new)

# 7. Ensure county/town filtering robust. current code already filters, but issue due whitespace maybe. patch with stripped temp.
old = '''            if has_town_lookup:
                county_options = sorted(town_lookup["縣市"].dropna().astype(str).unique())
                county = st.selectbox("縣市", county_options)
                town_options = sorted(
                    town_lookup.loc[town_lookup["縣市"].astype(str) == str(county), "鄉鎮市區"]
                    .dropna()
                    .astype(str)
                    .unique()
                )
                town = st.selectbox("鄉鎮市區", town_options)
'''
new = '''            if has_town_lookup:
                town_lookup_display = town_lookup.copy()
                town_lookup_display["縣市"] = town_lookup_display["縣市"].astype(str).str.strip()
                town_lookup_display["鄉鎮市區"] = town_lookup_display["鄉鎮市區"].astype(str).str.strip()

                county_options = sorted(town_lookup_display["縣市"].dropna().unique())
                county = st.selectbox("縣市", county_options)

                town_options = sorted(
                    town_lookup_display.loc[town_lookup_display["縣市"] == str(county).strip(), "鄉鎮市區"]
                    .dropna()
                    .unique()
                )
                town = st.selectbox("鄉鎮市區", town_options)
'''
code = code.replace(old, new)

out = Path("/mnt/data/app_updated_v2.py")
out.write_text(code, encoding="utf-8")
py_compile.compile(str(out), doraise=True)
print("created", out)
