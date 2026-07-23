import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO
import re
import os
import unicodedata

st.set_page_config(page_title="課表彙整系統", layout="wide")

# ============================================================
# 常數設定
# ============================================================
DAYS = range(1, 6)      # 週一 ~ 週五
PERIODS = range(1, 9)   # 第1節 ~ 第8節

TEMPLATE_FILES = {
    "class": "班級樣板.docx",
    "teacher": "教師樣板.docx",
}

DOWNLOAD_TEMPLATES = {
    "1. 配課表範本": "配課表.xlsx",
    "2. 課表範本": "課表.xlsx",
    "3. 教師排序表範本": "教師排序表.xlsx",
}

# 支援多種常見的星期表示法（全形數字未涵蓋，如有需要可再擴充）
DAY_MAP = {
    "一": 1, "二": 2, "三": 3, "四": 4, "五": 5,
    "週一": 1, "週二": 2, "週三": 3, "週四": 4, "週五": 5,
    "周一": 1, "周二": 2, "周三": 3, "周四": 4, "周五": 5,
    "1": 1, "2": 2, "3": 3, "4": 4, "5": 5,
    "mon": 1, "tue": 2, "wed": 3, "thu": 4, "fri": 5,
}


class DataValidationError(Exception):
    """
    資料格式或內容有問題時拋出。會被主流程攔截並顯示成使用者看得懂的訊息，
    而不是讓 Streamlit 顯示整串 Python 錯誤堆疊。
    """
    pass


# ============================================================
# 共用工具函式
# ============================================================
def normalize_name(name) -> str:
    """去除頭尾空白與全形空白，避免同一位教師因空白差異被誤判成不同人。"""
    if name is None:
        return ""
    return str(name).strip().replace("　", "").replace(" ", "")


def smart_read_csv(uploaded_file):
    """
    依序嘗試常見編碼讀取 CSV。
    台灣 Excel 匯出的 CSV 常見為 Big5／UTF-8-BOM，預設 utf-8 常會直接噴錯，
    這裡自動輪流嘗試，減少使用者因編碼問題而卡關。
    """
    encodings = ["utf-8-sig", "utf-8", "cp950", "big5"]
    last_err = None
    for enc in encodings:
        try:
            uploaded_file.seek(0)
            return pd.read_csv(uploaded_file, encoding=enc)
        except (UnicodeDecodeError, UnicodeError) as e:
            last_err = e
            continue
    raise DataValidationError(
        f"無法辨識 CSV 檔案編碼（已嘗試 UTF-8 / Big5 等格式）。"
        f"建議用 Excel 開啟後另存新檔為 .xlsx 再上傳。\n技術訊息：{last_err}"
    )


def read_uploaded_table(uploaded_file, file_label):
    """統一讀取上傳的 Excel/CSV，失敗時拋出好懂的錯誤訊息並清理欄名空白。"""
    try:
        if uploaded_file.name.lower().endswith(".csv"):
            df = smart_read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
    except DataValidationError:
        raise
    except Exception as e:
        raise DataValidationError(
            f"讀取「{file_label}」失敗，請確認檔案未損毀且格式正確。\n技術訊息：{e}"
        )

    df.columns = [str(c).strip() for c in df.columns]
    if df.empty:
        raise DataValidationError(f"「{file_label}」讀取成功，但內容是空的，請確認檔案內有資料。")
    return df


def validate_columns(df, required_cols, file_label):
    """檢查必要欄位是否存在，缺少時拋出清楚列出缺少哪些欄位的錯誤。"""
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        raise DataValidationError(
            f"「{file_label}」缺少必要欄位：{'、'.join(missing)}\n"
            f"目前偵測到的欄位為：{'、'.join(str(c) for c in df.columns)}"
        )


def load_default_template(file_name):
    """讀取內建樣板（編碼相容版），支援 NFC/NFD 檔名正規化比對。"""
    try:
        current_dir = os.path.dirname(os.path.abspath(__file__))
        full_path = os.path.join(current_dir, file_name)

        if os.path.exists(full_path):
            with open(full_path, "rb") as f:
                return f.read()

        target_filename_nfc = unicodedata.normalize("NFC", file_name)
        for actual_file in os.listdir(current_dir):
            actual_file_nfc = unicodedata.normalize("NFC", actual_file)
            if actual_file_nfc == target_filename_nfc:
                with open(os.path.join(current_dir, actual_file), "rb") as f:
                    return f.read()
    except Exception:
        pass
    return None


def apply_replacements(doc_obj, replacements: dict):
    """
    一次套用一整組佔位符取代，取代舊版「每個佔位符各自呼叫一次、
    每次都重新掃描整份文件」的做法。

    效能重點：
      1. 段落＋表格清單只蒐集一次（舊版每呼叫一次 master_replace 就重蒐集一次）。
      2. 對每個段落先做一次「{{」快速排除，沒有佔位符的段落直接跳過，
         不必逐一比對 80 幾個佔位符字串。
      3. 對合併大量班級／教師課表的批次匯出來說，效果最明顯。
    """
    targets = list(doc_obj.paragraphs)
    for table in doc_obj.tables:
        for row in table.rows:
            for cell in row.cells:
                targets.extend(cell.paragraphs)

    formatted = {}
    for old_text, new_text in replacements.items():
        if isinstance(new_text, (float, int)):
            formatted[old_text] = str(int(new_text))
        else:
            formatted[old_text] = str(new_text) if (new_text and str(new_text).strip() != "") else ""

    for p in targets:
        full_text = "".join(run.text for run in p.runs)
        if "{{" not in full_text:
            continue
        updated_text = full_text
        for old_text, new_val in formatted.items():
            if old_text in updated_text:
                updated_text = updated_text.replace(old_text, new_val)
        if updated_text != full_text:
            for i, run in enumerate(p.runs):
                run.text = updated_text if i == 0 else ""


def build_class_replacements(cname, class_data, tutors_map):
    """組出單一班級文件所需的所有佔位符取代字典。"""
    repl = {
        "{{CLASS}}": cname,
        "{{TUTOR}}": tutors_map.get(cname, "未設定"),
    }
    for d in DAYS:
        for p in PERIODS:
            v = class_data.get(cname, {}).get((d, p), {"subj": "", "teacher": ""})
            repl[f"{{{{SD{d}P{p}}}}}"] = v["subj"]
            repl[f"{{{{TD{d}P{p}}}}}"] = v["teacher"]
    return repl


def build_teacher_replacements(tname, teacher_data, base_hours, total_counts):
    """組出單一教師文件所需的所有佔位符取代字典。"""
    base = int(base_hours.get(tname, 0))
    total = int(total_counts.get(tname, 0))
    repl = {
        "{{TEACHER}}": tname,
        "{{BASE}}": base,
        "{{TOTAL}}": total,
        "{{EXTRA}}": total - base,
    }
    for d in DAYS:
        for p in PERIODS:
            v = teacher_data.get(tname, {}).get((d, p), {"subj": "", "class": ""})
            repl[f"{{{{CD{d}P{p}}}}}"] = v["class"]
            repl[f"{{{{SD{d}P{p}}}}}"] = v["subj"]
    return repl


def generate_document(template_bytes, replacements):
    doc = Document(BytesIO(template_bytes))
    apply_replacements(doc, replacements)
    return doc


def doc_to_bytes(doc):
    buf = BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ============================================================
# 資料解析
# ============================================================
def parse_assignment(df_assign):
    """
    解析配課表，支援兩種格式：
      模式 A：一維清單格式（需有「科目」「教師」欄位）
      模式 B：二維矩陣格式（科目當欄位名稱，儲存格內容為授課教師）

    回傳：
      assign_lookup: 明細列表 [{'c':班級,'s':科目,'t':教師}, ...]
      assign_dict:   {(班級,科目): [教師,...]}，供課表比對時 O(1) 查詢（效能優化）
      tutors:        {班級: 導師}
      all_teachers_db: 教師名稱集合
      warnings:      解析過程中的提醒訊息
    """
    warnings = []
    validate_columns(df_assign, ["班級"], "配課表")

    assign_lookup = []
    assign_dict = {}
    all_teachers_db = set()
    tutors = {}

    def register(c, s, t_raw):
        t_list = [normalize_name(x) for x in str(t_raw).split("/")]
        valid = [t for t in t_list if t and t != "nan"]
        for t in valid:
            assign_lookup.append({"c": c, "s": s, "t": t})
            all_teachers_db.add(t)
        if valid:
            assign_dict.setdefault((c, s), [])
            assign_dict[(c, s)].extend(valid)
        return valid

    if "科目" in df_assign.columns and "教師" in df_assign.columns:
        # 【模式 A】傳統一維清單格式
        for _, row in df_assign.iterrows():
            c = str(row["班級"]).strip()
            s = str(row["科目"]).strip()
            t_raw = str(row.get("教師", "")).strip()
            register(c, s, t_raw)
            if s == "班級":
                tutors[c] = t_raw
    else:
        # 【模式 B】二維矩陣表格格式
        if "導師" in df_assign.columns:
            for _, row in df_assign.iterrows():
                c = str(row["班級"]).strip()
                t_tutor = str(row["導師"]).strip()
                if t_tutor and t_tutor != "nan":
                    tutors[c] = t_tutor

        subject_cols = [col for col in df_assign.columns if col not in ["班級", "導師"]]
        if not subject_cols:
            raise DataValidationError("「配課表」除了班級／導師欄位外，找不到任何科目欄位，請確認檔案格式。")

        df_melted = pd.melt(df_assign, id_vars=["班級"], value_vars=subject_cols,
                             var_name="科目", value_name="教師")
        for _, row in df_melted.iterrows():
            c = str(row["班級"]).strip()
            s = str(row["科目"]).strip()
            t_raw = str(row["教師"]).strip() if pd.notna(row["教師"]) else ""
            if not t_raw or t_raw == "nan":
                continue
            register(c, s, t_raw)

    if not assign_lookup:
        raise DataValidationError("配課表中沒有解析出任何有效的教師配課資料，請確認欄位內容是否正確。")

    return assign_lookup, assign_dict, tutors, all_teachers_db, warnings


def parse_teacher_sort(df_sort, all_teachers_list):
    """
    解析教師排序暨時數表。若使用者未上傳，則以配課表中出現的教師依字母排序、
    應授時數預設為 0。教師姓名比對時會忽略頭尾／全形空白差異。
    """
    warnings = []
    ordered_teachers, base_hours = [], {}

    if df_sort is not None:
        if df_sort.shape[1] < 2:
            raise DataValidationError("「教師排序暨時數表」至少需要兩欄（教師姓名、應授時數）。")

        norm_to_actual = {normalize_name(t): t for t in all_teachers_list}
        for _, s_row in df_sort.iterrows():
            raw_name = str(s_row.iloc[0]).strip()
            key = normalize_name(raw_name)
            if key in norm_to_actual:
                t_name = norm_to_actual[key]
                if t_name not in ordered_teachers:
                    ordered_teachers.append(t_name)
                    try:
                        base_hours[t_name] = int(s_row.iloc[1])
                    except Exception:
                        base_hours[t_name] = 0
            elif raw_name and raw_name != "nan":
                warnings.append(f"教師排序表中的「{raw_name}」在配課表中查無對應資料，已略過。")

        for t in all_teachers_list:
            if t not in ordered_teachers:
                ordered_teachers.append(t)
                base_hours[t] = 0
    else:
        ordered_teachers = sorted(all_teachers_list)
        base_hours = {t: 0 for t in ordered_teachers}

    return ordered_teachers, base_hours, warnings


def parse_schedule(df_time, assign_dict):
    """
    解析課表，並與配課表（assign_dict）比對出每一節課的授課教師。
    使用字典查詢取代舊版逐筆掃描 assign_lookup 的線性比對，是效能優化的重點之一。
    """
    warnings = []
    validate_columns(df_time, ["班級", "科目", "星期", "節次"], "課表")

    class_data, teacher_data, total_counts = {}, {}, {}
    unmatched_format = 0

    for _, row in df_time.iterrows():
        c_raw = str(row["班級"]).strip()
        s_raw = str(row["科目"]).strip()
        day_key = str(row["星期"]).strip().lower()
        d = DAY_MAP.get(day_key, 0)
        p_match = re.search(r"\d+", str(row["節次"]))

        if not (p_match and d > 0):
            unmatched_format += 1
            continue
        p = int(p_match.group())
        if p not in PERIODS:
            continue

        if not s_raw or s_raw == "nan":
            display_t, s_raw, curr_t_list = "", "", []
        else:
            curr_t_list = assign_dict.get((c_raw, s_raw), [])
            display_t = "/".join(curr_t_list) if curr_t_list else "未知教師"
            if not curr_t_list:
                warnings.append(f"{c_raw} 週{d}第{p}節「{s_raw}」在配課表中查無對應教師。")

        class_data.setdefault(c_raw, {})[(d, p)] = {"subj": s_raw, "teacher": display_t}
        for t in curr_t_list:
            teacher_data.setdefault(t, {})[(d, p)] = {"subj": s_raw, "class": c_raw}
            total_counts[t] = total_counts.get(t, 0) + 1

    if unmatched_format:
        warnings.append(f"共有 {unmatched_format} 筆課表資料的「星期」或「節次」格式無法辨識，已自動略過。")

    if not class_data:
        raise DataValidationError("課表中沒有解析出任何有效的班級課程資料，請確認欄位內容與格式是否正確。")

    return class_data, teacher_data, total_counts, warnings


# ============================================================
# 側邊欄：資料管理
# ============================================================
with st.sidebar:
    st.header("⚙️ 資料管理")

    if "confirm_reset" not in st.session_state:
        st.session_state.confirm_reset = False

    if not st.session_state.confirm_reset:
        if st.button("🧹 清空重置系統"):
            st.session_state.confirm_reset = True
            st.rerun()
    else:
        st.warning("確定要清空所有已匯入的資料嗎？此動作無法復原。")
        cc1, cc2 = st.columns(2)
        if cc1.button("✅ 確定清空"):
            for key in list(st.session_state.keys()):
                del st.session_state[key]
            st.rerun()
        if cc2.button("↩️ 取消"):
            st.session_state.confirm_reset = False
            st.rerun()

    st.divider()
    st.subheader("📥 範本下載")
    for label, file_name in DOWNLOAD_TEMPLATES.items():
        try:
            current_dir = os.path.dirname(os.path.abspath(__file__))
            f_path = os.path.join(current_dir, file_name)
            if os.path.exists(f_path):
                with open(f_path, "rb") as f:
                    st.download_button(label=label, data=f, file_name=file_name, key=f"dl_{file_name}")
            else:
                st.caption(f"⚠️ 找不到 {file_name}")
        except Exception:
            st.caption(f"⚠️ 讀取 {file_name} 失敗")
    st.divider()

    st.subheader("📤 上傳資料檔")
    f_assign = st.file_uploader("1. 上傳【配課表】", type=["xlsx", "csv"])
    f_time = st.file_uploader("2. 上傳【課表】", type=["xlsx", "csv"])
    f_sort = st.file_uploader("3. 上傳【教師排序暨時數表】(選填)", type=["xlsx", "csv"])

    if f_assign and f_time and st.button("🚀 執行整合"):
        class_temp = load_default_template(TEMPLATE_FILES["class"])
        teacher_temp = load_default_template(TEMPLATE_FILES["teacher"])

        if not class_temp or not teacher_temp:
            current_dir = os.path.dirname(os.path.abspath(__file__))
            actual_files = os.listdir(current_dir) if os.path.exists(current_dir) else []
            st.error("❌ 系統錯誤：後台找不到「班級樣板.docx」或「教師樣板.docx」，請確認 GitHub 檔案。")
            st.info(f"🔍 雲端伺服器目前實際看到的檔案清單：\n`{actual_files}`")
        else:
            try:
                with st.spinner("同步內建樣板與解析資料中..."):
                    df_assign = read_uploaded_table(f_assign, "配課表")
                    df_time = read_uploaded_table(f_time, "課表")
                    df_sort = read_uploaded_table(f_sort, "教師排序暨時數表") if f_sort else None

                    assign_lookup, assign_dict, tutors, all_teachers_db, w1 = parse_assignment(df_assign)
                    all_teachers_list = list(all_teachers_db)
                    ordered_teachers, base_hours, w2 = parse_teacher_sort(df_sort, all_teachers_list)
                    class_data, teacher_data, total_counts, w3 = parse_schedule(df_time, assign_dict)
                    all_warnings = w1 + w2 + w3

                st.session_state.update({
                    "class_template": class_temp,
                    "teacher_template": teacher_temp,
                    "df_assign": df_assign,
                    "class_data": class_data,
                    "teacher_data": teacher_data,
                    "tutors_map": tutors,
                    "base_hours": base_hours,
                    "total_counts": total_counts,
                    "ordered_teachers": ordered_teachers,
                    "import_warnings": all_warnings,
                    "sel_class": sorted(class_data.keys())[0],
                    "sel_teacher": ordered_teachers[0],
                })
                st.session_state["last_import_summary"] = (
                    f"✅ 整合完成！共解析 {len(class_data)} 個班級、"
                    f"{len(ordered_teachers)} 位教師、{len(assign_lookup)} 筆配課紀錄。"
                )
                st.rerun()
            except DataValidationError as e:
                st.error(f"❌ 資料檢查未通過：\n\n{e}")
            except Exception as e:
                st.error(f"❌ 發生未預期的錯誤，請確認檔案格式是否正確。\n\n技術訊息：{e}")

# ============================================================
# 主介面與預覽
# ============================================================
if "class_data" in st.session_state:

    # 匯入成功摘要（只顯示一次，切換分頁後不再重複出現）
    summary = st.session_state.pop("last_import_summary", None)
    if summary:
        st.success(summary)

    import_warnings = st.session_state.get("import_warnings", [])
    if import_warnings:
        with st.expander(f"⚠️ 匯入時偵測到 {len(import_warnings)} 筆提醒事項，點擊展開查看", expanded=False):
            for w in import_warnings[:50]:
                st.write(f"- {w}")
            if len(import_warnings) > 50:
                st.caption(f"…等共 {len(import_warnings)} 筆，僅顯示前 50 筆。")

    tab1, tab2, tab3, tab4 = st.tabs(
        ["🏫 班級課表", "👩‍🏫 教師課表", "📋 配課總覽與分頁匯出", "🔍 課表搜尋"]
    )

    with tab1:
        classes = sorted(st.session_state.class_data.keys())
        curr_c = st.session_state.get("sel_class", classes[0])
        col1, col2, col3 = st.columns([1, 2, 1])
        if col1.button("⬅️ 上一班"):
            st.session_state.sel_class = classes[(classes.index(curr_c) - 1) % len(classes)]
            st.rerun()
        if col3.button("下一班 ➡️"):
            st.session_state.sel_class = classes[(classes.index(curr_c) + 1) % len(classes)]
            st.rerun()
        with col2:
            st.session_state.sel_class = st.selectbox("選取班級", classes, index=classes.index(curr_c))

        target_c = st.session_state.sel_class
        st.info(f"📍 班級：{target_c} ｜ 導師：{st.session_state.tutors_map.get(target_c, '未設定')}")

        c_preview = []
        for p in PERIODS:
            row = {"節次": f"第 {p} 節"}
            for d in DAYS:
                info = st.session_state.class_data[target_c].get((d, p))
                row[f"週{d}"] = f"{info['subj']}\n({info['teacher']})" if info else ""
            c_preview.append(row)
        st.table(pd.DataFrame(c_preview))

        bc1, bc2 = st.columns(2)
        with bc1:
            if st.button(f"📥 下載 {target_c} 課表"):
                repl = build_class_replacements(target_c, st.session_state.class_data, st.session_state.tutors_map)
                doc = generate_document(st.session_state.class_template, repl)
                st.download_button(f"💾 儲存 {target_c} 課表", doc_to_bytes(doc), f"{target_c}_班級課表.docx")
        with bc2:
            sel_c_batch = st.multiselect("勾選批次合併", classes, default=classes)
            if st.button("🚀 執行班級合併列印"):
                main_doc = None
                with st.spinner(f"正在產生 {len(sel_c_batch)} 個班級的課表…"):
                    for i, cname in enumerate(sel_c_batch):
                        repl = build_class_replacements(cname, st.session_state.class_data, st.session_state.tutors_map)
                        tmp = generate_document(st.session_state.class_template, repl)
                        if i == 0:
                            main_doc = tmp
                        else:
                            for el in tmp.element.body:
                                main_doc.element.body.append(el)
                if main_doc:
                    st.download_button("💾 下載班級彙整檔", doc_to_bytes(main_doc), "全校班級課表.docx")

    with tab2:
        teachers = st.session_state.ordered_teachers
        curr_t = st.session_state.get("sel_teacher", teachers[0])
        colt1, colt2, colt3 = st.columns([1, 2, 1])
        if colt1.button("⬅️ 前一位"):
            st.session_state.sel_teacher = teachers[(teachers.index(curr_t) - 1) % len(teachers)]
            st.rerun()
        if colt3.button("下一位 ➡️"):
            st.session_state.sel_teacher = teachers[(teachers.index(curr_t) + 1) % len(teachers)]
            st.rerun()
        with colt2:
            st.session_state.sel_teacher = st.selectbox("跳轉教師", teachers, index=teachers.index(curr_t))

        target_t = st.session_state.sel_teacher
        base = int(st.session_state.base_hours.get(target_t, 0))
        total = int(st.session_state.total_counts.get(target_t, 0))
        m1, m2, m3 = st.columns(3)
        m1.metric("應授時數", f"{base} 節")
        m2.metric("教學總時數", f"{total} 節")
        m3.metric("兼代課時數", f"{total - base} 節")

        t_prev = [
            {
                "節次": f"第 {p} 節",
                **{
                    f"週{d}": f"{st.session_state.teacher_data.get(target_t, {}).get((d, p), {}).get('class', '')} "
                              f"{st.session_state.teacher_data.get(target_t, {}).get((d, p), {}).get('subj', '')}".strip()
                    for d in DAYS
                },
            }
            for p in PERIODS
        ]
        st.table(pd.DataFrame(t_prev))

        bt1, bt2 = st.columns(2)
        with bt1:
            if st.button(f"📥 下載 {target_t} 課表"):
                repl = build_teacher_replacements(
                    target_t, st.session_state.teacher_data, st.session_state.base_hours, st.session_state.total_counts
                )
                doc = generate_document(st.session_state.teacher_template, repl)
                st.download_button(f"💾 儲存 {target_t} 課表", doc_to_bytes(doc), f"{target_t}_教師課表.docx")
        with bt2:
            sel_t_batch = st.multiselect("批次合併教師", teachers, default=teachers)
            if st.button("🚀 執行教師合併列印"):
                main_doc = None
                with st.spinner(f"正在產生 {len(sel_t_batch)} 位教師的課表…"):
                    for i, tname in enumerate(sel_t_batch):
                        repl = build_teacher_replacements(
                            tname, st.session_state.teacher_data, st.session_state.base_hours, st.session_state.total_counts
                        )
                        tmp = generate_document(st.session_state.teacher_template, repl)
                        if i == 0:
                            main_doc = tmp
                        else:
                            for el in tmp.element.body:
                                main_doc.element.body.append(el)
                if main_doc:
                    st.download_button("💾 下載教師彙整檔", doc_to_bytes(main_doc), "全校教師課表_彙整.docx")

    with tab3:
        st.header("📋 全校配課資料總覽")
        if "df_assign" in st.session_state:
            st.write(
                "💡 **提示**：下方顯示您所上傳的原始單一配課表。點擊最下方的按鈕，系統將自動依據"
                "「班級」欄位將資料拆分，產出含有**多個班級分頁**的 Excel 活頁簿。"
            )
            st.dataframe(st.session_state.df_assign, use_container_width=True)
            st.divider()
            st.subheader("📥 匯出「一班一工作表」Excel 檔案")

            buf_excel = BytesIO()
            with pd.ExcelWriter(buf_excel, engine="openpyxl") as writer:
                for cname, group in st.session_state.df_assign.groupby("班級"):
                    clean_sheet_name = str(cname).strip()
                    clean_sheet_name = re.sub(r"[\\/*?:\[\]]", "", clean_sheet_name)[:31]
                    if not clean_sheet_name:
                        clean_sheet_name = "未命名班級"
                    group.to_excel(writer, sheet_name=clean_sheet_name, index=False)

            st.download_button(
                label="💾 點我下載「各班級獨立分頁」配課明細表.xlsx",
                data=buf_excel.getvalue(),
                file_name="全校各班級配課明細表_分頁版.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    with tab4:
        st.header("🔍 課表關鍵字搜尋")
        st.caption("輸入班級、教師或科目名稱（支援部分關鍵字），快速找出所有符合的排課紀錄。")
        keyword = st.text_input("搜尋關鍵字", placeholder="例如：王小明、302、數學")

        if keyword:
            kw = keyword.strip()
            rows = []
            for c, periods in st.session_state.class_data.items():
                for (d, p), info in periods.items():
                    subj, teacher = info.get("subj", ""), info.get("teacher", "")
                    if kw in c or kw in subj or kw in teacher:
                        rows.append({
                            "班級": c, "_d": d, "_p": p,
                            "星期": f"週{d}", "節次": f"第{p}節",
                            "科目": subj, "教師": teacher,
                        })
            if rows:
                result_df = (
                    pd.DataFrame(rows)
                    .sort_values(["班級", "_d", "_p"])
                    .drop(columns=["_d", "_p"])
                    .reset_index(drop=True)
                )
                st.success(f"共找到 {len(result_df)} 筆符合「{keyword}」的紀錄")
                st.dataframe(result_df, use_container_width=True, hide_index=True)
            else:
                st.warning(f"找不到符合「{keyword}」的紀錄，請確認關鍵字是否正確。")
        else:
            st.info("請輸入關鍵字開始搜尋。")

else:
    st.info("👋 請上傳資料檔並點擊執行整合。")
