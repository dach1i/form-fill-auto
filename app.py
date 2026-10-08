import streamlit as st
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import datetime
import io
import os
import copy
import json
import re
import subprocess

st.set_page_config(page_title="DPR Generator", page_icon="🏗️", layout="wide")
st.title("🏗️ ყოველდღიური რეპორტის გენერატორი (DPR)")
st.caption("სითი მოლი საბურთალო — სრული ავტომატიზაცია")

target_file = None
for fname in ["template.docx", "CMC-CMS-DPR-001-20261001-Rev00.docx"]:
    if os.path.exists(fname):
        target_file = fname
        break

if not target_file:
    st.error("ვერ მოიძებნა 'template.docx' ან 'CMC-CMS-DPR-001-20261001-Rev00.docx'!")
    st.stop()

# ==================== AUTOSAVE DRAFT MANAGER ====================
AUTOSAVE_FILE = "dpr_autosave.json"

def load_draft():
    if os.path.exists(AUTOSAVE_FILE):
        try:
            with open(AUTOSAVE_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def save_draft(data):
    try:
        with open(AUTOSAVE_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

draft = load_draft()

with st.sidebar:
    st.header("⚙️ მართვის პანელი")
    if st.button("🔄 ახალი რეპორტის დაწყება (Reset Draft)", use_container_width=True):
        if os.path.exists(AUTOSAVE_FILE):
            os.remove(AUTOSAVE_FILE)
        st.session_state.clear()
        st.rerun()
    st.caption("ყველა ცვლილება ინახება ავტომატურად.")

if "docx_bytes" not in st.session_state:
    st.session_state["docx_bytes"] = None
    st.session_state["pdf_bytes"] = None
    st.session_state["doc_code"] = None

def clone_row(table, source_row):
    tr_copy = copy.deepcopy(source_row._tr)
    for t in tr_copy.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t'):
        t.text = ''
    table._tbl.append(tr_copy)
    return table.rows[-1]

# Bulletproof placeholder replacement for paragraphs (handles split XML runs)
def replace_in_p(p, pattern, replacement):
    if re.search(pattern, p.text):
        matched = False
        for r in p.runs:
            if re.search(pattern, r.text):
                r.text = re.sub(pattern, str(replacement), r.text)
                matched = True
        if not matched and re.search(pattern, p.text):
            p.text = re.sub(pattern, str(replacement), p.text)

# Sets a compact, single-block header cell to avoid duplicates and container stretching
def set_cell_header_block(cell, label, value):
    tc = cell._tc
    for child in list(tc):
        if not child.tag.endswith('tcPr'):
            tc.remove(child)
    p = cell.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.05
    
    r1 = p.add_run(f"{label}\n")
    r1.font.color.rgb = RGBColor(255, 255, 255)
    r1.font.size = Pt(7.5)
    r1.font.bold = False
    
    r2 = p.add_run(str(value) if value else "-")
    r2.font.color.rgb = RGBColor(255, 255, 255)
    r2.font.size = Pt(8.5)
    r2.font.bold = True

# Weather cell formatter: strict padding and proportional font sizes
def set_clean_weather_cell(cell, text, font_size=8.5, bold=False):
    tc = cell._tc
    for child in list(tc):
        if not child.tag.endswith('tcPr'):
            tc.remove(child)
    p = cell.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.0
    r = p.add_run(str(text))
    r.font.size = Pt(font_size)
    r.font.bold = bold
    r.font.color.rgb = RGBColor(0, 0, 0)

# Wind speed cell formatter
def set_clean_wind_cell(cell, wind_val):
    tc = cell._tc
    for child in list(tc):
        if not child.tag.endswith('tcPr'):
            tc.remove(child)
    p = cell.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.0
    lbl = p.add_run("ქარის სიჩქარე\nWind Speed:\n")
    lbl.font.size = Pt(7.5)
    r = p.add_run(str(wind_val))
    r.font.size = Pt(8.5)
    r.font.bold = True

# Generic compact cell populator for Data & Staff tables
def populate_compact_cell(cell, text, font_size=8.0, bold=False, align=WD_ALIGN_PARAGRAPH.LEFT):
    tc = cell._tc
    for child in list(tc):
        if not child.tag.endswith('tcPr'):
            tc.remove(child)
    p = cell.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = Pt(1.5)
    p.paragraph_format.space_after = Pt(1.5)
    p.paragraph_format.line_spacing = 1.05
    r = p.add_run(str(text))
    r.font.size = Pt(font_size)
    r.font.bold = bold
    r.font.color.rgb = RGBColor(0, 0, 0)

WEATHER_CONDITIONS = [
    "",
    "მზიანი/sunny",
    "ნაწ. მოღრ./p. cloudy",
    "მოღრუბლ./cloudy",
    "წვიმიანი/rainy",
    "თოვლიანი/snowy"
]

# ==================== INPUT FIELDS ====================
# 1. ძირითადი ინფორმაცია
st.subheader("1. ძირითადი ინფორმაცია / General Info")
col1, col2, col3, col4 = st.columns([1, 1, 1.5, 2])

rep_num = col1.text_input("რეპორტის №", value=draft.get("rep_num", ""))

raw_saved_date = draft.get("rep_date")
default_date = datetime.today().date()
if raw_saved_date:
    try:
        default_date = datetime.strptime(raw_saved_date, "%Y-%m-%d").date()
    except Exception:
        pass
rep_date = col2.date_input("თარიღი", value=default_date)
prepared_by = col3.text_input("მოამზადა / Prepared by", value=draft.get("prepared_by", ""))
custom_doc_code = col4.text_input("დოკუმენტის კოდი / Doc Code", value=draft.get("custom_doc_code", ""))

col_c1, col_c2 = st.columns(2)
client_name = col_c1.text_input("დამკვეთი / Client", value=draft.get("client_name", ""))
site_location = col_c2.text_input("ადგილმდებარეობა / Location", value=draft.get("site_location", ""))

st.divider()

# 2. ამინდი
st.subheader("2. ამინდი / Weather")
w1, w2, w3 = st.columns(3)

w1.markdown("#### 09:00")
t9 = w1.text_input("ტემპერატურა (09:00)", value=draft.get("t9", ""))
w9 = w1.text_input("ქარი (09:00)", value=draft.get("w9", ""))
saved_cond9 = draft.get("cond9", "")
cond9 = w1.selectbox("ამინდის პირობა (09:00)", WEATHER_CONDITIONS, index=WEATHER_CONDITIONS.index(saved_cond9) if saved_cond9 in WEATHER_CONDITIONS else 0)

w2.markdown("#### 14:00")
t14 = w2.text_input("ტემპერატურა (14:00)", value=draft.get("t14", ""))
w14 = w2.text_input("ქარი (14:00)", value=draft.get("w14", ""))
saved_cond14 = draft.get("cond14", "")
cond14 = w2.selectbox("ამინდის პირობა (14:00)", WEATHER_CONDITIONS, index=WEATHER_CONDITIONS.index(saved_cond14) if saved_cond14 in WEATHER_CONDITIONS else 0)

w3.markdown("#### 18:00")
t18 = w3.text_input("ტემპერატურა (18:00)", value=draft.get("t18", ""))
w18 = w3.text_input("ქარი (18:00)", value=draft.get("w18", ""))
saved_cond18 = draft.get("cond18", "")
cond18 = w3.selectbox("ამინდის პირობა (18:00)", WEATHER_CONDITIONS, index=WEATHER_CONDITIONS.index(saved_cond18) if saved_cond18 in WEATHER_CONDITIONS else 0)

st.divider()

# 3. სამუშაოები
st.subheader("3. სამუშაოების აღწერა / Daily Works")
saved_task_count = int(draft.get("task_count", 1))
task_count = st.number_input("სამუშაოების რაოდენობა", min_value=1, max_value=25, value=saved_task_count)

draft_tasks = draft.get("tasks", [])
tasks = []
for i in range(int(task_count)):
    st.markdown(f"**სამუშაო #{i+1}**")
    tc1, tc2, tc3, tc4, tc5 = st.columns([2, 1, 1, 2, 4])
    
    t_prev = draft_tasks[i] if i < len(draft_tasks) else {}
    def_name = t_prev.get("name", "")
    def_mp = t_prev.get("manpower", "")
    def_wbs = t_prev.get("wbs", "")
    def_co = t_prev.get("contractor", "")
    def_desc = t_prev.get("desc", "")

    w_name = tc1.text_input(f"დასახელება #{i+1}", value=def_name, key=f"wn_{i}")
    m_pow = tc2.text_input(f"პერსონალი #{i+1}", value=def_mp, key=f"mp_{i}")
    wbs_c = tc3.text_input(f"WBS #{i+1}", value=def_wbs, key=f"wbs_{i}")
    c_tor = tc4.text_input(f"კონტრაქტორი #{i+1}", value=def_co, key=f"co_{i}")
    desc = tc5.text_input(f"აღწერა #{i+1}", value=def_desc, key=f"de_{i}")
    tasks.append({"num": i+1, "contractor": c_tor, "manpower": m_pow, "wbs": wbs_c, "name": w_name, "desc": desc})

st.divider()

# 4. ფოტოები (Supports up to 100 photos)
st.subheader("4. სამუშაოების ფოტომასალა / Photos & Descriptions")
st.info("💡 თითო გვერდზე თავსდება 6 ფოტო. 7+ ფოტოს შემთხვევაში ავტომატურად შეიქმნება ახალი სრული გვერდი (მაქსიმუმ 100 ფოტო).")

saved_photo_slots = min(int(draft.get("photo_slots_count", 6)), 100)
photo_slots_count = st.number_input("რამდენი ფოტოს ატვირთვა გსურთ?", min_value=1, max_value=100, value=saved_photo_slots)
draft_caps = draft.get("captions", {})
photos_data = []

for pi in range(int(photo_slots_count)):
    p_c1, p_c2 = st.columns([2, 3])
    up_file = p_c1.file_uploader(f"ფოტო #{pi+1}", type=["jpg", "jpeg", "png"], key=f"photo_up_{pi}")
    def_cap = draft_caps.get(str(pi), "")
    p_caption = p_c2.text_input(f"აღწერა #{pi+1}", value=def_cap, key=f"photo_cap_{pi}")
    draft_caps[str(pi)] = p_caption
    if up_file is not None:
        photos_data.append({"file": up_file, "caption": p_caption, "num": pi + 1})

st.divider()

# 5. პერსონალი
st.subheader("5. სამშენებლო მოედანზე მომუშავე პერსონალი / Key Staff")
st.caption("ჩაწერეთ მხოლოდ დღეს მყოფი თანამშრომლები.")
saved_staff_count = int(draft.get("staff_count", 0))
staff_count = st.number_input("თანამშრომლების რაოდენობა", min_value=0, max_value=30, value=saved_staff_count)

draft_staff = draft.get("staff", [])
staff_members = []
for si in range(int(staff_count)):
    st.markdown(f"**თანამშრომელი #{si+1}**")
    sc1, sc2, sc3, sc4, sc5 = st.columns([2, 2, 2, 2, 3])
    st_prev = draft_staff[si] if si < len(draft_staff) else {}

    s_comp = sc1.text_input("კომპანია", value=st_prev.get("comp", ""), key=f"st_c_{si}")
    s_pos = sc2.text_input("პოზიცია", value=st_prev.get("pos", ""), key=f"st_p_{si}")
    s_name = sc3.text_input("სახელი და გვარი", value=st_prev.get("name", ""), key=f"st_n_{si}")
    s_phone = sc4.text_input("ტელ. ნომერი", value=st_prev.get("phone", ""), key=f"st_ph_{si}")
    s_mail = sc5.text_input("ელ. ფოსტა", value=st_prev.get("mail", ""), key=f"st_m_{si}")
    staff_members.append({"num": si+1, "comp": s_comp, "pos": s_pos, "name": s_name, "phone": s_phone, "mail": s_mail})

# Silent Autosave
current_state_to_save = {
    "rep_num": rep_num,
    "rep_date": rep_date.strftime("%Y-%m-%d"),
    "prepared_by": prepared_by,
    "custom_doc_code": custom_doc_code,
    "client_name": client_name,
    "site_location": site_location,
    "t9": t9, "w9": w9, "cond9": cond9,
    "t14": t14, "w14": w14, "cond14": cond14,
    "t18": t18, "w18": w18, "cond18": cond18,
    "task_count": int(task_count),
    "tasks": tasks,
    "photo_slots_count": int(photo_slots_count),
    "captions": draft_caps,
    "staff_count": int(staff_count),
    "staff": staff_members
}
save_draft(current_state_to_save)

st.write("")
generate_btn = st.button("🚀 რეპორტის შექმნა / Generate DPR", type="primary", use_container_width=True)

# ==================== DOCUMENT GENERATION ====================
if generate_btn:
    try:
        doc = docx.Document(target_file)
        formatted_date = rep_date.strftime("%Y/%m/%d")
        final_doc_code = custom_doc_code.strip() if custom_doc_code.strip() else f"CMC-CMS-DPR-{rep_num or '001'}-{rep_date.strftime('%Y%m%d')}-Rev00"

        # 1. Update Body Paragraphs
        for p in doc.paragraphs:
            if "CMC-CMS-DPR-" in p.text or "{{ doc_code }}" in p.text:
                p.text = ""
                r = p.add_run(final_doc_code)
                r.font.color.rgb = RGBColor(255, 255, 255)
                r.font.bold = True
                r.font.size = Pt(8.5)
            if p.text.strip() in ["001", "{{ rep_num }}", "111", str(rep_num)]:
                p.text = ""
                r = p.add_run(str(rep_num))
                r.font.color.rgb = RGBColor(255, 255, 255)
                r.font.bold = True
                r.font.size = Pt(9.5)

        # 2. Iterate Tables
        for table in doc.tables:
            t_text = " ".join([c.text for row in table.rows for c in row.cells])

            # A. HEADER CELLS (CLEAN SINGLE-BLOCK ASSIGNMENT, NO DUPLICATES)
            for row in table.rows:
                for cell in row.cells:
                    raw_c = cell.text.strip()

                    # Client
                    if ("დამკვეთი" in raw_c or "Client:" in raw_c) and ("ამინდი" not in raw_c and "Weather" not in raw_c):
                        set_cell_header_block(cell, "დამკვეთი / Client:", client_name)

                    # Date
                    elif ("თარიღი" in raw_c or "Date:" in raw_c) and ("ამინდი" not in raw_c and "Weather" not in raw_c):
                        set_cell_header_block(cell, "თარიღი / Date:", formatted_date)

                    # Location
                    elif "ადგილმდებარეობა" in raw_c or "Location:" in raw_c:
                        set_cell_header_block(cell, "ადგილმდებარეობა / Location:", site_location)

                    # Prepared by
                    elif "მოამზადა" in raw_c or "Prepared by" in raw_c:
                        set_cell_header_block(cell, "მოამზადა / Prepared by:", prepared_by)

                    # Document Code in table
                    elif "CMC-CMS-DPR-" in raw_c or "{{ doc_code }}" in raw_c:
                        for p in cell.paragraphs:
                            if "CMC-CMS-DPR-" in p.text or "{{ doc_code }}" in p.text:
                                p.text = ""
                                r = p.add_run(final_doc_code)
                                r.font.color.rgb = RGBColor(255, 255, 255)
                                r.font.bold = True
                                r.font.size = Pt(8.5)

                    # Report Number in table
                    elif raw_c in ["001", "{{ rep_num }}", "111", str(rep_num)]:
                        for p in cell.paragraphs:
                            if p.text.strip() in ["001", "{{ rep_num }}", "111", str(rep_num)]:
                                p.text = ""
                                r = p.add_run(str(rep_num))
                                r.font.color.rgb = RGBColor(255, 255, 255)
                                r.font.bold = True
                                r.font.size = Pt(9.5)

                    # Ensure white banner title text
                    if "ყოველდღიური რეპორტი" in raw_c or "Daily Progress Report" in raw_c:
                        for p in cell.paragraphs:
                            if "{{ rep_num }}" in p.text:
                                p.text = p.text.replace("{{ rep_num }}", str(rep_num))
                            for r in p.runs:
                                r.font.color.rgb = RGBColor(255, 255, 255)

            # B. WEATHER & WIND DIRECT REPLACEMENT (UNCONDITIONAL TAG SCANNER)
            for r_idx, row in enumerate(table.rows):
                row_raw = " ".join([c.text for c in row.cells])

                # Direct scan for tags
                for cell in row.cells:
                    c_txt = cell.text
                    if re.search(r'\{\{\s*w9\s*\}\}', c_txt):
                        set_clean_wind_cell(cell, w9)
                    elif re.search(r'\{\{\s*w14\s*\}\}', c_txt):
                        set_clean_wind_cell(cell, w14)
                    elif re.search(r'\{\{\s*w18\s*\}\}', c_txt):
                        set_clean_wind_cell(cell, w18)
                    elif re.search(r'\{\{\s*t9\s*\}\}', c_txt):
                        set_clean_weather_cell(cell, t9, font_size=9.5, bold=True)
                    elif re.search(r'\{\{\s*t14\s*\}\}', c_txt):
                        set_clean_weather_cell(cell, t14, font_size=9.5, bold=True)
                    elif re.search(r'\{\{\s*t18\s*\}\}', c_txt):
                        set_clean_weather_cell(cell, t18, font_size=9.5, bold=True)
                    elif re.search(r'\{\{\s*cond9\s*\}\}', c_txt):
                        set_clean_weather_cell(cell, cond9, font_size=8.5, bold=False)
                    elif re.search(r'\{\{\s*cond14\s*\}\}', c_txt):
                        set_clean_weather_cell(cell, cond14, font_size=8.5, bold=False)
                    elif re.search(r'\{\{\s*cond18\s*\}\}', c_txt):
                        set_clean_weather_cell(cell, cond18, font_size=8.5, bold=False)

                # Fallback row matching
                if ("wind speed" in row_raw.lower() or "ქარის სიჩქარე" in row_raw) and len(row.cells) >= 5:
                    set_clean_wind_cell(row.cells[0], w9)
                    set_clean_wind_cell(row.cells[2], w14)
                    set_clean_wind_cell(row.cells[4], w18)

                if any(h in row_raw for h in ["9:00", "09:00", "9·00"]) and "14:00" in row_raw:
                    if r_idx + 1 < len(table.rows):
                        t_row = table.rows[r_idx + 1]
                        if len(t_row.cells) >= 6:
                            set_clean_weather_cell(t_row.cells[0], t9, font_size=9.5, bold=True)
                            set_clean_weather_cell(t_row.cells[1], cond9, font_size=8.5, bold=False)
                            set_clean_weather_cell(t_row.cells[2], t14, font_size=9.5, bold=True)
                            set_clean_weather_cell(t_row.cells[3], cond14, font_size=8.5, bold=False)
                            set_clean_weather_cell(t_row.cells[4], t18, font_size=9.5, bold=True)
                            set_clean_weather_cell(t_row.cells[5], cond18, font_size=8.5, bold=False)

            # C. DAILY WORKS TABLE (HEADER COMPACTED & ROW PADDING BALANCED)
            if ("WBS" in t_text and ("სამუშაო" in t_text or "Code" in t_text)) or "{{ c1 }}" in t_text:
                header_labels = [
                    ("№", WD_ALIGN_PARAGRAPH.CENTER),
                    ("შემსრულებელი\nContractor", WD_ALIGN_PARAGRAPH.CENTER),
                    ("მუშახელი\nManpower", WD_ALIGN_PARAGRAPH.CENTER),
                    ("WBS კოდი\nCode", WD_ALIGN_PARAGRAPH.CENTER),
                    ("სამუშაოს დასახელება\nWork Description", WD_ALIGN_PARAGRAPH.CENTER),
                    ("აღწერა / დეტალები\nDescription", WD_ALIGN_PARAGRAPH.CENTER)
                ]
                for c_idx, (lbl, align) in enumerate(header_labels):
                    populate_compact_cell(table.rows[0].cells[c_idx], lbl, font_size=8.0, bold=True, align=align)

                while len(table.rows) > 1:
                    table._tbl.remove(table.rows[-1]._tr)

                for tsk in tasks:
                    new_r = table.add_row()
                    populate_compact_cell(new_r.cells[0], str(tsk["num"]), font_size=8.0, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
                    populate_compact_cell(new_r.cells[1], str(tsk["contractor"]), font_size=8.0, bold=False, align=WD_ALIGN_PARAGRAPH.CENTER)
                    populate_compact_cell(new_r.cells[2], str(tsk["manpower"]), font_size=8.0, bold=False, align=WD_ALIGN_PARAGRAPH.CENTER)
                    populate_compact_cell(new_r.cells[3], str(tsk["wbs"]), font_size=8.0, bold=False, align=WD_ALIGN_PARAGRAPH.CENTER)
                    populate_compact_cell(new_r.cells[4], str(tsk["name"]), font_size=8.0, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT)
                    populate_compact_cell(new_r.cells[5], str(tsk["desc"]), font_size=8.0, bold=False, align=WD_ALIGN_PARAGRAPH.LEFT)

            # D. PHOTOS TABLE
            if "დღიური სამუშაოების ამსახველი ფოტომასალა" in t_text or "სამუშაო პროცესი" in t_text:
                num_uploaded = len(photos_data)
                process_rows = [(1, 2, 3), (4, 5, 6), (7, 8, 9)]

                for proc_i, (h_idx, p_idx, c_idx) in enumerate(process_rows):
                    if h_idx < len(table.rows) and p_idx < len(table.rows) and c_idx < len(table.rows):
                        h_row = table.rows[h_idx]
                        p_row = table.rows[p_idx]
                        c_row = table.rows[c_idx]

                        slot_a = proc_i * 2
                        slot_b = proc_i * 2 + 1

                        p_row.cells[0].text = ""
                        c_row.cells[0].text = ""
                        if slot_a < num_uploaded:
                            p_run = p_row.cells[0].paragraphs[0].add_run()
                            p_run.add_picture(io.BytesIO(photos_data[slot_a]["file"].getvalue()), width=Inches(3.15))
                            populate_compact_cell(c_row.cells[0], photos_data[slot_a]["caption"], font_size=8.0, bold=False, align=WD_ALIGN_PARAGRAPH.CENTER)

                        p_row.cells[1].text = ""
                        c_row.cells[1].text = ""
                        if slot_b < num_uploaded:
                            p_run = p_row.cells[1].paragraphs[0].add_run()
                            p_run.add_picture(io.BytesIO(photos_data[slot_b]["file"].getvalue()), width=Inches(3.15))
                            populate_compact_cell(c_row.cells[1], photos_data[slot_b]["caption"], font_size=8.0, bold=False, align=WD_ALIGN_PARAGRAPH.CENTER)

                        if slot_a >= num_uploaded and slot_b >= num_uploaded:
                            h_row.cells[0].text = ""

                # Extra Photos (> 6) — Dynamic Pagination up to 100
                if num_uploaded > 6:
                    remaining_photos = photos_data[6:]
                    proc_counter = 4
                    sample_photo_row = table.rows[2]
                    sample_cap_row = table.rows[3]

                    for pair_start in range(0, len(remaining_photos), 2):
                        t_row = clone_row(table, table.rows[1])
                        t_cell = t_row.cells[0]
                        t_cell.merge(t_row.cells[1])
                        t_cell.text = f"სამუშაო პროცესი / Work Process №{proc_counter}"
                        if proc_counter == 4 or (proc_counter - 4) % 3 == 0:
                            t_cell.paragraphs[0].paragraph_format.page_break_before = True

                        p_row = clone_row(table, sample_photo_row)
                        c_row = clone_row(table, sample_cap_row)

                        p_row.cells[0].text = ""
                        c_row.cells[0].text = ""
                        p_row.cells[0].paragraphs[0].add_run().add_picture(io.BytesIO(remaining_photos[pair_start]["file"].getvalue()), width=Inches(3.15))
                        populate_compact_cell(c_row.cells[0], remaining_photos[pair_start]["caption"], font_size=8.0, bold=False, align=WD_ALIGN_PARAGRAPH.CENTER)

                        p_row.cells[1].text = ""
                        c_row.cells[1].text = ""
                        if pair_start + 1 < len(remaining_photos):
                            p_row.cells[1].paragraphs[0].add_run().add_picture(io.BytesIO(remaining_photos[pair_start + 1]["file"].getvalue()), width=Inches(3.15))
                            populate_compact_cell(c_row.cells[1], remaining_photos[pair_start + 1]["caption"], font_size=8.0, bold=False, align=WD_ALIGN_PARAGRAPH.CENTER)

                        proc_counter += 1

            # E. KEY STAFF TABLE (COMPACT PROPORTIONAL FONTS)
            if "სამშენებლო მოედანზე მომუშავე კომპანიები" in t_text or "Key Staff" in t_text:
                while len(table.rows) > 2:
                    table._tbl.remove(table.rows[-1]._tr)

                for sm in staff_members:
                    s_row = table.add_row()
                    populate_compact_cell(s_row.cells[0], str(sm["num"]), font_size=7.5, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
                    populate_compact_cell(s_row.cells[1], str(sm["comp"]), font_size=7.5, bold=False, align=WD_ALIGN_PARAGRAPH.CENTER)
                    populate_compact_cell(s_row.cells[2], str(sm["pos"]), font_size=7.5, bold=False, align=WD_ALIGN_PARAGRAPH.CENTER)
                    populate_compact_cell(s_row.cells[3], str(sm["name"]), font_size=7.5, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
                    populate_compact_cell(s_row.cells[4], str(sm["phone"]), font_size=7.5, bold=False, align=WD_ALIGN_PARAGRAPH.CENTER)
                    populate_compact_cell(s_row.cells[5], str(sm["mail"]), font_size=7.5, bold=False, align=WD_ALIGN_PARAGRAPH.CENTER)

        # 3. GLOBAL SWEEP: Replace any remaining placeholders anywhere in document
        global_tag_map = {
            r'\{\{\s*w9\s*\}\}': str(w9),
            r'\{\{\s*w14\s*\}\}': str(w14),
            r'\{\{\s*w18\s*\}\}': str(w18),
            r'\{\{\s*t9\s*\}\}': str(t9),
            r'\{\{\s*t14\s*\}\}': str(t14),
            r'\{\{\s*t18\s*\}\}': str(t18),
            r'\{\{\s*cond9\s*\}\}': str(cond9),
            r'\{\{\s*cond14\s*\}\}': str(cond14),
            r'\{\{\s*cond18\s*\}\}': str(cond18),
        }
        for p in doc.paragraphs:
            for pat, val in global_tag_map.items():
                replace_in_p(p, pat, val)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        for pat, val in global_tag_map.items():
                            replace_in_p(p, pat, val)

        # Save Word bytes
        bio = io.BytesIO()
        doc.save(bio)
        st.session_state["docx_bytes"] = bio.getvalue()
        st.session_state["doc_code"] = final_doc_code

        # ==================== PDF CONVERSION ENGINE ====================
        temp_doc_name = f"temp_{final_doc_code}.docx"
        temp_pdf_name = f"{final_doc_code}.pdf"
        abs_docx = os.path.abspath(temp_doc_name)
        abs_pdf = os.path.abspath(temp_pdf_name)

        with open(abs_docx, "wb") as f_tmp:
            f_tmp.write(bio.getvalue())

        pdf_converted = False
        conversion_error = None

        # 1. Try LibreOffice (Linux / Streamlit Cloud / local LibreOffice)
        try:
            cmd = ["soffice", "--headless", "--convert-to", "pdf", abs_docx, "--outdir", os.path.dirname(abs_docx)]
            subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=40)
            candidate_pdf = abs_docx.rsplit(".", 1)[0] + ".pdf"
            if os.path.exists(candidate_pdf):
                with open(candidate_pdf, "rb") as f_pdf:
                    st.session_state["pdf_bytes"] = f_pdf.read()
                pdf_converted = True
                if os.path.exists(candidate_pdf):
                    os.remove(candidate_pdf)
        except Exception as e:
            conversion_error = e

        # 2. Try Windows Microsoft Word (Local Windows PC)
        if not pdf_converted:
            try:
                import pythoncom
                pythoncom.CoInitialize()
                from docx2pdf import convert
                convert(abs_docx, abs_pdf)
                if os.path.exists(abs_pdf):
                    with open(abs_pdf, "rb") as f_pdf:
                        st.session_state["pdf_bytes"] = f_pdf.read()
                    pdf_converted = True
            except Exception:
                try:
                    import win32com.client
                    word_app = win32com.client.DispatchEx("Word.Application")
                    word_app.Visible = False
                    word_app.DisplayAlerts = False
                    w_doc = word_app.Documents.Open(abs_docx)
                    w_doc.SaveAs(abs_pdf, FileFormat=17)
                    w_doc.Close()
                    word_app.Quit()
                    if os.path.exists(abs_pdf):
                        with open(abs_pdf, "rb") as f_pdf:
                            st.session_state["pdf_bytes"] = f_pdf.read()
                        pdf_converted = True
                except Exception as e2:
                    conversion_error = e2
                    st.session_state["pdf_bytes"] = None

        # Clean up temporary disk files
        if os.path.exists(abs_docx):
            try:
                os.remove(abs_docx)
            except Exception:
                pass
        if os.path.exists(abs_pdf):
            try:
                os.remove(abs_pdf)
            except Exception:
                pass

        if not pdf_converted:
            st.warning(f"💡 შენიშვნა: PDF ვერ დაგენერირდა ({conversion_error}). გადმოწერეთ Word ფაილი.")

    except Exception as e:
        st.error(f"შეცდომა რეპორტის შექმნისას: {e}")

# ==================== DOWNLOAD BUTTONS ====================
if st.session_state["docx_bytes"] is not None:
    d_code = st.session_state["doc_code"]
    st.success(f"🎉 რეპორტი მზადაა: **{d_code}**")

    b_col1, b_col2 = st.columns(2)
    b_col1.download_button(
        label="📥 1. გადმოწერეთ Word (.docx)",
        data=st.session_state["docx_bytes"],
        file_name=f"{d_code}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True
    )

    if st.session_state["pdf_bytes"] is not None:
        b_col2.download_button(
            label="🖨️ 2. გადმოწერეთ დასაბეჭდი PDF (.pdf)",
            data=st.session_state["pdf_bytes"],
            file_name=f"{d_code}.pdf",
            mime="application/pdf",
            use_container_width=True
        )
    else:
        b_col2.info("💡 PDF ღილაკისთვის Streamlit Cloud-ზე GitHub რეპოზიტორიაში დაამატეთ `packages.txt` ჩანაწერით: `libreoffice`.")
