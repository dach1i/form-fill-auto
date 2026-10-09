import docx
import os
import glob

# 1. Detect base document
source_file = None
for f in ["CMC-CMS-DPR-007-20261008-Rev00.docx", "CMC-CMS-DPR-001-20261001-Rev00.docx", "template.docx"]:
    if os.path.exists(f):
        source_file = f
        break

if not source_file:
    print("❌ Error: No base .docx file found!")
    exit(1)

print(f"📖 Fixing template from: {source_file}")
doc = docx.Document(source_file)

def clear_cell_completely(cell):
    tc = cell._tc
    for child in list(tc):
        if not child.tag.endswith('tcPr'):
            tc.remove(child)
    cell.add_paragraph()

def safe_replace(p, old, new):
    if old in p.text:
        for r in p.runs:
            if old in r.text:
                r.text = r.text.replace(old, str(new))
        if old in p.text:
            p.text = p.text.replace(old, str(new))

# 2. REMOVE FLOATING TABLE PROPERTIES (CRITICAL FOR PDF OVERLAP FIX)
for tblp in doc._body._element.xpath('.//w:tblpPr'):
    tblp.getparent().remove(tblp)
print("  ✓ Stripped floating table positions (tables are now 100% inline).")

# 3. FIX EXACT ROW HEIGHTS TO atLeast (PREVENTS SQUISHING & COLLISION)
for trh in doc._body._element.xpath('.//w:trHeight'):
    if trh.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}hRule') == 'exact':
        trh.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}hRule', 'atLeast')
print("  ✓ Converted exact row heights to dynamic heights.")

# 4. UNWRAP SDT CONTROLS
for sdt in doc._body._element.xpath('.//w:sdt'):
    content = sdt.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sdtContent')
    parent = sdt.getparent()
    if content is not None and parent is not None:
        idx = parent.index(sdt)
        for child in list(content):
            parent.insert(idx, child)
            idx += 1
        parent.remove(sdt)

# 5. FIX TYPOS
misspellings = {
    "სამუშაეობის": "სამუშაოების",
    "სამუშაეობა": "სამუშაოები",
    "Click or tap here to enter text.": "-"
}
for p in doc.paragraphs:
    for bad, good in misspellings.items():
        safe_replace(p, bad, good)
for table in doc.tables:
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                for bad, good in misspellings.items():
                    safe_replace(p, bad, good)

# 6. HEADER PLACEHOLDERS
for table in doc.tables:
    t_text = " ".join([c.text for row in table.rows for c in row.cells])
    if "Daily Progress Report" in t_text or "ყოველდღიური რეპორტი" in t_text:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if p.text.strip() in ["007", "001", "111"]:
                        p.text = "{{ rep_num }}"
                    safe_replace(p, "მურად ახალაძე", "{{ prepared_by }}")
                    safe_replace(p, "2026/10/08", "{{ rep_date }}")
                    safe_replace(p, "2026/10/01", "{{ rep_date }}")
                    safe_replace(p, "შპს „აი ჯი დეველოფმენტ ჯორჯია“", "{{ client }}")
                    safe_replace(p, "ქ.თბილისი ქავთარაძის ქ. №1", "{{ location }}")

# 7. WEATHER & WIND PLACEHOLDERS
for table in doc.tables:
    for row in table.rows:
        row_txt = " ".join([c.text for c in row.cells])
        if "wind speed" in row_txt.lower() or "ქარის სიჩქარე" in row_txt:
            for c in row.cells:
                for p in c.paragraphs:
                    safe_replace(p, "1.1 მ/წ", "{{ w9 }}")
                    safe_replace(p, "1.3 მ/წ", "{{ w14 }}")
                    safe_replace(p, "1.0 მ/წ", "{{ w18 }}")

# 8. CLEAR OLD TASKS FROM WORKS TABLE (KEEP 6-COLUMN HEADER)
for table in doc.tables:
    t_text = " ".join([c.text for row in table.rows for c in row.cells])
    if "WBS" in t_text and ("Contractor" in t_text or "შემსრულებელი" in t_text):
        for r_idx, row in enumerate(table.rows):
            r_txt = " ".join([c.text for c in row.cells])
            if any(k in r_txt for k in ["ბკ ქონსთრაქშენი", "შპუნტების მონტაჟი", "{{ c2 }}", "ხიმინჯის"]):
                for c in row.cells:
                    clear_cell_completely(c)

# 9. WIPE ALL 6 EMBEDDED OLD PHOTOS AND CAPTIONS
for table in doc.tables:
    t_text = " ".join([c.text for row in table.rows for c in row.cells])
    if "ფოტომასალა" in t_text or "Daily Work Progress Photos" in t_text:
        wp_idx = None
        for r_idx, row in enumerate(table.rows):
            r_txt = " ".join([c.text for c in row.cells])
            if "სამუშაო პროცესი" in r_txt or "Work Process" in r_txt:
                wp_idx = r_idx
                break
        if wp_idx is not None:
            for r_offset in range(1, 7):
                if wp_idx + r_offset < len(table.rows):
                    for cell in table.rows[wp_idx + r_offset].cells:
                        clear_cell_completely(cell)
        print("  ✓ Wiped all 6 old photo slots.")

# 10. CLEAR STAFF TABLE
for table in doc.tables:
    t_text = " ".join([c.text for row in table.rows for c in row.cells])
    if any(k in t_text for k in ["კომპანიები", "Key Staff", "E-mail"]):
        if len(table.rows) > 2:
            for c in table.rows[2].cells:
                clear_cell_completely(c)
            while len(table.rows) > 3:
                table._tbl.remove(table.rows[-1]._tr)

doc.save("template.docx")
print("🎉 Success! Rebuilt clean 'template.docx' without floating table locks.")
