from datetime import datetime, timedelta
import io
import os
import re
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
import pandas as pd
import streamlit as st

# --- STREAMLIT PAGE CONFIG ---
st.set_page_config(
    page_title="Toastmasters Agenda & Intros", page_icon="🎤", layout="centered"
)

st.title("🎤 Toastmasters Document Generator")
st.write(
    "Upload the meeting plan to generate both the Agenda and Member Intros Word"
    " documents."
)


# --- SCRIPT 1 HELPER FUNCTIONS (AGENDA GENERATION) ---
def clean_name(raw_text):
  if not raw_text:
    return ""
  return re.sub(r"\([^)]*\)", "", raw_text).strip()


def parse_role_file(txt_content):
  lines = [
      line.strip()
      for line in txt_content.decode("utf-8").splitlines()
      if line.strip()
  ]

  data = {
      "theme": "",
      "meeting_num": "0",
      "date_str": "",
      "date_code": "",
      "roles": {},
  }

  full_text = "\n".join(lines)

  theme_match = re.search(r"THEME\s*:\s*(.*)", full_text, re.IGNORECASE)
  if theme_match:
    data["theme"] = theme_match.group(1).strip()

  num_match = re.search(r"Meeting:\s*#?(\d+)", full_text, re.IGNORECASE)
  if num_match:
    data["meeting_num"] = num_match.group(1)

  date_match = re.search(r"(\d{2}\.\d{2}\.\d{4})", full_text)
  if date_match:
    raw_date = date_match.group(1)
    dt = datetime.strptime(raw_date, "%d.%m.%Y")
    data["date_str"] = dt.strftime("%B %d, %Y")
    data["date_code"] = dt.strftime("%d%m%y")

  role_patterns = {
      "SAA": r"SAA\s*:\s*(.*)",
      "PO": r"PO\s*:\s*(.*)",
      "TMOD": r"TMOD\s*:\s*(.*)",
      "GE": r"GE\s*:\s*(.*)",
      "TTM": r"TTM\s*:\s*(.*)",
      "Speaker 1": r"Speaker 1\s*:\s*(.*)",
      "Evaluator 1": r"Evaluator 1\s*:\s*(.*)",
      "Speaker 2": r"Speaker 2\s*:\s*(.*)",
      "Evaluator 2": r"Evaluator 2\s*:\s*(.*)",
      "Timer": r"TIMER\s*:\s*(.*)",
      "Ah Counter": r"AH-COUNTER\s*:\s*(.*)",
      "Grammarian": r"GRAMMARIAN\s*:\s*(.*)",
      "Listener": r"Listener\s*:\s*(.*)",
  }

  for role_key, pattern in role_patterns.items():
    match = re.search(pattern, full_text, re.IGNORECASE)
    data["roles"][role_key] = clean_name(match.group(1)) if match else ""

  return data


def build_agenda_schedule(roles):
  has_speaker2 = bool(roles.get("Speaker 2"))
  has_listener = bool(roles.get("Listener"))

  master_items = [
      (15, "Networking & Technology setup", "ALL", True),
      (5, "SAA address", roles.get("SAA", ""), True),
      (
          10,
          "Presidential address TMOD introduction",
          f"{roles.get('PO', '')} / {roles.get('TMOD', '')}".strip(" /"),
          True,
      ),
      (
          7,
          "Meeting structure, Theme address and GE introduction",
          roles.get("TMOD", ""),
          True,
      ),
      (2, "GE address and TAG team Intro", roles.get("GE", ""), True),
      (2, "Timer", roles.get("Timer", ""), True),
      (2, "Ah Counter", roles.get("Ah Counter", ""), True),
      (2, "Grammarian", roles.get("Grammarian", ""), True),
      (2, "Listener", roles.get("Listener", ""), has_listener),
      (
          2,
          "Evaluator 1 intro – Speaker 1 objective – Speaker Intro",
          f"{roles.get('TMOD', '')} – {roles.get('Evaluator 1', '')}".strip(
              " –"
          ),
          True,
      ),
      (10, "Prepared Speech 1", roles.get("Speaker 1", ""), True),
      (
          2,
          "Evaluator 2 intro – Speaker 2 objective – Speaker Intro",
          f"{roles.get('TMOD', '')} – {roles.get('Evaluator 2', '')}".strip(
              " –"
          ),
          has_speaker2,
      ),
      (10, "Prepared Speech 2", roles.get("Speaker 2", ""), has_speaker2),
      (5, "TMOD interaction", roles.get("TMOD", ""), True),
      (15, "Table Topics Session", roles.get("TTM", ""), True),
      (4, "Evaluator 1 Report", roles.get("Evaluator 1", ""), True),
      (4, "Evaluator 2 Report", roles.get("Evaluator 2", ""), has_speaker2),
      (
          6,
          "TAG Team Report",
          (
              f"{roles.get('Timer', '')} / {roles.get('Ah Counter', '')} /"
              f" {roles.get('Grammarian', '')}"
          ),
          True,
      ),
      (5, "Listening Quiz Session", roles.get("Listener", ""), has_listener),
      (5, "GE Report", roles.get("GE", ""), True),
      (2, "TMOD Conclusion", roles.get("TMOD", ""), True),
      (8, "Voting – Meeting closure – Photo session", roles.get("PO", ""), True),
  ]

  active_schedule = []
  current_time = datetime.strptime("17:00", "%H:%M")

  for duration, role_text, role_taker, is_active in master_items:
    if is_active:
      time_str = f"{current_time.hour % 12 or 12}:{current_time.minute:02d}"
      active_schedule.append({
          "time": time_str,
          "duration": f"{duration:02d} mins",
          "role_text": role_text,
          "role_taker": role_taker,
      })
      current_time += timedelta(minutes=duration)

  return active_schedule


def set_paragraph_text_and_style(
    paragraph, text, font_size_pt=17, color_rgb=(255, 255, 255), bold=True
):
  paragraph.text = text
  for run in paragraph.runs:
    run.font.size = Pt(font_size_pt)
    run.font.color.rgb = RGBColor(*color_rgb)
    run.font.bold = bold


# --- SCRIPT 2 HELPER FUNCTIONS (INTRO GENERATION) ---
def load_member_database(excel_source):
  try:
    xls = pd.ExcelFile(excel_source)
    records = []
    for sheet_name in xls.sheet_names:
      df = pd.read_excel(excel_source, sheet_name=sheet_name)
      name_col, intro_col = None, None
      for col in df.columns:
        for val in df[col].dropna():
          val_str = str(val).strip().lower()
          if val_str == "name":
            name_col = col
          elif val_str == "introduction":
            intro_col = col

      if name_col is not None and intro_col is not None:
        header_idx = df[
            df[name_col].astype(str).str.strip().str.lower() == "name"
        ].index[0]
        clean_df = df.iloc[header_idx + 1 :].copy()
        clean_df = clean_df.rename(
            columns={name_col: "Name", intro_col: "Introduction"}
        )
        records.append(clean_df[["Name", "Introduction"]])

    if records:
      return pd.concat(records, ignore_index=True).dropna(subset=["Name"])
  except Exception:
    pass
  return pd.DataFrame(columns=["Name", "Introduction"])


def match_intro(name, db_df):
  if not name or name.strip().upper() in ["ALL", ""]:
    return "N/A"

  clean_name = name.strip().lower()

  for _, row in db_df.iterrows():
    db_name = str(row["Name"]).strip().lower()
    if db_name and (clean_name in db_name or db_name in clean_name):
      intro = str(row["Introduction"]).strip()
      if intro and intro.lower() != "nan":
        return intro

  tokens = [t for t in clean_name.split() if len(t) > 2]
  for _, row in db_df.iterrows():
    db_name = str(row["Name"]).strip().lower()
    if any(t in db_name for t in tokens):
      intro = str(row["Introduction"]).strip()
      if intro and intro.lower() != "nan":
        return intro

  return "Introduction not provided in database."


# --- USER INTERFACE & MAIN LOGIC ---
st.subheader("1. Upload Inputs")
txt_file = st.file_uploader(
    "Upload Role Players Plan (.txt)", type=["txt"], help="Mandatory file"
)
excel_file = st.file_uploader(
    "Upload Member Database (.xlsx) - OPTIONAL",
    type=["xlsx"],
    help="Optional. Fallback server database will be used if skipped.",
)

if txt_file is not None:
  # Determine Excel database source
  if excel_file is not None:
    db_df = load_member_database(excel_file)
    st.info("ℹ️ Using newly uploaded Excel database.")
  elif os.path.exists("Toastmasters Introduction .xlsx"):
    db_df = load_member_database("Toastmasters Introduction .xlsx")
    st.info("ℹ️ Using default server Excel database.")
  else:
    db_df = pd.DataFrame(columns=["Name", "Introduction"])
    st.warning("⚠️ No Excel database found. Placeholders will be used.")

  parsed = parse_role_file(txt_file.getvalue())
  st.success(
      f"Parsed Meeting #{parsed['meeting_num']} | Theme: {parsed['theme']}"
  )

  # --- GENERATE AGENDA DOC ---
  template_path = os.path.join(
      "template", "ToastMastersAgenda_template.docx"
  )
  if os.path.exists(template_path):
    doc = Document(template_path)

    set_paragraph_text_and_style(
        doc.paragraphs[0],
        f"Bosch Raconteurs Toastmasters Club Meeting #{parsed['meeting_num']}",
        font_size_pt=17,
        color_rgb=(255, 255, 255),
        bold=True,
    )
    set_paragraph_text_and_style(
        doc.paragraphs[1],
        f"THEME: {parsed['theme']}",
        font_size_pt=17,
        color_rgb=(255, 255, 255),
        bold=True,
    )
    doc.paragraphs[4].text = parsed["date_str"]

    active_schedule = build_agenda_schedule(parsed["roles"])
    table = doc.tables[0]

    for idx, item in enumerate(active_schedule, start=1):
      row_cells = table.rows[idx].cells
      row_cells[0].text = item["time"]
      row_cells[1].text = item["duration"]
      row_cells[2].text = item["role_text"]
      row_cells[3].text = item["role_taker"]

    total_needed_rows = len(active_schedule) + 1
    while len(table.rows) > total_needed_rows:
      extra_row = table.rows[-1]
      extra_row._tr.getparent().remove(extra_row._tr)

    agenda_buffer = io.BytesIO()
    doc.save(agenda_buffer)
    agenda_buffer.seek(0)

    # --- GENERATE INTROS DOC ---
    intro_doc = Document()
    for s in intro_doc.sections:
      s.top_margin = Inches(1)
      s.bottom_margin = Inches(1)
      s.left_margin = Inches(1)
      s.right_margin = Inches(1)

    p_head = intro_doc.add_paragraph()
    r_head = p_head.add_run(
        f"Bosch Raconteurs Toastmasters meeting #{parsed['meeting_num']} member"
        " intros"
    )
    r_head.font.name = "Calibri"
    r_head.font.size = Pt(20)
    r_head.font.bold = True
    r_head.font.color.rgb = RGBColor(31, 78, 121)
    p_head.paragraph_format.space_after = Pt(18)

    role_map = parsed["roles"]
    tmod_roles = []
    if role_map.get("Speaker 1"):
      tmod_roles.append(("Speaker 1", role_map["Speaker 1"]))
    if role_map.get("Evaluator 1"):
      tmod_roles.append(("Evaluator 1", role_map["Evaluator 1"]))
    if role_map.get("Speaker 2"):
      tmod_roles.append(("Speaker 2", role_map["Speaker 2"]))
    if role_map.get("Evaluator 2"):
      tmod_roles.append(("Evaluator 2", role_map["Evaluator 2"]))

    tmod_roles.append(
        ("Table Topics Master (TTM)", role_map.get("TTM", "N/A"))
    )
    tmod_roles.append(("General Evaluator (GE)", role_map.get("GE", "N/A")))

    sections_schema = [
        (
            "1. Introduction of PO",
            [("Presiding Officer (PO)", role_map.get("PO", "N/A"))],
        ),
        (
            "2. Introduction of TMOD",
            [("Toastmaster of the Day (TMOD)", role_map.get("TMOD", "N/A"))],
        ),
        ("3. Roles Introduced by TMOD", tmod_roles),
        (
            "4. Roles Introduced by GE",
            [
                ("Timer", role_map.get("Timer", "N/A")),
                ("Ah-Counter", role_map.get("Ah Counter", "N/A")),
                ("Grammarian", role_map.get("Grammarian", "N/A")),
                ("Listener", role_map.get("Listener", "N/A")),
            ],
        ),
    ]

    for cat_title, role_tuples in sections_schema:
      cp = intro_doc.add_paragraph()
      c_run = cp.add_run(cat_title)
      c_run.font.name = "Calibri"
      c_run.font.size = Pt(14)
      c_run.font.bold = True
      c_run.font.color.rgb = RGBColor(46, 117, 182)
      cp.paragraph_format.space_before = Pt(14)
      cp.paragraph_format.space_after = Pt(6)

      for role_label_text, taker_name in role_tuples:
        intro_text = match_intro(taker_name, db_df)

        rp = intro_doc.add_paragraph()
        rp.paragraph_format.left_indent = Inches(0.2)
        rp.paragraph_format.space_before = Pt(4)
        rp.paragraph_format.space_after = Pt(2)

        r_lbl = rp.add_run(f"• {role_label_text}: ")
        r_lbl.font.name = "Calibri"
        r_lbl.font.size = Pt(11)
        r_lbl.font.bold = True

        r_name = rp.add_run(taker_name)
        r_name.font.name = "Calibri"
        r_name.font.size = Pt(11)
        r_name.font.bold = True
        r_name.font.color.rgb = RGBColor(31, 78, 121)

        ip = intro_doc.add_paragraph()
        ip.paragraph_format.left_indent = Inches(0.4)
        ip.paragraph_format.space_after = Pt(8)

        irun = ip.add_run(intro_text)
        irun.font.name = "Calibri"
        irun.font.size = Pt(10.5)

        if "not provided" in intro_text.lower() or intro_text == "N/A":
          irun.font.bold = True
          irun.font.italic = False
          irun.font.color.rgb = RGBColor(192, 0, 0)
        else:
          irun.font.bold = False
          irun.font.italic = True
          irun.font.color.rgb = RGBColor(89, 89, 89)

    intro_buffer = io.BytesIO()
    intro_doc.save(intro_buffer)
    intro_buffer.seek(0)

    # --- DOWNLOAD BUTTONS ---
    st.subheader("2. Download Generated Documents")
    col1, col2 = st.columns(2)
    with col1:
      st.download_button(
          label="📄 Download Meeting Agenda",
          data=agenda_buffer,
          file_name=(
              f"ToastMastersAgenda_{parsed['meeting_num']}_{parsed['date_code']}.docx"
          ),
          mime=(
              "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
          ),
      )
    with col2:
      st.download_button(
          label="📝 Download Member Intros",
          data=intro_buffer,
          file_name=(
              f"Bosch_Raconteurs_Meeting_{parsed['meeting_num']}_Member_Intros.docx"
          ),
          mime=(
              "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
          ),
      )
  else:
    st.error(
        "Template missing on server! Ensure"
        " 'template/ToastMastersAgenda_template.docx' exists."
    )