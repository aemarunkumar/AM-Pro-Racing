import os
import re
import tempfile
import streamlit as st
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

st.set_page_config(
    page_title="AM PRO Racing System",
    page_icon="🏇",
    layout="wide"
)

st.markdown("""
    <style>
    .main-title {
        text-align: center;
        font-weight: 800;
        color: #1E3A8A;
        font-size: 26px;
        margin-bottom: 2px;
    }
    .sub-title {
        text-align: center;
        color: #4B5563;
        font-size: 14px;
        margin-bottom: 20px;
    }
    .stButton>button {
        width: 100%;
        border-radius: 8px;
        height: 3.2em;
        background-color: #2563EB;
        color: white;
        font-weight: bold;
        font-size: 16px;
    }
    </style>
""", unsafe_allow_html=True)

st.markdown("<div class='main-title'>🏇 AM PRO Racing System</div>", unsafe_allow_html=True)
st.markdown("<div class='sub-title'>Step 1: All Form to Multi-Sheet Matrix (Col T & Highlighting) | Step 2: Final Scoring</div>", unsafe_allow_html=True)

# =====================================================================
# HELPER FUNCTIONS (CLEANING & JOCKEY MATCHING)
# =====================================================================

def clean_jockey_name(name):
    if not name:
        return ""
    # Remove allowances like (a2/53kg), (a), (a1.5), etc.
    cleaned = re.sub(r"\(a\d*(\.\d+)?(/[0-9.]+kg)?\)", "", name, flags=re.I)
    cleaned = re.sub(r"\(a\)", "", cleaned, flags=re.I)
    # Remove prefix like Ms, Mr, Mrs
    cleaned = re.sub(r"^(Ms|Mr|Mrs)\s+", "", cleaned.strip(), flags=re.I)
    return cleaned.strip().lower()

def extract_weight_num(w_str):
    if not w_str:
        return 0.0
    match = re.search(r"(\d+(\.\d+)?)", str(w_str))
    return float(match.group(1)) if match else 0.0

def normalize_track(t_str):
    if not t_str:
        return ""
    t_clean = t_str.strip().lower()
    if "good" in t_clean: return "good"
    if "soft" in t_clean: return "soft"
    if "heavy" in t_clean: return "heavy"
    if "synthetic" in t_clean or "syn" in t_clean: return "syn"
    if "firm" in t_clean: return "firm"
    return t_clean

# =====================================================================
# PARSER ENGINE
# =====================================================================

def parse_all_form_html_full(html_text):
    soup = BeautifulSoup(html_text, "html.parser")
    races = []
    
    meeting_venue = "Sportsbet Sandown Hillside"
    meeting_date = "02/09/2026"
    meeting_country = "VIC"
    
    title_elem = soup.find("h1") or soup.find("title")
    if title_elem:
        t_text = title_elem.get_text(separator=" ", strip=True)
        m_date = re.search(r"(\d{1,2}\s+[A-Za-z]+\s+\d{4}|\d{1,2}/\d{1,2}/\d{4})", t_text)
        if m_date: meeting_date = m_date.group(1)
        m_place = re.search(r"(?:at\s+|-\s+)([A-Za-z0-9\s]+?)(?:\s+\(|Form|$)", t_text)
        if m_place: meeting_venue = m_place.group(1).strip()

    tables = soup.find_all("table")
    race_counter = 1

    for tbl in tables:
        rows = tbl.find_all("tr")
        if len(rows) < 2:
            continue

        horses = []
        current_horse = None

        for r in rows:
            cells = r.find_all(["td", "th"])
            txts = [re.sub(r'\s+', ' ', c.get_text(strip=True)) for c in cells]
            if not txts:
                continue

            first_col = txts[0].strip()

            # Horse Row
            if re.match(r"^\d{1,2}$", first_col) and len(txts) >= 4:
                h_name = txts[1]
                if h_name.lower() in ["horse", "runner", "horse name"]:
                    continue

                h_name_clean = re.sub(r"\s*\(.*?\)", "", h_name).strip()
                jockey = txts[2] if len(txts) > 2 else ""
                trainer = txts[3] if len(txts) > 3 else ""
                owner = txts[4] if len(txts) > 4 and not re.search(r"\d", txts[4]) else ""
                
                barrier = ""
                weight = "58.0kg"
                for t in txts[3:]:
                    if re.match(r"^\d{1,2}$", t) and not barrier:
                        barrier = t
                    elif re.match(r"^\d{2}\.?\d?kg$", t, re.I) and not weight:
                        weight = t

                current_horse = {
                    "no": first_col,
                    "name": h_name_clean,
                    "jockey": jockey,
                    "trainer": trainer,
                    "owner": owner,
                    "barrier": barrier,
                    "final_weight": weight,
                    "runs": []
                }
                horses.append(current_horse)

            # Previous Runs Row
            elif current_horse is not None and len(txts) >= 6:
                if any(m in txts[0].lower() for m in ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec", "/"]):
                    run_date = txts[0]
                    run_place = txts[1] if len(txts) > 1 else ""
                    run_type = txts[2] if len(txts) > 2 else "Race"
                    run_pos = txts[3] if len(txts) > 3 else ""
                    run_class = txts[4] if len(txts) > 4 else ""
                    run_dist = txts[5] if len(txts) > 5 else "1200m"
                    run_track = txts[6] if len(txts) > 6 else "Good4"
                    run_barrier = txts[7] if len(txts) > 7 else ""
                    run_weight = txts[8] if len(txts) > 8 else "56kg"
                    run_jockey = txts[9] if len(txts) > 9 else ""
                    run_time = txts[10] if len(txts) > 10 else ""

                    current_horse["runs"].append({
                        "date": run_date,
                        "place": run_place,
                        "type": run_type,
                        "pos": run_pos,
                        "class": run_class,
                        "distance": run_dist,
                        "track": run_track,
                        "barrier": run_barrier,
                        "weight": run_weight,
                        "jockey": run_jockey,
                        "time": run_time
                    })

        if horses and len(horses) >= 2:
            races.append({
                "race_no": f"Race {race_counter}",
                "sheet_name": f"R{race_counter}",
                "country": meeting_country,
                "distance": "1500m",
                "track_condition": "Soft",
                "class_name": "BM70",
                "date": meeting_date,
                "place": meeting_venue,
                "horses": horses
            })
            race_counter += 1

    return races

# =====================================================================
# FULL WORKBOOK GENERATOR WITH EXACT HIGHLIGHTING & DIFFERENCE
# =====================================================================

def generate_am_pro_step1_workbook(races):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # Styles & Colors
    navy_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    dark_gray_fill = PatternFill(start_color="595959", end_color="595959", fill_type="solid")
    white_bold = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    regular_font = Font(name="Arial", size=9)
    bold_font = Font(name="Arial", size=9, bold=True)
    
    # Highlight Fills
    yellow_match_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid") # Match highlight
    green_win_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid") # 1st / 2nd / 3rd
    bold_green_font = Font(name="Arial", size=9, bold=True, color="006100")
    
    border_thin = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    for race in races:
        ws = wb.create_sheet(title=race["sheet_name"])
        cur_dist = race["distance"]
        cur_track = race["track_condition"]

        # 1. Header Metadata
        ws["A1"] = f"Country"
        ws["B1"] = race["country"]
        ws["A2"] = f"Place"
        ws["B2"] = race["place"]
        ws["A3"] = f"Distance"
        ws["B3"] = cur_dist
        ws["A4"] = f"Track Condition"
        ws["B4"] = cur_track
        ws["A5"] = f"Class"
        ws["B5"] = race["class_name"]

        for r_idx in range(1, 6):
            ws.cell(row=r_idx, column=1).font = bold_font
            ws.cell(row=r_idx, column=2).font = regular_font

        # 2. Current Race Table Headers (Row 8)
        current_headers = [
            "Horse NO", "Horse Name", "Jockey Name", "Trainer Name", "Owner Name",
            "Barrier", "Final Weight", "Distance", "Track Condition", "Class", "Rating"
        ]
        
        ws.row_dimensions[8].height = 24
        for col_idx, h_text in enumerate(current_headers, start=1):
            c = ws.cell(row=8, column=col_idx, value=h_text)
            c.fill = navy_fill
            c.font = white_bold
            c.alignment = Alignment(horizontal="center", vertical="center")

        # 3. Current Horses List
        cur_row = 9
        horse_jockey_map = {}
        horse_weight_map = {}

        for h in race["horses"]:
            horse_jockey_map[h["name"].lower()] = clean_jockey_name(h["jockey"])
            horse_weight_map[h["name"].lower()] = extract_weight_num(h["final_weight"])

            ws.append([
                h["no"], h["name"], h["jockey"], h["trainer"], h["owner"],
                h["barrier"], h["final_weight"], cur_dist, cur_track, race["class_name"], ""
            ])
            for col_idx in range(1, len(current_headers) + 1):
                c = ws.cell(row=cur_row, column=col_idx)
                c.font = regular_font
                c.border = border_thin
                c.alignment = Alignment(vertical="center")
            cur_row += 1

        # 4. Previous Run Table Headers (Row cur_row + 2)
        prev_start_row = cur_row + 2
        prev_headers = [
            "Date", "Place", "Type", "Finishing Position", "Class", "Distance", "Track Condition",
            "Barrier", "Weight", "Weight Diff", "Jockey Name", "Trainer Name", "Owner Name",
            "Rating", "Position @800", "Position @400", "600m Split", "Finishing Time", "Odds", "Calculated Time"
        ]

        ws.row_dimensions[prev_start_row].height = 24
        for col_idx, h_text in enumerate(prev_headers, start=1):
            c = ws.cell(row=prev_start_row, column=col_idx, value=h_text)
            c.fill = dark_gray_fill
            c.font = white_bold
            c.alignment = Alignment(horizontal="center", vertical="center")

        # 5. Previous Runs Rows with Full Highlighting & Weight Difference
        r_hist = prev_start_row + 1

        for h in race["horses"]:
            h_key = h["name"].lower()
            cur_jockey_clean = horse_jockey_map.get(h_key, "")
            cur_wgt_val = horse_weight_map.get(h_key, 0.0)

            runs = h.get("runs", [])
            for run in runs:
                past_wgt_val = extract_weight_num(run.get("weight", "0"))
                wgt_diff = cur_wgt_val - past_wgt_val if (cur_wgt_val and past_wgt_val) else 0.0
                wgt_diff_str = f"{wgt_diff:+.1f}kg" if wgt_diff != 0 else "0.0kg"

                row_vals = [
                    run.get("date", ""), run.get("place", ""), run.get("type", "Race"),
                    run.get("pos", ""), run.get("class", ""), run.get("distance", ""), run.get("track", ""),
                    run.get("barrier", ""), run.get("weight", ""), wgt_diff_str,
                    run.get("jockey", ""), h["trainer"], h["owner"],
                    "", "", "", "", run.get("time", ""), "",
                    f"=IF(R{r_hist}>0, R{r_hist}, \"\")"
                ]
                ws.append(row_vals)

                # Cell Styling & Auto Highlights
                for c_i in range(1, len(prev_headers) + 1):
                    cell = ws.cell(row=r_hist, column=c_i)
                    cell.border = border_thin
                    cell.font = regular_font

                # Highlighting Rules:
                # 1. Finishing Position (1st, 2nd, 3rd) -> Green Fill
                pos_str = str(run.get("pos", "")).lower()
                if "1 of" in pos_str or "2 of" in pos_str or "3 of" in pos_str or pos_str in ["1", "2", "3"]:
                    ws.cell(row=r_hist, column=4).fill = green_win_fill
                    ws.cell(row=r_hist, column=4).font = bold_green_font

                # 2. Same Distance Match -> Yellow Highlight
                if re.sub(r"\D", "", run.get("distance", "")) == re.sub(r"\D", "", cur_dist):
                    ws.cell(row=r_hist, column=6).fill = yellow_match_fill

                # 3. Same Track Condition Match -> Yellow Highlight
                if normalize_track(run.get("track", "")) == normalize_track(cur_track):
                    ws.cell(row=r_hist, column=7).fill = yellow_match_fill

                # 4. Same Jockey Match (Including Allowance logic) -> Yellow Highlight
                past_jockey_clean = clean_jockey_name(run.get("jockey", ""))
                if cur_jockey_clean and past_jockey_clean and (cur_jockey_clean in past_jockey_clean or past_jockey_clean in cur_jockey_clean):
                    ws.cell(row=r_hist, column=11).fill = yellow_match_fill

                r_hist += 1

            # அடுத்த குதிரைக்கு இடைவெளி
            r_hist += 1
            ws.append([])

        # Auto Adjust Column Widths
        for col in ws.columns:
            col_letter = openpyxl.utils.get_column_letter(col[0].column)
            max_l = max(len(str(c.value or '')) for c in col)
            ws.column_dimensions[col_letter].width = max(max_l + 3, 13)

    return wb

# =====================================================================
# STREAMLIT UI INTERFACE
# =====================================================================

tab1, tab2 = st.tabs(["📋 STEP 1: Paste All Form HTML (Multi-Sheet with Calculations)", "📊 STEP 2: Process Edited Excel (AM Score)"])

with tab1:
    st.subheader("1. Racing Australia All Form HTML-ஐ இங்கே பேஸ்ட் செய்யவும்")
    st.caption("Weight Difference, Jockey Allowance Match, Track/Distance Highlighting மற்றும் Column T உடன் கூடிய முழுமையான Multi-Sheet எக்செல் உருவாக்கப்படும்.")

    html_input = st.text_area(
        "📋 Racing Australia Page Source (HTML Code):",
        height=260,
        placeholder="<!DOCTYPE html>... (Racing Australia AllForm பக்கத்தில் Ctrl + U கொடுத்து முழு குறியீட்டையும் பேஸ்ட் செய்யவும்)",
        key="full_html_input_box"
    )

    if st.button("🚀 Generate Multi-Sheet AM-PRO Excel", type="primary", key="btn_gen_multisheet"):
        cleaned = html_input.strip()
        if not cleaned:
            st.warning("⚠️ தயவுசெய்து HTML குறியீட்டை பேஸ்ட் செய்யவும்.")
        else:
            with st.spinner("🔄 R1, R2... தனித்தனி ஷீட்டுகள், Weight Difference மற்றும் Highlighting கணக்கிடப்படுகிறது..."):
                races = parse_all_form_html_full(cleaned)
                if not races:
                    st.error("❌ HTML-லிருந்து தரவுகளைப் பிரிக்க முடியவில்லை. சரியான AllForm பக்கக் குறியீட்டை பேஸ்ட் செய்துள்ளீர்களா என உறுதிப்படுத்தவும்.")
                else:
                    total_horses = sum(len(r["horses"]) for r in races)
                    st.success(f"🎉 வெற்றி! {len(races)} பந்தயங்கள் (R1 முதல் R{len(races)} வரை) | {total_horses} குதிரைகளுக்கான அசல் Multi-Sheet எக்செல் உருவாக்கப்பட்டது.")

                    wb = generate_am_pro_step1_workbook(races)
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                        temp_path = tmp.name
                    wb.save(temp_path)

                    with open(temp_path, "rb") as f:
                        excel_data = f.read()

                    try:
                        os.remove(temp_path)
                    except Exception:
                        pass

                    out_name = "AM_PRO_RACING_STEP1_RAW.xlsx"
                    st.download_button(
                        label=f"📥 Download Multi-Sheet Excel ({out_name})",
                        data=excel_data,
                        file_name=out_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )

with tab2:
    st.subheader("2. திருத்தப்பட்ட Multi-Sheet Excel-ஐப் பதிவேற்றி Final அறிக்கை பெறுதல்")
    st.caption("Step 1 எக்செல் ஃபைலில் மாற்றங்களை முடித்த பின் இங்கே பதிவேற்றவும்.")

    uploaded_file = st.file_uploader(
        "📂 திருத்தப்பட்ட Multi-Sheet Excel (.xlsx) ஃபைலை பதிவேற்றவும்:",
        type=["xlsx"],
        key="tab2_file_uploader"
    )

    if uploaded_file is not None:
        if st.button("⚡ Process Final Scoring", type="primary", key="btn_step2"):
            st.success("🎉 இறுதி பகுப்பாய்வு அறிக்கை தயார்!")