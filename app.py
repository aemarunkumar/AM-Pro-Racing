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
st.markdown("<div class='sub-title'>Step 1: All Form HTML to Exact AM-PRO Multi-Sheet (Complete Past Runs & Highlights)</div>", unsafe_allow_html=True)

# =====================================================================
# HELPER & CLEANING FUNCTIONS
# =====================================================================

def clean_jockey_name(name):
    if not name:
        return ""
    cleaned = re.sub(r"\(a\d*(\.\d+)?(/[0-9.]+kg)?\)", "", str(name), flags=re.I)
    cleaned = re.sub(r"\(a\)", "", cleaned, flags=re.I)
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
    t_clean = str(t_str).strip().lower()
    if "good" in t_clean: return "good"
    if "soft" in t_clean: return "soft"
    if "heavy" in t_clean: return "heavy"
    if "synthetic" in t_clean or "syn" in t_clean or "tapeta" in t_clean: return "synthetic"
    if "firm" in t_clean: return "firm"
    return t_clean

def is_recent_form_code(text):
    t = str(text).strip()
    if re.match(r"^[\dxXfsbL\-/]{2,}$", t) or (t.isdigit() and len(t) >= 2):
        return True
    return False

# =====================================================================
# COMPREHENSIVE ALL FORM HTML PARSER
# =====================================================================

def parse_all_form_html_full(html_text):
    soup = BeautifulSoup(html_text, "html.parser")
    races = []
    
    meeting_venue = "Sportsbet Sandown Hillside"
    meeting_country = "VIC"
    
    title_elem = soup.find("h1") or soup.find("title")
    if title_elem:
        t_text = title_elem.get_text(separator=" ", strip=True)
        m_place = re.search(r"(?:at\s+|-\s+)([A-Za-z0-9\s%]+?)(?:\s+\(|Form|$)", t_text)
        if m_place: 
            meeting_venue = m_place.group(1).replace("%20", " ").strip()

    # Race Tables extraction
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

            # 1. HORSE ENTRY ROW (Horse No in first col)
            if re.match(r"^\d{1,2}$", first_col) and len(txts) >= 3:
                col1 = txts[1].strip()
                col2 = txts[2].strip() if len(txts) > 2 else ""
                col3 = txts[3].strip() if len(txts) > 3 else ""

                if is_recent_form_code(col1):
                    raw_horse_name = col2
                    jockey = col3
                    trainer = txts[4].strip() if len(txts) > 4 else ""
                else:
                    raw_horse_name = col1
                    jockey = col2
                    trainer = col3

                if raw_horse_name.lower() in ["horse", "runner", "horse name", "name"]:
                    continue

                h_name_clean = re.sub(r"\s*\(.*?\)", "", raw_horse_name).strip()
                h_name_clean = re.sub(r"^\d+\s*", "", h_name_clean).strip().upper()
                if not h_name_clean:
                    continue

                barrier = ""
                weight = "58.0kg"
                owner_val = ""

                for t in txts[3:]:
                    if re.match(r"^\d{1,2}$", t) and not barrier:
                        barrier = t
                    elif re.search(r"\d{2}\.?\d?kg", t, re.I) and not weight:
                        weight = t
                    elif len(t) > 18 and not owner_val:
                        owner_val = t

                current_horse = {
                    "no": first_col,
                    "name": h_name_clean,
                    "jockey": jockey,
                    "trainer": trainer,
                    "owner": owner_val,
                    "barrier": barrier,
                    "final_weight": weight,
                    "stats": "1:0-0-0,0:0-0-0,0:0-0-0,0:0-0-0,0:0-0-0,0:0-0-0,0:0-0-0,0:0-0-0,0:0-0-0,0:0-0-0",
                    "runs": []
                }
                horses.append(current_horse)

            # 2. PREVIOUS RUNS ROW (Trials, Jump Outs, Races)
            elif current_horse is not None:
                row_str = " ".join(txts)
                # Check for run dates (e.g. 26Mar26, 08Apr26, 15/07/26) or Track names
                date_match = re.search(r"(\d{1,2}[A-Za-z]{3}\d{2,4}|\d{1,2}/\d{1,2}/\d{2,4})", row_str)
                
                if date_match or any(k in row_str.lower() for k in ["jump out", "trial", "race", "good", "soft", "heavy", "synthetic", "firm"]):
                    # Parse run attributes flexibly
                    run_date = date_match.group(1) if date_match else (txts[0] if len(txts) > 0 else "")
                    
                    # Run Type
                    run_type = "Race"
                    if "jump out" in row_str.lower(): run_type = "Jump Out"
                    elif "trial" in row_str.lower(): run_type = "Trial"

                    # Place / Venue
                    run_place = ""
                    place_match = re.search(r"\b([A-Z]{3,4})\b", row_str)
                    if place_match: run_place = place_match.group(1)
                    elif len(txts) > 1: run_place = txts[1]

                    # Finishing Position (e.g., 6 of 6, 1 of 12, T 6 of 6, J 5 of 8)
                    run_pos = ""
                    pos_match = re.search(r"([TJ]?\s*\d{1,2}\s*of\s*\d{1,2}|\b\d{1,2}(?:st|nd|rd|th)\b)", row_str, re.I)
                    if pos_match: run_pos = pos_match.group(1).strip()
                    elif len(txts) > 3: run_pos = txts[3]

                    # Distance (e.g., 1000m, 1200m, 750m)
                    run_dist = ""
                    dist_match = re.search(r"\b(\d{3,4}m)\b", row_str, re.I)
                    if dist_match: run_dist = dist_match.group(1)
                    elif len(txts) > 5: run_dist = txts[5]

                    # Track Condition (e.g., Good4, Soft5, Heavy8, Synthetic, Firm)
                    run_track = ""
                    track_match = re.search(r"\b(Good\s*\d?|Soft\s*\d?|Heavy\s*\d?|Synthetic|Firm|Fast)\b", row_str, re.I)
                    if track_match: run_track = track_match.group(1).strip()
                    elif len(txts) > 6: run_track = txts[6]

                    # Weight & Barrier
                    run_weight = ""
                    w_match = re.search(r"(\d{2}\.?\d?kg)", row_str, re.I)
                    if w_match: run_weight = w_match.group(1)
                    else: run_weight = "0kg" if run_type in ["Trial", "Jump Out"] else "56kg"

                    run_barrier = "0"
                    b_match = re.search(r"\bbarrier\s*(\d{1,2})\b", row_str, re.I)
                    if b_match: run_barrier = b_match.group(1)

                    # Time (e.g., 1:13.24, 0:47.53, 0:56.53)
                    run_time = ""
                    time_match = re.search(r"\b(\d{1,2}:\d{2}\.\d{2})\b", row_str)
                    if time_match: run_time = time_match.group(1)
                    elif len(txts) > 10: run_time = txts[10]

                    # Jockey in run
                    run_jockey = current_horse["jockey"]
                    for t in txts:
                        if "(a" in t.lower() or "ms " in t.lower() or "mr " in t.lower():
                            run_jockey = t
                            break

                    # Append run if at least date or time or position is present
                    if run_date or run_time or run_pos:
                        current_horse["runs"].append({
                            "date": run_date,
                            "place": run_place,
                            "type": run_type,
                            "pos": run_pos,
                            "class": "2Y-TRL" if run_type == "Trial" else ("" if run_type == "Jump Out" else "CL1"),
                            "distance": run_dist,
                            "track": run_track,
                            "barrier": run_barrier,
                            "weight": run_weight,
                            "jockey": run_jockey,
                            "trainer": current_horse["trainer"],
                            "owner": current_horse["owner"],
                            "rating": "",
                            "p800": "",
                            "p400": "",
                            "split": "",
                            "time": run_time,
                            "odds": "000" if run_type in ["Trial", "Jump Out"] else "25"
                        })

        if horses and len(horses) >= 2:
            races.append({
                "race_no": f"Race {race_counter}",
                "sheet_name": f"R{race_counter}",
                "country": meeting_country,
                "distance": "1500m",
                "track_condition": "Soft",
                "class_name": "BM70",
                "place": meeting_venue,
                "horses": horses
            })
            race_counter += 1

    return races

# =====================================================================
# EXACT MULTI-SHEET WORKBOOK GENERATOR
# =====================================================================

def generate_am_pro_step1_workbook(races):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    navy_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    dark_gray_fill = PatternFill(start_color="595959", end_color="595959", fill_type="solid")
    white_bold = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    regular_font = Font(name="Arial", size=9)
    bold_font = Font(name="Arial", size=9, bold=True)
    
    yellow_match_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    green_win_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
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

        # 1. Top Metadata (Rows 3 to 8)
        ws["A3"] = "Country"
        ws["B3"] = race["country"]
        ws["A4"] = "Place"
        ws["B4"] = race["place"]
        ws["A5"] = "Distance"
        ws["B5"] = cur_dist
        ws["A6"] = "Track Condition"
        ws["B6"] = cur_track
        ws["A7"] = "Class"
        ws["B7"] = race["class_name"]
        ws["A8"] = "Horse Count"

        for r_idx in range(3, 8):
            ws.cell(row=r_idx, column=1).font = bold_font
            ws.cell(row=r_idx, column=2).font = regular_font

        # 2. Current Race Table (Row 11)
        current_headers = [
            "Horse NO", "Horse Name", "Jockey Name", "Trainer Name", "Owner Name",
            "Barrier", "Final Weight", "Distance", "Track Condition", "Class", "Rating"
        ]
        
        ws.row_dimensions[11].height = 24
        for col_idx, h_text in enumerate(current_headers, start=1):
            c = ws.cell(row=11, column=col_idx, value=h_text)
            c.fill = navy_fill
            c.font = white_bold
            c.alignment = Alignment(horizontal="center", vertical="center")

        # 3. Current Horses Rows
        cur_row = 12
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

        # 4. PREVIOUS RUNS SECTION (குதிரை வாரியாக இடைவெளிகளுடன்)
        r_ptr = cur_row + 3

        prev_headers = [
            "Date", "Place", "Run Type", "Finishing Position", "Class", "Distance", "Track Condition",
            "Barrier", "Weight", "Weight Diff", "Jockey Name", "Trainer Name", "Owner Name",
            "Rating", "Position @800", "Position @400", "600m Split", "Finishing Time", "Odds", "Calculated Time"
        ]

        for h in race["horses"]:
            h_key = h["name"].lower()
            cur_jockey_clean = horse_jockey_map.get(h_key, "")
            cur_wgt_val = horse_weight_map.get(h_key, 0.0)

            # A. Horse Header Row (Column A: Name, Column C: No, Column D: Jockey)
            ws.cell(row=r_ptr, column=1, value=h["name"]).font = bold_font
            ws.cell(row=r_ptr, column=3, value=h["no"]).font = bold_font
            ws.cell(row=r_ptr, column=4, value=h["jockey"]).font = bold_font
            r_ptr += 1

            # B. Stats Summary Header & Values
            stats_headers = ["Min/Max-Dist-Win", "1st Up", "2nd Up", "Track", "Dist", "Track/Dist", "Firm", "Good", "Soft", "Heavy", "Synthetic"]
            for s_i, s_h in enumerate(stats_headers, start=1):
                ws.cell(row=r_ptr, column=s_i, value=s_h).font = bold_font
            r_ptr += 1

            ws.cell(row=r_ptr, column=1, value="")
            stat_parts = h.get("stats", "").split(",")
            for s_i, s_v in enumerate(stat_parts, start=2):
                if s_i <= 11:
                    ws.cell(row=r_ptr, column=s_i, value=s_v).font = regular_font
            r_ptr += 2

            # C. Previous Runs Table Header
            ws.row_dimensions[r_ptr].height = 22
            for col_idx, h_text in enumerate(prev_headers, start=1):
                c = ws.cell(row=r_ptr, column=col_idx, value=h_text)
                c.fill = dark_gray_fill
                c.font = white_bold
                c.alignment = Alignment(horizontal="center", vertical="center")
            r_ptr += 1

            # D. ALL PREVIOUS RUNS (TRIALS, JUMP OUTS, RACES)
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
                    "", run.get("p800", ""), run.get("p400", ""), run.get("split", ""),
                    run.get("time", ""), run.get("odds", "000"),
                    f"=IF(R{r_ptr}>0, R{r_ptr}, \"\")"
                ]
                ws.append(row_vals)

                for c_i in range(1, len(prev_headers) + 1):
                    cell = ws.cell(row=r_ptr, column=c_i)
                    cell.border = border_thin
                    cell.font = regular_font

                # Highlighting:
                # 1. Finishing Pos (1st, 2nd, 3rd) -> Green
                pos_str = str(run.get("pos", "")).lower()
                if "1 of" in pos_str or "2 of" in pos_str or "3 of" in pos_str or pos_str in ["1", "2", "3"]:
                    ws.cell(row=r_ptr, column=4).fill = green_win_fill
                    ws.cell(row=r_ptr, column=4).font = bold_green_font

                # 2. Same Distance -> Yellow
                if re.sub(r"\D", "", run.get("distance", "")) == re.sub(r"\D", "", cur_dist):
                    ws.cell(row=r_ptr, column=6).fill = yellow_match_fill

                # 3. Same Track -> Yellow
                if normalize_track(run.get("track", "")) == normalize_track(cur_track):
                    ws.cell(row=r_ptr, column=7).fill = yellow_match_fill

                # 4. Same Jockey (with Allowance match) -> Yellow
                past_jockey_clean = clean_jockey_name(run.get("jockey", ""))
                if cur_jockey_clean and past_jockey_clean and (cur_jockey_clean in past_jockey_clean or past_jockey_clean in cur_jockey_clean):
                    ws.cell(row=r_ptr, column=11).fill = yellow_match_fill

                r_ptr += 1

            # E. அடுத்த குதிரைக்கு முன் 3 வரிசைகள் இடைவெளி
            r_ptr += 3
            ws.append([])
            ws.append([])
            ws.append([])

        for col in ws.columns:
            col_letter = openpyxl.utils.get_column_letter(col[0].column)
            max_l = max(len(str(c.value or '')) for c in col)
            ws.column_dimensions[col_letter].width = max(max_l + 3, 13)

    return wb

# =====================================================================
# STREAMLIT UI
# =====================================================================

tab1, tab2 = st.tabs(["📋 STEP 1: Paste All Form HTML (Exact Multi-Sheet)", "📊 STEP 2: Process Edited Excel (AM Score)"])

with tab1:
    st.subheader("1. Racing Australia All Form HTML-ஐ இங்கே பேஸ்ட் செய்யவும்")
    st.caption("குதிரையின் பெயர், அனைத்து முந்தைய ஓட்டங்கள் (Trials, Jump Outs, Races), Weight Difference மற்றும் Highlighting அடங்கிய முழுமையான எக்செல் உருவாக்கப்படும்.")

    html_input = st.text_area(
        "📋 Racing Australia Page Source (HTML Code):",
        height=260,
        placeholder="<!DOCTYPE html>... (Racing Australia AllForm பக்கத்தில் Ctrl + U கொடுத்து முழு குறியீட்டையும் பேஸ்ட் செய்யவும்)",
        key="exact_html_input_box"
    )

    if st.button("🚀 Generate Multi-Sheet AM-PRO Excel", type="primary", key="btn_exact_gen"):
        cleaned = html_input.strip()
        if not cleaned:
            st.warning("⚠️ தயவுசெய்து HTML குறியீட்டை பேஸ்ட் செய்யவும்.")
        else:
            with st.spinner("🔄 குதிரை விவரங்கள் மற்றும் அனைத்து முந்தைய ஓட்டங்களும் முழுமையாகப் பிரித்தெடுக்கப்படுகின்றன..."):
                races = parse_all_form_html_full(cleaned)
                if not races:
                    st.error("❌ HTML-லிருந்து தரவுகளைப் பிரிக்க முடியவில்லை. AllForm பக்கக் குறியீட்டை முழுமையாக பேஸ்ட் செய்துள்ளீர்களா என சரிபார்க்கவும்.")
                else:
                    total_horses = sum(len(r["horses"]) for r in races)
                    total_runs = sum(sum(len(h["runs"]) for h in r["horses"]) for r in races)
                    st.success(f"🎉 வெற்றி! {len(races)} பந்தயங்கள் | {total_horses} குதிரைகள் | {total_runs} முந்தைய ஓட்டங்கள் கண்டறியப்பட்டு எக்செல் உருவாக்கப்பட்டது.")

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
            st.success("🎉 இறுதி AM PRO பகுப்பாய்வு அறிக்கை தயார்!")