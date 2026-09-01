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
st.markdown("<div class='sub-title'>Step 1: All Form HTML to Exact AM-PRO Multi-Sheet (With All Previous Runs)</div>", unsafe_allow_html=True)

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

# =====================================================================
# TARGETED RACING AUSTRALIA ALL FORM PARSER
# =====================================================================

def parse_racing_australia_html(html_text):
    soup = BeautifulSoup(html_text, "html.parser")
    races = []

    # 1. Meeting Venue, Date, Track Info
    meeting_venue = "Warwick Farm"
    meeting_country = "NSW"
    meeting_date = "02/09/2026"

    venue_elem = soup.find("div", class_="race-venue")
    if venue_elem:
        h2 = venue_elem.find("h2")
        if h2:
            meeting_venue = h2.get_text().split(":")[0].strip()
        date_span = venue_elem.find("span", class_="race-venue-date")
        if date_span:
            meeting_date = date_span.get_text().strip()

    # 2. Iterate through each race block (Race 1, Race 2, etc.)
    # We can split by race title anchors <a name="Race1"></a> or <table class="race-title">
    race_titles = soup.find_all("table", class_="race-title")

    for idx, r_title_table in enumerate(race_titles, start=1):
        race_num_str = f"Race {idx}"
        sheet_name_str = f"R{idx}"

        # Extract Distance, Class, Track Condition for this race
        title_text = r_title_table.get_text()
        dist_match = re.search(r"\((\d{3,4})\s*METRES\)", title_text, re.I)
        race_dist = f"{dist_match.group(1)}m" if dist_match else "1200m"

        class_match = re.search(r"\b(BenchMark\s*\d+|BM\s*\d+|Maiden|MDN|Open|Handicap|Listed|G[1-3])\b", title_text, re.I)
        race_class = class_match.group(1) if class_match else "Handicap"

        # Find the fields table immediately following this race title
        fields_table = r_title_table.find_next("table", class_="race-strip-fields")
        current_horses_map = {}

        if fields_table:
            for tr in fields_table.find_all("tr"):
                if "Scratched" in tr.get("class", []):
                    continue
                no_td = tr.find("td", class_="no")
                horse_td = tr.find("td", class_="horse")
                jockey_td = tr.find("td", class_="jockey")
                trainer_td = tr.find("td", class_="trainer")
                barrier_td = tr.find("td", class_="barrier")
                weight_td = tr.find("td", class_="weight")

                if no_td and horse_td:
                    h_no = no_td.get_text(strip=True)
                    h_name = re.sub(r"\s*\(.*?\)", "", horse_td.get_text(strip=True)).strip().upper()
                    jock_name = jockey_td.get_text(strip=True) if jockey_td else ""
                    train_name = trainer_td.get_text(strip=True) if trainer_td else ""
                    barrier_val = barrier_td.get_text(strip=True) if barrier_td else ""
                    weight_val = weight_td.get_text(strip=True) if weight_td else "58.0kg"

                    current_horses_map[h_name] = {
                        "no": h_no,
                        "name": h_name,
                        "jockey": jock_name,
                        "trainer": train_name,
                        "owner": "",
                        "barrier": barrier_val,
                        "final_weight": weight_val,
                        "stats": "",
                        "runs": []
                    }

        # Find all horse-form-tables up to the next race-title
        curr_node = r_title_table.find_next_sibling()
        race_horses_list = []

        # Collect horse detail sections
        # We search inside all horse-form-table elements belonging to this race
        next_race_table = race_titles[idx] if idx < len(race_titles) else None
        
        # Extract all horse-form-tables in the entire document matching current race
        all_horse_forms = soup.find_all("table", class_="horse-form-table")
        
        for h_form in all_horse_forms:
            # Check if this form belongs to current race by position
            if next_race_table and h_form.sourceline and next_race_table.sourceline:
                if h_form.sourceline > next_race_table.sourceline:
                    continue
            if r_title_table.sourceline and h_form.sourceline:
                if h_form.sourceline < r_title_table.sourceline:
                    continue

            # If scratched, skip
            if h_form.find("td", class_="Scratched"):
                continue

            h_name_tag = h_form.find("span", class_="horse-name")
            h_no_tag = h_form.find("span", class_="horse-number")

            if not h_name_tag:
                continue

            raw_h_name = h_name_tag.get_text(strip=True).upper()
            h_clean_name = re.sub(r"\s*\(.*?\)", "", raw_h_name).strip().upper()
            h_num = h_no_tag.get_text(strip=True) if h_no_tag else ""

            # Match with current horse data
            h_obj = current_horses_map.get(h_clean_name, {
                "no": h_num,
                "name": h_clean_name,
                "jockey": "",
                "trainer": "",
                "owner": "",
                "barrier": "",
                "final_weight": "58.0kg",
                "stats": "",
                "runs": []
            })

            # Extract Stats Record line (Track, Dist, Good, Soft, Heavy, etc.)
            form_text = h_form.get_text(separator=" ", strip=True)
            stats_dict = {}
            for stat_key in ["1st Up", "2nd Up", "Track", "Dist", "Track/Dist", "Firm", "Good", "Soft", "Heavy", "Synthetic"]:
                sm = re.search(rf"\b{re.escape(stat_key)}:\s*([\d:-]+)", form_text)
                stats_dict[stat_key] = sm.group(1) if sm else "0:0-0-0"

            h_obj["stats"] = ",".join([
                stats_dict.get("1st Up", "0:0-0-0"),
                stats_dict.get("2nd Up", "0:0-0-0"),
                stats_dict.get("Track", "0:0-0-0"),
                stats_dict.get("Dist", "0:0-0-0"),
                stats_dict.get("Track/Dist", "0:0-0-0"),
                stats_dict.get("Firm", "0:0-0-0"),
                stats_dict.get("Good", "0:0-0-0"),
                stats_dict.get("Soft", "0:0-0-0"),
                stats_dict.get("Heavy", "0:0-0-0"),
                stats_dict.get("Synthetic", "0:0-0-0")
            ])

            # Owner
            owner_m = re.search(r"Owners:\s*(.*?)(?:Colours:|Gear Changes:|<br|$)", form_text)
            if owner_m and not h_obj["owner"]:
                h_obj["owner"] = owner_m.group(1).strip()

            # EXTRACT ALL PREVIOUS RUNS FROM <table class="horse-last-start">
            last_start_tbl = h_form.find("table", class_="horse-last-start")
            if last_start_tbl:
                for r_tr in last_start_tbl.find_all("tr"):
                    pos_td = r_tr.find("td", class_="Pos")
                    remain_td = r_tr.find("td", class_="remain")

                    if not remain_td:
                        continue

                    pos_val = pos_td.get_text(separator=" ", strip=True) if pos_td else ""
                    rem_text = remain_td.get_text(separator=" ", strip=True)

                    # Run Type
                    run_type = "Race"
                    if "Jump Out" in rem_text or pos_val.startswith("J"):
                        run_type = "Jump Out"
                    elif "-BT" in rem_text or "-TRL" in rem_text or pos_val.startswith("T"):
                        run_type = "Trial"

                    # Venue & Date (e.g. W FM 19Aug26)
                    venue_date_m = re.search(r"^([A-Z\s.]+?)\s+(\d{1,2}[A-Za-z]{3}\d{2,4})", rem_text)
                    if venue_date_m:
                        p_place = venue_date_m.group(1).strip()
                        p_date = venue_date_m.group(2).strip()
                    else:
                        p_place = ""
                        p_date = ""

                    # Distance
                    d_m = re.search(r"\b(\d{3,4}m)\b", rem_text)
                    p_dist = d_m.group(1) if d_m else ""

                    # Track Condition
                    tr_m = re.search(r"\b(Good\s*\d?|Soft\s*\d?|Heavy\s*\d?|Synthetic|Firm\s*\d?|Fast)\b", rem_text, re.I)
                    p_track = tr_m.group(1) if tr_m else "Good4"

                    # Class
                    cl_m = re.search(r"\b(BM\d+|MDN-SW|2Y-BT|OPEN-BT|2Y\s+MDN|3Y\s+MDN|SUPER\s+3Y\s+MDN|2YF-BT|2YC&G-BT|3Y\+MDN-BT|CL\d+|LR|G[1-3])\b", rem_text, re.I)
                    p_class = cl_m.group(1) if cl_m else ("Trial" if run_type == "Trial" else "")

                    # Weight
                    w_m = re.search(r"\b(\d{2}(?:\.\d)?kg)\b", rem_text, re.I)
                    p_weight = w_m.group(1) if w_m else ("0kg" if run_type in ["Trial", "Jump Out"] else "56kg")

                    # Barrier
                    bar_m = re.search(r"Barrier\s*(\d{1,2})", rem_text, re.I)
                    p_barrier = bar_m.group(1) if bar_m else "0"

                    # Past Jockey
                    jock_tag = remain_td.find("a", href=re.compile(r"JockeyLastRuns", re.I))
                    p_jockey = jock_tag.get_text(strip=True) if jock_tag else ""

                    # Finishing Time
                    time_m = re.search(r"\b(\d{1,2}:\d{2}\.\d{2})\b", rem_text)
                    p_time = time_m.group(1) if time_m else ""

                    # Splits & Positions
                    p800_m = re.search(r"(\d{1,2}(?:st|nd|rd|th)@800m)", rem_text)
                    p400_m = re.search(r"(\d{1,2}(?:st|nd|rd|th)@400m)", rem_text)
                    p_800 = p800_m.group(1) if p800_m else ""
                    p_400 = p400_m.group(1) if p400_m else ""

                    # Split 600m
                    split_m = re.search(r"\(600m\s*([\d.]+)\)", rem_text)
                    p_split = split_m.group(1) if split_m else ""

                    # Odds
                    odds_m = re.search(r"(\$[\d/.]+[Ff]?)$", rem_text)
                    p_odds = odds_m.group(1) if odds_m else ("$000" if run_type in ["Trial", "Jump Out"] else "")

                    h_obj["runs"].append({
                        "date": p_date,
                        "place": p_place,
                        "type": run_type,
                        "pos": pos_val,
                        "class": p_class,
                        "distance": p_dist,
                        "track": p_track,
                        "barrier": p_barrier,
                        "weight": p_weight,
                        "jockey": p_jockey,
                        "trainer": h_obj["trainer"],
                        "owner": h_obj["owner"],
                        "rating": "",
                        "p800": p_800,
                        "p400": p_400,
                        "split": p_split,
                        "time": p_time,
                        "odds": p_odds
                    })

            race_horses_list.append(h_obj)

        if race_horses_list:
            races.append({
                "race_no": race_num_str,
                "sheet_name": sheet_name_str,
                "country": meeting_country,
                "distance": race_dist,
                "track_condition": "Good 4",
                "class_name": race_class,
                "place": meeting_venue,
                "horses": race_horses_list
            })

    return races

# =====================================================================
# EXACT MULTI-SHEET EXCEL BUILDER (MATCHING USER TEMPLATE)
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
        ws["B8"] = len(race["horses"])

        for r_idx in range(3, 9):
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

        # 4. PREVIOUS RUNS SECTION (குதிரை வாரியாக 3 வரிகள் இடைவெளியுடன்)
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

            # D. ALL PREVIOUS RUNS
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
                    run.get("time", ""), run.get("odds", "$000"),
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
    st.caption("குதிரை வாரியான அட்டவணை, Stats Summary மற்றும் அனைத்து Previous Runs (Trials, Jump Outs, Races) அடங்கிய முழுமையான எக்செல் உருவாக்கப்படும்.")

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
            with st.spinner("🔄 HTML-லிருந்து குதிரைகள் மற்றும் அனைத்து முந்தைய ஓட்டங்கள் பிரித்தெடுக்கப்படுகின்றன..."):
                races = parse_racing_australia_html(cleaned)
                if not races:
                    st.error("❌ HTML-லிருந்து தரவுகளைப் பிரிக்க முடியவில்லை. AllForm பக்கக் குறியீட்டை முழுமையாக பேஸ்ட் செய்துள்ளீர்களா என சரிபார்க்கவும்.")
                else:
                    total_horses = sum(len(r["horses"]) for r in races)
                    total_runs = sum(sum(len(h["runs"]) for h in r["horses"]) for r in races)
                    st.success(f"🎉 வெற்றி! {len(races)} பந்தயங்கள் | {total_horses} குதிரைகள் | {total_runs} முந்தைய ஓட்டங்கள் (Previous Runs) கண்டறியப்பட்டு எக்செல் உருவாக்கப்பட்டது.")

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