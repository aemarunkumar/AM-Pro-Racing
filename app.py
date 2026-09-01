import streamlit as st
import requests
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
import tempfile
import os
import re
from datetime import datetime

st.set_page_config(
    page_title="AM PRO Racing Mobile",
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
st.markdown("<div class='sub-title'>Step 1: URL to Excel (Scraper & Col T) | Step 2: Process Edited Excel (AM Score & Highlighting)</div>", unsafe_allow_html=True)

# =====================================================================
# ENGINE: SCRAPING & CALCULATIONS (SELF CONTAINED)
# =====================================================================

def parse_time_to_seconds(time_str):
    if not time_str or not isinstance(time_str, str):
        return None
    time_str = time_str.strip()
    try:
        if ":" in time_str:
            parts = time_str.split(":")
            mins = float(parts[0])
            secs = float(parts[1])
            return (mins * 60.0) + secs
        else:
            return float(time_str)
    except Exception:
        return None

def format_seconds_to_time(seconds):
    if seconds is None:
        return ""
    mins = int(seconds // 60)
    secs = seconds % 60
    if mins > 0:
        return f"{mins}:{secs:05.2f}"
    return f"{secs:.2f}"

def scrape_meeting_and_history(url_input):
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    }
    
    # URL Format Setting
    all_form_url = url_input.strip()
    if "Form.aspx" in all_form_url and "AllForm.aspx" not in all_form_url:
        all_form_url = all_form_url.replace("Form.aspx", "AllForm.aspx")

    resp = requests.get(all_form_url, headers=headers, timeout=25)
    resp.raise_encoding = 'utf-8'
    soup = BeautifulSoup(resp.content, "html.parser")

    # Venue & Date Info
    title_text = soup.find("h1") or soup.find("title")
    title_str = title_text.get_text(strip=True) if title_text else "Racing Australia"

    meeting_data = {
        "title": title_str,
        "url": all_form_url,
        "races": []
    }

    # Finding Race Tables
    race_tables = soup.find_all("table", class_=lambda c: c and "race-fields" in c) or soup.find_all("table")

    current_race_idx = 1
    for tbl in race_tables:
        rows = tbl.find_all("tr")
        if not rows or len(rows) < 2:
            continue

        race_horses = []
        for r in rows[1:]:
            cols = [td.get_text(strip=True) for td in r.find_all(["td", "th"])]
            if len(cols) >= 5:
                horse_name = cols[1] if len(cols) > 1 else ""
                jockey_name = cols[2] if len(cols) > 2 else ""
                trainer_name = cols[3] if len(cols) > 3 else ""
                weight = cols[4] if len(cols) > 4 else ""

                if horse_name and not horse_name.lower().startswith("horse"):
                    race_horses.append({
                        "horse_no": cols[0],
                        "horse_name": horse_name,
                        "jockey": jockey_name,
                        "trainer": trainer_name,
                        "weight": weight,
                        "runs": []
                    })

        if race_horses:
            meeting_data["races"].append({
                "race_no": f"Race {current_race_idx}",
                "horses": race_horses
            })
            current_race_idx += 1

    return meeting_data

def generate_step1_workbook(meeting_data):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Master Racing Sheet"

    # Styling
    header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    header_font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
    border_thin = Border(
        left=Side(style='thin', color='D1D5DB'),
        right=Side(style='thin', color='D1D5DB'),
        top=Side(style='thin', color='D1D5DB'),
        bottom=Side(style='thin', color='D1D5DB')
    )

    headers = [
        "Race No", "Horse No", "Horse Name", "Jockey", "Trainer", "Weight", 
        "Date", "Place", "Track", "Dist", "Class", "Pos", "Margin", 
        "Actual Time", "Col T (Calculated Time)", "Track Cond", "AM Score"
    ]
    ws.append(headers)

    for col_idx, cell in enumerate(ws[1], start=1):
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    current_row = 2
    for race in meeting_data["races"]:
        r_name = race["race_no"]
        for h in race["horses"]:
            # Default empty calculated time formula / logic for Column T
            ws.append([
                r_name, h["horse_no"], h["horse_name"], h["jockey"], h["trainer"], h["weight"],
                "", "", "", "", "", "", "", "", "=IF(N{0}>0, N{0}, \"\")".format(current_row), "", ""
            ])
            for col in range(1, len(headers) + 1):
                cell = ws.cell(row=current_row, column=col)
                cell.border = border_thin
                cell.alignment = Alignment(vertical="center")
            current_row += 1

    # Auto-adjust column widths
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    return wb

def apply_step2_scoring(wb):
    ws = wb.active
    high_fill = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid") # Soft Green
    bold_font = Font(name="Arial", size=11, bold=True, color="166534")

    for row in range(2, ws.max_row + 1):
        # Sample Scoring & Rule Highlighting
        track_val = ws.cell(row=row, column=9).value # Track
        pos_val = ws.cell(row=row, column=12).value # Pos

        score = 0
        if str(pos_val) in ["1", "2", "3"]:
            score += 50
            ws.cell(row=row, column=12).fill = high_fill
            ws.cell(row=row, column=12).font = bold_font

        ws.cell(row=row, column=17).value = score # AM Score column

# =====================================================================
# UI TABS
# =====================================================================

tab1, tab2 = st.tabs(["🌐 STEP 1: URL Scraper (Col T)", "📊 STEP 2: Process Edited Excel (AM Score)"])

# ----------------- TAB 1 -----------------
with tab1:
    st.subheader("1. Web URL-லிருந்து Excel உருவாக்குதல்")
    st.caption("Racing Australia All Form URL-ஐ உள்ளிட்டால் Column T கணக்கீட்டுடன் கூடிய ஆரம்ப எக்செல் கிடைக்கும்.")

    input_url = st.text_input(
        "🔗 Racing Australia All Form URL:",
        placeholder="https://racingaustralia.horse/FreeFields/AllForm.aspx?Key=...",
        key="tab1_all_form_url"
    )

    if st.button("🚀 Scrape & Download Step 1 Excel", type="primary", key="btn_step1"):
        cleaned_url = input_url.strip()
        if not cleaned_url:
            st.warning("⚠️ தயவுசெய்து சரியான Racing Australia URL-ஐ உள்ளிடவும்.")
        else:
            status_box = st.status("🔄 தரவுகள் சேகரிக்கப்படுகின்றன...", expanded=True)
            try:
                status_box.write("🌐 1. பந்தய விவரங்கள் ஸ்கிராப் செய்யப்படுகின்றன...")
                meeting_data = scrape_meeting_and_history(cleaned_url)

                if not meeting_data["races"]:
                    status_box.update(label="❌ ரேஸ் விவரங்கள் கிடைக்கவில்லை! URL-ஐ சரிபார்க்கவும்.", state="error")
                    st.error("தரவுகள் கிடைக்கவில்லை. URL சரியானதா என உறுதிப்படுத்தவும்.")
                else:
                    status_box.write(f"✅ {len(meeting_data['races'])} பந்தயங்கள் கண்டறியப்பட்டன.")
                    status_box.write("📊 2. Calculated Time (Col T) எக்செல் உருவாக்கப்படுகிறது...")
                    
                    wb = generate_step1_workbook(meeting_data)

                    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                        temp_path = tmp.name
                    wb.save(temp_path)

                    with open(temp_path, "rb") as f:
                        excel_data = f.read()

                    try:
                        os.remove(temp_path)
                    except Exception:
                        pass

                    match = re.search(r"Key=([^&#]+)", cleaned_url)
                    file_key = match.group(1).replace("%2C", "_") if match else "RACING_DATA"
                    out_name = f"{file_key}_STEP1_RAW.xlsx"

                    status_box.update(label="🎉 Step 1 Excel தயாராகிவிட்டது!", state="complete", expanded=False)
                    st.success("✅ முதல் நிலை எக்செல் ஃபைல் தயார்! டவுன்லோட் செய்து தேவையான திருத்தங்களை செய்யவும்.")
                    st.download_button(
                        label=f"📥 Download {out_name}",
                        data=excel_data,
                        file_name=out_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
            except Exception as e:
                status_box.update(label="❌ செயலாக்கத்தில் பிழை ஏற்பட்டது!", state="error")
                st.error(f"Error Details: {str(e)}")

# ----------------- TAB 2 -----------------
with tab2:
    st.subheader("2. எடிட் செய்த Excel-ஐ பதிவேற்றி Final அறிக்கை பெறுதல்")
    st.caption("Step 1 எக்செல் ஃபைலில் மாற்றங்களை முடித்த பின் இங்கே பதிவேற்றவும்.")

    uploaded_file = st.file_uploader(
        "📂 திருத்தப்பட்ட Excel (.xlsx) ஃபைலை பதிவேற்றவும்:",
        type=["xlsx"],
        key="tab2_file_uploader"
    )

    if uploaded_file is not None:
        if st.button("⚡ Process Final Scoring & Highlights", type="primary", key="btn_step2"):
            with st.spinner("🔄 AM Score மற்றும் ஹைலைட்டிங் கணக்கிடப்படுகிறது..."):
                try:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                        tmp.write(uploaded_file.read())
                        temp_in = tmp.name

                    wb = openpyxl.load_workbook(temp_in, data_only=False)
                    apply_step2_scoring(wb)
                    wb.save(temp_in)

                    with open(temp_in, "rb") as f:
                        final_data = f.read()

                    try:
                        os.remove(temp_in)
                    except Exception:
                        pass

                    base_name = os.path.splitext(uploaded_file.name)[0].replace("_STEP1_RAW", "")
                    out_name = f"{base_name}_AM_PRO_FINAL.xlsx"

                    st.success("🎉 இறுதி பகுப்பாய்வு அறிக்கை தயார்!")
                    st.download_button(
                        label=f"📥 Download {out_name}",
                        data=final_data,
                        file_name=out_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                except Exception as e:
                    st.error(f"Error: {str(e)}")