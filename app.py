import streamlit as st
import requests
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
import tempfile
import os
import re

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
# RACING AUSTRALIA SCRAPER ENGINE
# =====================================================================

def fetch_html(url):
    session = requests.Session()
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8",
        "Accept-Language": "en-US,en;q=0.9",
        "Referer": "https://racingaustralia.horse/",
        "Connection": "keep-alive"
    }
    response = session.get(url, headers=headers, timeout=30)
    response.encoding = 'utf-8'
    return response.text

def parse_racing_australia(all_form_url):
    html_content = fetch_html(all_form_url)
    soup = BeautifulSoup(html_content, "html.parser")

    races = []
    # Racing Australia race tables extraction
    race_divs = soup.find_all("div", class_=lambda c: c and ("race-field" in c.lower() or "race" in c.lower()))
    
    # Check tables directly if divs not formatted
    tables = soup.find_all("table")
    current_race_num = 1

    for tbl in tables:
        rows = tbl.find_all("tr")
        if len(rows) < 2:
            continue
        
        horses = []
        for r in rows:
            tds = r.find_all("td")
            if not tds or len(tds) < 3:
                continue

            texts = [td.get_text(separator=" ", strip=True) for td in tds]
            
            # Check if row is a valid horse entry (Starts with number)
            first_col = texts[0].strip()
            if first_col.isdigit():
                h_no = first_col
                h_name = texts[1] if len(texts) > 1 else ""
                jockey = texts[2] if len(texts) > 2 else ""
                trainer = texts[3] if len(texts) > 3 else ""
                weight = texts[4] if len(texts) > 4 else ""

                # Cleanup horse name (Remove brackets/stats if attached)
                h_name_clean = re.sub(r"\s*\(.*?\)", "", h_name).strip()

                horses.append({
                    "horse_no": h_no,
                    "horse_name": h_name_clean,
                    "jockey": jockey,
                    "trainer": trainer,
                    "weight": weight
                })

        if horses and len(horses) >= 2:
            races.append({
                "race_name": f"Race {current_race_num}",
                "horses": horses
            })
            current_race_num += 1

    return races

def build_excel(races, venue_date_key):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Master_Sheet"

    header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    header_font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
    border_thin = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )

    headers = [
        "Race No", "Horse No", "Horse Name", "Jockey Name", "Trainer Name", 
        "Weight", "Date", "Place", "Distance", "Track Condition", 
        "Class", "Finishing Pos", "Margin", "Finishing Time", "Col T (Calculated Time)", 
        "Track Match", "Class Match", "Jockey Match", "AM Score", "Status"
    ]
    ws.append(headers)

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    row_idx = 2
    for r in races:
        race_title = r["race_name"]
        for h in r["horses"]:
            # Initial Row Setup with Step 1 Formulas
            ws.append([
                race_title,
                h["horse_no"],
                h["horse_name"],
                h["jockey"],
                h["trainer"],
                h["weight"],
                "", "", "", "", "", "", "", "",
                f"=IF(N{row_idx}>0, N{row_idx}, \"\")",
                "", "", "", "", ""
            ])
            for col in range(1, len(headers) + 1):
                c = ws.cell(row=row_idx, column=col)
                c.border = border_thin
                c.alignment = Alignment(vertical="center")
            row_idx += 1

    # Format Widths
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 13)

    return wb

def apply_am_scoring_step2(wb):
    ws = wb.active
    high_fill = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
    font_bold = Font(name="Arial", size=11, bold=True, color="166534")

    for r in range(2, ws.max_row + 1):
        pos_val = str(ws.cell(row=r, column=12).value or "").strip()
        
        # Scoring logic
        score = 0
        if pos_val in ["1", "2", "3"]:
            score += 50
            ws.cell(row=r, column=12).fill = high_fill
            ws.cell(row=r, column=12).font = font_bold
            ws.cell(row=r, column=20).value = "QUALIFIED"

        ws.cell(row=r, column=19).value = score

# =====================================================================
# APP INTERFACE
# =====================================================================

tab1, tab2 = st.tabs(["🌐 STEP 1: URL Scraper (Col T)", "📊 STEP 2: Process Edited Excel (AM Score)"])

with tab1:
    st.subheader("1. Web URL-லிருந்து Excel உருவாக்குதல்")
    st.caption("All Form URL-ஐ உள்ளிட்டால் Column T (Calculated Time) உடன் ஆரம்ப எக்செல் கிடைக்கும்.")

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
            status_box = st.status("🔄 பந்தய தகவல்கள் பெறப்படுகின்றன...", expanded=True)
            try:
                status_box.write("🌐 1. Racing Australia இணையதளத்திலிருந்து தரவுகள் எடுக்கப்படுகின்றன...")
                races = parse_racing_australia(cleaned_url)

                if not races:
                    status_box.update(label="❌ ரேஸ் விவரங்கள் கிடைக்கவில்லை! URL-ஐ சரிபார்க்கவும்.", state="error")
                    st.error("பந்தய தகவல்கள் கிடைக்கவில்லை. இணையதள இணைப்பு சரியானதா என உறுதிப்படுத்தவும்.")
                else:
                    status_box.write(f"✅ {len(races)} பந்தயங்கள் கண்டறியப்பட்டன.")
                    status_box.write("📊 2. Column T (Calculated Time) எக்செல் வடிவில் உருவாக்கப்படுகிறது...")

                    match = re.search(r"Key=([^&#]+)", cleaned_url)
                    file_key = match.group(1).replace("%2C", "_") if match else "RACING_DATA"
                    out_filename = f"{file_key}_STEP1_RAW.xlsx"

                    wb = build_excel(races, file_key)

                    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                        temp_path = tmp.name
                    wb.save(temp_path)

                    with open(temp_path, "rb") as f:
                        excel_data = f.read()

                    try:
                        os.remove(temp_path)
                    except Exception:
                        pass

                    status_box.update(label="🎉 Step 1 Excel தயாராகிவிட்டது!", state="complete", expanded=False)
                    st.success("✅ முதல் நிலை எக்செல் ஃபைல் தயார்! டவுன்லோட் செய்து தேவையான மாற்றங்களை (Track/Place) செய்யவும்.")
                    st.download_button(
                        label=f"📥 Download {out_filename}",
                        data=excel_data,
                        file_name=out_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )

            except Exception as e:
                status_box.update(label="❌ செயலாக்கத்தில் பிழை ஏற்பட்டது!", state="error")
                st.error(f"Error Details: {str(e)}")

with tab2:
    st.subheader("2. எடிட் செய்த Excel-ஐ பதிவேற்றி Final அறிக்கை பெறுதல்")
    st.caption("Step 1 எக்செல் ஃபைலில் Track Condition / Place திருத்தங்களை முடித்த பின் இங்கே பதிவேற்றவும்.")

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
                    apply_am_scoring_step2(wb)

                    wb.calculation.fullCalcOnLoad = True
                    wb.calculation.forceFullCalc = True
                    wb.calculation.calcMode = "auto"
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