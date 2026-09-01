import os
import re
import tempfile
import streamlit as st
import cloudscraper
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

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
# RACING AUSTRALIA SCRAPER (CLOUDFLARE BYPASS ENGINE)
# =====================================================================

def fetch_racing_data_bypassed(url):
    target_url = url.strip()
    if "racingaustralia.horse" in target_url and "www.racingaustralia.horse" not in target_url:
        target_url = target_url.replace("racingaustralia.horse", "www.racingaustralia.horse")

    # Cloudflare மற்றும் Bot Protection-ஐ கடந்து செல்லும் Scraper Session
    scraper = cloudscraper.create_scraper(
        browser={
            'browser': 'chrome',
            'platform': 'windows',
            'desktop': True
        }
    )

    headers = {
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8",
        "Accept-Language": "en-US,en;q=0.9",
        "Referer": "https://www.racingaustralia.horse/",
        "Sec-Fetch-Dest": "document",
        "Sec-Fetch-Mode": "navigate",
        "Sec-Fetch-Site": "same-origin",
        "Sec-Fetch-User": "?1",
        "Upgrade-Insecure-Requests": "1"
    }

    response = scraper.get(target_url, headers=headers, timeout=40)
    response.encoding = 'utf-8'
    return response.text, response.status_code, target_url

def parse_racing_html(html_text):
    soup = BeautifulSoup(html_text, "html.parser")
    races = []

    # Strategy 1: Race Tables Parsing
    all_tables = soup.find_all("table")
    current_race_idx = 1

    for tbl in all_tables:
        rows = tbl.find_all("tr")
        if len(rows) < 2:
            continue

        horses = []
        for r in rows:
            cells = r.find_all(["td", "th"])
            if len(cells) < 4:
                continue

            c_texts = [re.sub(r'\s+', ' ', c.get_text(strip=True)) for c in cells]
            first_col = c_texts[0].strip()

            # Horse Row கண்டறிதல் (எண்ணில் தொடங்குவது)
            if re.match(r"^\d{1,2}$", first_col):
                h_no = first_col
                h_name = c_texts[1] if len(c_texts) > 1 else ""
                
                if h_name.lower() in ["horse", "horse name", "name", "runner"]:
                    continue

                h_name_clean = re.sub(r"\s*\([A-Z0-9a-z\s]+\)$", "", h_name).strip()
                jockey = c_texts[2] if len(c_texts) > 2 else ""
                trainer = c_texts[3] if len(c_texts) > 3 else ""
                weight = c_texts[4] if len(c_texts) > 4 else ""

                horses.append({
                    "horse_no": h_no,
                    "horse_name": h_name_clean,
                    "jockey": jockey,
                    "trainer": trainer,
                    "weight": weight
                })

        if len(horses) >= 2:
            races.append({
                "race_name": f"Race {current_race_idx}",
                "horses": horses
            })
            current_race_idx += 1

    # Strategy 2: Links Fallback
    if not races:
        horse_links = soup.find_all("a", href=re.compile(r"Horse\.aspx", re.I))
        if horse_links:
            horses = []
            for idx, a in enumerate(horse_links, start=1):
                h_name = a.get_text(strip=True)
                if h_name:
                    horses.append({
                        "horse_no": str(idx),
                        "horse_name": h_name,
                        "jockey": "",
                        "trainer": "",
                        "weight": ""
                    })
            if horses:
                races.append({
                    "race_name": "Race 1",
                    "horses": horses
                })

    return races

def build_excel_workbook(races):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "AM_PRO_RACING"

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
        cell.alignment = Alignment(horizontal="center", vertical="center")

    row_idx = 2
    for r in races:
        race_title = r["race_name"]
        for h in r["horses"]:
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
        score = 0
        if pos_val in ["1", "2", "3"]:
            score += 50
            ws.cell(row=r, column=12).fill = high_fill
            ws.cell(row=r, column=12).font = font_bold
            ws.cell(row=r, column=20).value = "QUALIFIED"
        ws.cell(row=r, column=19).value = score

# =====================================================================
# UI TABS
# =====================================================================

tab1, tab2 = st.tabs(["🌐 STEP 1: URL Scraper (Col T)", "📊 STEP 2: Process Edited Excel (AM Score)"])

with tab1:
    st.subheader("1. Web URL-லிருந்து Excel உருவாக்குதல்")
    st.caption("Racing Australia All Form URL-ஐ உள்ளிட்டால் Column T கணக்கீட்டுடன் கூடிய ஆரம்ப எக்செல் கிடைக்கும்.")

    input_url = st.text_input(
        "🔗 Racing Australia All Form URL:",
        placeholder="https://www.racingaustralia.horse/FreeFields/AllForm.aspx?Key=...",
        key="tab1_all_form_url"
    )

    if st.button("🚀 Scrape & Download Step 1 Excel", type="primary", key="btn_step1"):
        cleaned_url = input_url.strip()
        if not cleaned_url:
            st.warning("⚠️ தயவுசெய்து சரியான Racing Australia URL-ஐ உள்ளிடவும்.")
        else:
            status_box = st.status("🔄 பந்தய தகவல்கள் பெறப்படுகின்றன...", expanded=True)
            try:
                status_box.write("🌐 1. Cloudflare பைபாஸ் மூலம் இணையதள இணைப்பு பெறப்படுகிறது...")
                html_text, status_code, final_url = fetch_racing_data_bypassed(cleaned_url)

                status_box.write("📋 2. குதிரை மற்றும் பந்தய விவரங்கள் பிரித்தெடுக்கப்படுகின்றன...")
                races = parse_racing_html(html_text)

                if not races:
                    status_box.update(label="❌ ரேஸ் விவரங்கள் கண்டறியப்படவில்லை!", state="error")
                    st.error(f"இணையதளத்திலிருந்து தரவுகள் கிடைக்கவில்லை (HTTP Status: {status_code}).")
                else:
                    total_horses = sum(len(r["horses"]) for r in races)
                    status_box.write(f"✅ {len(races)} பந்தயங்கள் | {total_horses} குதிரைகள் கண்டறியப்பட்டன.")
                    status_box.write("📊 3. Column T (Calculated Time) எக்செல் உருவாக்கப்படுகிறது...")

                    match = re.search(r"Key=([^&#]+)", cleaned_url)
                    file_key = match.group(1).replace("%2C", "_") if match else "RACING_DATA"
                    out_filename = f"{file_key}_STEP1_RAW.xlsx"

                    wb = build_excel_workbook(races)

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
                    st.success(f"✅ {len(races)} பந்தயங்களுக்கான ஆரம்ப எக்செல் ஃபைல் தயார்! டவுன்லோட் செய்து திருத்தங்களை மேற்கொள்ளவும்.")
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