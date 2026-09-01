import os
import sys
import tempfile
import re
import traceback
import openpyxl
import streamlit as st

# 1. அனைத்து கோப்பகப் பாதைகளையும் சேர்க்கவும்
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
for path_item in [BASE_DIR, os.path.join(BASE_DIR, "scraper"), os.path.join(BASE_DIR, "engine"), os.path.join(BASE_DIR, "database")]:
    if os.path.exists(path_item) and path_item not in sys.path:
        sys.path.insert(0, path_item)

# 2. அனைத்து சப்-ஃபோல்டர்களையும் ஆழமாக இணைத்தல்
for root, dirs, files in os.walk(BASE_DIR):
    if root not in sys.path and ".git" not in root and "__pycache__" not in root:
        sys.path.insert(0, root)

# 3. தொகுதிகளை இறக்குமதி செய்தல்
import_errors = {}

try:
    from scraper.racing_australia_scraper import RacingAustraliaScraper
except Exception as e1:
    try:
        from racing_australia_scraper import RacingAustraliaScraper
    except Exception as e2:
        RacingAustraliaScraper = None
        import_errors["RacingAustraliaScraper"] = f"Folder import: {e1} | Root import: {e2}"

try:
    from engine.horse_history_collector import HorseHistoryCollector
except Exception as e1:
    try:
        from horse_history_collector import HorseHistoryCollector
    except Exception as e2:
        HorseHistoryCollector = None
        import_errors["HorseHistoryCollector"] = f"Folder import: {e1} | Root import: {e2}"

try:
    from engine.excel_exporter import ExcelExporter
except Exception as e1:
    try:
        from excel_exporter import ExcelExporter
    except Exception as e2:
        ExcelExporter = None
        import_errors["ExcelExporter"] = f"Folder import: {e1} | Root import: {e2}"

try:
    from engine.am_score_engine import AMScoreEngine
except Exception as e1:
    try:
        from am_score_engine import AMScoreEngine
    except Exception as e2:
        AMScoreEngine = None
        import_errors["AMScoreEngine"] = f"Folder import: {e1} | Root import: {e2}"

st.set_page_config(
    page_title="AM PRO Racing Mobile",
    page_icon="🏇",
    layout="centered"
)

st.markdown("""
    <style>
    .main-title {
        text-align: center;
        font-weight: 800;
        color: #1E3A8A;
        font-size: 24px;
        margin-bottom: 2px;
    }
    .sub-title {
        text-align: center;
        color: #6B7280;
        font-size: 13px;
        margin-bottom: 20px;
    }
    .stButton>button {
        width: 100%;
        border-radius: 8px;
        height: 3.2em;
        background-color: #2563EB;
        color: white;
        font-weight: bold;
        font-size: 15px;
    }
    </style>
""", unsafe_allow_html=True)

st.markdown("<div class='main-title'>🏇 AM PRO Racing System</div>", unsafe_allow_html=True)
st.markdown("<div class='sub-title'>Step 1: URL to Excel (Col T) | Step 2: Excel to Final Scoring & Highlights</div>", unsafe_allow_html=True)

tab1, tab2 = st.tabs(["🌐 STEP 1: URL Scraper (Col T)", "📊 STEP 2: Process Edited Excel (AM Score)"])

# =====================================================================
# TAB 1: STEP 1 (Single All Form URL Input -> Initial Excel with Col T)
# =====================================================================
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
        elif RacingAustraliaScraper is None or HorseHistoryCollector is None or ExcelExporter is None:
            st.error("❌ கோப்பகத்தில் உள்ள பைல் பெயர்கள்:")
            try:
                st.code("Files in current workspace:\n" + "\n".join(os.listdir(BASE_DIR)))
            except Exception:
                pass
            st.error("சில மாட்யூல்களை லோட் செய்ய முடியவில்லை:")
            for mod, err in import_errors.items():
                st.code(f"{mod}: {err}")
        else:
            if "AllForm.aspx" in cleaned_url:
                all_form_url = cleaned_url
                meeting_url = cleaned_url.replace("AllForm.aspx", "Form.aspx")
            elif "Form.aspx" in cleaned_url:
                meeting_url = cleaned_url
                all_form_url = cleaned_url.replace("Form.aspx", "AllForm.aspx")
            else:
                all_form_url = cleaned_url
                meeting_url = cleaned_url

            status_box = st.status("🔄 பந்தய தரவுகள் சேகரிக்கப்படுகின்றன...", expanded=True)
            try:
                status_box.write("🌐 1. பந்தய அட்டவணை ஸ்கிராப் செய்யப்படுகிறது...")
                scraper = RacingAustraliaScraper()
                races = scraper.collect_meeting(meeting_url)

                if not races:
                    status_box.update(label="❌ ரேஸ் விவரங்கள் கிடைக்கவில்லை! URL-ஐ சரிபார்க்கவும்.", state="error")
                    st.error("தரவுகள் கிடைக்கவில்லை. URL சரியானதா என உறுதிப்படுத்தவும்.")
                else:
                    status_box.write(f"✅ {len(races)} பந்தயங்கள் கண்டறியப்பட்டன.")

                    status_box.write("📋 2. முந்தைய பந்தய வரலாறு (All Form) சேகரிக்கப்படுகிறது...")
                    collector = HorseHistoryCollector()
                    races_with_history = collector.collect_meeting_history(races, all_form_url)

                    status_box.write("📊 3. Calculated Time (Col T) கணக்கிடப்பட்டு எக்செல் உருவாக்கப்படுகிறது...")
                    exporter = ExcelExporter()

                    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                        temp_excel_path = tmp.name

                    exporter.export_meeting(
                        template_path="",
                        output_path=temp_excel_path,
                        race_collection=races_with_history
                    )

                    status_box.update(label="🎉 Step 1 Excel தயாராகிவிட்டது!", state="complete", expanded=False)

                    with open(temp_excel_path, "rb") as f:
                        excel_data = f.read()

                    try:
                        os.remove(temp_excel_path)
                    except Exception:
                        pass

                    match = re.search(r"Key=([^&#]+)", input_url)
                    file_key = match.group(1).replace("%2C", "_") if match else "RACING_DATA"
                    output_filename = f"{file_key}_STEP1_RAW.xlsx"

                    st.success("✅ முதல் நிலை எக்செல் ஃபைல் தயார்! டவுன்லோட் செய்து தேவையான மாற்றங்களை (Track/Place) செய்யவும்.")
                    st.download_button(
                        label=f"📥 Download {output_filename}",
                        data=excel_data,
                        file_name=output_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )

            except Exception as e:
                status_box.update(label="❌ செயலாக்கத்தில் பிழை ஏற்பட்டது!", state="error")
                st.error(f"Error Details: {str(e)}")
                st.code(traceback.format_exc())

# =====================================================================
# TAB 2: STEP 2 (Upload Modified Excel -> AM Score & Re-Highlighting)
# =====================================================================
with tab2:
    st.subheader("2. எடிட் செய்த Excel-ஐ பதிவேற்றி Final அறிக்கை பெறுதல்")
    st.caption("Step 1-ல் பெற்ற எக்செல் ஃபைலில் Track Condition / Place திருத்தங்களை முடித்த பின் இங்கே பதிவேற்றவும்.")

    uploaded_file = st.file_uploader(
        "📂 திருத்தப்பட்ட Excel (.xlsx) ஃபைலை பதிவேற்றவும்:",
        type=["xlsx"],
        key="tab2_file_uploader"
    )

    if uploaded_file is not None:
        if st.button("⚡ Process Final Scoring & Highlights", type="primary", key="btn_step2"):
            if AMScoreEngine is None:
                st.error("❌ AMScoreEngine மாட்யூலை லோட் செய்ய முடியவில்லை:")
                st.code(import_errors.get("AMScoreEngine", "Unknown error"))
            else:
                with st.spinner("🔄 Track, Class, Jockey அடிப்படையில் ஹைலைட் மற்றும் AM Score கணக்கிடப்படுகிறது..."):
                    try:
                        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                            tmp.write(uploaded_file.read())
                            temp_in_path = tmp.name

                        wb = openpyxl.load_workbook(temp_in_path, data_only=False)
                        AMScoreEngine.apply_am_score_and_formatting(wb)

                        wb.calculation.fullCalcOnLoad = True
                        wb.calculation.forceFullCalc = True
                        wb.calculation.calcMode = "auto"
                        wb.save(temp_in_path)

                        with open(temp_in_path, "rb") as f:
                            final_data = f.read()

                        try:
                            os.remove(temp_in_path)
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
                        st.code(traceback.format_exc())