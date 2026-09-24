import io
import json
import loggingd
import os
import subprocess
import sys
import zipfile
from datetime import date, datetime
from pathlib import Path

import pandas as pd
import streamlit as st

# -----------------------------------------------------------------------------
# Configuration & Constants
# -----------------------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent
CONFIG_PATH = BASE_DIR / "source.json"
OUTPUT_BASIC_EXCEL = BASE_DIR / "2026Q1_GitAnalysis.xlsx"
OUTPUT_ADV_EXCEL = BASE_DIR / "2026Q1_GitAnalysis_Advanced.xlsx"
OUTPUT_PDF = BASE_DIR / "GitHub_Executive_Report.pdf"
OUTPUT_TECHM_EXCEL = BASE_DIR / "output" / "2026Q1_GitAnalysis_TechM_Comments.xlsx"

st.set_page_config(
    page_title="GitHub PR Analytics Suite",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# -----------------------------------------------------------------------------
# Helper Functions
# -----------------------------------------------------------------------------
def run_python_script(script_name: str, args: list = None) -> tuple[int, str]:
    """Execute a python script in a subprocess and return returncode & logs."""
    cmd = [sys.executable, str(BASE_DIR / script_name)] + (args or [])
    try:
        process = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            cwd=str(BASE_DIR),
        )
        logs = []
        log_container = st.empty()

        while True:
            line = process.stdout.readline()
            if not line and process.poll() is not None:
                break
            if line:
                logs.append(line)
                # Show last 10 lines of progress
                log_container.code("".join(logs[-10:]), language="bash")

        process.wait()
        return process.returncode, "".join(logs)
    except Exception as exc:
        return 1, f"Execution failed: {exc}"


def build_zip_package() -> io.BytesIO:
    """Bundle all available generated report files into a zip buffer."""
    zip_buffer = io.BytesIO()
    files_to_zip = [
        ("2026Q1_GitAnalysis.xlsx", OUTPUT_BASIC_EXCEL),
        ("2026Q1_GitAnalysis_Advanced.xlsx", OUTPUT_ADV_EXCEL),
        ("GitHub_Executive_Report.pdf", OUTPUT_PDF),
        ("2026Q1_GitAnalysis_TechM_Comments.xlsx", OUTPUT_TECHM_EXCEL),
    ]

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for arcname, filepath in files_to_zip:
            if filepath.exists():
                zf.write(filepath, arcname=arcname)

    zip_buffer.seek(0)
    return zip_buffer


# -----------------------------------------------------------------------------
# Sidebar: Setup & Configuration
# -----------------------------------------------------------------------------
with st.sidebar:
    st.header("⚙️ Configuration")

    # GitHub Token
    github_token = st.text_input(
        "GitHub Personal Access Token",
        type="password",
        value=os.environ.get("GITHUB_TOKEN", os.environ.get("PAYMENTS_TOKEN", "")),
        help="Used to authenticate against the GitHub REST API.",
    )
    if github_token:
        os.environ["GITHUB_TOKEN"] = github_token
        os.environ["PAYMENTS_TOKEN"] = github_token

    st.markdown("---")
    st.subheader("📁 `source.json` Configuration")

    # Config File Uploader / Editor
    uploaded_config = st.file_uploader("Upload `source.json`", type=["json"])
    if uploaded_config is not None:
        try:
            cfg_data = json.load(uploaded_config)
            with open(CONFIG_PATH, "w", encoding="utf-8") as f:
                json.dump(cfg_data, f, indent=2)
            st.success("`source.json` updated successfully!")
        except Exception as e:
            st.error(f"Error reading JSON: {e}")

    # Display / Edit current config
    if CONFIG_PATH.exists():
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            current_json = f.read()
        with st.expander("Inspect / Edit JSON directly"):
            edited_json = st.text_area("source.json content", value=current_json, height=220)
            if st.button("Save Changes"):
                try:
                    parsed = json.loads(edited_json)
                    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
                        json.dump(parsed, f, indent=2)
                    st.success("Saved `source.json`")
                except json.JSONDecodeError as err:
                    st.error(f"Invalid JSON format: {err}")
    else:
        st.warning("⚠️ No `source.json` found in working directory.")

# -----------------------------------------------------------------------------
# Main Panel
# -----------------------------------------------------------------------------
st.title("🚀 GitHub Pull Request Analytics Dashboard")
st.markdown(
    "Run automated analytics across your GitHub repositories, calculate merge latency, "
    "classify review comments, quantify rework ratio, and produce executive reports."
)

st.markdown("---")

# Action Execution Tabs
tab_run, tab_view, tab_download = st.tabs(["⚡ Execution Pipeline", "📈 Results Preview", "📥 Downloads"])

with tab_run:
    st.subheader("Execute Analysis Scripts")

    col1, col2 = st.columns([2, 1])

    with col1:
        st.info("You can run the complete pipeline or execute individual scripts.")

        # Run Complete Pipeline
        if st.button("🚀 Run Complete Pipeline (All 3 Scripts)", type="primary", use_container_width=True):
            if not github_token:
                st.warning("⚠️ Please provide a GitHub Token in the sidebar before executing.")
            else:
                progress_bar = st.progress(0, text="Starting pipeline...")

                # Step 1
                progress_bar.progress(15, text="1/3: Running Basic PR Analysis (`github_pr_analysis_improved.py`)...")
                ret1, log1 = run_python_script("github_pr_analysis_improved.py")

                if ret1 != 0:
                    st.error("❌ Step 1 (Basic Analysis) failed. Check logs below.")
                    st.expander("Error Logs", expanded=True).code(log1)
                else:
                    st.success("✅ Step 1: Basic PR Analysis completed.")

                    # Step 2
                    progress_bar.progress(50, text="2/3: Running Advanced PR Risk Analyzer (`github_pr_advanced_analyzer.py`)...")
                    ret2, log2 = run_python_script("github_pr_advanced_analyzer.py", ["--excel", str(OUTPUT_BASIC_EXCEL)])

                    if ret2 != 0:
                        st.error("❌ Step 2 (Advanced Analysis) failed.")
                        st.expander("Error Logs", expanded=True).code(log2)
                    else:
                        st.success("✅ Step 2: Advanced Risk & Executive Report generated.")

                        # Step 3
                        progress_bar.progress(85, text="3/3: Running TechM Comments Analysis (`github_pr_techm_comments_analysis.py`)...")
                        ret3, log3 = run_python_script("github_pr_techm_comments_analysis.py")

                        if ret3 != 0:
                            st.error("❌ Step 3 (TechM Comments Analysis) failed.")
                            st.expander("Error Logs", expanded=True).code(log3)
                        else:
                            progress_bar.progress(100, text="Pipeline Completed Successfully!")
                            st.balloons()
                            st.success("🎉 All scripts finished successfully! Head over to the **Results Preview** or **Downloads** tabs.")

    with col2:
        st.markdown("#### Individual Actions")

        if st.button("1. Run Basic Analysis", use_container_width=True):
            with st.spinner("Running basic PR analysis..."):
                code, out = run_python_script("github_pr_analysis_improved.py")
                if code == 0:
                    st.success("Basic analysis completed.")
                else:
                    st.error(f"Failed with code {code}")
                    st.code(out)

        if st.button("2. Run Advanced Analyzer", use_container_width=True):
            with st.spinner("Running advanced PR analysis..."):
                code, out = run_python_script("github_pr_advanced_analyzer.py")
                if code == 0:
                    st.success("Advanced analysis completed.")
                else:
                    st.error(f"Failed with code {code}")
                    st.code(out)

        if st.button("3. Run TechM Comments", use_container_width=True):
            with st.spinner("Running TechM comment analysis..."):
                code, out = run_python_script("github_pr_techm_comments_analysis.py")
                if code == 0:
                    st.success("TechM analysis completed.")
                else:
                    st.error(f"Failed with code {code}")
                    st.code(out)

# -----------------------------------------------------------------------------
# Results Preview Tab
# -----------------------------------------------------------------------------
with tab_view:
    st.subheader("Data & Analysis Previews")

    preview_option = st.selectbox(
        "Select Report to Inspect",
        [
            "Advanced PR Metrics & Risk (2026Q1_GitAnalysis_Advanced.xlsx)",
            "Basic PR & Repository Overview (2026Q1_GitAnalysis.xlsx)",
            "TechM Comments & Review Summary (2026Q1_GitAnalysis_TechM_Comments.xlsx)",
        ],
    )

    if preview_option.startswith("Advanced") and OUTPUT_ADV_EXCEL.exists():
        excel_file = pd.ExcelFile(OUTPUT_ADV_EXCEL)
        sheet = st.selectbox("Select Sheet", excel_file.sheet_names)
        df = pd.read_excel(OUTPUT_ADV_EXCEL, sheet_name=sheet)

        # Summary KPI cards if detailed metrics
        if sheet == "Detailed_PR_Metrics" and not df.empty:
            kpi1, kpi2, kpi3, kpi4 = st.columns(4)
            kpi1.metric("Total PRs Analyzed", len(df))
            kpi2.metric("Critical Risk PRs", len(df[df.get("RISK_LEVEL") == "CRITICAL"]))
            kpi3.metric("High Risk PRs", len(df[df.get("RISK_LEVEL") == "HIGH"]))
            kpi4.metric("Avg Rework %", f"{df.get('REWORK_RATIO_PCT', pd.Series([0])).mean():.1f}%")

        st.dataframe(df, use_container_width=True)

    elif preview_option.startswith("Basic") and OUTPUT_BASIC_EXCEL.exists():
        excel_file = pd.ExcelFile(OUTPUT_BASIC_EXCEL)
        sheet = st.selectbox("Select Sheet", excel_file.sheet_names)
        df = pd.read_excel(OUTPUT_BASIC_EXCEL, sheet_name=sheet)
        st.dataframe(df, use_container_width=True)

    elif preview_option.startswith("TechM") and OUTPUT_TECHM_EXCEL.exists():
        excel_file = pd.ExcelFile(OUTPUT_TECHM_EXCEL)
        sheet = st.selectbox("Select Sheet", excel_file.sheet_names)
        df = pd.read_excel(OUTPUT_TECHM_EXCEL, sheet_name=sheet)
        st.dataframe(df, use_container_width=True)

    else:
        st.info("ℹ️ No generated output found for this selection yet. Run the pipeline first.")

# -----------------------------------------------------------------------------
# Downloads Tab
# -----------------------------------------------------------------------------
with tab_download:
    st.subheader("Generated Output Files")

    d_col1, d_col2 = st.columns(2)

    with d_col1:
        st.markdown("### 📦 Bulk Download")
        zip_data = build_zip_package()
        st.download_button(
            label="📥 Download All Output Files (.ZIP)",
            data=zip_data,
            file_name=f"GitHub_PR_Analysis_{datetime.now().strftime('%Y%m%d_%H%M')}.zip",
            mime="application/zip",
            type="primary",
            use_container_width=True,
        )

    with d_col2:
        st.markdown("### 📄 Individual File Downloads")

        # 1. Advanced Excel
        if OUTPUT_ADV_EXCEL.exists():
            with open(OUTPUT_ADV_EXCEL, "rb") as f:
                st.download_button(
                    label="📊 Download Advanced Analysis Excel (`2026Q1_GitAnalysis_Advanced.xlsx`)",
                    data=f.read(),
                    file_name=OUTPUT_ADV_EXCEL.name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

        # 2. Executive PDF Report
        if OUTPUT_PDF.exists():
            with open(OUTPUT_PDF, "rb") as f:
                st.download_button(
                    label="📑 Download Executive Report (`GitHub_Executive_Report.pdf`)",
                    data=f.read(),
                    file_name=OUTPUT_PDF.name,
                    mime="application/pdf",
                    use_container_width=True,
                )

        # 3. Basic Excel
        if OUTPUT_BASIC_EXCEL.exists():
            with open(OUTPUT_BASIC_EXCEL, "rb") as f:
                st.download_button(
                    label="📋 Download Basic Analysis Excel (`2026Q1_GitAnalysis.xlsx`)",
                    data=f.read(),
                    file_name=OUTPUT_BASIC_EXCEL.name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

        # 4. TechM Comments Excel
        if OUTPUT_TECHM_EXCEL.exists():
            with open(OUTPUT_TECHM_EXCEL, "rb") as f:
                st.download_button(
                    label="💬 Download TechM Comments Excel (`2026Q1_GitAnalysis_TechM_Comments.xlsx`)",
                    data=f.read(),
                    file_name=OUTPUT_TECHM_EXCEL.name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )