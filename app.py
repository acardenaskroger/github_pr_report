import os
import sys
import json
import logging
import subprocess
from pathlib import Path
import pandas as pd
import streamlit as st

# Optional support for YAML and .env
try:
    import yaml
    YAML_AVAILABLE = True
except ImportError:
    YAML_AVAILABLE = False

try:
    from dotenv import load_dotenv, set_key
    DOTENV_AVAILABLE = True
except ImportError:
    DOTENV_AVAILABLE = False

# Streamlit Page Configuration
st.set_page_config(
    page_title="GitHub PR Advanced Analyzer",
    page_icon="📊",
    layout="wide"
)

# Base directories and file paths
BASE_DIR = Path(__file__).resolve().parent
ANALYZER_SCRIPT = BASE_DIR / "github_pr_advanced_analyzer.py"
ENV_FILE = BASE_DIR / ".env"
PDF_REPORT_NAME = "GitHub_Executive_Report.pdf"
EXCEL_REPORT_NAME = "2026Q1_GitAnalysis_Advanced.xlsx"

CONFIG_YML = BASE_DIR / "source.yml"
CONFIG_YAML = BASE_DIR / "source.yaml"
CONFIG_JSON = BASE_DIR / "source.json"

pdf_path = BASE_DIR / PDF_REPORT_NAME
excel_path = BASE_DIR / EXCEL_REPORT_NAME

# Load environment variables from .env if available
if ENV_FILE.exists():
    if DOTENV_AVAILABLE:
        load_dotenv(dotenv_path=ENV_FILE, override=True)
    else:
        with open(ENV_FILE, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith("#") and "=" in line:
                    k, v = line.split("=", 1)
                    os.environ[k.strip()] = v.strip().strip("'\"")

# Determine active configuration file
if CONFIG_YML.exists():
    active_config_path = CONFIG_YML
elif CONFIG_YAML.exists():
    active_config_path = CONFIG_YAML
else:
    active_config_path = CONFIG_JSON

# Initialize session state for persistence
if "analysis_completed" not in st.session_state:
    st.session_state.analysis_completed = False
if "last_logs" not in st.session_state:
    st.session_state.last_logs = ""
if "last_error" not in st.session_state:
    st.session_state.last_error = ""

st.title("📊 GitHub Pull Request Advanced Analyzer")
st.markdown("Analyze PR lead time, review taxonomy, cross-team rework ratios, and generate Executive PDF & Excel reports.")

# =============================================================================
# Sidebar: Unified Token Configuration & .env Storage
# =============================================================================
st.sidebar.header("⚙️ Configuration")

# Get current token from environment
current_env_token = os.environ.get("GITHUB_TOKEN", "")

github_token = st.sidebar.text_input(
    "GitHub Token (`GITHUB_TOKEN`):",
    value=current_env_token,
    type="password",
    help="Unified Personal Access Token used for all GitHub API requests across squads."
)

col_token_save, col_token_status = st.sidebar.columns([1, 1])
with col_token_save:
    if st.button("💾 Save to .env", use_container_width=True):
        if github_token:
            try:
                if DOTENV_AVAILABLE:
                    set_key(str(ENV_FILE), "GITHUB_TOKEN", github_token)
                else:
                    # Fallback manual key-value write
                    env_lines = []
                    token_found = False
                    if ENV_FILE.exists():
                        with open(ENV_FILE, "r", encoding="utf-8") as f:
                            for line in f:
                                if line.strip().startswith("GITHUB_TOKEN="):
                                    env_lines.append(f"GITHUB_TOKEN={github_token}\n")
                                    token_found = True
                                else:
                                    env_lines.append(line)
                    if not token_found:
                        env_lines.append(f"GITHUB_TOKEN={github_token}\n")

                    with open(ENV_FILE, "w", encoding="utf-8") as f:
                        f.writelines(env_lines)

                os.environ["GITHUB_TOKEN"] = github_token
                st.sidebar.success("Token saved in `.env`")
            except Exception as env_err:
                st.sidebar.error(f"Failed to save: {env_err}")
        else:
            st.sidebar.warning("Please enter a token before saving.")

input_mode = st.sidebar.radio(
    "Select Input Mode:",
    ["Configuration File (source.yml / source.json)", "Existing Excel (Raw_Data)"]
)

excel_file_arg = None
if input_mode == "Existing Excel (Raw_Data)":
    uploaded_excel = st.sidebar.file_uploader("Upload Excel with PR_URL column", type=["xlsx", "xls"])
    if uploaded_excel:
        temp_excel_path = BASE_DIR / uploaded_excel.name
        with open(temp_excel_path, "wb") as f:
            f.write(uploaded_excel.getbuffer())
        excel_file_arg = str(temp_excel_path)

# =============================================================================
# In-App Configuration Editor (source.yml / source.json)
# =============================================================================
if input_mode == "Configuration File (source.yml / source.json)":
    with st.sidebar.expander("📝 Edit source.yml / source.json", expanded=False):
        default_yaml_content = """teams:
  - team: payments
    org: krogertechnology
    begin_date: "2026-05-23"
    end_date: "2026-08-25"
    repos:
      - payments-tokenization-api
"""
        current_content = ""
        if active_config_path.exists():
            with open(active_config_path, "r", encoding="utf-8") as f:
                current_content = f.read()
        else:
            current_content = default_yaml_content

        edited_config = st.text_area(
            f"Editing: `{active_config_path.name}`",
            value=current_content,
            height=250,
            help="Directly edit repositories, date ranges, and organizations."
        )

        if st.button("💾 Save Configuration", use_container_width=True):
            try:
                if active_config_path.suffix in [".yml", ".yaml"]:
                    if YAML_AVAILABLE:
                        yaml.safe_load(edited_config)
                    active_config_path = CONFIG_YML
                else:
                    json.loads(edited_config)
                    active_config_path = CONFIG_JSON

                with open(active_config_path, "w", encoding="utf-8") as f:
                    f.write(edited_config)
                st.success(f"Configuration saved to `{active_config_path.name}`")
            except Exception as parse_err:
                st.error(f"Invalid format: {parse_err}")

# =============================================================================
# Execution Control Panel
# =============================================================================
col1, col2 = st.columns([2, 1])

with col1:
    st.subheader("🚀 Execution Control")
    run_button = st.button("Run PR Advanced Analyzer", type="primary", use_container_width=True)

    if run_button:
        if not ANALYZER_SCRIPT.exists():
            st.error(f"❌ Could not find `{ANALYZER_SCRIPT.name}` in `{BASE_DIR}`.")
        else:
            cmd = [sys.executable, str(ANALYZER_SCRIPT)]

            if input_mode == "Existing Excel (Raw_Data)":
                if excel_file_arg:
                    cmd.extend(["--excel", excel_file_arg])
                else:
                    st.warning("⚠️ No Excel uploaded. Looking for default `2026Q1_GitAnalysis.xlsx`.")
            else:
                if active_config_path.exists():
                    cmd.extend(["--config", str(active_config_path)])
                else:
                    st.warning(f"⚠️ `{active_config_path.name}` not found. Using default paths.")

            env = os.environ.copy()
            if github_token:
                env["GITHUB_TOKEN"] = github_token

            with st.spinner("⏳ Analyzing Pull Requests and generating executive reports..."):
                try:
                    result = subprocess.run(
                        cmd,
                        cwd=str(BASE_DIR),
                        env=env,
                        capture_output=True,
                        text=True,
                        check=False
                    )

                    st.session_state.last_logs = result.stdout
                    st.session_state.last_error = result.stderr

                    if result.returncode == 0:
                        st.session_state.analysis_completed = True
                        st.success("✅ Analysis completed successfully!")
                    else:
                        st.session_state.analysis_completed = False
                        st.error(f"❌ Process exited with error code {result.returncode}.")

                except Exception as e:
                    st.session_state.analysis_completed = False
                    st.session_state.last_error = str(e)
                    st.error(f"❌ Execution failed: {e}")

with col2:
    st.subheader("📥 Download Generated Reports")

    if pdf_path.exists():
        with open(pdf_path, "rb") as f_pdf:
            st.download_button(
                label="📄 Download Executive Report (PDF)",
                data=f_pdf,
                file_name=PDF_REPORT_NAME,
                mime="application/pdf",
                use_container_width=True
            )
    else:
        st.info("ℹ️ PDF report not available yet.")

    if excel_path.exists():
        with open(excel_path, "rb") as f_excel:
            st.download_button(
                label="📊 Download Detailed Metrics (Excel)",
                data=f_excel,
                file_name=EXCEL_REPORT_NAME,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
    else:
        st.info("ℹ️ Excel report not available yet.")

# =============================================================================
# Interactive Dashboard & Analytics Viewer
# =============================================================================
if excel_path.exists():
    st.markdown("---")
    st.header("📈 Interactive Dashboard & Analytics")

    try:
        xls = pd.ExcelFile(excel_path)
        df_details = pd.read_excel(xls, sheet_name="Detailed_PR_Metrics")
        df_users = pd.read_excel(xls, sheet_name="User_Breakdown") if "User_Breakdown" in xls.sheet_names else pd.DataFrame()
        df_teams = pd.read_excel(xls, sheet_name="Team_Rework_Breakdown") if "Team_Rework_Breakdown" in xls.sheet_names else pd.DataFrame()
        df_rework = pd.read_excel(xls, sheet_name="Top_5_Rework") if "Top_5_Rework" in xls.sheet_names else pd.DataFrame()
        df_types = pd.read_excel(xls, sheet_name="PR_Type_Breakdown") if "PR_Type_Breakdown" in xls.sheet_names else pd.DataFrame()

        # High-Level KPI Summary Cards
        kpi1, kpi2, kpi3, kpi4 = st.columns(4)
        with kpi1:
            st.metric("Total PRs Analyzed", len(df_details))
        with kpi2:
            merged_cnt = len(df_details[df_details["STATUS"] == "MERGED"])
            st.metric("Merged PRs", merged_cnt)
        with kpi3:
            valid_leads = pd.to_numeric(df_details["LEAD_TIME_HRS"], errors="coerce").dropna()
            avg_lead = f"{valid_leads.mean():.1f}h" if not valid_leads.empty else "N/A"
            st.metric("Avg Lead Time", avg_lead)
        with kpi4:
            valid_rework = pd.to_numeric(df_details["REWORK_RATIO_PCT"], errors="coerce").dropna()
            avg_rw = f"{valid_rework.mean():.1f}%" if not valid_rework.empty else "0.0%"
            st.metric("Avg Rework Ratio", avg_rw)

        st.markdown("###")

        # Multi-tab Breakdown
        tab1, tab2, tab3, tab4, tab5 = st.tabs([
            "📊 Contributor Breakdown", 
            "🏢 Cross-Team Rework",
            "🏷️ PR Work Types", 
            "⚠️ Top Rework PRs", 
            "📋 Full Detailed Metrics"
        ])

        with tab1:
            if not df_users.empty:
                st.dataframe(df_users, use_container_width=True, hide_index=True)
            else:
                st.info("No contributor data available.")

        with tab2:
            if not df_teams.empty:
                col_tchart, col_ttbl = st.columns([1, 1])
                with col_tchart:
                    st.bar_chart(df_teams.set_index("TEAM")["AVG_REWORK_PCT"])
                with col_ttbl:
                    st.dataframe(df_teams, use_container_width=True, hide_index=True)
            else:
                st.info("No team breakdown data available.")

        with tab3:
            if not df_types.empty:
                col_chart, col_tbl = st.columns([1, 1])
                with col_chart:
                    st.bar_chart(df_types.set_index("PR_TYPE")["COUNT"])
                with col_tbl:
                    st.dataframe(df_types, use_container_width=True, hide_index=True)

        with tab4:
            if not df_rework.empty:
                st.dataframe(df_rework, use_container_width=True, hide_index=True)

        with tab5:
            risk_filter = st.multiselect(
                "Filter by Risk Level:", 
                options=df_details["RISK_LEVEL"].unique(), 
                default=df_details["RISK_LEVEL"].unique()
            )
            filtered_df = df_details[df_details["RISK_LEVEL"].isin(risk_filter)]
            st.dataframe(filtered_df, use_container_width=True, hide_index=True)

    except Exception as e:
        st.warning(f"Could not load in-page metrics view: {e}")

# =============================================================================
# Execution Diagnostics & Logs
# =============================================================================
st.markdown("---")
st.subheader("📋 Execution Output & Diagnostics")

if st.session_state.last_error:
    with st.expander("⚠️ View Errors / Stderr Output", expanded=True):
        st.code(st.session_state.last_error, language="bash")

if st.session_state.last_logs:
    with st.expander("📜 View Full Process Logs", expanded=False):
        st.code(st.session_state.last_logs, language="text")
