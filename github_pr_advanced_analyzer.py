# =============================================================================
# GitHub PR Advanced Analyzer - Dual Input Version (REST API & Excel Input)
# =============================================================================
"""
GitHub PR Advanced Analyzer

Features:
1. Dual Input Modes:
   - Direct REST API mode (from source.yml / source.json)
   - Excel input mode (loads existing 2026Q1_GitAnalysis.xlsx)
2. Unified Token Resolution (single GITHUB_TOKEN environment variable / .env)
3. Human-Centric Risk Calibration & Small-PR Guardrails
4. Comment Taxonomy & Keyword Risk Analysis (Defects, Arch, Refactor, Lint, Test)
5. Reviewer tracking (REVIEWERS)
6. PR Type Classification (PR_TYPE: Feature, Bugfix, Refactor, Chore, Docs, etc.)
7. PR Type Breakdown Pie Chart in Executive PDF
8. Cross-Team Rework Ratio Comparison Bar Chart in PDF & Excel
9. Metrics Glossary & Methodology final page in PDF and Excel
"""

import argparse
import json
import logging
import os
import re
import statistics
import time
import warnings
from datetime import datetime, time as datetime_time, timezone
from pathlib import Path
from typing import Any, Optional, Tuple, Dict, List

# Suppress LibreSSL/urllib3 compatibility warning on macOS
warnings.filterwarnings("ignore", category=UserWarning, module="urllib3")

import pandas as pd
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# Optional .env support
try:
    from dotenv import load_dotenv
    DOTENV_AVAILABLE = True
except ImportError:
    DOTENV_AVAILABLE = False

# Optional YAML support
try:
    import yaml
    YAML_AVAILABLE = True
except ImportError:
    YAML_AVAILABLE = False

# ReportLab imports for PDF generation and visual charts
try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.units import inch
    from reportlab.platypus import (
        HRFlowable,
        KeepTogether,
        PageBreak,
        Paragraph,
        SimpleDocTemplate,
        Spacer,
        Table,
        TableStyle,
    )
    from reportlab.graphics.shapes import Drawing
    from reportlab.graphics.charts.piecharts import Pie
    from reportlab.graphics.charts.barcharts import VerticalBarChart
    from reportlab.graphics.charts.legends import Legend
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False

# ---------------------------------------------------------------------------
# Logging Setup
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Constants & Defaults
# ---------------------------------------------------------------------------
SCRIPT_DIR = Path(__file__).resolve().parent
DEFAULT_CONFIG_PATH = SCRIPT_DIR / "source.json"
DEFAULT_INPUT_EXCEL = SCRIPT_DIR / "2026Q1_GitAnalysis.xlsx"
OUTPUT_EXCEL_PATH = SCRIPT_DIR / "2026Q1_GitAnalysis_Advanced.xlsx"
OUTPUT_PDF_PATH = SCRIPT_DIR / "GitHub_Executive_Report.pdf"

# Load .env if present
ENV_PATH = SCRIPT_DIR / ".env"
if ENV_PATH.exists() and DOTENV_AVAILABLE:
    load_dotenv(dotenv_path=ENV_PATH, override=True)

GITHUB_API_BASE = "https://api.github.com"
GITHUB_API_VERSION = "2022-11-28"
MAIN_BRANCHES = {"main", "master", "dev", "develop", "trunk"}
DATE_FMT = "%Y-%m-%d"
GH_DATE_FMT = "%Y-%m-%dT%H:%M:%SZ"
DEFAULT_BEGIN_DATE = "2025-02-01"
DEFAULT_END_DATE = "2026-08-25"
MAX_RETRIES = 3
BACKOFF_FACTOR = 1.5
REQUEST_TIMEOUT = 30

KNOWN_BOTS = {
    "dependabot[bot]",
    "dependabot",
    "renovate[bot]",
    "renovate",
    "github-actions[bot]",
    "codecov-commenter",
    "atlassian-bitbucket-cloud[bot]",
}

TAXONOMY_PATTERNS = {
    "DEFECT_CRITICAL": re.compile(
        r"\b(bug|error|fail|failure|broken|defect|vulnerability|security|crash|exception|nullpointer|leak|outage|flaw|regression)\b",
        re.IGNORECASE,
    ),
    "ARCH_DESIGN": re.compile(
        r"\b(architecture|design|pattern|performance|scalable|database|query|async|concurrency|deadlock|race condition|latency)\b",
        re.IGNORECASE,
    ),
    "REFACTOR": re.compile(
        r"\b(refactor|clean|structure|naming|simplify|redundant|duplication|decouple|extract)\b",
        re.IGNORECASE,
    ),
    "STYLE_LINT": re.compile(
        r"\b(style|lint|format|indent|typo|convention|nit|prettier|spelling|whitespace)\b",
        re.IGNORECASE,
    ),
    "TEST_DOC": re.compile(
        r"\b(test|coverage|doc|readme|assert|spec|mock|unit test)\b",
        re.IGNORECASE,
    ),
}

CHART_PALETTE = [
    colors.HexColor('#2B6CB0'),  # Blue (Feature)
    colors.HexColor('#C53030'),  # Red (Bugfix)
    colors.HexColor('#319795'),  # Teal (Refactor)
    colors.HexColor('#D69E2E'),  # Yellow/Orange (Chore)
    colors.HexColor('#805AD5'),  # Purple (Dependency)
    colors.HexColor('#38A169'),  # Green (Testing)
    colors.HexColor('#DD6B20'),  # Orange (Docs)
    colors.HexColor('#4A5568'),  # Gray (Other)
]


# ---------------------------------------------------------------------------
# PR Type Classifier
# ---------------------------------------------------------------------------
def classify_pr_type(title: str, branch: str, labels: List[str], is_bot: bool) -> str:
    """Classify PR into standardized engineering categories using title, branch, and labels."""
    if is_bot:
        return "DEPENDENCY_UPDATE"

    labels_str = " ".join(labels).lower()
    combined_text = f"{title} {branch} {labels_str}".lower()

    if re.search(r"(^|\b)(feat|feature|story)([\/\(\:\s_-]|$)", combined_text):
        return "FEATURE"
    if re.search(r"(^|\b)(fix|bug|bugfix|hotfix|patch)([\/\(\:\s_-]|$)", combined_text):
        return "BUGFIX"
    if re.search(r"(^|\b)(refactor|cleanup|clean-up)([\/\(\:\s_-]|$)", combined_text):
        return "REFACTOR"
    if re.search(r"(^|\b)(perf|performance|optimize|optimization)([\/\(\:\s_-]|$)", combined_text):
        return "PERFORMANCE"
    if re.search(r"(^|\b)(test|tests|testing|coverage|cypress|jest|pytest)([\/\(\:\s_-]|$)", combined_text):
        return "TESTING"
    if re.search(r"(^|\b)(doc|docs|documentation|readme)([\/\(\:\s_-]|$)", combined_text):
        return "DOCUMENTATION"
    if re.search(r"(^|\b)(revert)([\/\(\:\s_-]|$)", combined_text):
        return "REVERT"
    if re.search(r"(^|\b)(chore|build|ci|cd|pipeline|deps|dependency|bump|release|config)([\/\(\:\s_-]|$)", combined_text):
        return "CHORE / MAINTENANCE"

    return "OTHER / GENERAL"


def _build_session(token: str) -> requests.Session:
    """Build pre-configured requests session with retry strategies for GitHub API."""
    session = requests.Session()
    if token:
        session.headers.update({"Authorization": f"Bearer {token}"})
    session.headers.update(
        {
            "X-GitHub-Api-Version": GITHUB_API_VERSION,
            "Accept": "application/vnd.github+json",
        }
    )
    retry_strategy = Retry(
        total=MAX_RETRIES,
        backoff_factor=BACKOFF_FACTOR,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["GET"],
        raise_on_status=False,
    )
    adapter = HTTPAdapter(max_retries=retry_strategy)
    session.mount("https://", adapter)
    return session


def github_get(session: requests.Session, url: str, params: Optional[dict] = None, label: str = "") -> requests.Response:
    """Execute timed GET request with rate limit handling and automatic wait-and-retry."""
    start_time = time.perf_counter()
    logger.debug("GET START %s params=%s", label or url, params)

    try:
        response = session.get(url, params=params, timeout=REQUEST_TIMEOUT)
    except requests.RequestException as exc:
        elapsed = time.perf_counter() - start_time
        logger.error("GET FAILED %s after %.2fs: %s", label or url, elapsed, exc)
        raise

    elapsed = time.perf_counter() - start_time
    if response.status_code == 403 and "rate limit" in response.text.lower():
        reset_ts = int(response.headers.get("X-RateLimit-Reset", time.time() + 60))
        wait_secs = max(reset_ts - int(time.time()), 1)
        logger.warning("Rate limit hit during %s. Waiting %ds...", label, wait_secs)
        time.sleep(wait_secs)
        return github_get(session, url, params=params, label=label)

    response.raise_for_status()
    return response


def parse_pr_url(pr_url: str) -> Tuple[str, str, int]:
    """Extract org, repo_name, and pr_number from a GitHub PR URL."""
    match = re.search(r"github\.com/([^/]+)/([^/]+)/pull/(\d+)", pr_url)
    if not match:
        raise ValueError(f"Invalid GitHub PR URL format: {pr_url}")
    return match.group(1), match.group(2), int(match.group(3))


# ---------------------------------------------------------------------------
# Deep PR Detail Extractor
# ---------------------------------------------------------------------------
def enrich_pr_details(session: requests.Session, org: str, repo_name: str, pr_number: int, base_info: dict) -> dict:
    """Fetch reviews, taxonomy comments, commits, additions/deletions, reviewers, and compute risk."""
    full_repo = f"{org}/{repo_name}"
    pr_label = f"{full_repo}#{pr_number}"

    pr_data = github_get(session, f"{GITHUB_API_BASE}/repos/{full_repo}/pulls/{pr_number}", label=f"PR {pr_label}").json()
    reviews = github_get(session, f"{GITHUB_API_BASE}/repos/{full_repo}/pulls/{pr_number}/reviews", label=f"Reviews {pr_label}").json()
    review_comments = github_get(session, f"{GITHUB_API_BASE}/repos/{full_repo}/pulls/{pr_number}/comments", label=f"Comments {pr_label}").json()
    issue_comments = github_get(session, f"{GITHUB_API_BASE}/repos/{full_repo}/issues/{pr_number}/comments", label=f"Issue comments {pr_label}").json()
    commits = github_get(session, f"{GITHUB_API_BASE}/repos/{full_repo}/pulls/{pr_number}/commits", label=f"Commits {pr_label}").json()

    pr_title = pr_data.get("title", "")
    head_branch = pr_data.get("head", {}).get("ref", "")
    labels = [lb.get("name", "") for lb in pr_data.get("labels", []) if isinstance(lb, dict)]

    author = pr_data.get("user", {}).get("login", base_info.get("CREATED_USER_NAME", "unknown"))
    author_type = pr_data.get("user", {}).get("type", "User")
    is_bot = (author.lower() in KNOWN_BOTS) or (author_type == "Bot") or ("[bot]" in author.lower())

    pr_type = classify_pr_type(pr_title, head_branch, labels, is_bot)

    created_at = datetime.strptime(pr_data["created_at"], GH_DATE_FMT)
    merged_at_str = pr_data.get("merged_at")
    closed_at_str = pr_data.get("closed_at")

    merged_at = datetime.strptime(merged_at_str, GH_DATE_FMT) if merged_at_str else None
    closed_at = datetime.strptime(closed_at_str, GH_DATE_FMT) if closed_at_str else None
    end_time = merged_at or closed_at or datetime.now(timezone.utc)

    lead_time_hrs = round((end_time - created_at).total_seconds() / 3600, 2)

    human_reviews = [r for r in reviews if r.get("user", {}).get("type") != "Bot" and r.get("user", {}).get("login") != author]
    human_inline_comments = [c for c in review_comments if c.get("user", {}).get("type") != "Bot" and c.get("user", {}).get("login") != author]
    human_issue_comments = [c for c in issue_comments if c.get("user", {}).get("type") != "Bot" and c.get("user", {}).get("login") != author]

    all_review_timestamps = []
    for r in human_reviews:
        if r.get("submitted_at"):
            all_review_timestamps.append(datetime.strptime(r["submitted_at"], GH_DATE_FMT))
    for c in human_inline_comments:
        if c.get("created_at"):
            all_review_timestamps.append(datetime.strptime(c["created_at"], GH_DATE_FMT))

    first_review_at = min(all_review_timestamps) if all_review_timestamps else None
    ttfr_hrs = round((first_review_at - created_at).total_seconds() / 3600, 2) if first_review_at else None

    approval_timestamps = [
        datetime.strptime(r["submitted_at"], GH_DATE_FMT)
        for r in human_reviews
        if r.get("state") == "APPROVED" and r.get("submitted_at")
    ]
    first_approval_at = min(approval_timestamps) if approval_timestamps else None

    review_cycle_hrs = round((first_approval_at - first_review_at).total_seconds() / 3600, 2) if (first_approval_at and first_review_at and first_approval_at >= first_review_at) else None
    merge_delay_hrs = round((end_time - first_approval_at).total_seconds() / 3600, 2) if (first_approval_at and end_time >= first_approval_at) else None

    additions = pr_data.get("additions", 0)
    deletions = pr_data.get("deletions", 0)
    changed_files = pr_data.get("changed_files", 0)
    total_lines = additions + deletions

    unique_reviewers_set = {
        r.get("user", {}).get("login") for r in human_reviews if r.get("user", {}).get("login")
    }.union({
        c.get("user", {}).get("login") for c in human_inline_comments if c.get("user", {}).get("login")
    })

    unique_reviewers_list = sorted([u for u in unique_reviewers_set if u and u != author])
    reviewers_str = ", ".join(unique_reviewers_list) if unique_reviewers_list else "None"
    reviewers_count = len(unique_reviewers_list)
    review_rounds = len([r for r in human_reviews if r.get("state") in ("APPROVED", "CHANGES_REQUESTED")])

    total_commits = len(commits)
    commits_after_first_review = 0
    if first_review_at and commits:
        for commit_item in commits:
            commit_date_str = commit_item.get("commit", {}).get("committer", {}).get("date") or commit_item.get("commit", {}).get("author", {}).get("date")
            if commit_date_str:
                c_date = datetime.strptime(commit_date_str, GH_DATE_FMT)
                if c_date > first_review_at:
                    commits_after_first_review += 1

    rework_ratio_pct = round((commits_after_first_review / max(total_commits, 1)) * 100, 1)

    all_human_comment_bodies = [c.get("body", "") for c in human_inline_comments + human_issue_comments]
    taxonomy_counts = {key: 0 for key in TAXONOMY_PATTERNS}
    matched_keywords = {key: set() for key in TAXONOMY_PATTERNS}

    for body in all_human_comment_bodies:
        for cat, pattern in TAXONOMY_PATTERNS.items():
            matches = pattern.findall(body)
            if matches:
                taxonomy_counts[cat] += len(matches)
                matched_keywords[cat].update([m.lower() for m in matches])

    defect_comments_count = taxonomy_counts["DEFECT_CRITICAL"]
    arch_comments_count = taxonomy_counts["ARCH_DESIGN"]
    defect_density = round((defect_comments_count / max(total_lines, 1)) * 100, 2)

    if commits_after_first_review > 0:
        if defect_comments_count > 0 or arch_comments_count > 0:
            rework_category = "CRITICAL_REWORK"
        else:
            rework_category = "NORMAL_REWORK"
    else:
        rework_category = "NO_REWORK"

    severity_score = round(
        (total_lines * 0.1)
        + (changed_files * 2.0)
        + (total_commits * 1.5)
        + (defect_comments_count * 15.0),
        1
    )

    if is_bot:
        risk_level = "LOW"
        risk_explanation = "LOW RISK: AUTOMATED UPDATE: Bot dependency bump (isolated from human metrics)."
    elif total_lines <= 50 and changed_files <= 3 and defect_comments_count == 0:
        risk_level = "LOW"
        risk_explanation = f"LOW RISK: Small, clean changeset ({total_lines} lines in {changed_files} files) merged in {lead_time_hrs}h with 0 defect comments."
    elif defect_comments_count >= 3 or total_lines > 1000 or changed_files > 25:
        risk_level = "CRITICAL"
        kw_str = ", ".join(list(matched_keywords["DEFECT_CRITICAL"])[:3])
        risk_explanation = f"CRITICAL RISK: High change scope ({total_lines} lines / {changed_files} files) or multiple defect comments ({defect_comments_count} defects: '{kw_str}') with {rework_ratio_pct}% rework."
    elif defect_comments_count > 0 or total_lines > 400 or changed_files > 10 or rework_ratio_pct > 50:
        risk_level = "HIGH"
        kw_str = ", ".join(list(matched_keywords["DEFECT_CRITICAL"])[:3]) if matched_keywords["DEFECT_CRITICAL"] else "none"
        risk_explanation = f"HIGH RISK: Elevated defect flags ({defect_comments_count} defects: '{kw_str}') or post-review iteration ({commits_after_first_review} post-review commits)."
    elif total_lines > 150 or changed_files > 5 or rework_ratio_pct > 25:
        risk_level = "MEDIUM"
        risk_explanation = f"MEDIUM RISK: Moderate changeset size ({total_lines} lines in {changed_files} files) with {rework_ratio_pct}% rework."
    else:
        risk_level = "LOW"
        risk_explanation = f"LOW RISK: Standard PR ({total_lines} lines in {changed_files} files) merged smoothly with no defect flags."

    return {
        "TEAM": base_info.get("TEAM", "Unknown"),
        "REPO_NAME": repo_name,
        "PR_NUMBER": pr_number,
        "PR_TITLE": pr_title,
        "PR_TYPE": pr_type,
        "PR_ID": pr_data.get("id"),
        "PR_URL": pr_data.get("html_url"),
        "IS_BOT": is_bot,
        "AUTHOR": author,
        "REVIEWERS": reviewers_str,
        "STATUS": "MERGED" if merged_at else "CLOSED_UNMERGED",
        "PR_CREATED_AT": pr_data["created_at"],
        "PR_MERGED_AT": merged_at_str or "",
        "PR_CLOSED_AT": closed_at_str or "",
        "LEAD_TIME_HRS": lead_time_hrs,
        "TTFR_HRS": ttfr_hrs if ttfr_hrs is not None else "N/A",
        "REVIEW_CYCLE_HRS": review_cycle_hrs if review_cycle_hrs is not None else "N/A",
        "MERGE_DELAY_HRS": merge_delay_hrs if merge_delay_hrs is not None else "N/A",
        "ADDITIONS": additions,
        "DELETIONS": deletions,
        "TOTAL_LINES": total_lines,
        "CHANGED_FILES": changed_files,
        "TOTAL_COMMITS": total_commits,
        "COMMITS_AFTER_REVIEW": commits_after_first_review,
        "REWORK_RATIO_PCT": rework_ratio_pct,
        "REVIEWERS_COUNT": reviewers_count,
        "REVIEW_ROUNDS": review_rounds,
        "DEFECT_COMMENTS": defect_comments_count,
        "ARCH_COMMENTS": arch_comments_count,
        "REFACTOR_COMMENTS": taxonomy_counts["REFACTOR"],
        "STYLE_COMMENTS": taxonomy_counts["STYLE_LINT"],
        "TEST_COMMENTS": taxonomy_counts["TEST_DOC"],
        "DEFECT_DENSITY_PER_100L": defect_density,
        "REWORK_CATEGORY": rework_category,
        "SEVERITY_SCORE": severity_score,
        "RISK_LEVEL": risk_level,
        "RISK_EXPLANATION": risk_explanation,
    }


# ---------------------------------------------------------------------------
# Data Loaders (Excel & Direct API)
# ---------------------------------------------------------------------------
def load_prs_from_excel(excel_path: Path, token: str) -> List[dict]:
    """Read PRs directly from the Raw_Data sheet of 2026Q1_GitAnalysis.xlsx."""
    if not excel_path.exists():
        raise FileNotFoundError(f"Input Excel file not found: {excel_path}")

    logger.info("Loading input PR list from Excel: %s", excel_path)
    df_raw = pd.read_excel(excel_path, sheet_name="Raw_Data")

    session = _build_session(token)
    enriched_prs = []
    total_prs = len(df_raw)

    for index, row in df_raw.iterrows():
        pr_url = str(row["PR_URL"])
        team_name = str(row.get("TEAM", "payments"))

        try:
            org, repo_name, pr_number = parse_pr_url(pr_url)
            logger.info("Processing Excel PR %d/%d: %s/%s#%d", index + 1, total_prs, org, repo_name, pr_number)

            base_info = {
                "TEAM": team_name,
                "CREATED_USER_NAME": str(row.get("CREATED_USER_NAME", "")),
            }

            enriched = enrich_pr_details(session, org, repo_name, pr_number, base_info)
            enriched_prs.append(enriched)
        except Exception as exc:
            logger.error("Error enriching PR from Excel row %d (%s): %s", index + 1, pr_url, exc)

    return enriched_prs


def load_prs_from_api(config_path: Path) -> List[dict]:
    """Load configuration from source.yml or source.json and query GitHub API."""
    if not config_path.exists():
        for alt_ext in [".yml", ".yaml", ".json"]:
            alt_path = config_path.with_suffix(alt_ext)
            if alt_path.exists():
                config_path = alt_path
                break
        else:
            raise FileNotFoundError(f"Configuration file not found: {config_path}")

    with config_path.open("r", encoding="utf-8") as f:
        if config_path.suffix in [".yml", ".yaml"]:
            if not YAML_AVAILABLE:
                raise ImportError("PyYAML is required to read .yml files. Run: pip install pyyaml")
            source_data = yaml.safe_load(f)
        else:
            source_data = json.load(f)

    global_env_token = os.environ.get("GITHUB_TOKEN", "")

    team_configs = []
    if isinstance(source_data, dict) and "teams" in source_data:
        for t in source_data["teams"]:
            team_name = t["team"]
            org = t.get("org", "krogertechnology")
            begin_date = t.get("begin_date", DEFAULT_BEGIN_DATE)
            end_date = t.get("end_date", DEFAULT_END_DATE)
            resolved_token = t.get("token") or global_env_token
            repos = t.get("repos", [])

            for r in repos:
                repo_name = r if isinstance(r, str) else r.get("repo")
                team_configs.append({
                    "TEAM": team_name,
                    "ORG": org,
                    "REPO_NAME": repo_name,
                    "BEGIN_DATE": begin_date,
                    "END_DATE": end_date,
                    "TOKEN": resolved_token,
                })

    enriched_prs = []

    for cfg in team_configs:
        org = cfg["ORG"]
        repo_name = cfg["REPO_NAME"]
        token = cfg["TOKEN"]
        team_name = cfg["TEAM"]
        full_repo = f"{org}/{repo_name}"

        begin_dt = datetime.combine(datetime.strptime(cfg["BEGIN_DATE"], DATE_FMT).date(), datetime_time.min)
        end_dt = datetime.combine(datetime.strptime(cfg["END_DATE"], DATE_FMT).date(), datetime_time.max)

        logger.info("Querying GitHub API for %s (Range: %s to %s)...", full_repo, cfg["BEGIN_DATE"], cfg["END_DATE"])

        if not token:
            logger.warning("No GITHUB_TOKEN provided for %s. Requests might hit rate limits or fail for private repos.", full_repo)

        session = _build_session(token)
        page = 1
        matching_prs = []

        while True:
            url = f"{GITHUB_API_BASE}/repos/{full_repo}/pulls?state=closed&per_page=100&page={page}"
            res = github_get(session, url, label=f"Closed PRs page {page}")
            prs_page = res.json()

            if not prs_page:
                break

            for pr_item in prs_page:
                merged_at_str = pr_item.get("merged_at")
                base_ref = pr_item.get("base", {}).get("ref")

                if not merged_at_str or base_ref not in MAIN_BRANCHES:
                    continue

                m_date = datetime.strptime(merged_at_str, GH_DATE_FMT)
                if begin_dt <= m_date <= end_dt:
                    matching_prs.append(pr_item)

            if len(prs_page) < 100:
                break
            page += 1

        logger.info("Found %d merged PRs in date range for %s", len(matching_prs), full_repo)

        for idx, pr_item in enumerate(matching_prs, start=1):
            pr_num = pr_item["number"]
            logger.info("Enriching PR %d/%d: %s#%d", idx, len(matching_prs), full_repo, pr_num)
            base_info = {"TEAM": team_name}
            enriched = enrich_pr_details(session, org, repo_name, pr_num, base_info)
            enriched_prs.append(enriched)

    return enriched_prs


# ---------------------------------------------------------------------------
# Aggregators
# ---------------------------------------------------------------------------
def compute_team_rework_breakdown(enriched_prs: List[dict]) -> pd.DataFrame:
    """Compute cross-team rework and cycle metrics."""
    df = pd.DataFrame(enriched_prs)
    if df.empty:
        return pd.DataFrame()

    team_rows = []
    for team_name, group in df.groupby("TEAM"):
        total_prs = len(group)
        human_group = group[~group["IS_BOT"]]
        human_prs = len(human_group)

        valid_lead = [p["LEAD_TIME_HRS"] for p in group.to_dict("records") if isinstance(p["LEAD_TIME_HRS"], (int, float))]
        avg_lead = round(statistics.mean(valid_lead), 1) if valid_lead else 0.0

        avg_rework = round(human_group["REWORK_RATIO_PCT"].mean(), 1) if not human_group.empty else 0.0
        crit_rework = len(human_group[human_group["REWORK_CATEGORY"] == "CRITICAL_REWORK"])
        total_defects = group["DEFECT_COMMENTS"].sum()

        team_rows.append({
            "TEAM": team_name,
            "TOTAL_PRS": total_prs,
            "HUMAN_PRS": human_prs,
            "AVG_REWORK_PCT": avg_rework,
            "CRITICAL_REWORK_PRS": crit_rework,
            "AVG_LEAD_HRS": avg_lead,
            "TOTAL_DEFECTS": total_defects,
        })

    return pd.DataFrame(team_rows).sort_values(by="AVG_REWORK_PCT", ascending=False)


def compute_user_breakdown(enriched_prs: List[dict]) -> pd.DataFrame:
    """Compute individual contributor performance metrics."""
    human_prs = [p for p in enriched_prs if not p["IS_BOT"]]
    if not human_prs:
        return pd.DataFrame()

    df = pd.DataFrame(human_prs)
    user_rows = []

    for author, group in df.groupby("AUTHOR"):
        total_prs = len(group)
        merged_prs = len(group[group["STATUS"] == "MERGED"])
        unmerged_prs = total_prs - merged_prs

        valid_lead = [p["LEAD_TIME_HRS"] for p in group.to_dict("records") if isinstance(p["LEAD_TIME_HRS"], (int, float))]
        valid_ttfr = [p["TTFR_HRS"] for p in group.to_dict("records") if isinstance(p["TTFR_HRS"], (int, float))]

        avg_lead = round(statistics.mean(valid_lead), 1) if valid_lead else "N/A"
        avg_ttfr = round(statistics.mean(valid_ttfr), 1) if valid_ttfr else "N/A"
        avg_rework_pct = round(group["REWORK_RATIO_PCT"].mean(), 1)

        defects = group["DEFECT_COMMENTS"].sum()
        critical_rework_prs = len(group[group["REWORK_CATEGORY"] == "CRITICAL_REWORK"])

        c_cnt = len(group[group["RISK_LEVEL"] == "CRITICAL"])
        h_cnt = len(group[group["RISK_LEVEL"] == "HIGH"])
        m_cnt = len(group[group["RISK_LEVEL"] == "MEDIUM"])
        l_cnt = len(group[group["RISK_LEVEL"] == "LOW"])
        risk_profile = f"{c_cnt}C / {h_cnt}H / {m_cnt}M / {l_cnt}L"

        user_rows.append({
            "AUTHOR": author,
            "TOTAL_PRS": total_prs,
            "MERGED": merged_prs,
            "UNMERGED": unmerged_prs,
            "AVG_LEAD_TIME_HRS": avg_lead,
            "AVG_TTFR_HRS": avg_ttfr,
            "AVG_REWORK_PCT": avg_rework_pct,
            "TOTAL_DEFECT_COMMENTS": defects,
            "CRITICAL_REWORK_PRS": critical_rework_prs,
            "RISK_PROFILE": risk_profile,
        })

    return pd.DataFrame(user_rows).sort_values(by="TOTAL_PRS", ascending=False)


def compute_top_5_rework(enriched_prs: List[dict]) -> pd.DataFrame:
    """Identify top 5 PRs with highest post-review rework iterations."""
    df = pd.DataFrame(enriched_prs)
    if df.empty:
        return pd.DataFrame()

    df_rework = df.sort_values(by=["COMMITS_AFTER_REVIEW", "REWORK_RATIO_PCT", "DEFECT_COMMENTS"], ascending=False).head(5)
    cols = [
        "PR_NUMBER", "PR_TYPE", "REPO_NAME", "AUTHOR", "REVIEWERS", "LEAD_TIME_HRS",
        "TOTAL_COMMITS", "COMMITS_AFTER_REVIEW", "REWORK_RATIO_PCT",
        "DEFECT_COMMENTS", "REWORK_CATEGORY", "RISK_LEVEL"
    ]
    available_cols = [c for c in cols if c in df_rework.columns]
    return df_rework[available_cols]


def compute_pr_type_breakdown(enriched_prs: List[dict]) -> pd.DataFrame:
    """Compute aggregated metrics grouped by PR type."""
    df = pd.DataFrame(enriched_prs)
    if df.empty:
        return pd.DataFrame()

    type_rows = []
    for pr_type, group in df.groupby("PR_TYPE"):
        total_count = len(group)
        pct = round((total_count / len(df)) * 100, 1)
        total_lines = group["TOTAL_LINES"].sum()

        valid_lead = [p["LEAD_TIME_HRS"] for p in group.to_dict("records") if isinstance(p["LEAD_TIME_HRS"], (int, float))]
        avg_lead = round(statistics.mean(valid_lead), 1) if valid_lead else 0.0
        total_defects = group["DEFECT_COMMENTS"].sum()

        type_rows.append({
            "PR_TYPE": pr_type,
            "COUNT": total_count,
            "SHARE_PCT": pct,
            "TOTAL_LINES": total_lines,
            "AVG_LEAD_HRS": avg_lead,
            "DEFECTS": total_defects,
        })

    return pd.DataFrame(type_rows).sort_values(by="COUNT", ascending=False)


# ---------------------------------------------------------------------------
# Visual Chart Generators & PDF Exporter (ReportLab)
# ---------------------------------------------------------------------------
def create_pr_type_pie_chart(df_types: pd.DataFrame) -> Drawing:
    """Create styled ReportLab Pie Chart for PR type distribution."""
    d = Drawing(480, 140)
    if df_types.empty:
        return d

    pc = Pie()
    pc.x = 20
    pc.y = 5
    pc.width = 120
    pc.height = 120
    pc.data = df_types["COUNT"].tolist()
    pc.labels = [f"{cnt}" for cnt in df_types["COUNT"].tolist()]
    pc.simpleLabels = 0

    for i, color in enumerate(CHART_PALETTE[:len(df_types)]):
        pc.slices[i].fillColor = color
        pc.slices[i].strokeColor = colors.white
        pc.slices[i].strokeWidth = 1

    legend = Legend()
    legend.x = 180
    legend.y = 125
    legend.dx = 10
    legend.dy = 10
    legend.fontName = "Helvetica"
    legend.fontSize = 8
    legend.boxAnchor = "nw"
    legend.columnMaximum = 8
    legend.strokeWidth = 0.5
    legend.strokeColor = colors.transparent

    legend_items = []
    for i, (_, row) in enumerate(df_types.iterrows()):
        legend_items.append((
            CHART_PALETTE[i % len(CHART_PALETTE)],
            f"{row['PR_TYPE']} — {row['COUNT']} PRs ({row['SHARE_PCT']}%)"
        ))
    legend.colorNamePairs = legend_items

    d.add(pc)
    d.add(legend)
    return d


def create_team_rework_bar_chart(df_teams: pd.DataFrame) -> Drawing:
    """Create styled Vertical Bar Chart comparing average rework % by team."""
    d = Drawing(480, 140)
    if df_teams.empty:
        return d

    bc = VerticalBarChart()
    bc.x = 40
    bc.y = 25
    bc.height = 95
    bc.width = 400

    rework_values = [float(v) for v in df_teams["AVG_REWORK_PCT"].tolist()]
    bc.data = [rework_values]

    bc.categoryAxis.categoryNames = [str(name)[:16] for name in df_teams["TEAM"].tolist()]
    bc.categoryAxis.labels.fontSize = 8
    bc.categoryAxis.labels.dy = -10
    bc.categoryAxis.labels.fontName = "Helvetica-Bold"

    max_val = max(rework_values) if rework_values else 100
    bc.valueAxis.valueMin = 0
    bc.valueAxis.valueMax = max(100, int(max_val * 1.25))
    bc.valueAxis.valueStep = 25
    bc.valueAxis.labels.fontSize = 7.5
    bc.valueAxis.labels.fontName = "Helvetica"

    bc.bars[0].fillColor = colors.HexColor('#C53030')
    bc.bars[0].strokeColor = colors.white
    bc.bars[0].strokeWidth = 0.5

    d.add(bc)
    return d


def generate_pdf_report(enriched_prs: List[dict], pdf_path: Path):
    """Generate multi-page Executive PDF report with visual charts and methodology glossary."""
    if not REPORTLAB_AVAILABLE:
        logger.warning("ReportLab is not installed. Skipping PDF generation.")
        return

    doc = SimpleDocTemplate(
        str(pdf_path),
        pagesize=letter,
        leftMargin=0.5 * inch,
        rightMargin=0.5 * inch,
        topMargin=0.5 * inch,
        bottomMargin=0.5 * inch,
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle("DocTitle", parent=styles["Heading1"], fontName="Helvetica-Bold", fontSize=18, leading=22, textColor=colors.HexColor("#1A365D"))
    h2_style = ParagraphStyle("SectionH2", parent=styles["Heading2"], fontName="Helvetica-Bold", fontSize=12, leading=15, textColor=colors.HexColor("#2B6CB0"), spaceBefore=8, spaceAfter=4)
    body_style = ParagraphStyle("BodyTextCustom", parent=styles["Normal"], fontName="Helvetica", fontSize=8, leading=10.5, textColor=colors.HexColor("#2D3748"))
    table_cell = ParagraphStyle("TableCell", parent=styles["Normal"], fontName="Helvetica", fontSize=7.5, leading=9.5)
    table_header = ParagraphStyle("TableHeader", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=7.5, leading=9.5, textColor=colors.white)

    story = []
    story.append(Paragraph("GitHub Engineering Executive Analysis Dashboard", title_style))
    story.append(Paragraph(f"<b>Generated:</b> {datetime.now().strftime('%Y-%m-%d %H:%M')} | <b>Token:</b> Unified GITHUB_TOKEN", body_style))
    story.append(Spacer(1, 8))

    df = pd.DataFrame(enriched_prs)
    total_prs = len(df)
    merged_prs = len(df[df["STATUS"] == "MERGED"])
    human_df = df[~df["IS_BOT"]]

    valid_leads = [p["LEAD_TIME_HRS"] for p in enriched_prs if isinstance(p["LEAD_TIME_HRS"], (int, float))]
    valid_ttfrs = [p["TTFR_HRS"] for p in human_df.to_dict("records") if isinstance(p["TTFR_HRS"], (int, float))]

    avg_lead = round(statistics.mean(valid_leads), 1) if valid_leads else 0.0
    avg_ttfr = round(statistics.mean(valid_ttfrs), 1) if valid_ttfrs else 0.0
    avg_rework = round(human_df["REWORK_RATIO_PCT"].mean(), 1) if not human_df.empty else 0.0
    crit_rework_cnt = len(human_df[human_df["REWORK_CATEGORY"] == "CRITICAL_REWORK"])

    kpi_data = [
        [
            Paragraph(f"<b>Total PRs Analyzed:</b> {total_prs}", table_cell),
            Paragraph(f"<b>Merged PRs:</b> {merged_prs}", table_cell),
            Paragraph(f"<b>Avg Lead Time:</b> {avg_lead}h", table_cell),
        ],
        [
            Paragraph(f"<b>Avg Time to 1st Review:</b> {avg_ttfr}h", table_cell),
            Paragraph(f"<b>Avg Rework Ratio:</b> {avg_rework}%", table_cell),
            Paragraph(f"<b>Critical Rework PRs:</b> {crit_rework_cnt}", table_cell),
        ]
    ]
    t_kpi = Table(kpi_data, colWidths=[2.4 * inch, 2.4 * inch, 2.4 * inch])
    t_kpi.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#EDF2F7')),
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#CBD5E0')),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E0')),
        ('PADDING', (0,0), (-1,-1), 4),
    ]))
    story.append(t_kpi)
    story.append(Spacer(1, 8))

    # Section 1: Cross-Team Rework Benchmark
    story.append(Paragraph("1. Cross-Team Rework Ratio Benchmark (%)", h2_style))
    df_teams = compute_team_rework_breakdown(enriched_prs)
    if not df_teams.empty:
        story.append(create_team_rework_bar_chart(df_teams))
        story.append(Spacer(1, 4))
        tm_headers = ["Team", "Total PRs", "Human PRs", "Avg Rework (%)", "Critical Rework", "Avg Lead (h)", "Defects"]
        tm_rows = [[Paragraph(h, table_header) for h in tm_headers]]
        for _, tm in df_teams.iterrows():
            tm_rows.append([
                Paragraph(f"<b>{tm['TEAM']}</b>", table_cell),
                Paragraph(str(tm["TOTAL_PRS"]), table_cell),
                Paragraph(str(tm["HUMAN_PRS"]), table_cell),
                Paragraph(f"<b>{tm['AVG_REWORK_PCT']}%</b>", table_cell),
                Paragraph(str(tm["CRITICAL_REWORK_PRS"]), table_cell),
                Paragraph(f"{tm['AVG_LEAD_HRS']}h", table_cell),
                Paragraph(str(tm["TOTAL_DEFECTS"]), table_cell),
            ])
        t_teams = Table(tm_rows, colWidths=[1.8*inch, 0.8*inch, 0.8*inch, 1.1*inch, 1.1*inch, 0.9*inch, 0.7*inch])
        t_teams.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#2B6CB0')),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E2E8F0')),
            ('PADDING', (0,0), (-1,-1), 3),
        ]))
        story.append(t_teams)

    story.append(Spacer(1, 8))

    # Section 2: PR Type Distribution
    story.append(Paragraph("2. Pull Request Work Type Breakdown & Composition", h2_style))
    df_types = compute_pr_type_breakdown(enriched_prs)
    if not df_types.empty:
        story.append(create_pr_type_pie_chart(df_types))
        story.append(Spacer(1, 4))
        pt_headers = ["PR Type", "Count", "Share (%)", "Total Lines Changed", "Avg Lead Time", "Defects"]
        pt_rows = [[Paragraph(h, table_header) for h in pt_headers]]
        for _, tr in df_types.iterrows():
            pt_rows.append([
                Paragraph(f"<b>{tr['PR_TYPE']}</b>", table_cell),
                Paragraph(str(tr["COUNT"]), table_cell),
                Paragraph(f"{tr['SHARE_PCT']}%", table_cell),
                Paragraph(f"{tr['TOTAL_LINES']:,} lines", table_cell),
                Paragraph(f"{tr['AVG_LEAD_HRS']}h", table_cell),
                Paragraph(str(tr["DEFECTS"]), table_cell),
            ])
        t_types = Table(pt_rows, colWidths=[1.8*inch, 0.7*inch, 0.9*inch, 1.4*inch, 1.2*inch, 0.8*inch])
        t_types.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#2B6CB0')),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E2E8F0')),
            ('PADDING', (0,0), (-1,-1), 3),
        ]))
        story.append(t_types)

    story.append(Spacer(1, 8))

    # Section 3: Contributor Performance
    story.append(Paragraph("3. Human Contributor Performance Breakdown", h2_style))
    df_user = compute_user_breakdown(enriched_prs)
    if not df_user.empty:
        u_headers = ["Author", "PRs", "Merged", "Avg Lead", "Avg TTFR", "Rework %", "Defects", "Crit Rework", "Risk Profile"]
        u_rows = [[Paragraph(h, table_header) for h in u_headers]]
        for _, u in df_user.iterrows():
            u_rows.append([
                Paragraph(str(u["AUTHOR"]), table_cell),
                Paragraph(str(u["TOTAL_PRS"]), table_cell),
                Paragraph(str(u["MERGED"]), table_cell),
                Paragraph(f"{u['AVG_LEAD_TIME_HRS']}h", table_cell),
                Paragraph(f"{u['AVG_TTFR_HRS']}h", table_cell),
                Paragraph(f"{u['AVG_REWORK_PCT']}%", table_cell),
                Paragraph(str(u["TOTAL_DEFECT_COMMENTS"]), table_cell),
                Paragraph(str(u["CRITICAL_REWORK_PRS"]), table_cell),
                Paragraph(str(u["RISK_PROFILE"]), table_cell),
            ])
        t_user = Table(u_rows, colWidths=[1.1*inch, 0.5*inch, 0.6*inch, 0.7*inch, 0.7*inch, 0.7*inch, 0.6*inch, 0.8*inch, 1.5*inch])
        t_user.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#2B6CB0')),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E2E8F0')),
            ('PADDING', (0,0), (-1,-1), 3),
        ]))
        story.append(t_user)

    story.append(Spacer(1, 8))

    # Section 4: Top 5 High Rework PRs
    story.append(Paragraph("4. Top 5 High-Rework Pull Requests", h2_style))
    df_top_rework = compute_top_5_rework(enriched_prs)
    if not df_top_rework.empty:
        r_headers = ["PR #", "Type", "Repo", "Author", "Reviewers", "Lead (h)", "Total Commits", "Post Commits", "Rework %", "Risk"]
        r_rows = [[Paragraph(h, table_header) for h in r_headers]]
        for _, r in df_top_rework.iterrows():
            r_rows.append([
                Paragraph(f"#{r['PR_NUMBER']}", table_cell),
                Paragraph(str(r.get('PR_TYPE', 'OTHER')), table_cell),
                Paragraph(str(r['REPO_NAME']), table_cell),
                Paragraph(str(r['AUTHOR']), table_cell),
                Paragraph(str(r.get('REVIEWERS', 'None')), table_cell),
                Paragraph(f"{r['LEAD_TIME_HRS']}h", table_cell),
                Paragraph(str(r['TOTAL_COMMITS']), table_cell),
                Paragraph(str(r['COMMITS_AFTER_REVIEW']), table_cell),
                Paragraph(f"{r['REWORK_RATIO_PCT']}%", table_cell),
                Paragraph(str(r['RISK_LEVEL']), table_cell),
            ])
        t_top = Table(r_rows, colWidths=[0.5*inch, 0.9*inch, 1.1*inch, 0.8*inch, 0.9*inch, 0.5*inch, 0.5*inch, 0.5*inch, 0.6*inch, 0.6*inch])
        t_top.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#C53030')),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E2E8F0')),
            ('PADDING', (0,0), (-1,-1), 3),
        ]))
        story.append(t_top)

    story.append(Spacer(1, 8))

    # Section 5: Risk Insights & Explanations
    story.append(Paragraph("5. Detailed Executive Risk Insights & Explanations", h2_style))
    d_headers = ["PR #", "Type", "Author", "Lines/Files", "Defects", "Risk", "Plain-Language Risk Explanation"]
    d_rows = [[Paragraph(h, table_header) for h in d_headers]]
    for p in enriched_prs[:12]:
        d_rows.append([
            Paragraph(f"#{p['PR_NUMBER']}", table_cell),
            Paragraph(str(p['PR_TYPE']), table_cell),
            Paragraph(str(p['AUTHOR']), table_cell),
            Paragraph(f"{p['TOTAL_LINES']}L / {p['CHANGED_FILES']}F", table_cell),
            Paragraph(str(p['DEFECT_COMMENTS']), table_cell),
            Paragraph(str(p['RISK_LEVEL']), table_cell),
            Paragraph(str(p['RISK_EXPLANATION']), table_cell),
        ])
    t_detail = Table(d_rows, colWidths=[0.5*inch, 0.8*inch, 0.8*inch, 0.8*inch, 0.5*inch, 0.6*inch, 3.2*inch])
    t_detail.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#2D3748')),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E2E8F0')),
        ('PADDING', (0,0), (-1,-1), 3),
    ]))
    story.append(t_detail)

    # PAGE BREAK -> METRICS GLOSSARY PAGE
    story.append(PageBreak())
    story.append(Paragraph("Metrics Calculation Methodology & Glossary", title_style))
    story.append(Paragraph("Comprehensive definitions, mathematical formulas, guardrails, and risk threshold rules.", body_style))
    story.append(Spacer(1, 8))

    g_headers = ["Metric Name", "Mathematical Formula / Rule", "Business Meaning & Purpose"]
    g_rows = [[Paragraph(h, table_header) for h in g_headers]]

    glossary_items = [
        ("PR Lead Time (hrs)", "t_merged/closed - t_created", "Total cycle time elapsed from PR creation until final merge or closure."),
        ("Time to 1st Review (TTFR)", "t_first_review - t_created", "Measures peer reviewer responsiveness and latency in starting code reviews."),
        ("Cross-Team Rework Ratio", "Mean Rework % per Engineering Squad", "Benchmarks review efficiency and iteration overhead across distinct squads."),
        ("PR Type Classification", "Regex matching on title, branch & labels", "Categorizes PRs into FEATURE, BUGFIX, REFACTOR, DEPENDENCY_UPDATE, CHORE, etc."),
        ("Small-PR Guardrail Rule", "Lines <= 50 & Files <= 3 & Defects == 0 -> LOW RISK", "Prevents small, clean PRs from being falsely flagged as HIGH risk due to high file multipliers."),
        ("Critical Rework vs Normal Rework", "Post-review commits + (Defect or Arch Comments > 0) -> CRITICAL_REWORK", "Isolates bug/defect-driven rework iterations from routine formatting, linting, and style nits."),
        ("Human-Only Comment Taxonomy", "Regex matching on human reviewer comments only (Bots excluded)", "Classifies feedback into DEFECT_CRITICAL, ARCH_DESIGN, REFACTOR, STYLE_LINT, and TEST_DOC."),
        ("Review Defect Density", "(Defect Comments / Total Lines Changed) * 100", "Normalized defect frequency per 100 lines of changed code."),
        ("Severity Score", "(Lines*0.1) + (Files*2.0) + (Commits*1.5) + (Defects*15.0)", "Unified complexity and risk score weighting code volume, structural changes, and review defects."),
    ]

    for m_name, m_formula, m_meaning in glossary_items:
        g_rows.append([
            Paragraph(f"<b>{m_name}</b>", table_cell),
            Paragraph(f"<code>{m_formula}</code>", table_cell),
            Paragraph(m_meaning, table_cell),
        ])

    t_glossary = Table(g_rows, colWidths=[1.8*inch, 2.4*inch, 3.0*inch])
    t_glossary.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1A365D')),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E0')),
        ('PADDING', (0,0), (-1,-1), 4),
    ]))
    story.append(t_glossary)

    doc.build(story)
    logger.info("Generated PDF report: %s", pdf_path)


# ---------------------------------------------------------------------------
# Excel Exporter
# ---------------------------------------------------------------------------
def generate_excel_report(enriched_prs: List[dict], excel_path: Path):
    """Write multi-tab Excel workbook including PR_TYPE, Team Breakdown, and Glossary."""
    df_detail = pd.DataFrame(enriched_prs)
    df_teams = compute_team_rework_breakdown(enriched_prs)
    df_user = compute_user_breakdown(enriched_prs)
    df_rework = compute_top_5_rework(enriched_prs)
    df_types = compute_pr_type_breakdown(enriched_prs)

    glossary_data = [
        {"Metric Name": "PR Lead Time (hrs)", "Formula / Rule": "t_merged/closed - t_created", "Business Purpose": "Total cycle time from PR creation to merge."},
        {"Metric Name": "Cross-Team Rework Ratio", "Formula / Rule": "Mean Rework % across squad", "Business Purpose": "Benchmarks team review iteration efficiency."},
        {"Metric Name": "PR Type Classification", "Formula / Rule": "Regex on Title/Branch/Labels", "Business Purpose": "Categorizes work into Feature, Bugfix, Refactor, Chore, etc."},
        {"Metric Name": "Time to 1st Review (TTFR)", "Formula / Rule": "t_first_review - t_created", "Business Purpose": "Reviewer responsiveness and latency."},
        {"Metric Name": "Small-PR Guardrail", "Formula / Rule": "Lines <= 50 & Files <= 3 & Defects == 0 -> LOW", "Business Purpose": "Prevents small, clean PRs from false high risk flags."},
        {"Metric Name": "Critical Rework", "Formula / Rule": "Post-review commits + Defect/Arch comments", "Business Purpose": "Isolates bug-driven rework from style nits."},
        {"Metric Name": "Rework Ratio (%)", "Formula / Rule": "(Post-Review Commits / Total Commits) * 100", "Business Purpose": "Percentage of rework post-review."},
        {"Metric Name": "Severity Score", "Formula / Rule": "(Lines*0.1) + (Files*2.0) + (Commits*1.5) + (Defects*15)", "Business Purpose": "Unified complexity score."},
    ]
    df_glossary = pd.DataFrame(glossary_data)

    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        df_detail.to_excel(writer, sheet_name="Detailed_PR_Metrics", index=False)
        if not df_teams.empty:
            df_teams.to_excel(writer, sheet_name="Team_Rework_Breakdown", index=False)
        if not df_types.empty:
            df_types.to_excel(writer, sheet_name="PR_Type_Breakdown", index=False)
        if not df_user.empty:
            df_user.to_excel(writer, sheet_name="User_Breakdown", index=False)
        if not df_rework.empty:
            df_rework.to_excel(writer, sheet_name="Top_5_Rework", index=False)
        df_glossary.to_excel(writer, sheet_name="Metrics_Glossary", index=False)

    logger.info("Generated Excel report: %s", excel_path)


# ---------------------------------------------------------------------------
# Main Execution Pipeline
# ---------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(description="GitHub PR Advanced Analyzer (Dual Input Version)")
    parser.add_argument("--excel", "-e", type=Path, help="Path to input Excel file (e.g. 2026Q1_GitAnalysis.xlsx)")
    parser.add_argument("--config", "-c", type=Path, default=DEFAULT_CONFIG_PATH, help="Path to JSON/YML config file (default: source.json)")
    args = parser.parse_args()

    pipeline_start = time.perf_counter()
    logger.info("Starting GitHub PR Advanced Analyzer...")

    token = os.environ.get("GITHUB_TOKEN", "")
    enriched_prs = []

    if args.excel or DEFAULT_INPUT_EXCEL.exists():
        input_excel = args.excel or DEFAULT_INPUT_EXCEL

        if not token and DEFAULT_CONFIG_PATH.exists():
            try:
                with open(DEFAULT_CONFIG_PATH, "r") as f:
                    cfg = json.load(f)
                    if "teams" in cfg and cfg["teams"]:
                        token = cfg["teams"][0].get("token", "")
            except Exception:
                pass

        logger.info("INPUT MODE: EXCEL INPUT (%s)", input_excel)
        if not token:
            logger.warning("No GITHUB_TOKEN found. API requests may fail if repos are private or rate-limited.")

        enriched_prs = load_prs_from_excel(input_excel, token)
    else:
        logger.info("INPUT MODE: DIRECT REST API (from %s)", args.config)
        enriched_prs = load_prs_from_api(args.config)

    if not enriched_prs:
        logger.error("No PRs were analyzed. Exiting.")
        return

    logger.info("Successfully analyzed %d PRs.", len(enriched_prs))

    # Generate Reports
    generate_excel_report(enriched_prs, OUTPUT_EXCEL_PATH)
    generate_pdf_report(enriched_prs, OUTPUT_PDF_PATH)

    total_time = round(time.perf_counter() - pipeline_start, 2)
    logger.info("Analysis Pipeline finished in %.2fs (%.2f mins)", total_time, total_time / 60)
    logger.info("Outputs created:")
    logger.info(" - Excel: %s", OUTPUT_EXCEL_PATH)
    logger.info(" - PDF:   %s", OUTPUT_PDF_PATH)


if __name__ == "__main__":
    main()
