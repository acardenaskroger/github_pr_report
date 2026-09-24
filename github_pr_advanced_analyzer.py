# =============================================================================
# GitHub PR Advanced Analyzer - Dual Input Version (REST API & Excel Input)
# =============================================================================
"""
GitHub PR Advanced Analyzer

Features:
1. Dual Input Modes:
   - Direct REST API mode (from source.json)
   - Excel input mode (loads existing 2026Q1_GitAnalysis.xlsx output from basic script)
2. Human-Centric Risk Calibration & Small-PR Guardrails
3. Comment Taxonomy & Keyword Risk Analysis (Defects, Arch, Refactor, Lint, Test)
4. Bot Account Isolation (Dependabot, Renovate, etc.)
5. Metrics Glossary & Methodology final page in PDF and Excel
"""

import argparse
import json
import logging
import os
import re
import statistics
import time
from datetime import datetime, time as datetime_time, timezone
from pathlib import Path
from typing import Any, Optional, Tuple, Dict, List

import pandas as pd
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# ReportLab imports for PDF generation
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
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False

# ---------------------------------------------------------------------------
# Logging setup
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
OUTPUT_HTML_PATH = SCRIPT_DIR / "GitHub_Executive_Report.html"

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

# ---------------------------------------------------------------------------
# Keyword Taxonomy Patterns (Word-Boundary Strict)
# ---------------------------------------------------------------------------
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


# ---------------------------------------------------------------------------
# Helper Functions
# ---------------------------------------------------------------------------
def _build_session(token: str) -> requests.Session:
    """Build a pre-configured requests session for GitHub API."""
    session = requests.Session()
    session.headers.update(
        {
            "Authorization": f"Bearer {token}",
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


def github_get(
    session: requests.Session,
    url: str,
    params: Optional[dict] = None,
    label: str = "",
) -> requests.Response:
    """Execute timed GET request with rate limit inspection and error handling."""
    start_time = time.perf_counter()
    logger.debug("GET START %s params=%s", label or url, params)

    try:
        response = session.get(url, params=params, timeout=REQUEST_TIMEOUT)
    except requests.RequestException as exc:
        elapsed = time.perf_counter() - start_time
        logger.error("GET FAILED %s after %.2fs: %s", label or url, elapsed, exc)
        raise

    elapsed = time.perf_counter() - start_time
    rem = response.headers.get("X-RateLimit-Remaining", "?")
    limit = response.headers.get("X-RateLimit-Limit", "?")

    logger.debug(
        "GET END %s status=%d elapsed=%.2fs rate_limit=%s/%s",
        label or url,
        response.status_code,
        elapsed,
        rem,
        limit,
    )

    if response.status_code == 403 and "rate limit" in response.text.lower():
        reset_ts = int(response.headers.get("X-RateLimit-Reset", time.time() + 60))
        wait_secs = max(reset_ts - int(time.time()), 1)
        logger.warning("Rate limit hit during %s. Waiting %ds...", label, wait_secs)
        time.sleep(wait_secs)
        return github_get(session, url, params=params, label=label)

    response.raise_for_status()
    return response


def parse_pr_url(pr_url: str) -> Tuple[str, str, int]:
    """
    Extract org, repo_name, and pr_number from a GitHub PR URL.
    E.g., https://github.com/krogertechnology/payments-tokenization-api/pull/272
    -> ('krogertechnology', 'payments-tokenization-api', 272)
    """
    match = re.search(r"github\.com/([^/]+)/([^/]+)/pull/(\d+)", pr_url)
    if not match:
        raise ValueError(f"Invalid GitHub PR URL format: {pr_url}")
    return match.group(1), match.group(2), int(match.group(3))


# ---------------------------------------------------------------------------
# Deep PR Detail Extractor
# ---------------------------------------------------------------------------
def enrich_pr_details(
    session: requests.Session,
    org: str,
    repo_name: str,
    pr_number: int,
    base_info: dict,
) -> dict:
    """
    Fetch deep PR metrics: reviews, comments, commits, additions, deletions,
    human taxonomy analysis, rework counts, and risk classification.
    """
    full_repo = f"{org}/{repo_name}"
    pr_label = f"{full_repo}#{pr_number}"

    # 1. PR Detailed Object
    pr_detail_url = f"{GITHUB_API_BASE}/repos/{full_repo}/pulls/{pr_number}"
    res = github_get(session, pr_detail_url, label=f"PR details {pr_label}")
    pr_data = res.json()

    # 2. Reviews
    reviews_url = f"{GITHUB_API_BASE}/repos/{full_repo}/pulls/{pr_number}/reviews"
    reviews_res = github_get(session, reviews_url, label=f"Reviews {pr_label}")
    reviews = reviews_res.json()

    # 3. Inline Review Comments
    review_comments_url = f"{GITHUB_API_BASE}/repos/{full_repo}/pulls/{pr_number}/comments"
    rev_comments_res = github_get(session, review_comments_url, label=f"Review comments {pr_label}")
    review_comments = rev_comments_res.json()

    # 4. PR Issue Comments (General)
    issue_comments_url = f"{GITHUB_API_BASE}/repos/{full_repo}/issues/{pr_number}/comments"
    issue_comments_res = github_get(session, issue_comments_url, label=f"Issue comments {pr_label}")
    issue_comments = issue_comments_res.json()

    # 5. Commits
    commits_url = f"{GITHUB_API_BASE}/repos/{full_repo}/pulls/{pr_number}/commits"
    commits_res = github_get(session, commits_url, label=f"Commits {pr_label}")
    commits = commits_res.json()

    # Base Fields
    author = pr_data.get("user", {}).get("login", base_info.get("CREATED_USER_NAME", "unknown"))
    author_type = pr_data.get("user", {}).get("type", "User")
    is_bot = (author.lower() in KNOWN_BOTS) or (author_type == "Bot") or ("[bot]" in author.lower())

    created_at = datetime.strptime(pr_data["created_at"], GH_DATE_FMT)
    merged_at_str = pr_data.get("merged_at")
    closed_at_str = pr_data.get("closed_at")

    merged_at = datetime.strptime(merged_at_str, GH_DATE_FMT) if merged_at_str else None
    closed_at = datetime.strptime(closed_at_str, GH_DATE_FMT) if closed_at_str else None
    end_time = merged_at or closed_at or datetime.now(timezone.utc)

    # Lead Time
    lead_time_hrs = round((end_time - created_at).total_seconds() / 3600, 2)

    # First Review & TTFR
    human_reviews = [
        r for r in reviews
        if r.get("user", {}).get("type") != "Bot" and r.get("user", {}).get("login") != author
    ]
    human_inline_comments = [
        c for c in review_comments
        if c.get("user", {}).get("type") != "Bot" and c.get("user", {}).get("login") != author
    ]
    human_issue_comments = [
        c for c in issue_comments
        if c.get("user", {}).get("type") != "Bot" and c.get("user", {}).get("login") != author
    ]

    all_review_timestamps = []
    for r in human_reviews:
        if r.get("submitted_at"):
            all_review_timestamps.append(datetime.strptime(r["submitted_at"], GH_DATE_FMT))
    for c in human_inline_comments:
        if c.get("created_at"):
            all_review_timestamps.append(datetime.strptime(c["created_at"], GH_DATE_FMT))

    first_review_at = min(all_review_timestamps) if all_review_timestamps else None
    ttfr_hrs = round((first_review_at - created_at).total_seconds() / 3600, 2) if first_review_at else None

    # Review Cycle Time & Merge Delay
    approval_timestamps = [
        datetime.strptime(r["submitted_at"], GH_DATE_FMT)
        for r in human_reviews
        if r.get("state") == "APPROVED" and r.get("submitted_at")
    ]
    first_approval_at = min(approval_timestamps) if approval_timestamps else None

    review_cycle_hrs = round((first_approval_at - first_review_at).total_seconds() / 3600, 2) if (first_approval_at and first_review_at and first_approval_at >= first_review_at) else None
    merge_delay_hrs = round((end_time - first_approval_at).total_seconds() / 3600, 2) if (first_approval_at and end_time >= first_approval_at) else None

    # Code Size
    additions = pr_data.get("additions", 0)
    deletions = pr_data.get("deletions", 0)
    changed_files = pr_data.get("changed_files", 0)
    total_lines = additions + deletions

    # Reviewers Count & Review Rounds
    unique_reviewers = {
        r.get("user", {}).get("login") for r in human_reviews if r.get("user", {}).get("login")
    }.union({
        c.get("user", {}).get("login") for c in human_inline_comments if c.get("user", {}).get("login")
    })
    reviewers_count = len(unique_reviewers)
    review_rounds = len([r for r in human_reviews if r.get("state") in ("APPROVED", "CHANGES_REQUESTED")])

    # Commits & Post-Review Rework
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

    # Human-Only Comment Taxonomy Keyword Parsing
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

    # Review Defect Density (defects per 100 lines)
    defect_density = round((defect_comments_count / max(total_lines, 1)) * 100, 2)

    # Rework Classification
    if commits_after_first_review > 0:
        if defect_comments_count > 0 or arch_comments_count > 0:
            rework_category = "CRITICAL_REWORK"
        else:
            rework_category = "NORMAL_REWORK"
    else:
        rework_category = "NO_REWORK"

    # Human-Centric Risk Calibration & Small-PR Guardrail
    severity_score = round(
        (total_lines * 0.1)
        + (changed_files * 2.0)
        + (total_commits * 1.5)
        + (defect_comments_count * 15.0),
        1
    )

    # Risk Level Evaluation
    if is_bot:
        risk_level = "LOW"
        risk_explanation = "LOW RISK: AUTOMATED UPDATE: Dependabot/Bot dependency bump (excluded from human defect & team review metrics)."
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
        "PR_ID": pr_data.get("id"),
        "PR_URL": pr_data.get("html_url"),
        "IS_BOT": is_bot,
        "AUTHOR": author,
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
# Input Handler: Excel Input Mode
# ---------------------------------------------------------------------------
def load_prs_from_excel(excel_path: Path, token: str) -> List[dict]:
    """
    Read PRs directly from the Raw_Data sheet of 2026Q1_GitAnalysis.xlsx,
    then call GitHub API for deep metrics.
    """
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


# ---------------------------------------------------------------------------
# Input Handler: Direct REST API Mode
# ---------------------------------------------------------------------------
def load_prs_from_api(config_path: Path) -> List[dict]:
    """
    Load repos from source.json, query GitHub for closed PRs, pre-filter by date,
    and fetch deep metrics.
    """
    if not config_path.exists():
        raise FileNotFoundError(f"Config file not found: {config_path}")

    with config_path.open("r", encoding="utf-8") as f:
        source_data = json.load(f)

    # Normalize teams schema
    team_configs = []
    if isinstance(source_data, dict) and "teams" in source_data:
        for t in source_data["teams"]:
            team_name = t["team"]
            org = t.get("org", "krogertechnology")
            begin_date = t.get("begin_date", DEFAULT_BEGIN_DATE)
            end_date = t.get("end_date", DEFAULT_END_DATE)
            token = t.get("token", "")
            repos = t["repos"]

            # Token resolution from ENV
            env_key = team_name.upper().replace(" ", "_").replace("/", "_") + "_TOKEN"
            resolved_token = os.environ.get(env_key, os.environ.get("GITHUB_TOKEN", token))

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
            logger.error("Missing token for %s. Set PAYMENTS_TOKEN or GITHUB_TOKEN environment variable.", full_repo)
            continue

        session = _build_session(token)

        # List Closed PRs
        page = 1
        matching_prs = []

        while True:
            url = f"{GITHUB_API_BASE}/repos/{full_repo}/pulls?state=closed&per_page=100&page={page}"
            res = github_get(session, url, label=f"Fetch closed PRs page {page}")
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

        logger.info("Found %d matching merged PRs in date range for %s", len(matching_prs), full_repo)

        for idx, pr_item in enumerate(matching_prs, start=1):
            pr_num = pr_item["number"]
            logger.info("Enriching PR %d/%d: %s#%d", idx, len(matching_prs), full_repo, pr_num)
            base_info = {"TEAM": team_name}
            enriched = enrich_pr_details(session, org, repo_name, pr_num, base_info)
            enriched_prs.append(enriched)

    return enriched_prs


# ---------------------------------------------------------------------------
# Analytics Aggregators
# ---------------------------------------------------------------------------
def compute_user_breakdown(enriched_prs: List[dict]) -> pd.DataFrame:
    """Compute human contributor performance breakdown table."""
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

        # Risk Profile Breakdown string (C/H/M/L)
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
    """Identify top 5 PRs with highest rework ratio and commits post-review."""
    df = pd.DataFrame(enriched_prs)
    if df.empty:
        return pd.DataFrame()

    df_rework = df.sort_values(by=["COMMITS_AFTER_REVIEW", "REWORK_RATIO_PCT", "DEFECT_COMMENTS"], ascending=False).head(5)
    cols = [
        "PR_NUMBER", "REPO_NAME", "AUTHOR", "LEAD_TIME_HRS",
        "TOTAL_COMMITS", "COMMITS_AFTER_REVIEW", "REWORK_RATIO_PCT",
        "DEFECT_COMMENTS", "REWORK_CATEGORY", "RISK_LEVEL"
    ]
    return df_rework[cols]


# ---------------------------------------------------------------------------
# PDF Exporter (ReportLab with Final Glossary Page)
# ---------------------------------------------------------------------------
def generate_pdf_report(enriched_prs: List[dict], pdf_path: Path):
    """Generate multi-page Executive PDF report with Metrics Glossary final page."""
    if not REPORTLAB_AVAILABLE:
        logger.warning("ReportLab not installed. Skipping PDF generation.")
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
    title_style = ParagraphStyle(
        "DocTitle",
        parent=styles["Heading1"],
        fontName="Helvetica-Bold",
        fontSize=20,
        leading=24,
        textColor=colors.HexColor("#1A365D"),
    )
    h2_style = ParagraphStyle(
        "SectionH2",
        parent=styles["Heading2"],
        fontName="Helvetica-Bold",
        fontSize=13,
        leading=17,
        textColor=colors.HexColor("#2B6CB0"),
        spaceBefore=10,
        spaceAfter=6,
    )
    body_style = ParagraphStyle(
        "BodyTextCustom",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=8.5,
        leading=11,
        textColor=colors.HexColor("#2D3748"),
    )
    table_cell = ParagraphStyle(
        "TableCell",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=8,
        leading=10,
    )
    table_header = ParagraphStyle(
        "TableHeader",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=8,
        leading=10,
        textColor=colors.white,
    )

    story = []

    # Title & Banner
    story.append(Paragraph("GitHub Engineering Executive Analysis Dashboard", title_style))
    story.append(Paragraph(f"<b>Generated:</b> {datetime.now().strftime('%Y-%m-%d %H:%M')} | <b>Period:</b> Q1 2026 Analysis", body_style))
    story.append(Spacer(1, 10))

    # KPI Summary Cards
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
        ('PADDING', (0,0), (-1,-1), 6),
    ]))
    story.append(t_kpi)
    story.append(Spacer(1, 12))

    # 1. Contributor Performance Table
    story.append(Paragraph("1. Human Contributor Performance Breakdown", h2_style))
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
            ('PADDING', (0,0), (-1,-1), 4),
        ]))
        story.append(t_user)

    story.append(Spacer(1, 12))

    # 2. Top 5 High Rework Table
    story.append(Paragraph("2. Top 5 High-Rework Pull Requests", h2_style))
    df_top_rework = compute_top_5_rework(enriched_prs)
    if not df_top_rework.empty:
        r_headers = ["PR #", "Repo", "Author", "Lead (h)", "Total Commits", "Post Commits", "Rework %", "Defects", "Category", "Risk"]
        r_rows = [[Paragraph(h, table_header) for h in r_headers]]

        for _, r in df_top_rework.iterrows():
            r_rows.append([
                Paragraph(f"#{r['PR_NUMBER']}", table_cell),
                Paragraph(str(r['REPO_NAME']), table_cell),
                Paragraph(str(r['AUTHOR']), table_cell),
                Paragraph(f"{r['LEAD_TIME_HRS']}h", table_cell),
                Paragraph(str(r['TOTAL_COMMITS']), table_cell),
                Paragraph(str(r['COMMITS_AFTER_REVIEW']), table_cell),
                Paragraph(f"{r['REWORK_RATIO_PCT']}%", table_cell),
                Paragraph(str(r['DEFECT_COMMENTS']), table_cell),
                Paragraph(str(r['REWORK_CATEGORY']), table_cell),
                Paragraph(str(r['RISK_LEVEL']), table_cell),
            ])

        t_top = Table(r_rows, colWidths=[0.5*inch, 1.4*inch, 0.9*inch, 0.6*inch, 0.7*inch, 0.7*inch, 0.7*inch, 0.5*inch, 1.1*inch, 0.6*inch])
        t_top.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#C53030')),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E2E8F0')),
            ('PADDING', (0,0), (-1,-1), 4),
        ]))
        story.append(t_top)

    story.append(Spacer(1, 12))

    # 3. Detailed Risk Insights Explanations
    story.append(Paragraph("3. Detailed Executive Risk Insights & Explanations", h2_style))
    d_headers = ["PR #", "Author", "Lines/Files", "Defects", "Risk Level", "Plain-Language Risk Explanation"]
    d_rows = [[Paragraph(h, table_header) for h in d_headers]]

    for p in enriched_prs[:15]:  # Top 15 PRs
        d_rows.append([
            Paragraph(f"#{p['PR_NUMBER']}", table_cell),
            Paragraph(str(p['AUTHOR']), table_cell),
            Paragraph(f"{p['TOTAL_LINES']}L / {p['CHANGED_FILES']}F", table_cell),
            Paragraph(str(p['DEFECT_COMMENTS']), table_cell),
            Paragraph(str(p['RISK_LEVEL']), table_cell),
            Paragraph(str(p['RISK_EXPLANATION']), table_cell),
        ])

    t_detail = Table(d_rows, colWidths=[0.5*inch, 0.9*inch, 0.8*inch, 0.5*inch, 0.7*inch, 3.8*inch])
    t_detail.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#2D3748')),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E2E8F0')),
        ('PADDING', (0,0), (-1,-1), 4),
    ]))
    story.append(t_detail)

    # -----------------------------------------------------------------------
    # PAGE BREAK -> DEDICATED METRICS GLOSSARY PAGE
    # -----------------------------------------------------------------------
    story.append(PageBreak())

    story.append(Paragraph("Metrics Calculation Methodology & Glossary", title_style))
    story.append(Paragraph("Comprehensive definitions, mathematical formulas, guardrails, and risk threshold rules.", body_style))
    story.append(Spacer(1, 10))

    g_headers = ["Metric Name", "Mathematical Formula / Rule", "Business Meaning & Purpose"]
    g_rows = [[Paragraph(h, table_header) for h in g_headers]]

    glossary_items = [
        ("PR Lead Time (hrs)", "t_merged/closed - t_created", "Total cycle time elapsed from PR creation until final merge or closure."),
        ("Time to 1st Review (TTFR)", "t_first_review - t_created", "Measures peer reviewer responsiveness and latency in starting code reviews."),
        ("Review Cycle Time", "t_first_approval - t_first_review", "Active iteration duration spent in code review cycles between author and reviewers."),
        ("Merge Delay", "t_merged - t_first_approval", "Idle latency between receiving final code approval and executing the merge."),
        ("Rework Ratio (%)", "(Commits After 1st Review / Total Commits) * 100", "Proportion of total commits pushed after code review feedback was received."),
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
        ('PADDING', (0,0), (-1,-1), 5),
    ]))
    story.append(t_glossary)

    doc.build(story)
    logger.info("Generated PDF report: %s", pdf_path)


# ---------------------------------------------------------------------------
# Excel Exporter
# ---------------------------------------------------------------------------
def generate_excel_report(enriched_prs: List[dict], excel_path: Path):
    """Write multi-tab Excel workbook including Metrics Glossary tab."""
    df_detail = pd.DataFrame(enriched_prs)
    df_user = compute_user_breakdown(enriched_prs)
    df_rework = compute_top_5_rework(enriched_prs)

    glossary_data = [
        {"Metric Name": "PR Lead Time (hrs)", "Formula / Rule": "t_merged/closed - t_created", "Business Purpose": "Total cycle time from PR creation to merge."},
        {"Metric Name": "Time to 1st Review (TTFR)", "Formula / Rule": "t_first_review - t_created", "Business Purpose": "Reviewer responsiveness and latency."},
        {"Metric Name": "Small-PR Guardrail", "Formula / Rule": "Lines <= 50 & Files <= 3 & Defects == 0 -> LOW", "Business Purpose": "Prevents small, clean PRs from false high risk flags."},
        {"Metric Name": "Critical Rework", "Formula / Rule": "Post-review commits + Defect/Arch comments", "Business Purpose": "Isolates bug-driven rework from style nits."},
        {"Metric Name": "Rework Ratio (%)", "Formula / Rule": "(Post-Review Commits / Total Commits) * 100", "Business Purpose": "Percentage of rework post-review."},
        {"Metric Name": "Severity Score", "Formula / Rule": "(Lines*0.1) + (Files*2.0) + (Commits*1.5) + (Defects*15)", "Business Purpose": "Unified complexity score."},
    ]
    df_glossary = pd.DataFrame(glossary_data)

    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        df_detail.to_excel(writer, sheet_name="Detailed_PR_Metrics", index=False)
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
    parser.add_argument("--config", "-c", type=Path, default=DEFAULT_CONFIG_PATH, help="Path to JSON config file (default: source.json)")
    args = parser.parse_args()

    pipeline_start = time.perf_counter()
    logger.info("Starting GitHub PR Advanced Analyzer...")

    enriched_prs = []

    # Decision Logic for Input Mode
    if args.excel or DEFAULT_INPUT_EXCEL.exists():
        input_excel = args.excel or DEFAULT_INPUT_EXCEL
        token = os.environ.get("PAYMENTS_TOKEN", os.environ.get("GITHUB_TOKEN", ""))

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
            logger.warning("No GitHub token found in env (PAYMENTS_TOKEN/GITHUB_TOKEN). API requests may fail if repos are private.")

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
