"""
GitHub PR TechM Comment Analysis

Builds a report focused on PRs authored by configured techM members, then
collects PR comment details and tech lead participation metrics.
"""

import logging
import os
import time
from datetime import datetime
from pathlib import Path
from typing import Any, Optional

import pandas as pd
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

from github_pr_analysis_improved import (
    CONFIG_PATH,
    DATE_FMT,
    GITHUB_API_BASE,
    GITHUB_API_VERSION,
    REQUEST_TIMEOUT,
    load_request_data,
)

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
# Constants
# ---------------------------------------------------------------------------
SCRIPT_DIR = Path(__file__).resolve().parent
OUTPUT_PATH = SCRIPT_DIR / "output" / "2026Q1_GitAnalysis_TechM_Comments.xlsx"
DEFAULT_BEGIN_DATE = "2026-02-01"
DEFAULT_END_DATE = "2026-05-23"
MAX_RETRIES = 3
BACKOFF_FACTOR = 1.5


def _build_session(token: str) -> requests.Session:
    """Build a requests session pre-configured for GitHub API requests."""
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


def _normalize_logins(values: Any) -> list[str]:
    """Normalize user login arrays by trimming blanks and lowercasing."""
    if not isinstance(values, list):
        return []
    return [str(value).strip().lower() for value in values if str(value).strip()]


def fetch_pull_requests(repo: str, session: requests.Session) -> list[dict]:
    """
    Fetch all pull requests for a repository using pagination.

    Note: state=all is used intentionally per reporting requirements.
    """
    all_prs: list[dict] = []
    page = 1

    while True:
        url = f"{GITHUB_API_BASE}/repos/{repo}/pulls?state=all&per_page=100&page={page}"
        logger.info("Fetching PR list for %s (page %d)", repo, page)
        response = session.get(url, timeout=REQUEST_TIMEOUT)

        if response.status_code == 403 and "rate limit" in response.text.lower():
            reset_ts = int(response.headers.get("X-RateLimit-Reset", time.time() + 60))
            wait_secs = max(reset_ts - int(time.time()), 1)
            logger.warning("Rate limit hit. Sleeping %d seconds", wait_secs)
            time.sleep(wait_secs)
            continue

        response.raise_for_status()
        page_data: list[dict] = response.json()
        if not page_data:
            break

        all_prs.extend(page_data)
        if len(page_data) < 100:
            break

        page += 1

    logger.info("Total PRs fetched for %s: %d", repo, len(all_prs))
    return all_prs


def fetch_pr_comments(
    session: requests.Session,
    pr_api_url: str,
    fallback_comments_url: Optional[str],
) -> list[dict]:
    """
    Fetch PR comments from PR API url + '/comments' with an API fallback.

    The first call follows the requested URL format; if response is non-JSON,
    the API comments endpoint from the PR payload is used.
    """
    primary_url = f"{pr_api_url}/comments"
    response = session.get(primary_url, timeout=REQUEST_TIMEOUT)

    if response.status_code == 403 and "rate limit" in response.text.lower():
        reset_ts = int(response.headers.get("X-RateLimit-Reset", time.time() + 60))
        wait_secs = max(reset_ts - int(time.time()), 1)
        logger.warning("Rate limit hit while fetching comments. Sleeping %d seconds", wait_secs)
        time.sleep(wait_secs)
        response = session.get(primary_url, timeout=REQUEST_TIMEOUT)

    try:
        response.raise_for_status()
        payload = response.json()
        if isinstance(payload, list):
            return payload
    except (requests.RequestException, ValueError):
        pass

    if fallback_comments_url:
        fallback_response = session.get(fallback_comments_url, timeout=REQUEST_TIMEOUT)
        fallback_response.raise_for_status()
        payload = fallback_response.json()
        if isinstance(payload, list):
            return payload

    return []


def _is_within_date_window(created_at_raw: Optional[str], begin: datetime, end: datetime) -> bool:
    """Return True when PR created_at is inside inclusive date range."""
    if not created_at_raw:
        return False
    created_at = datetime.strptime(created_at_raw, "%Y-%m-%dT%H:%M:%SZ")
    return begin.date() <= created_at.date() <= end.date()


def _resolve_report_date_range(begin_date_str: str, end_date_str: str) -> tuple[datetime, datetime]:
    """Resolve report begin/end date strings, falling back to default constants."""
    begin_date = datetime.strptime((begin_date_str or DEFAULT_BEGIN_DATE).strip(), DATE_FMT)
    end_date = datetime.strptime((end_date_str or DEFAULT_END_DATE).strip(), DATE_FMT)
    return begin_date, end_date


def build_report_rows(
    team_name: str,
    repo_name: str,
    begin_date_str: str,
    end_date_str: str,
    techm_members: list[str],
    tech_leads: list[str],
    pull_requests: list[dict],
    session: requests.Session,
) -> list[dict]:
    """Build detailed rows for PRs authored by configured techM members."""
    begin_date, end_date = _resolve_report_date_range(begin_date_str, end_date_str)
    rows: list[dict] = []

    for pr in pull_requests:
        author_login = str(pr.get("user", {}).get("login", "")).strip().lower()
        if not author_login or author_login not in techm_members:
            continue

        try:
            if not _is_within_date_window(pr.get("created_at"), begin_date, end_date):
                continue
        except ValueError:
            continue

        pr_api_url = str(pr.get("url", "")).strip()
        if not pr_api_url:
            continue

        comments = fetch_pr_comments(session, pr_api_url, pr.get("comments_url"))
        total_records = len(comments)

        lead_comment_records: list[str] = []
        for comment in comments:
            comment_user = str(comment.get("user", {}).get("login", "")).strip().lower()
            if comment_user not in tech_leads:
                continue

            lead_comment_records.append(
                f"{comment_user} : {comment.get('body', '')} : {comment.get('created_at', '')}"
            )

        rows.append(
            {
                "TEAM": team_name,
                "REPO_NAME": repo_name,
                "TECHM_MEMBER": author_login,
                "PR_ID": pr.get("id"),
                "PR_NUMBER": pr.get("number"),
                "PR_TITLE": pr.get("title"),
                "PR_STATE": pr.get("state"),
                "PR_CREATED_AT": pr.get("created_at"),
                "PR_UPDATED_AT": pr.get("updated_at"),
                "PR_CLOSED_AT": pr.get("closed_at"),
                "PR_URL": pr_api_url,
                "TOTAL_COMMENT_RECORDS": total_records,
                "TECH_LEAD_COMMENT_RECORDS": len(lead_comment_records),
                "TECH_LEAD_COMMENT_DETAILS": "\n".join(lead_comment_records),
            }
        )

    return rows


def main() -> None:
    """Execute techM-member PR analysis and export results to Excel."""
    request_data = load_request_data(CONFIG_PATH)

    detailed_rows: list[dict] = []
    skipped_entries: list[dict] = []

    for element in request_data.index:
        team_name = str(request_data.loc[element, "TEAM"])
        repo_name = str(request_data.loc[element, "REPO_NAME"])
        begin_date = str(request_data.loc[element, "BEGIN_DATE"])
        end_date = str(request_data.loc[element, "END_DATE"])
        env_key = team_name.upper().replace(" ", "_").replace("/", "_") + "_TOKEN"
        token = str(os.environ.get(env_key, request_data.loc[element, "GIT_TOKEN"]))

        techm_members = _normalize_logins(request_data.loc[element, "TECHM_MEMBERS"])
        tech_leads = _normalize_logins(request_data.loc[element, "TECH_LEAD"])

        if not techm_members:
            skipped_entries.append(
                {
                    "TEAM": team_name,
                    "REPO_NAME": repo_name,
                    "REASON": "Skipped: techm_members is empty",
                }
            )
            logger.info("Skipping %s/%s because techm_members is empty", team_name, repo_name)
            continue

        if not token:
            skipped_entries.append(
                {
                    "TEAM": team_name,
                    "REPO_NAME": repo_name,
                    "REASON": "Skipped: missing GitHub token",
                }
            )
            logger.warning("Skipping %s/%s because token is missing", team_name, repo_name)
            continue

        repo = f"krogertechnology/{repo_name}"
        session = _build_session(token)

        try:
            pull_requests = fetch_pull_requests(repo, session)
            report_rows = build_report_rows(
                team_name=team_name,
                repo_name=repo_name,
                begin_date_str=begin_date,
                end_date_str=end_date,
                techm_members=techm_members,
                tech_leads=tech_leads,
                pull_requests=pull_requests,
                session=session,
            )
            detailed_rows.extend(report_rows)
        except (requests.RequestException, ValueError) as exc:
            skipped_entries.append(
                {
                    "TEAM": team_name,
                    "REPO_NAME": repo_name,
                    "REASON": f"Skipped due to API/config error: {exc}",
                }
            )
            logger.error("Failed processing %s: %s", repo, exc)

    detailed_df = pd.DataFrame(detailed_rows)

    if detailed_df.empty:
        repo_member_summary_df = pd.DataFrame(
            columns=[
                "TEAM",
                "REPO_NAME",
                "TECHM_MEMBER",
                "MATCHED_PRS",
                "TOTAL_COMMENT_RECORDS",
                "TECH_LEAD_COMMENT_RECORDS",
            ]
        )
        repo_summary_df = pd.DataFrame(
            columns=[
                "TEAM",
                "REPO_NAME",
                "MATCHED_PRS",
                "TOTAL_COMMENT_RECORDS",
                "TECH_LEAD_COMMENT_RECORDS",
            ]
        )
    else:
        repo_member_summary_df = (
            detailed_df.groupby(["TEAM", "REPO_NAME", "TECHM_MEMBER"], as_index=False)
            .agg(
                MATCHED_PRS=("PR_ID", "count"),
                TOTAL_COMMENT_RECORDS=("TOTAL_COMMENT_RECORDS", "sum"),
                TECH_LEAD_COMMENT_RECORDS=("TECH_LEAD_COMMENT_RECORDS", "sum"),
            )
            .sort_values(["TEAM", "REPO_NAME", "TECHM_MEMBER"])
        )
        repo_summary_df = (
            detailed_df.groupby(["TEAM", "REPO_NAME"], as_index=False)
            .agg(
                MATCHED_PRS=("PR_ID", "count"),
                TOTAL_COMMENT_RECORDS=("TOTAL_COMMENT_RECORDS", "sum"),
                TECH_LEAD_COMMENT_RECORDS=("TECH_LEAD_COMMENT_RECORDS", "sum"),
            )
            .sort_values(["TEAM", "REPO_NAME"])
        )

    skipped_df = pd.DataFrame(skipped_entries)

    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(OUTPUT_PATH, engine="openpyxl") as writer:
        repo_member_summary_df.to_excel(writer, sheet_name="Repo_TechM_Summary", index=False)
        repo_summary_df.to_excel(writer, sheet_name="Repo_Summary", index=False)
        detailed_df.to_excel(writer, sheet_name="PR_Details", index=False)
        skipped_df.to_excel(writer, sheet_name="Skipped", index=False)

    logger.info("TechM comment analysis written to %s", OUTPUT_PATH)
    logger.info("Rows in PR_Details: %d", len(detailed_df))


if __name__ == "__main__":
    main()
