# =============================================================================
# GitHub PR Analysis - Improved Version
# =============================================================================

import json
import logging
import os
import statistics
import time
from collections import OrderedDict
from datetime import datetime
from pathlib import Path
from typing import Any, Optional

import pandas as pd
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

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
CONFIG_PATH = SCRIPT_DIR / "source.json" 
OUTPUT_PATH = SCRIPT_DIR / "2026Q1_GitAnalysis.xlsx"
GITHUB_API_BASE = "https://api.github.com"
GITHUB_API_VERSION = "2022-11-28"
MAIN_BRANCHES = {"main", "master", "dev", "develop", "trunk"}
DATE_FMT = "%Y-%m-%d"
GH_DATE_FMT = "%Y-%m-%dT%H:%M:%SZ"
DEFAULT_BEGIN_DATE = "2025-02-01"
DEFAULT_END_DATE = "2026-05-23"
MAX_RETRIES = 3
BACKOFF_FACTOR = 1.5
REQUEST_TIMEOUT = 30
LEGACY_REQUIRED_KEYS = {"team", "repo"}
TEAM_SCHEMA_REQUIRED_KEYS = {"team", "repos"}


def _validate_date(value: str, field_name: str, context: str) -> None:
    """Validate date format for config fields."""
    try:
        datetime.strptime(value, DATE_FMT)
    except ValueError as exc:
        raise ValueError(
            f"Invalid {field_name} '{value}' in {context}; expected format {DATE_FMT}"
        ) from exc


def _resolve_date_range(
    begin_date: Optional[str], end_date: Optional[str], context: str
) -> tuple[str, str]:
    """Resolve begin/end dates from config or Python defaults."""
    resolved_begin = begin_date or DEFAULT_BEGIN_DATE
    resolved_end = end_date or DEFAULT_END_DATE

    _validate_date(resolved_begin, "begin_date", context)
    _validate_date(resolved_end, "end_date", context)

    if resolved_begin > resolved_end:
        raise ValueError(
            f"Invalid date range in {context}: begin_date {resolved_begin} is after end_date {resolved_end}"
        )

    return resolved_begin, resolved_end


def _normalize_legacy_rows(source_data: list[Any]) -> list[dict]:
    """Normalize legacy row-based source.json entries."""
    normalized_rows: list[dict] = []

    for index, item in enumerate(source_data):
        if not isinstance(item, dict):
            raise ValueError(f"Config entry at index {index} must be a JSON object")

        missing_keys = sorted(LEGACY_REQUIRED_KEYS - item.keys())
        if missing_keys:
            raise ValueError(
                f"Config entry at index {index} is missing required keys: {', '.join(missing_keys)}"
            )

        context = f"legacy config entry index {index}"
        begin_date, end_date = _resolve_date_range(
            item.get("begin_date"), item.get("end_date"), context
        )

        normalized_rows.append(
            {
                "TEAM": item["team"],
                "REPO_NAME": item["repo"],
                "BEGIN_DATE": begin_date,
                "END_DATE": end_date,
                "GIT_TOKEN": item.get("token", ""),
                "TECHM_MEMBERS": item.get("techm_members", []),
                "TECH_LEAD": item.get("tech_lead", []),
            }
        )

    return normalized_rows


def _normalize_team_centric(source_data: dict[str, Any]) -> list[dict]:
    """Normalize team-centric source.json entries."""
    teams = source_data.get("teams")
    if not isinstance(teams, list) or not teams:
        raise ValueError("Team-centric source.json must contain a non-empty 'teams' list")

    normalized_rows: list[dict] = []

    for team_index, team_item in enumerate(teams):
        if not isinstance(team_item, dict):
            raise ValueError(f"Team entry at index {team_index} must be a JSON object")

        missing_team_keys = sorted(TEAM_SCHEMA_REQUIRED_KEYS - team_item.keys())
        if missing_team_keys:
            raise ValueError(
                f"Team entry at index {team_index} is missing keys: {', '.join(missing_team_keys)}"
            )

        team_name = team_item["team"]
        team_begin_date = team_item.get("begin_date")
        team_end_date = team_item.get("end_date")
        team_token = team_item.get("token", "")
        team_members = team_item.get("techm_members", [])
        team_lead = team_item.get("tech_lead", [])
        team_repos = team_item["repos"]

        team_begin_date, team_end_date = _resolve_date_range(
            team_begin_date, team_end_date, f"team '{team_name}'"
        )

        if not isinstance(team_repos, list) or not team_repos:
            raise ValueError(f"Team '{team_name}' must define a non-empty repos list")

        for repo_index, repo_item in enumerate(team_repos):
            repo_context = f"team '{team_name}', repo index {repo_index}"
            repo_name: str

            if isinstance(repo_item, str):
                repo_name = repo_item
                repo_begin_date = team_begin_date
                repo_end_date = team_end_date
                repo_token = team_token
                repo_members = team_members
                repo_lead = team_lead
            elif isinstance(repo_item, dict):
                repo_name = repo_item.get("repo") or repo_item.get("name", "")
                repo_begin_date = repo_item.get("begin_date", team_begin_date)
                repo_end_date = repo_item.get("end_date", team_end_date)
                repo_token = repo_item.get("token", team_token)
                repo_members = repo_item.get("techm_members", team_members)
                repo_lead = repo_item.get("tech_lead", team_lead)
            else:
                raise ValueError(
                    f"Repo entry in {repo_context} must be a string or object, got {type(repo_item).__name__}"
                )

            if not repo_name:
                raise ValueError(f"Missing repo name in {repo_context}")

            repo_begin_date, repo_end_date = _resolve_date_range(
                repo_begin_date, repo_end_date, repo_context
            )

            normalized_rows.append(
                {
                    "TEAM": team_name,
                    "REPO_NAME": repo_name,
                    "BEGIN_DATE": repo_begin_date,
                    "END_DATE": repo_end_date,
                    "GIT_TOKEN": repo_token,
                    "TECHM_MEMBERS": repo_members,
                    "TECH_LEAD": repo_lead,
                }
            )

    return normalized_rows


def load_request_data(config_path: Path) -> pd.DataFrame:
    """
    Load repository configuration from source.json into a DataFrame.

    Supported schemas:
    - Legacy row-based list of repo entries
    - Team-centric schema with teams[] and nested repos[]
    - begin_date and end_date are optional in source and default to Python constants

    Args:
        config_path: Path to the JSON configuration file.

    Returns:
        DataFrame with normalized column names used by the analysis pipeline.

    Raises:
        FileNotFoundError: If the config file does not exist.
        ValueError: If the config file structure is invalid.
    """
    if not config_path.exists():
        raise FileNotFoundError(f"Config file not found: {config_path}")

    with config_path.open("r", encoding="utf-8") as config_file:
        source_data = json.load(config_file)

    normalized_rows: list[dict]
    if isinstance(source_data, list):
        if not source_data:
            raise ValueError("source.json must contain a non-empty list of repository configs")
        normalized_rows = _normalize_legacy_rows(source_data)
    elif isinstance(source_data, dict) and "teams" in source_data:
        normalized_rows = _normalize_team_centric(source_data)
    else:
        raise ValueError(
            "Unsupported source.json schema. Expected legacy list or team-centric object with 'teams'."
        )

    duplicate_tracker: OrderedDict[tuple[str, str], int] = OrderedDict()
    for row in normalized_rows:
        key = (row["TEAM"], row["REPO_NAME"])
        duplicate_tracker[key] = duplicate_tracker.get(key, 0) + 1

    duplicates = [
        f"{team}/{repo} (x{count})"
        for (team, repo), count in duplicate_tracker.items()
        if count > 1
    ]
    if duplicates:
        raise ValueError(f"Duplicate team/repo entries found: {', '.join(duplicates)}")

    request_data = pd.DataFrame(normalized_rows)
    logger.info("Loaded %d repo entries from %s", len(request_data), config_path)
    return request_data


def _build_session(token: str) -> requests.Session:
    """
    Build a requests.Session pre-configured for the GitHub API.

    Args:
        token: GitHub personal access token.

    Returns:
        Configured requests.Session.
    """
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


def fetch_pull_requests(repo: str, token: str) -> list[dict]:
    """
    Fetch all closed pull requests for a GitHub repository using pagination.

    Args:
        repo: Full repository path, e.g. "org/repo-name".
        token: GitHub personal access token.

    Returns:
        List of pull request dicts as returned by the GitHub API.

    Raises:
        ValueError: If no token is available for the repository.
        requests.HTTPError: If any API call returns a non-2xx status.
    """
    if not token:
        raise ValueError(f"Missing GitHub token for repository {repo}")

    session = _build_session(token)
    all_prs: list[dict] = []
    page = 1

    while True:
        url = (
            f"{GITHUB_API_BASE}/repos/{repo}/pulls"
            f"?state=closed&per_page=1000&page={page}"
        )
        logger.info("Fetching %s (page %d) ...", repo, page)
        response = session.get(url, timeout=REQUEST_TIMEOUT)
        logger.info(response)

        if response.status_code == 403 and "rate limit" in response.text.lower():
            reset_ts = int(response.headers.get("X-RateLimit-Reset", time.time() + 60))
            wait_secs = max(reset_ts - int(time.time()), 1)
            logger.warning("Rate limit hit. Sleeping %d seconds ...", wait_secs)
            time.sleep(wait_secs)
            continue

        response.raise_for_status()
        page_data: list[dict] = response.json()

        if not page_data:
            break

        all_prs.extend(page_data)
        logger.info("  -> %d PRs fetched so far for %s", len(all_prs), repo)

        if len(page_data) < 1000:
            break

        page += 1

    logger.info("Total PRs fetched for %s: %d", repo, len(all_prs))
    return all_prs


def _filter_prs(
    pull_requests: list[dict],
    begin_date: datetime,
    end_date: datetime,
) -> list[tuple[dict, datetime, datetime]]:
    """
    Filter pull requests to those merged into a main branch within the date range.

    Args:
        pull_requests: Raw list of PR dicts from the GitHub API.
        begin_date: Start of the date range, inclusive.
        end_date: End of the date range, inclusive.

    Returns:
        List of tuples containing the PR object and parsed timestamps.
    """
    filtered: list[tuple[dict, datetime, datetime]] = []

    for index, pull_request in enumerate(pull_requests):
        try:
            merged_at_str: Optional[str] = pull_request.get("merged_at")
            base_ref: Optional[str] = pull_request.get("base", {}).get("ref")

            if not merged_at_str or base_ref not in MAIN_BRANCHES:
                continue

            merged_at = datetime.strptime(merged_at_str, GH_DATE_FMT)
            created_at = datetime.strptime(pull_request["created_at"], GH_DATE_FMT)

            if begin_date <= merged_at <= end_date:
                filtered.append((pull_request, created_at, merged_at))

        except (KeyError, ValueError) as exc:
            logger.warning("Skipping PR at index %d due to error: %s", index, exc)

    return filtered


def store_calc_data(
    pull_requests: list[dict],
    begin_date_str: str,
    end_date_str: str,
    team_name: str,
) -> pd.DataFrame:
    """
    Build a detailed DataFrame of individual PRs merged within the date range.

    Args:
        pull_requests: Raw list of PR dicts from the GitHub API.
        begin_date_str: Start date string in YYYY-MM-DD format.
        end_date_str: End date string in YYYY-MM-DD format.
        team_name: Team label to include in the output.

    Returns:
        DataFrame with one row per qualifying PR.
    """
    begin_date = datetime.strptime(begin_date_str, DATE_FMT)
    end_date = datetime.strptime(end_date_str, DATE_FMT)

    rows: list[dict] = []
    for pull_request, created_at, merged_at in _filter_prs(pull_requests, begin_date, end_date):
        difference_hours = round((merged_at - created_at).total_seconds() / 3600, 2)
        rows.append(
            {
                "TEAM": team_name,
                "REPO_NAME": pull_request["base"]["repo"]["name"],
                "PR_ID": pull_request["id"],
                "PR_URL": pull_request["html_url"],
                "PR_MERGED_AT": merged_at.strftime(GH_DATE_FMT),
                "PR_CREATED_AT": pull_request["created_at"],
                "DIFFERENCE_HOURS": difference_hours,
                "CREATED_USER_NAME": pull_request["user"]["login"],
                "CREATED_USER_ID": pull_request["user"]["id"],
            }
        )

    return pd.DataFrame(rows)


def calculate_average_merge_time(
    pull_requests: list[dict],
    begin_date_str: str,
    end_date_str: str,
) -> Optional[float]:
    """
    Calculate the mean PR merge time in hours for PRs merged within the date range.

    Args:
        pull_requests: Raw list of PR dicts from the GitHub API.
        begin_date_str: Start date string in YYYY-MM-DD format.
        end_date_str: End date string in YYYY-MM-DD format.

    Returns:
        Mean merge time in hours rounded to 2 decimal places,
        or None if no qualifying PRs exist.
    """
    begin_date = datetime.strptime(begin_date_str, DATE_FMT)
    end_date = datetime.strptime(end_date_str, DATE_FMT)

    merge_times = [
        (merged_at - created_at).total_seconds() / 3600
        for _, created_at, merged_at in _filter_prs(pull_requests, begin_date, end_date)
    ]

    if not merge_times:
        logger.warning(
            "No qualifying PRs found between %s and %s.", begin_date_str, end_date_str
        )
        return None

    return round(statistics.mean(merge_times), 2)


def main() -> None:
    """
    Read config JSON, fetch PRs, compute metrics, and write the Excel report.
    """
    request_data = load_request_data(CONFIG_PATH)
    raw_data_frames: list[pd.DataFrame] = []

    for element in request_data.index:
        begin_date: str = request_data.loc[element, "BEGIN_DATE"]
        end_date: str = request_data.loc[element, "END_DATE"]
        repo_name: str = request_data.loc[element, "REPO_NAME"]
        team_name: str = request_data.loc[element, "TEAM"]

        env_key = team_name.upper().replace(" ", "_").replace("/", "_") + "_TOKEN"
        token: str = os.environ.get(env_key, request_data.loc[element, "GIT_TOKEN"])
        repo = f"krogertechnology/{repo_name}"

        try:
            pull_requests = fetch_pull_requests(repo, token)
        except (requests.RequestException, ValueError) as exc:
            logger.error("Failed to fetch PRs for %s: %s", repo, exc)
            request_data.loc[element, "AVR_MERGED_TIME_HRS"] = None
            continue

        average_merge_time = calculate_average_merge_time(pull_requests, begin_date, end_date)
        request_data.loc[element, "AVR_MERGED_TIME_HRS"] = average_merge_time

        data_calc = store_calc_data(pull_requests, begin_date, end_date, team_name)
        raw_data_frames.append(data_calc)

    raw_data = pd.concat(raw_data_frames, ignore_index=True) if raw_data_frames else pd.DataFrame()
    output_data = request_data[
        ["TEAM", "REPO_NAME", "BEGIN_DATE", "END_DATE", "AVR_MERGED_TIME_HRS"]
    ]

    with pd.ExcelWriter(OUTPUT_PATH, engine="openpyxl") as writer:
        output_data.to_excel(writer, sheet_name="Output_Data", index=False)
        raw_data.to_excel(writer, sheet_name="Raw_Data", index=False)

    logger.info("Results written to %s", OUTPUT_PATH)
    logger.info("Summary:\n%s", output_data.to_string(index=False))


if __name__ == "__main__":
    main()
