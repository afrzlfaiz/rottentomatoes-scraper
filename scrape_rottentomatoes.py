import argparse
import csv
import json
import os
import re
import sys
from pathlib import Path
from typing import Any
from urllib.parse import urlencode, urlparse
from urllib.error import HTTPError
from urllib.request import Request, urlopen

from openpyxl import Workbook
from colorama import init, Fore, Style, Back
from datetime import datetime
import time

# Initialize colorama
init(autoreset=True)


MOVIE_REVIEWS_URL = "https://www.rottentomatoes.com/napi/rtcf/v1/movies/{movie_id}/reviews"
TV_SEASON_REVIEWS_URL = "https://www.rottentomatoes.com/napi/rtcf/v1/tv/seasons/{movie_id}/reviews"
PROPS_PATTERN = re.compile(
    r'<script\s+data-json="props"\s+type="application/json">\s*(\{.*?\})\s*</script>',
    re.DOTALL,
)
CLI_HEADER = r"""$$$$$$$\             $$\     $$\                          $$$$$$\  $$\       $$$$$$\
$$  __$$\            $$ |    $$ |                        $$  __$$\ $$ |      \_$$  _|
$$ |  $$ | $$$$$$\ $$$$$$\ $$$$$$\    $$$$$$\  $$$$$$$\  $$ /  \__|$$ |        $$ |
$$$$$$$  |$$  __$$\\_$$  _|\_$$  _|  $$  __$$\ $$  __$$\ $$ |      $$ |        $$ |
$$  __$$< $$ /  $$ | $$ |    $$ |    $$$$$$$$ |$$ |  $$ |$$ |      $$ |        $$ |
$$ |  $$ |$$ |  $$ | $$ |$$\ $$ |$$\ $$   ____|$$ |  $$ |$$ |  $$\ $$ |        $$ |
$$ |  $$ |\$$$$$$  | \$$$$  |\$$$$  |\$$$$$$$\ $$ |  $$ |\$$$$$$  |$$$$$$$$\ $$$$$$\
\__|  \__| \______/   \____/  \____/  \_______|\__|  \__| \______/ \________|\______|  """

# Color schemes
class Colors:
    HEADER = Fore.MAGENTA + Style.BRIGHT
    SUCCESS = Fore.GREEN + Style.BRIGHT
    WARNING = Fore.YELLOW + Style.BRIGHT
    ERROR = Fore.RED + Style.BRIGHT
    INFO = Fore.CYAN + Style.BRIGHT
    DEBUG = Fore.BLUE + Style.BRIGHT
    PROCESS = Fore.WHITE + Back.BLUE
    HIGHLIGHT = Fore.YELLOW + Back.BLACK
    RESET_ALL = Style.RESET_ALL

# Logger class for better output formatting
class Logger:
    def __init__(self):
        self.start_time = datetime.now()

    def log(self, level: str, message: str, emoji: str = ""):
        timestamp = datetime.now().strftime("%H:%M:%S")
        level_colors = {
            "INFO": Colors.INFO,
            "SUCCESS": Colors.SUCCESS,
            "WARNING": Colors.WARNING,
            "ERROR": Colors.ERROR,
            "PROCESS": Colors.PROCESS,
            "DEBUG": Colors.DEBUG,
            "HEADER": Colors.HEADER,
        }

        color = level_colors.get(level, Colors.INFO)
        emoji_prefix = f" {emoji}" if emoji else ""

        print(f"{color}[{timestamp}] {level}{emoji_prefix}: {message}{Style.RESET_ALL}")

    def log_progress(self, current: int, total: int, suffix: str = ""):
        percent = current / total * 100
        bar_length = 30
        filled_length = int(bar_length * current // total)
        bar = '█' * filled_length + '-' * (bar_length - filled_length)

        print(f"{Colors.PROCESS}\r[{bar}] {current}/{total} ({percent:.1f}%) {suffix}{Style.RESET_ALL}", end='\r')
        if current == total:
            print()  # New line when complete

    def log_with_spinner(self, message: str, duration: float = 0.5):
        spinner = ['⠋', '⠙', '⠹', '⠸', '⠼', '⠴', '⠦', '⠧', '⠇', '⠏']
        i = 0
        start = time.time()

        while time.time() - start < duration:
            print(f"{Colors.INFO}\r{message} {spinner[i % len(spinner)]}{Style.RESET_ALL}", end='\r')
            time.sleep(0.1)
            i += 1
        print()  # New line

# Initialize logger
logger = Logger()


def star_to_number(value: str | None) -> float | None:
    if not value or not value.startswith("STAR_"):
        return None
    return float(value.removeprefix("STAR_").replace("_", "."))


def fetch_page(
    movie_id: str,
    content_type: str,
    review_type: str,
    after: str = "",
    page_count: int = 20,
    verified: bool = False,
    top_only: bool = False,
) -> dict[str, Any]:
    params = {
        "after": after,
        "before": "",
        "pageCount": page_count,
        "topOnly": str(top_only).lower(),
        "type": review_type,
        "verified": str(verified).lower(),
    }
    if content_type == "tv":
        url = f"{TV_SEASON_REVIEWS_URL.format(movie_id=movie_id)}?{urlencode(params)}"
    else:
        url = f"{MOVIE_REVIEWS_URL.format(movie_id=movie_id)}?{urlencode(params)}"

    request = Request(
        url,
        headers={
            "User-Agent": "Mozilla/5.0",
            "Accept": "application/json",
        },
    )

    try:
        with urlopen(request) as response:
            data = json.load(response)
            if isinstance(data, list):
                data = {
                    "reviews": [item for item in data if isinstance(item, dict)],
                    "pageInfo": {"hasNextPage": False, "endCursor": ""},
                }
            return data
    except HTTPError as e:
        logger.log("ERROR", f"HTTP Error {e.code}: {e.reason}", "❌")
        raise
    except Exception as e:
        logger.log("ERROR", f"Request failed: {str(e)}", "❌")
        raise


def fetch_text(url: str) -> str:
    request = Request(
        url,
        headers={
            "User-Agent": "Mozilla/5.0",
            "Accept": "text/html,application/xhtml+xml",
        },
    )

    with urlopen(request) as response:
        return response.read().decode("utf-8", errors="replace")


def normalize_rt_url(rt_url: str) -> str:
    parsed = urlparse(rt_url.strip())
    if not parsed.scheme or not parsed.netloc:
        raise ValueError("Please enter a valid Rotten Tomatoes URL.")
    if parsed.netloc not in {"www.rottentomatoes.com", "rottentomatoes.com"}:
        raise ValueError("URL must be from rottentomatoes.com")

    path = parsed.path.rstrip("/")
    if not path or path == "/":
        raise ValueError("Invalid Rotten Tomatoes URL path.")

    return f"https://www.rottentomatoes.com{path}"


def is_tv_url(rt_url: str) -> bool:
    path = urlparse(normalize_rt_url(rt_url)).path.rstrip("/")
    return path.startswith("/tv/")


def tv_url_has_season(rt_url: str) -> bool:
    path = urlparse(normalize_rt_url(rt_url)).path.rstrip("/")
    return bool(re.search(r"/s\d{1,2}$", path))


def append_tv_season(rt_url: str, season: int) -> str:
    base_url = normalize_rt_url(rt_url)
    if season <= 0:
        raise ValueError("Season number must be greater than 0.")
    return f"{base_url}/s{season:02d}"


def ensure_tv_season_url(rt_url: str, interactive: bool) -> str:
    normalized = normalize_rt_url(rt_url)
    if not is_tv_url(normalized) or tv_url_has_season(normalized):
        return normalized

    if not interactive:
        raise ValueError("TV show URL must include season path (example: /s01).")

    season = prompt_positive_int("TV show detected. Which season number? ")
    season_url = append_tv_season(normalized, season)
    logger.log("INFO", f"Using TV season URL: {season_url}", "📺")
    return season_url


def resolve_movie_id_from_url(rt_url: str) -> tuple[str, str, str | None]:
    base_url = normalize_rt_url(rt_url)
    reviews_url = f"{base_url}/reviews/all-audience"

    logger.log("INFO", f"Accessing URL: {reviews_url}")
    logger.log_with_spinner("Fetching web page", 1.0)

    html = fetch_text(reviews_url)

    match = PROPS_PATTERN.search(html)
    if not match:
        logger.log("ERROR", "Failed to find JSON props on audience review page")
        raise ValueError("Failed to find JSON props on audience review page.")

    props = json.loads(match.group(1))
    media = props.get("media") or {}
    movie_id = media.get("emsId")
    title = media.get("title")
    if not movie_id:
        logger.log("ERROR", "emsId not found on audience review page")
        raise ValueError("emsId not found on audience review page.")

    logger.log("SUCCESS", f"Content ID found: {movie_id}", "🎯")
    if title:
        logger.log("SUCCESS", f"Title found: {title}", "🎬")
    return str(movie_id), reviews_url, title


def slug_from_url(rt_url: str) -> str:
    path = urlparse(normalize_rt_url(rt_url)).path.rstrip("/")
    return path.rsplit("/", 1)[-1]


def normalize_review(review: dict[str, Any], review_type: str) -> dict[str, Any]:
    if review_type == "audience":
        return {
            "id": review.get("ratingId"),
            "type": review_type,
            "rating": review.get("rating"),
            "rating_value": star_to_number(review.get("rating")),
            "verified": review.get("isVerified"),
            "display_name": review.get("displayName"),
            "review": review.get("review"),
            "created_at": review.get("createDate"),
        }

    critic = review.get("critic") or {}
    publication = review.get("publication") or {}
    return {
        "id": review.get("reviewId"),
        "type": review_type,
        "score_sentiment": review.get("scoreSentiment"),
        "original_score": review.get("originalScore"),
        "is_top_review": review.get("isTopReview"),
        "critic_name": critic.get("displayName"),
        "publication_name": publication.get("name"),
        "review_quote": review.get("reviewQuote"),
        "created_at": review.get("createDate"),
        "review_url": review.get("publicationReviewUrl"),
    }


def fetch_reviews(
    movie_id: str,
    content_type: str,
    review_type: str,
    max_reviews: int,
    verified: bool,
    top_only: bool,
) -> list[dict[str, Any]]:
    after = ""
    reviews: list[dict[str, Any]] = []

    logger.log("INFO", f"Starting to fetch {max_reviews} {review_type} reviews...", "📥")
    logger.log_with_spinner(f"Fetching {review_type} reviews", 1.0)

    page_num = 1
    while len(reviews) < max_reviews:
        page_count = min(20, max_reviews - len(reviews))

        # Show progress bar
        logger.log_progress(len(reviews), max_reviews, f"Review {review_type}")

        payload = fetch_page(
            movie_id=movie_id,
            content_type=content_type,
            review_type=review_type,
            after=after,
            page_count=page_count,
            verified=verified,
            top_only=top_only,
        )

        page_reviews = payload.get("reviews") or []
        reviews.extend(normalize_review(review, review_type) for review in page_reviews)

        if len(reviews) >= max_reviews:
            break

        page_info = payload.get("pageInfo") or {}
        if not page_info.get("hasNextPage"):
            logger.log("WARNING", "No next page")
            break

        after = page_info.get("endCursor") or ""
        if not after:
            logger.log("WARNING", "Empty cursor, stopping")
            break

        page_num += 1
        time.sleep(0.5)  # Small delay to avoid rate limiting

    logger.log("SUCCESS", f"Completed: {len(reviews)}/{max_reviews} {review_type} reviews fetched", "✅")
    return reviews[:max_reviews]


def clear_screen() -> None:
    os.system('cls' if os.name == 'nt' else 'clear')


def render_header() -> None:
    clear_screen()
    print(Colors.HEADER + CLI_HEADER + Style.RESET_ALL)
    print(Colors.HIGHLIGHT + "=" * 80 + Style.RESET_ALL)
    print()


def prompt_non_empty(message: str) -> str:
    error_message = ""
    while True:
        render_header()
        if error_message:
            logger.log("ERROR", error_message, "❌")
            print()
        value = input(f"{Colors.INFO}{message}{Style.RESET_ALL}").strip()
        if value:
            return value
        error_message = "Input cannot be empty."
        time.sleep(0.1)  # Small delay to prevent rapid rendering


def prompt_choice(message: str, options: dict[str, str]) -> str:
    error_message = ""
    while True:
        render_header()
        if error_message:
            logger.log("ERROR", error_message, "❌")
            print()

        print(f"{Colors.INFO}{message}{Style.RESET_ALL}")
        for key, label in options.items():
            print(f"{Colors.HIGHLIGHT}{key}. {Style.RESET_ALL}{label}")
        choice = input(f"{Colors.INFO}Choose a number: {Style.RESET_ALL}").strip()
        if choice in options:
            return choice
        error_message = "Invalid choice."


def prompt_positive_int(message: str) -> int:
    error_message = ""
    while True:
        render_header()
        if error_message:
            logger.log("ERROR", error_message, "❌")
            print()

        value = input(f"{Colors.INFO}{message}{Style.RESET_ALL}").strip()
        try:
            number = int(value)
        except ValueError:
            error_message = "Enter an integer."
            continue
        if number > 0:
            logger.log("SUCCESS", f"Input: {number}", "✓")
            return number
        error_message = "Number must be greater than 0."


def ensure_parent_dir(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)


def write_json(output_path: Path, payload: dict[str, Any]) -> None:
    logger.log("INFO", f"Writing JSON to {output_path}...", "💾")
    ensure_parent_dir(output_path)
    output_path.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")
    logger.log("SUCCESS", "JSON file written successfully", "✅")


def write_csv(base_path: Path, payload: dict[str, Any]) -> list[Path]:
    written_files: list[Path] = []
    base_stem = base_path.stem
    suffix = base_path.suffix if base_path.suffix.lower() == ".csv" else ".csv"

    for review_type, rows in payload["reviews"].items():
        logger.log("INFO", f"Writing CSV for {review_type} reviews...", "📊")
        output_path = base_path.with_name(f"{base_stem}_{review_type}{suffix}")
        ensure_parent_dir(output_path)

        enriched_rows = [
            {
                "title": payload.get("title"),
                "movie_id": payload.get("movie_id"),
                **row,
            }
            for row in rows
        ]

        fieldnames = list(enriched_rows[0].keys()) if enriched_rows else ["title", "movie_id", "type"]
        with output_path.open("w", newline="", encoding="utf-8-sig") as file:
            writer = csv.DictWriter(file, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(enriched_rows)
        written_files.append(output_path)
        logger.log("SUCCESS", f"CSV written: {output_path}", "✅")

    return written_files


def write_excel(output_path: Path, payload: dict[str, Any]) -> None:
    logger.log("INFO", f"Writing Excel to {output_path}...", "📈")
    ensure_parent_dir(output_path)
    workbook = Workbook()
    default_sheet = workbook.active
    workbook.remove(default_sheet)

    for review_type, rows in payload["reviews"].items():
        logger.log("INFO", f"Creating sheet for {review_type} reviews...", "📊")
        sheet = workbook.create_sheet(title=review_type[:31])

        enriched_rows = [
            {
                "title": payload.get("title"),
                "movie_id": payload.get("movie_id"),
                **row,
            }
            for row in rows
        ]

        headers = list(enriched_rows[0].keys()) if enriched_rows else ["title", "movie_id", "type"]
        sheet.append(headers)
        for row in enriched_rows:
            sheet.append([row.get(header) for header in headers])

    workbook.save(output_path)
    logger.log("SUCCESS", "Excel file written successfully", "✅")


def export_reviews(output_base: str, output_format: str, payload: dict[str, Any]) -> list[Path]:
    base_path = Path(output_base)
    written_files: list[Path] = []

    if output_format in {"json", "all"}:
        logger.log("INFO", "Exporting JSON files...", "📄")
        json_base = base_path.with_suffix("") if base_path.suffix.lower() == ".json" else base_path
        for review_type, rows in payload["reviews"].items():
            json_path = json_base.with_name(f"{json_base.name}_{review_type}").with_suffix(".json")
            json_payload = {
                **payload,
                "review_types": [review_type],
                "reviews": {review_type: rows},
            }
            write_json(json_path, json_payload)
            written_files.append(json_path)
            logger.log("SUCCESS", f"JSON file saved: {json_path}", "✅")

    if output_format in {"csv", "all"}:
        logger.log("INFO", "Exporting CSV files...", "📊")
        csv_before = len(written_files)
        written_files.extend(write_csv(base_path, payload))
        logger.log("SUCCESS", f"CSV files saved: {len(written_files) - csv_before} files", "✅")

    if output_format in {"excel", "all"}:
        logger.log("INFO", "Exporting Excel files...", "📈")
        excel_base = base_path.with_suffix("") if base_path.suffix.lower() == ".xlsx" else base_path
        for review_type, rows in payload["reviews"].items():
            excel_path = excel_base.with_name(f"{excel_base.name}_{review_type}").with_suffix(".xlsx")
            excel_payload = {
                **payload,
                "review_types": [review_type],
                "reviews": {review_type: rows},
            }
            write_excel(excel_path, excel_payload)
            written_files.append(excel_path)
            logger.log("SUCCESS", f"Excel file saved: {excel_path}", "✅")

    return written_files


def prompt_interactive_inputs() -> dict[str, Any]:
    source_url = prompt_non_empty("Enter Rotten Tomatoes movie URL: ")
    # Resolve movie ID immediately after URL input
    try:
        source_url = ensure_tv_season_url(source_url, interactive=True)
        movie_id, _, title = resolve_movie_id_from_url(source_url)
        logger.log("SUCCESS", f"Content ID: {movie_id}", "🎯")
        if title:
            logger.log("SUCCESS", f"Title: {title}", "🎬")
    except Exception as e:
        logger.log("ERROR", f"Failed to resolve movie ID: {e}", "❌")
        sys.exit(1)
    review_type_choice = prompt_choice(
        "Select review type:",
        {"1": "audience", "2": "critic", "3": "both"},
    )
    output_format_choice = prompt_choice(
        "Select output format:",
        {"1": "json", "2": "csv", "3": "excel", "4": "all"},
    )
    count = prompt_positive_int("Number of reviews to fetch: ")
    output_base = prompt_non_empty("Base output file name/path (e.g., output/review): ")

    review_type_map = {"1": "audience", "2": "critic", "3": "both"}
    output_format_map = {"1": "json", "2": "csv", "3": "excel", "4": "all"}
    return {
        "movie_id": movie_id,
        "title": title,
        "url": source_url,
        "type": review_type_map[review_type_choice],
        "count": count,
        "output": output_base,
        "format": output_format_map[output_format_choice],
    }


def show_interactive_summary(inputs: dict[str, Any]) -> None:
    render_header()
    logger.log("HEADER", "Configuration Summary:", "📋")
    print(f"{Colors.INFO}• URL: {Style.RESET_ALL}{inputs['url']}")
    print(f"{Colors.INFO}• Title: {Style.RESET_ALL}{inputs.get('title') or '-'}")
    print(f"{Colors.INFO}• Review type: {Style.RESET_ALL}{inputs['type']}")
    print(f"{Colors.INFO}• Number of reviews: {Style.RESET_ALL}{inputs['count']}")
    print(f"{Colors.INFO}• Base output: {Style.RESET_ALL}{inputs['output']}")
    print(f"{Colors.INFO}• Format output: {Style.RESET_ALL}{inputs['format']}")
    print()

    logger.log("INFO", "Press Enter to start scraping...", "🚀")
    input()


def main() -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")

    # Show welcome message
    render_header()
    logger.log("HEADER", "🍅 Rotten Tomatoes Scraper - Modern Version", "🎬")
    logger.log("INFO", "Script to scrape reviews from Rotten Tomatoes", "📊")
    print(Colors.HIGHLIGHT + "=" * 80 + Style.RESET_ALL)
    time.sleep(1)

    parser = argparse.ArgumentParser(
        description="Scrape Rotten Tomatoes reviews via public JSON endpoint."
    )
    parser.add_argument(
        "movie_id",
        nargs="?",
        help="Rotten Tomatoes Movie ID. If empty, script will prompt for movie URL interactively.",
    )
    parser.add_argument("--url", help="Rotten Tomatoes movie URL, example: https://www.rottentomatoes.com/m/lee_cronins_the_mummy")
    parser.add_argument(
        "--type",
        choices=["audience", "critic", "both"],
        default="both",
        help="Type of reviews to fetch",
    )
    parser.add_argument("--count", type=int, default=20, help="Number of reviews to fetch per type")
    parser.add_argument("--verified", action="store_true", help="Audience only: fetch verified reviews only")
    parser.add_argument("--top-only", action="store_true", help="Fetch top reviews only if available")
    parser.add_argument("--output", help="Base output file path. If empty in interactive mode, will be prompted")
    parser.add_argument(
        "--format",
        choices=["json", "csv", "excel", "all"],
        default="json",
        help="Output file format",
    )
    args = parser.parse_args()

    interactive_mode = not args.movie_id and not args.url
    interactive_inputs: dict[str, Any] = {}
    if interactive_mode:
        logger.log("INFO", "Interactive mode enabled", "🔄")
        interactive_inputs = prompt_interactive_inputs()
        show_interactive_summary(interactive_inputs)
        # Use the resolved movie ID from interactive inputs
        movie_id = interactive_inputs.get("movie_id")
    else:
        movie_id = args.movie_id
    source_url = args.url or interactive_inputs.get("url")
    movie_title = interactive_inputs.get("title")
    if source_url and not interactive_mode:
        source_url = ensure_tv_season_url(source_url, interactive=False)
    content_type = "tv" if source_url and is_tv_url(source_url) else "movie"
    review_type_arg = interactive_inputs.get("type", args.type)
    count = interactive_inputs.get("count", args.count)
    output_base = interactive_inputs.get("output", args.output)
    output_format = interactive_inputs.get("format", args.format)

    logger.log("INFO", f"Content type: {content_type.upper()}", "📺")
    logger.log("INFO", f"Review types: {review_type_arg}", "📝")
    logger.log("INFO", f"Request count: {count} per type", "🔢")
    logger.log("INFO", f"Output format: {output_format}", "💾")

    try:
        if not movie_id:
            logger.log("INFO", "Fetching movie ID from URL...", "🔍")
            movie_id, reviews_url, movie_title = resolve_movie_id_from_url(source_url)
            logger.log("SUCCESS", f"Content ID: {movie_id}", "🎯")
            if movie_title:
                logger.log("SUCCESS", f"Title: {movie_title}", "🎬")
            logger.log("INFO", f"Source: {reviews_url}", "📄")
        else:
            reviews_url = None
            logger.log("INFO", "Using provided movie ID", "🔢")

        review_types = [review_type_arg] if review_type_arg != "both" else ["audience", "critic"]
        reviews = {
            review_type: fetch_reviews(
                movie_id=movie_id,
                content_type=content_type,
                review_type=review_type,
                max_reviews=count,
                verified=args.verified if review_type == "audience" else False,
                top_only=args.top_only,
            )
            for review_type in review_types
        }
    except HTTPError as error:
        if error.code == 404:
            logger.log("ERROR", "Film/TV show has no rating on Rotten Tomatoes", "❌")
        else:
            logger.log("ERROR", f"HTTP Error {error.code}: {error.reason}", "❌")
        return
    except Exception as error:
        logger.log("ERROR", f"Error: {str(error)}", "❌")
        return

    result = {
        "movie_id": movie_id,
        "title": movie_title,
        "movie_url": source_url,
        "source_reviews_url": reviews_url,
        "requested_count": count,
        "review_types": review_types,
        "reviews": reviews,
    }

    if output_base:
        logger.log("INFO", "Saving results to file...", "💾")
        logger.log_with_spinner("Exporting data", 1.0)

        written_files = export_reviews(output_base, output_format, result)

        logger.log("SUCCESS", "Output files created:", "✅")
        for file_path in written_files:
            print(f"  {Colors.SUCCESS}✓ {file_path}{Style.RESET_ALL}")

        total_reviews = sum(len(reviews[rt]) for rt in reviews)
        logger.log("SUCCESS", f"Total reviews: {total_reviews}", "📊")

        # Show completion animation
        clear_screen()
        print(Colors.SUCCESS + COMPLETION_ART + Style.RESET_ALL)
        logger.log("HEADER", "✅ Scraping completed!", "🎬")
        logger.log("SUCCESS", "Thank you for using the Rotten Tomatoes Scraper!", "🍅")
        return

    # Print JSON result with colors
    logger.log("INFO", "Displaying results in JSON format:", "📋")
    print(json.dumps(result, indent=2, ensure_ascii=False))

    # Show completion message
    total_reviews = sum(len(reviews[rt]) for rt in reviews)
    clear_screen()
    print(Colors.SUCCESS + COMPLETION_ART + Style.RESET_ALL)
    logger.log("HEADER", "✅ Scraping completed!", "🎬")
    logger.log("SUCCESS", f"Total reviews: {total_reviews}", "📊")
    logger.log("SUCCESS", "Thank you for using the Rotten Tomatoes Scraper!", "🍅")


# ASCII art for completion
COMPLETION_ART = """
    🎬     🍅     ✨
   ( ●  ● )   ( ●  ● )  ( ●  ● )
  --------  --------  --------
   ROTTEN   TOMATOES   SCRAPER
   --------  --------  --------
        🎉   COMPLETED!   🎉
"""

if __name__ == "__main__":
    main()
