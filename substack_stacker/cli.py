import argparse
import re
import sys

from substack_stacker.client import fetch_post_list
from substack_stacker.builder import build_document


def extract_subdomain(url):
    url = url.strip().rstrip("/")

    # Handle full URLs: https://example.substack.com/...
    match = re.match(r"https?://([a-zA-Z0-9-]+)\.substack\.com", url)
    if match:
        return match.group(1).lower()

    # Handle bare substack.com URLs without protocol
    match = re.match(r"([a-zA-Z0-9-]+)\.substack\.com", url)
    if match:
        return match.group(1).lower()

    # Handle bare subdomain (no dots)
    if re.match(r"^[a-zA-Z0-9-]+$", url):
        return url.lower()

    raise ValueError(
        f"Could not parse '{url}'. Use a Substack URL "
        "(e.g., https://example.substack.com) or just the subdomain name (e.g., example)."
    )


def main():
    parser = argparse.ArgumentParser(
        prog="substack-stacker",
        description="Compile a Substack publication into a single .docx book.",
    )
    parser.add_argument(
        "url",
        help="Substack URL or subdomain name (e.g., https://example.substack.com or just 'example')",
    )
    parser.add_argument(
        "-o", "--output",
        help="Output filename (default: {subdomain}_posts.docx)",
    )
    parser.add_argument(
        "--oldest-first",
        action="store_true",
        help="Order posts chronologically (oldest first)",
    )
    parser.add_argument(
        "-n", "--limit",
        type=int,
        default=None,
        help="Maximum number of posts to fetch",
    )
    parser.add_argument(
        "--delay",
        type=float,
        default=1.5,
        help="Seconds to wait between API requests (default: 1.5)",
    )

    args = parser.parse_args()

    try:
        subdomain = extract_subdomain(args.url)
    except ValueError as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)

    output = args.output or f"{subdomain}_posts.docx"

    print(f"Fetching posts from {subdomain}.substack.com...")

    try:
        posts = fetch_post_list(subdomain, limit=args.limit, delay=args.delay)
    except Exception as e:
        print(f"Error fetching posts: {e}", file=sys.stderr)
        sys.exit(1)

    if not posts:
        print("No posts found.")
        sys.exit(0)

    print(f"Found {len(posts)} posts.")

    if args.oldest_first:
        posts.reverse()

    build_document(subdomain, posts, output, delay=args.delay)


if __name__ == "__main__":
    main()
