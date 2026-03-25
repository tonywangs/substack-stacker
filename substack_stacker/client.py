import time
import requests
from tqdm import tqdm

ARCHIVE_URL = "https://{subdomain}.substack.com/api/v1/archive"
POST_URL = "https://{subdomain}.substack.com/api/v1/posts/{slug}"
PAGE_SIZE = 12
USER_AGENT = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
)


def _make_session():
    session = requests.Session()
    session.headers.update({
        "User-Agent": USER_AGENT,
        "Accept": "application/json",
    })
    return session


def _request_with_retry(session, url, params=None, max_retries=1, backoff=3):
    for attempt in range(max_retries + 1):
        try:
            resp = session.get(url, params=params, timeout=30)
            if resp.status_code == 429:
                if attempt < max_retries:
                    wait = backoff * (attempt + 1)
                    print(f"  Rate limited, waiting {wait}s...")
                    time.sleep(wait)
                    continue
                resp.raise_for_status()
            resp.raise_for_status()
            return resp
        except requests.RequestException:
            if attempt < max_retries:
                time.sleep(backoff)
                continue
            raise
    return None


def fetch_post_list(subdomain, limit=None, delay=1.5):
    session = _make_session()
    url = ARCHIVE_URL.format(subdomain=subdomain)
    posts = []
    offset = 0

    # First request to check if publication exists
    resp = _request_with_retry(session, url, params={
        "sort": "new", "offset": 0, "limit": PAGE_SIZE
    })
    if resp.status_code == 404:
        raise ValueError(f"Publication '{subdomain}' not found")

    batch = resp.json()
    if not batch:
        return []

    posts.extend(batch)
    offset += PAGE_SIZE

    pbar = tqdm(desc="Fetching post list", unit=" posts", initial=len(posts))

    while len(batch) == PAGE_SIZE and (limit is None or len(posts) < limit):
        time.sleep(delay)
        resp = _request_with_retry(session, url, params={
            "sort": "new", "offset": offset, "limit": PAGE_SIZE
        })
        batch = resp.json()
        posts.extend(batch)
        offset += PAGE_SIZE
        pbar.update(len(batch))

    pbar.update(0)
    pbar.close()

    if limit is not None:
        posts = posts[:limit]

    return posts


def fetch_post_body(session, subdomain, slug):
    url = POST_URL.format(subdomain=subdomain, slug=slug)
    try:
        resp = _request_with_retry(session, url)
        data = resp.json()
        return data.get("body_html", "")
    except requests.RequestException as e:
        print(f"  Warning: failed to fetch post '{slug}': {e}")
        return ""


def download_image(session, url):
    if url.startswith("//"):
        url = "https:" + url
    try:
        resp = session.get(url, timeout=15)
        resp.raise_for_status()
        if len(resp.content) > 10 * 1024 * 1024:  # skip >10MB
            return None
        return resp.content
    except requests.RequestException:
        return None
