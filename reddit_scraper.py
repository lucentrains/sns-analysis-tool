import argparse
import csv
import time
import requests
import pandas as pd
import re
from typing import List, Dict, Tuple


def prepare_tokens(query: str) -> Tuple[List[str], Dict[str, List[re.Pattern]]]:
    """Return token list and compiled regex patterns for fuzzy matching."""
    base_tokens = re.findall(r"\w+", query.lower())
    patterns: Dict[str, List[re.Pattern]] = {}
    for tok in base_tokens:
        variations = {tok, tok.replace("-", " "), tok.replace(" ", "")}
        if re.search(r"\d", tok):
            variations.add(re.sub(r"(\D)(\d+)", r"\1 \2", tok))
        patterns[tok] = [re.compile(r"\b" + re.escape(v) + r"\b", re.I) for v in variations if v]
    return base_tokens, patterns


def token_match_ratio(text: str, patterns: Dict[str, List[re.Pattern]]) -> float:
    """Return the fraction of tokens that appear in *text* using regex patterns."""
    count = 0
    for pats in patterns.values():
        if any(p.search(text) for p in pats):
            count += 1
    return count / len(patterns) if patterns else 0.0


def fetch_comments(query: str, limit: int = 1000) -> List[Dict[str, str]]:
    tokens, token_patterns = prepare_tokens(query)
    search_q = "|".join(tokens)
    url = "https://api.pushshift.io/reddit/search/comment/"
    params = {"q": search_q, "size": 100, "sort": "desc"}
    comments: List[Dict[str, str]] = []
    last_ts = None
    while len(comments) < limit:
        if last_ts:
            params["before"] = last_ts
        r = requests.get(url, params=params, timeout=10)
        if r.status_code != 200:
            print("Request failed", r.status_code, r.text[:100])
            break
        data = r.json().get("data", [])
        if not data:
            break
        for c in data:
            last_ts = c.get("created_utc", 0) - 1
            body = c.get("body", "")
            if token_match_ratio(body.lower(), token_patterns) >= 0.7:
                comments.append(
                    {
                        "comment_id": c.get("id"),
                        "author": c.get("author"),
                        "subreddit": c.get("subreddit"),
                        "created_utc": c.get("created_utc"),
                        "body": body.replace("\n", " ").strip(),
                        "permalink": f"https://www.reddit.com{c.get('permalink')}" if c.get("permalink") else None,
                    }
                )
                if len(comments) >= limit:
                    break
        time.sleep(1)
    return comments


def write_output(records: List[Dict[str, str]], path: str) -> None:
    df = pd.DataFrame.from_records(records)
    df.to_csv(path, index=False, quoting=csv.QUOTE_NONNUMERIC)
    print(f"Saved {len(df)} rows -> {path}")


def main() -> None:
    ap = argparse.ArgumentParser(description="Fetch Reddit comments mentioning a product")
    ap.add_argument("--query", required=True, help="Product name to search for")
    ap.add_argument("--limit", type=int, default=1000, help="Max comments to fetch")
    ap.add_argument("--outfile", default="comments.csv", help="Output CSV file")
    args = ap.parse_args()

    comments = fetch_comments(args.query, limit=args.limit)
    if not comments:
        print("No comments found for that query.")
        return
    write_output(comments, args.outfile)


if __name__ == "__main__":
    main()
