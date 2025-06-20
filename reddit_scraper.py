import argparse
import csv
import time
import pandas as pd
import re
from typing import List, Dict, Tuple
import praw

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

def fetch_comments_praw(query: str, limit: int, reddit) -> List[Dict[str, str]]:
    tokens, token_patterns = prepare_tokens(query)
    comments: List[Dict[str, str]] = []
    for comment in reddit.subreddit("all").search(query, sort="new", syntax="lucene", limit=limit*2):
        # commentはSubmission型なので、コメントを取得
        comment.comments.replace_more(limit=0)
        for c in comment.comments.list():
            body = c.body
            if token_match_ratio(body.lower(), token_patterns) >= 0.7:
                comments.append(
                    {
                        "comment_id": c.id,
                        "author": str(c.author),
                        "subreddit": c.subreddit.display_name,
                        "created_utc": int(c.created_utc),
                        "body": body.replace("\n", " ").strip(),
                        "permalink": f"https://www.reddit.com{c.permalink}",
                    }
                )
                if len(comments) >= limit:
                    return comments
        if len(comments) >= limit:
            break
    return comments

def write_output(records: List[Dict[str, str]], path: str) -> None:
    df = pd.DataFrame.from_records(records)
    df.to_csv(path, index=False, quoting=csv.QUOTE_NONNUMERIC)
    print(f"Saved {len(df)} rows -> {path}")

def main() -> None:
    ap = argparse.ArgumentParser(description="Fetch Reddit comments mentioning a product (using Reddit official API)")
    ap.add_argument("--query", required=True, help="Product name to search for")
    ap.add_argument("--limit", type=int, default=1000, help="Max comments to fetch")
    ap.add_argument("--outfile", default="comments.csv", help="Output CSV file")
    ap.add_argument("--client_id", required=True, help="Reddit API client_id")
    ap.add_argument("--client_secret", required=True, help="Reddit API client_secret")
    ap.add_argument("--user_agent", required=True, help="Reddit API user_agent")
    args = ap.parse_args()

    reddit = praw.Reddit(
        client_id=args.client_id,
        client_secret=args.client_secret,
        user_agent=args.user_agent,
    )

    comments = fetch_comments_praw(args.query, limit=args.limit, reddit=reddit)
    if not comments:
        print("No comments found for that query.")
        return
    write_output(comments, args.outfile)

if __name__ == "__main__":
    main()
