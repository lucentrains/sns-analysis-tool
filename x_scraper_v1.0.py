#!/usr/bin/env python3
"""
twitter_scraper.py – Fetch tweets mentioning a keyword and save them for analysis.

Features
========
* Search the Twitter/X API v2 for tweets containing any keyword or advanced query string.
* Works with **recent search** (last 7 days) or **full-archive search** if your account has academic/enterprise access.
* Captures useful analytic fields in one go:
    • tweet_id, conversation_id, created_at, text (cleaned), language
    • author_id, username, verified, followers_count
    • like_count, retweet_count, reply_count, quote_count
    • possibly referenced_tweets + source if desired
* Handles pagination & rate-limit back-off automatically.
* Saves data to CSV (default), JSON, or Parquet for downstream work in Pandas.

Quick start
===========
1. ``pip install tweepy pandas python-dotenv tqdm``
2. Create an env file or export ``TWITTER_BEARER_TOKEN=<your-token>``.
3. ``python twitter_scraper.py --query "iPhone 16" --start "2025-05-12" --end "2025-05-19" --outfile iphone16.csv``

See ``python twitter_scraper.py -h`` for all options.
"""

import argparse
import csv
import os
import sys
import time
from datetime import datetime, timezone, timedelta
from typing import Iterable, List, Dict, Any

import tweepy
import pandas as pd
from tqdm import tqdm

# -----------------------------------------------------------------------------
# Helpers & Setup
# -----------------------------------------------------------------------------

def load_client(bearer_token: str | None = None) -> tweepy.Client:
    """Return an authenticated Tweepy v2 Client."""
    token = bearer_token or os.getenv("TWITTER_BEARER_TOKEN")
    if not token:
        sys.exit("Error: bearer token not supplied (set TWITTER_BEARER_TOKEN env var)")
    return tweepy.Client(bearer_token=token, wait_on_rate_limit=True)


def parse_date(date_str: str | None) -> str | None:
    """Return ISO-8601 for the Twitter API given YYYY-MM-DD[THH:MM:SS[Z]]."""
    if date_str is None:
        return None
    try:
        # Allow plain date (midnight UTC) or already-formatted ISO.
        if len(date_str) == 10:
            dt = datetime.strptime(date_str, "%Y-%m-%d").replace(tzinfo=timezone.utc)
        else:
            dt = datetime.fromisoformat(date_str.replace("Z", "+00:00"))
        return dt.isoformat().replace("+00:00", "Z")
    except ValueError as e:
        sys.exit(f"Invalid date format: {e}")


# -----------------------------------------------------------------------------
# Core fetch logic
# -----------------------------------------------------------------------------

def fetch_tweets(
    client: tweepy.Client,
    query: str,
    start_time: str | None = None,
    end_time: str | None = None,
    max_results: int = 100,
    limit: int | None = None,
) -> List[Dict[str, Any]]:
    """Fetch tweets matching *query*.

    If *limit* is None we retrieve everything until the API is exhausted.
    """

    tweet_fields = [
        "id",
        "text",
        "author_id",
        "conversation_id",
        "created_at",
        "public_metrics",
        "lang",
        "source",
    ]
    user_fields = [
        "id",
        "name",
        "username",
        "verified",
        "public_metrics",
    ]

    paginator = tweepy.Paginator(
        client.search_recent_tweets,
        query=query,
        start_time=start_time,
        end_time=end_time,
        tweet_fields=tweet_fields,
        user_fields=user_fields,
        expansions=["author_id"],
        max_results=max_results,
    )

    # Build a mapping for author lookup.
    records: List[Dict[str, Any]] = []
    pbar = tqdm(total=limit if limit else None, desc="Tweets")

    try:
        for page in paginator:
            if page.data is None:
                break
            users = {u.id: u for u in page.includes.get("users", [])}
            for t in page.data:
                user = users.get(t.author_id)
                data = {
                    "tweet_id": t.id,
                    "conversation_id": t.conversation_id,
                    "created_at": t.created_at,
                    "text": t.text.replace("\n", " ").strip(),
                    "lang": t.lang,
                    "author_id": t.author_id,
                    "username": user.username if user else None,
                    "display_name": user.name if user else None,
                    "verified": user.verified if user else None,
                    "followers_count": user.public_metrics.get("followers_count") if user else None,
                    "following_count": user.public_metrics.get("following_count") if user else None,
                    "tweet_count": user.public_metrics.get("tweet_count") if user else None,
                    "like_count": t.public_metrics.get("like_count"),
                    "retweet_count": t.public_metrics.get("retweet_count"),
                    "reply_count": t.public_metrics.get("reply_count"),
                    "quote_count": t.public_metrics.get("quote_count"),
                    "source": t.source,
                }
                records.append(data)
                pbar.update(1)
                if limit and len(records) >= limit:
                    pbar.close()
                    return records
    finally:
        pbar.close()
    return records


# -----------------------------------------------------------------------------
# CLI
# -----------------------------------------------------------------------------

def build_argparser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="Fetch tweets mentioning a keyword for analysis.")
    p.add_argument("--query", required=True, help="Search query or keyword (use quotes)")
    p.add_argument("--start", help="Start time YYYY-MM-DD or ISO-8601 (UTC)")
    p.add_argument("--end", help="End time YYYY-MM-DD or ISO-8601 (UTC)")
    p.add_argument("--outfile", default="tweets.csv", help="Output file path (CSV/JSON/parquet)")
    p.add_argument("--limit", type=int, help="Max tweets to fetch (default: all available)")
    p.add_argument("--max-results", type=int, default=100, choices=range(10,101), metavar="[10-100]", help="Results per request (API max 100)")
    p.add_argument("--bearer-token", help="Override BEARER token env var")
    return p


def write_output(records: List[Dict[str, Any]], path: str) -> None:
    ext = os.path.splitext(path)[1].lower()
    df = pd.DataFrame.from_records(records)
    if ext in {".parquet", ".pq"}:
        df.to_parquet(path, index=False)
    elif ext == ".json":
        df.to_json(path, orient="records", lines=True, force_ascii=False)
    else:
        df.to_csv(path, index=False, quoting=csv.QUOTE_NONNUMERIC)
    print(f"Saved {len(df)} rows → {path}")


def main():
    args = build_argparser().parse_args()

    client = load_client(args.bearer_token)
    records = fetch_tweets(
        client,
        query=args.query,
        start_time=parse_date(args.start),
        end_time=parse_date(args.end),
        max_results=args.max_results,
        limit=args.limit,
    )
    if not records:
        print("No tweets found for that query.")
        return
    write_output(records, args.outfile)


if __name__ == "__main__":
    main()
