# sns-analysis-tool
*A modular command-line toolkit for scraping, cleaning, and analyzing social-media data.*

## reddit_scraper.py
Fetch comments from Reddit mentioning a product. Uses the Pushshift API and simple fuzzy matching so that slight variations of the product name are also captured.

### Usage
```bash
python reddit_scraper.py --query "Bravia Theatre Bar9" --limit 100 --outfile comments.csv
```

The resulting CSV will include comment ID, author, subreddit, timestamp, body text, and a permalink to the comment.
