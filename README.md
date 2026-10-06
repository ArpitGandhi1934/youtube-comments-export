# Export YouTube Comments to CSV and Excel (Python + Node.js)

[![YouTube Scraper on Apify][badge]][store]

**Export YouTube comments** to **Excel (.xlsx), CSV or JSONL** for one video, a list of videos, a channel's latest uploads or a search: author, comment text, likes, reply count, date, replies if you want them, and the video each comment belongs to. The Python and Node.js scripts in this repo call the [YouTube Scraper][store] Actor on Apify, which **bills per video, not per comment**: 1,000 top comments from one video cost $0.002. No YouTube API key, no Google Cloud project, no quota.

Guide: [YouTube comments scraper, billed per video](https://yugenox-data.vercel.app/youtube/comments-scraper).

## What you get

Real output from one run on 2026-10-06: the top comments on [Rick Astley - Never Gonna Give You Up](https://www.youtube.com/watch?v=dQw4w9WgXcQ) with 3 replies each on the most-replied ones, from `--max-comments 1000 --sort top --replies 3`. The sample keeps the first 15 comments and their 18 replies; commenters are pseudonymized and text is cut to one sentence ([how](samples/redact.jq)).

| author | text | likes | reply_count | published | is_reply |
|---|---|---:|---:|---|---|
| commenter_01 | can confirm: he never gave us up | 324000 | 962 | 1 year ago | false |
| commenter_02 | YOUTUBE AND ONE LIKE WOOHAAAAH | 7500 | 0 | 1 year ago | true |
| commenter_05 | Gonna flag this for nudity so I can rick roll the YouTube staff | 567000 | 708 | 6 years ago | false |
| commenter_06 | yeah, it's big brain time | 63000 | 0 | 6 years ago | true |
| commenter_09 | My algebra teacher said that this was a link to the test answers | 68000 | 611 | 7 years ago | false |

The same 33 rows as [Excel](samples/rick-astley-top-comments.xlsx), [CSV](samples/rick-astley-top-comments.csv) and the Actor's [raw JSON row](samples/rick-astley-top-comments.json) (one row per video, comments nested inside). The full run returned 1,000 top comments and 60 replies for this video and cost $0.002.

| Column | Meaning |
|---|---|
| `author` | The commenter's YouTube handle, without the leading `@` |
| `text` | The comment text |
| `likes` | Likes as YouTube shows them: exact below 1,000, rounded above (`4.8M` becomes `4800000`). Empty when YouTube shows no number, which is how it shows a comment nobody has liked |
| `reply_count` | Replies the comment has on YouTube |
| `published` | YouTube's own date text, for example `6 years ago (edited)` |
| `published_approx` | The run date minus that age, as `YYYY-MM-DD` (`6 years ago` on 2026-10-06 is `2020-10-06`). Good for sorting, not an exact date; a real date cell in the .xlsx |
| `is_reply`, `parent_id` | `true` and the parent's `comment_id` for replies (`--replies`) |
| `author_is_creator` | `true` when the video's own channel wrote it |
| `is_pinned`, `is_hearted` | Pinned by the creator, hearted by the creator |
| `author_is_verified` | The commenter has a verified badge |
| `comment_id` | YouTube's comment id (`--append` uses it to skip comments you already have) |
| `video_id`, `video_title`, `video_url`, `channel_name` | The video the comment is on |

## Quick start

1. Create a free Apify account and copy your API token ([sign up][signup], then Console > Settings > API & Integrations).
2. `cp .env.example .env` and paste the token after `APIFY_TOKEN=`.

**Node.js** (20.6 or newer):

```bash
npm install
node --env-file=.env export.mjs "https://www.youtube.com/watch?v=dQw4w9WgXcQ" --max-comments 1000 --out comments.xlsx
node --env-file=.env export.mjs --file videos.example.txt --max-comments 200 --replies 3
node --env-file=.env export.mjs --channel @NASA --videos 20 --max-comments 200 --out nasa.csv
```

**Python** (3.10 or newer):

```bash
pip install -r requirements.txt
python export.py "https://www.youtube.com/watch?v=dQw4w9WgXcQ" --max-comments 1000 --out comments.xlsx
python export.py --search "standing desk review" --videos 50 --out reviews.csv
python export.py --file videos.txt --sort newest --max-comments 200 --append --out comments.csv
```

Quote video URLs in the shell: zsh (the macOS default) treats the `?` in `watch?v=` as a wildcard. Both scripts take the same options (`--help` prints them) and write byte-identical CSV and JSONL:

| Option | What it does | Default |
|---|---|---|
| video URLs | `watch?v=`, `youtu.be/` and `/shorts/` links or 11-character video ids. A playlist URL takes its first `--videos` videos, a channel URL its latest | |
| `--file videos.txt` | One URL per line; lines starting with `#` are skipped | |
| `--channel NAME` | A channel's latest videos, by name, `@handle` or channel URL. Repeat it for several | |
| `--search TERMS` | Videos that match a YouTube search. Repeat it for several | |
| `--videos N` | Videos per channel, playlist or search. Above 40, a search keeps paging until it has N | `10` |
| `--since WHEN` | Only channel videos published on or after a date (`2026-09-01`) or within a period (`7 days`). Older videos are skipped and not billed | |
| `--max-comments N` | Top-level comments per video, 1 to 5,000 | `100` |
| `--sort top\|newest` | YouTube's "Top comments" order, or newest first | `top` |
| `--replies N` | Also export up to N replies per comment, most-replied comments first (see [Limits](#limits)) | `0` |
| `--out FILE` | `.xlsx`, `.csv` or `.jsonl` | `comments.csv` |
| `--append` | Add only comments whose `comment_id` isn't in `--out` yet (`.csv` or `.jsonl`) | |
| `--max-charge USD` | Hard cost cap for the run, enforced by Apify. $2 covers about 1,000 videos; the script warns before a run that could need more | `2` |
| `--dataset ID` | Export the dataset of a run you already made in Apify Console. No new run starts | |
| `--from FILE` | Convert a JSON or JSONL file of the Actor's rows (Console export) or of this script's rows. Needs no token | |

After a run, each video gets one line: `1000 comments + 60 replies of 2,458,378 on YouTube  Rick Astley - ...`. The CSV is UTF-8 with a BOM so Excel shows emoji and accents correctly, and a cell that starts with `=`, `+`, `-`, `@`, a tab or a carriage return gets a leading `'` so a spreadsheet shows it as text instead of running it as a formula. The .xlsx stores every comment as plain text.

## Recipes

**A channel's last N videos.** `--channel @NASA --videos 50 --max-comments 200` reads the 50 newest uploads and up to 200 top comments on each: 50 videos, $0.10. `--since "30 days"` keeps only the videos from the last 30 days, and the older ones aren't billed.

**Everything people say about a product.** `--search "standing desk review" --videos 100 --max-comments 300` exports comments from the first 100 search results (above 40, the search keeps paging). Filter for questions and complaints in the spreadsheet: the Actor returns the top or newest comments, not keyword matches.

**Replies.** `--replies 5` adds up to 5 replies under each comment that has them, as rows with `is_reply = true` and the parent's id in `parent_id`. YouTube serves replies in pages of about 10, and the Actor spends at most 20 reply pages per video (about 200 replies), starting with the most-replied comments. With `--replies 3` each of those threads takes one page, so the test run got 3 replies on each of the 20 most-replied comments: 60 per video.

**Monitor new comments every day.** Newest first, appended to one file, on your latest videos:

```bash
# crontab -e   (every morning at 8)
0 8 * * * cd /path/to/youtube-comments-export && node --env-file=.env export.mjs --channel @yourchannel --videos 10 --sort newest --max-comments 200 --append --out comments.csv
```

That run bills 10 videos a day, $0.02, about $0.60 a month. `--append` skips comments already in the file, so each day adds only the new ones; raise `--max-comments` if a video gets more than 200 comments a day. To run it on Apify instead of your machine, save the input as a task in Apify Console, attach a daily Schedule, and convert each day's result:

```bash
curl -s "https://api.apify.com/v2/actor-tasks/<TASK_ID>/runs/last/dataset/items?status=SUCCEEDED&format=json" \
  -H "Authorization: Bearer $APIFY_TOKEN" -o last.json
node export.mjs --from last.json --append --out comments.csv
```

**No script at all.** One HTTP call starts a run, waits for it (up to 5 minutes) and returns one CSV line per top-level comment, with the video's `id` and `title` on each line. Apify's `unwind` option does the flattening; replies need the scripts:

```bash
curl -s -X POST "https://api.apify.com/v2/acts/yugenox~youtube-scraper/run-sync-get-dataset-items?format=csv&unwind=comments&fields=id,title,comments" \
  -H "Authorization: Bearer $APIFY_TOKEN" -H "Content-Type: application/json" \
  -d '{"startUrls":["https://www.youtube.com/watch?v=dQw4w9WgXcQ"],"maxComments":500}' -o comments.csv
```

**From Claude, ChatGPT or Cursor.** Apify's hosted MCP server can be pinned to this one Actor: `https://mcp.apify.com/?tools=fetch-actor-details,yugenox/youtube-scraper`. You sign in to your own Apify account with OAuth; never put a token in the URL. Setup per client: [yugenox-mcp](https://github.com/ArpitGandhi1934/yugenox-mcp) and the [guide](https://yugenox-data.vercel.app/ai/use-in-chatgpt-claude-cursor). n8n and Make can run the Actor with their Apify node.

## Price: per video, not per comment

The Actor charges $0.002 for each video row, on every Apify plan, and that row carries its comments and replies. There's no per-comment fee. Prices from the Actor's pricing in the Apify API on 2026-10-06 (in force since 2026-07-14); the [Store page][store] always shows the current ones.

| Job | Videos billed | Comments | Cost |
|---|---:|---:|---:|
| 1 video, 1,000 top comments | 1 | 1,000 | $0.002 |
| 20 videos, 1,000 top comments each | 20 | 20,000 | $0.04 |
| A channel's last 100 videos, 200 comments each | 100 | 20,000 | $0.20 |
| 1,000 videos, 100 comments each | 1,000 | 100,000 | $2.00 |

- Plus $0.00005 per run (the Actor start, per GB of memory; 1 GB is enough).
- A video with comments turned off is still one row ($0.002) with no comments.
- Not charged: an input that returns nothing (unknown channel, empty search, private or deleted video; the script prints these), inputs YouTube temporarily blocked, and channel videos older than `--since`.
- Apify lists a pricing update for this Actor from 2026-10-10 19:24 UTC. It adds one new event that these scripts never trigger; the per-video price stays $0.002.

Other YouTube comment Actors on the Apify Store bill per comment. From the Apify API on 2026-10-06 (each price in force since March 2026):

| Actor | Free | Starter | Scale | Business | 20 videos × 1,000 comments, Free | Same, Business |
|---|---:|---:|---:|---:|---:|---:|
| **yugenox/youtube-scraper** (this repo) | $0.002 per video | same | same | same | $0.04 | $0.04 |
| [streamers/youtube-comments-scraper][streamers] | $2.00 per 1,000 comments | $1.50 | $1.20 | $0.90 | $40.00 | $18.00 |
| [apidojo/youtube-comments-scraper][apidojo] | $0.50 per 1,000 comments + $0.001 per video | same | same | same | $10.02 | $10.02 |
| [api-ninja/youtube-comments-scraper][apininja] | $4.00 per 1,000 results + $0.01 per run | $0.40 | $0.40 | $0.40 | $80.01 | $8.01 |

The gap comes from the billing unit: it is widest with many comments per video and closes with few. At 5 comments per video, 1,000 videos cost $2.00 here, $10.00 on streamers' Free plan, $3.50 on apidojo's and $2.01 on api-ninja's on a paid plan. At 2 comments per video apidojo's costs the same as this one, and below that it costs less. If you need more than 5,000 comments from a single video, this Actor stops at 5,000: a per-comment Actor (try a small run first) or the YouTube Data API is the better fit.

## YouTube comments API: the Data API or this?

For comments alone, the official YouTube Data API is a good free option: [`commentThreads.list`](https://developers.google.com/youtube/v3/docs/commentThreads/list) returns up to 100 comment threads per call and costs 1 quota unit, from a default of 10,000 units a day ([quota costs](https://developers.google.com/youtube/v3/determine_quota_cost), Google's pages last updated 2026-09-14 and 2026-09-15). If you already have a Google Cloud project and an API key, and you know which videos you want, use it.

This repo is for the other cases: no Google Cloud project, key or quota; videos picked by channel name or search in the same run (the Data API gives `search.list` its own allowance of 100 calls a day); each video's views, likes and upload date in the same run's dataset; and a run you can schedule on Apify. For one video, now and then, on your own machine, the free [youtube-comment-downloader](https://pypi.org/project/youtube-comment-downloader/) library is a fine choice too.

## Limits

- **Up to 5,000 top-level comments per video** (`--max-comments`), in YouTube's "Top comments" order or newest first.
- **Fewer than you asked for** happens two ways. When YouTube stops handing out new comments first, the row is just shorter; in the test run below, one video stopped at 983 of 1,000. When paging stops early (a page YouTube wouldn't serve, or the run's timeout), the Actor sets `commentsTruncated` on that video and the script prints `[stopped early]` next to it.
- **Replies:** at most 20 reply pages per video, about 200 replies, most-replied comments first.
- **Dates are relative.** YouTube shows comment ages ("2 years ago"), not timestamps; `published_approx` is computed from that age.
- **No filters inside the Actor** for keywords, likes or dates of comments: it returns the top or newest N per video. Filter in your spreadsheet.
- **Public videos only.** Nothing here signs in to YouTube, so private videos and members-only content are out of reach.

Depth test, one run on 2026-10-06 (Actor build 0.1.14, `--max-comments 1000 --sort top --replies 3`): "Me at the zoo" (10,634,741 comments on YouTube), Rick Astley (2,458,378) and PSY's Gangnam Style (5,464,623) returned 1,000, 1,000 and 983 top-level comments plus 60 replies each, in 28 seconds, with no `commentsTruncated`. One run is one data point; run your own videos with a small `--max-comments` first.

## Privacy

Comments are public, but commenters' handles and comment text are personal data under PIPEDA (Canada), GDPR (EU), CCPA and similar laws. You are responsible for having a lawful reason to collect them, for how long you keep them and for what you do with them. Don't send commercial messages to commenters without consent (CASL, CAN-SPAM), and follow YouTube's terms. See [Is web scraping legal?][legal]

## More

- Guides on [yugenox-data.vercel.app](https://yugenox-data.vercel.app): [YouTube comments scraper, billed per video](https://yugenox-data.vercel.app/youtube/comments-scraper), [a channel's videos to CSV](https://yugenox-data.vercel.app/youtube/channel-videos-to-csv), [use it in ChatGPT, Claude and Cursor](https://yugenox-data.vercel.app/ai/use-in-chatgpt-claude-cursor)
- Search, channel, playlist, Shorts and comments inputs, with Node, Python, curl, CLI and Google Sheets examples: [youtube-instagram-scraper-examples](https://github.com/ArpitGandhi1934/youtube-instagram-scraper-examples)
- Estimated dislike counts in bulk, from [Return YouTube Dislike](https://returnyoutubedislike.com): [youtube-dislike-count-bulk](https://github.com/ArpitGandhi1934/youtube-dislike-count-bulk)
- Instagram comments to Excel: [instagram-comments-export](https://github.com/ArpitGandhi1934/instagram-comments-export) ([Instagram Comments Scraper][store-igc])
- The Actor itself, with its input form, output table and reviews: [YouTube Scraper on Apify][store]. Other ways to call it (Python and JS clients, OpenAPI, MCP) are on its [API page][api].

## Samples, testing and legal

- The sample comes from one run on 2026-10-06 (Actor build 0.1.14) through `export.mjs`; the .csv and .xlsx were then written from the pseudonymized JSON with `--from`, and `export.py` writes the same CSV byte for byte. Input keys checked against the Actor's live input schema the same day.
- Disclosure: the YouTube Scraper Actor these scripts call is built by the author of this repo.
- Not affiliated with or endorsed by YouTube or Google. YouTube is a trademark of Google LLC.
- MIT licensed. Made by Yugenox Corporation.

<!-- All apify.com links for this README live below. When the Apify affiliate id exists, append ?fpr=<id> to these URLs only. -->
[store]: https://apify.com/yugenox/youtube-scraper
[api]: https://apify.com/yugenox/youtube-scraper/api
[badge]: https://apify.com/actor-badge?actor=yugenox/youtube-scraper
[store-igc]: https://apify.com/yugenox/instagram-comments-scraper
[streamers]: https://apify.com/streamers/youtube-comments-scraper
[apidojo]: https://apify.com/apidojo/youtube-comments-scraper
[apininja]: https://apify.com/api-ninja/youtube-comments-scraper
[signup]: https://console.apify.com/sign-up
[legal]: https://blog.apify.com/is-web-scraping-legal/
