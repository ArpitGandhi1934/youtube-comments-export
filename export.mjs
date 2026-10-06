#!/usr/bin/env node
// Export YouTube comments to CSV, Excel (.xlsx) or JSONL.
//
// Calls the yugenox/youtube-scraper Actor on Apify for video URLs, a channel's latest videos or a
// search, then writes one row per comment (and per reply, with --replies): author, text, likes,
// reply count, date and the video it belongs to. The Actor bills per video, not per comment:
// $0.002 per video whether it carries 10 comments or 5,000.
//
// Run with --help for every option. Requires Node 20.6+; APIFY_TOKEN in the environment (or in
// .env, loaded with `node --env-file=.env`) for new runs and --dataset.

import { appendFileSync, existsSync, readFileSync, statSync, writeFileSync } from 'node:fs';

const ACTOR = 'yugenox/youtube-scraper';
const PRICE_PER_VIDEO = 0.002; // USD per video row, every Apify plan (checked 2026-10-06)
const START_FEE = 0.00005; // USD per run start (1 GB of memory)

const HELP = `Export YouTube comments to CSV, Excel (.xlsx) or JSONL.

Usage:
  node --env-file=.env export.mjs <video URL>... [options]
  node --env-file=.env export.mjs --file videos.txt --out comments.xlsx
  node --env-file=.env export.mjs --channel @NASA --videos 20 --max-comments 200
  node --env-file=.env export.mjs --search "standing desk review" --videos 50
  node --env-file=.env export.mjs --dataset <datasetId> --out comments.xlsx
  node export.mjs --from rows.json --out comments.xlsx

Which videos
  <URL>...              Video URLs (watch?v=, youtu.be/, /shorts/) or 11-character video ids.
                        A playlist URL takes its first --videos videos, a channel URL its latest.
  --file <path>         One URL per line. Blank lines and lines starting with # are skipped.
  --channel <name>      A channel's latest videos: name, @handle or channel URL. Repeat it
                        for several channels.
  --search <terms>      Videos that match a YouTube search. Repeat it for several searches.
  --videos <n>          Videos per channel, playlist or search (default 10). Above 40, a
                        search keeps paging until it has n videos.
  --since <when>        Channel videos published on or after a date (2026-09-01) or within a
                        period ("7 days", "24 hours"). Older videos are skipped and not billed.

Comments
  --max-comments <n>    Top-level comments per video, 1 to 5000 (default 100).
  --sort <top|newest>   YouTube's "Top comments" order (default) or newest first.
  --replies <n>         Also export up to n replies per comment (default 0). The most-replied
                        comments go first, up to about 200 replies per video in total.

Output
  --out <path>          .csv (default comments.csv), .xlsx or .jsonl
  --append              Add only comments whose comment_id is not in --out yet (.csv or
                        .jsonl). For a daily export into one growing file.
  --max-charge <usd>    Hard cost cap for the run, enforced by Apify (default 2, about
                        1,000 videos).

Re-export without a new run
  --dataset <id>        Export the dataset of a run you already made (Apify Console >
                        Storage > Datasets). No new run starts.
  --from <file>         Convert a JSON or JSONL file of the Actor's rows (Console export)
                        or of this script's rows. Needs no token.

  -h, --help            Show this help.

Price: $0.002 per video ($2.00 per 1,000 videos) on every Apify plan, with its comments and
replies included, plus $0.00005 per run (checked 2026-10-06).`;

// Column order for CSV, XLSX and JSONL.
const COLUMNS = ['author', 'text', 'likes', 'reply_count', 'published', 'published_approx', 'is_reply',
    'parent_id', 'author_is_creator', 'is_pinned', 'is_hearted', 'author_is_verified', 'comment_id',
    'video_id', 'video_title', 'video_url', 'channel_name'];

function parseArgs(argv) {
    const o = { urls: [], channels: [], searches: [], videos: 10, maxComments: 100, replies: 0,
        out: 'comments.csv', maxCharge: 2, append: false };
    const int = (flag, v, min, max) => {
        const n = Number(v);
        if (v === undefined || !Number.isInteger(n) || n < min || n > max) {
            throw new Error(`${flag} needs a whole number from ${min} to ${max}`);
        }
        return n;
    };
    const val = (flag, v) => {
        if (v === undefined || v.startsWith('--')) throw new Error(`${flag} needs a value`);
        return v;
    };
    for (let i = 0; i < argv.length; i++) {
        const a = argv[i];
        if (a === '-h' || a === '--help') o.help = true;
        else if (a === '--file') o.urls.push(...readList(val(a, argv[++i])));
        else if (a === '--channel') o.channels.push(val(a, argv[++i]));
        else if (a === '--search') o.searches.push(val(a, argv[++i]));
        else if (a === '--videos') o.videos = int(a, argv[++i], 1, 20000);
        else if (a === '--since') o.since = val(a, argv[++i]);
        else if (a === '--max-comments') o.maxComments = int(a, argv[++i], 1, 5000);
        else if (a === '--sort') o.sort = val(a, argv[++i]);
        else if (a === '--replies') o.replies = int(a, argv[++i], 0, 500);
        else if (a === '--out') o.out = val(a, argv[++i]);
        else if (a === '--append') o.append = true;
        else if (a === '--max-charge') {
            o.maxCharge = Number(argv[++i]);
            if (!(o.maxCharge > 0)) throw new Error('--max-charge needs an amount in USD, for example 2');
        } else if (a === '--dataset') o.dataset = val(a, argv[++i]);
        else if (a === '--from') o.from = val(a, argv[++i]);
        else if (a.startsWith('-')) throw new Error(`Unknown option ${a} (see --help)`);
        else o.urls.push(a);
    }
    if (o.help) return o;
    if (o.sort && !['top', 'newest'].includes(o.sort)) throw new Error('--sort is top or newest');
    if (!/\.(csv|xlsx|jsonl)$/i.test(o.out)) throw new Error('--out must end in .csv, .xlsx or .jsonl');
    if (o.append && /\.xlsx$/i.test(o.out)) throw new Error('--append works with .csv or .jsonl (Excel opens the CSV)');
    const fresh = o.urls.length + o.channels.length + o.searches.length > 0;
    const sources = [fresh, Boolean(o.dataset), Boolean(o.from)].filter(Boolean).length;
    if (sources === 0) throw new Error('Give video URLs, --file, --channel, --search, --dataset <id> or --from <file> (see --help)');
    if (sources > 1) throw new Error('Use one of: URLs/--file/--channel/--search, --dataset, --from');
    return o;
}

function readList(path) {
    return readFileSync(path, 'utf8').split('\n').map((s) => s.trim()).filter((s) => s && !s.startsWith('#'));
}

// A single video (counts as one) or a list (a playlist or channel URL, up to --videos each).
const isSingleVideo = (s) => /^[A-Za-z0-9_-]{11}$/.test(s) || /youtu\.be\/|[?&]v=|\/shorts\/|\/embed\//.test(s);

// The Actor's comment author is the @handle; the leading @ is dropped so that a CSV cell can't
// start with it (Excel can read a cell that starts with @ as a formula).
const handle = (a) => (typeof a === 'string' ? a.replace(/^@/, '') : null);

function commentRow(c, video, parentId) {
    return {
        author: handle(c.author),
        text: c.text ?? null,
        likes: c.likes ?? null,
        reply_count: c.replyCount ?? null,
        published: c.publishedTime ?? null,
        published_approx: typeof c.publishedAtApprox === 'string' ? c.publishedAtApprox.slice(0, 10) : null,
        is_reply: parentId != null,
        parent_id: parentId ?? null,
        author_is_creator: c.authorIsCreator ?? null,
        is_pinned: c.isPinned ?? null,
        is_hearted: c.isHearted ?? null,
        author_is_verified: c.authorIsVerified ?? null,
        comment_id: c.commentId ?? null,
        video_id: video.id ?? null,
        video_title: video.title ?? null,
        video_url: video.url ?? null,
        channel_name: video.channelName ?? null,
    };
}

// One Actor row (a video with its comments) -> one output row per comment and per reply.
// A row this script wrote earlier passes through unchanged. Other rows (errors) give nothing.
function toRows(r) {
    if ('comment_id' in r) return [Object.fromEntries(COLUMNS.map((c) => [c, r[c] ?? null]))];
    if (!Array.isArray(r.comments)) return [];
    const rows = [];
    for (const c of r.comments) {
        rows.push(commentRow(c, r, null));
        for (const reply of c.replies ?? []) rows.push(commentRow(reply, r, c.commentId ?? null));
    }
    return rows;
}

// CSV cell. Text that starts with = + - @, a tab or a carriage return gets a leading ' so Excel and
// Google Sheets show it as text instead of running it as a formula (comments are written by strangers).
const GUARDED = new Set(['author', 'text', 'video_title', 'channel_name']);
function csvCell(v, col) {
    if (v == null) return '';
    let s = String(v);
    if (GUARDED.has(col) && /^[=+\-@\t\r]/.test(s)) s = `'${s}`;
    return /[",\n\r]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
}

// Minimal RFC 4180 reader, enough to find the comment_id column of a CSV this script wrote.
function parseCsv(text) {
    const rows = [];
    let row = [], cell = '', q = false;
    for (let i = 0; i < text.length; i++) {
        const ch = text[i];
        if (q) {
            if (ch === '"' && text[i + 1] === '"') { cell += '"'; i++; }
            else if (ch === '"') q = false;
            else cell += ch;
        } else if (ch === '"') q = true;
        else if (ch === ',') { row.push(cell); cell = ''; }
        else if (ch === '\n' || ch === '\r') {
            if (ch === '\r' && text[i + 1] === '\n') i++;
            row.push(cell); rows.push(row); row = []; cell = '';
        } else cell += ch;
    }
    if (cell || row.length) { row.push(cell); rows.push(row); }
    return rows;
}

function existingIds(path) {
    if (!existsSync(path) || statSync(path).size === 0) return new Set();
    const text = readFileSync(path, 'utf8').replace(/^﻿/, '');
    if (/\.jsonl$/i.test(path)) {
        return new Set(text.split('\n').filter(Boolean).map((l) => JSON.parse(l).comment_id).filter(Boolean));
    }
    const [header, ...rows] = parseCsv(text);
    if (header.join(',') !== COLUMNS.join(',')) {
        throw new Error(`${path} has different columns than this script writes; use a new --out file`);
    }
    const idx = COLUMNS.indexOf('comment_id');
    return new Set(rows.map((r) => r[idx]).filter(Boolean));
}

async function writeXlsx(path, rows) {
    // Loaded only for .xlsx so CSV/JSONL users don't need it.
    const { default: writeExcelFile } = await import('write-excel-file/node');
    // Strip characters that are not allowed in XLSX (XML) cells.
    const clean = (s) => (s == null ? null : String(s).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, ''));
    const str = (v, extra = {}) => (v == null ? null : { value: clean(v), type: String, ...extra });
    const n = (v) => (v == null ? null : { value: Number(v), type: Number });
    const bool = (v) => (v == null ? null : { value: Boolean(v), type: Boolean });
    const bold = (value) => ({ value, fontWeight: 'bold' });
    const data = [
        COLUMNS.map(bold),
        ...rows.map((r) => [
            str(r.author),
            str(r.text, { wrap: true }),
            n(r.likes),
            n(r.reply_count),
            str(r.published),
            // Midnight UTC of the approximate day; Excel has no time zones.
            r.published_approx ? { value: new Date(`${r.published_approx}T00:00:00Z`), type: Date, format: 'yyyy-mm-dd' } : null,
            bool(r.is_reply),
            str(r.parent_id, { format: '@' }),
            bool(r.author_is_creator),
            bool(r.is_pinned),
            bool(r.is_hearted),
            bool(r.author_is_verified),
            str(r.comment_id, { format: '@' }),
            str(r.video_id, { format: '@' }),
            str(r.video_title),
            str(r.video_url),
            str(r.channel_name),
        ]),
    ];
    const widths = [22, 70, 9, 11, 14, 15, 9, 28, 16, 9, 10, 17, 28, 13, 40, 44, 20];
    await writeExcelFile(data, { columns: widths.map((width) => ({ width })), stickyRowsCount: 1, sheet: 'Comments' })
        .toFile(path);
}

function writeOut(o, rows) {
    if (/\.xlsx$/i.test(o.out)) return writeXlsx(o.out, rows);
    if (/\.jsonl$/i.test(o.out)) {
        const body = rows.map((r) => JSON.stringify(r)).join('\n') + (rows.length ? '\n' : '');
        if (o.append) appendFileSync(o.out, body);
        else writeFileSync(o.out, body);
        return;
    }
    const lines = rows.map((r) => COLUMNS.map((c) => csvCell(r[c], c)).join(','));
    const fresh = !o.append || !existsSync(o.out) || statSync(o.out).size === 0;
    // UTF-8 with a BOM, so Excel shows emoji and accents correctly when it opens the file.
    if (fresh) writeFileSync(o.out, `﻿${[COLUMNS.join(','), ...lines].join('\r\n')}\r\n`);
    else if (lines.length) appendFileSync(o.out, `${lines.join('\r\n')}\r\n`);
}

function readFromFile(path) {
    const text = readFileSync(path, 'utf8').replace(/^﻿/, '').trim();
    if (!text) return [];
    const data = text.startsWith('[') ? JSON.parse(text) : text.split('\n').filter(Boolean).map((l) => JSON.parse(l));
    return Array.isArray(data) ? data : [data];
}

// One line per video: how many comments came back, YouTube's total, and whether paging stopped early.
function printVideoSummary(items) {
    for (const v of items) {
        if (v.error) {
            console.error(`  ${v.error}  ${v.fromYTUrl ?? v.url ?? v.input ?? ''}${v.note ? ` (${v.note})` : ''}`);
            continue;
        }
        if (!Array.isArray(v.comments)) continue;
        const replies = v.comments.reduce((s, c) => s + (c.replies?.length ?? 0), 0);
        const total = v.totalCommentsCount != null ? ` of ${v.totalCommentsCount.toLocaleString('en-US')} on YouTube` : '';
        const more = replies ? ` + ${replies} replies` : '';
        const early = v.commentsTruncated ? '  [stopped early]' : '';
        console.error(`  ${v.comments.length} comments${more}${total}  ${v.title ?? v.id}${early}`);
    }
}

async function printRunNotes(client, run) {
    if (run.statusMessage) console.error(`  ${run.statusMessage}`);
    // ERRORS (in the run's key-value store) lists inputs that returned nothing: an unknown
    // channel, an empty search, a private or deleted video. They are not charged.
    try {
        const rec = await client.keyValueStore(run.defaultKeyValueStoreId).getRecord('ERRORS');
        for (const e of Array.isArray(rec?.value) ? rec.value : []) {
            console.error(`  ${e.error ?? '?'}  ${e.input ?? e.url ?? ''}${e.note ? ` (${e.note})` : ''}, not charged`);
        }
    } catch { /* the record is a convenience; the export doesn't depend on it */ }
}

function estimateVideos(o) {
    const singles = o.urls.filter(isSingleVideo).length;
    const lists = o.urls.length - singles + o.channels.length + o.searches.length;
    return singles + lists * o.videos;
}

async function main() {
    const o = parseArgs(process.argv.slice(2));
    if (o.help) { console.log(HELP); return; }

    let items;
    if (o.from) {
        items = readFromFile(o.from);
    } else {
        if (!process.env.APIFY_TOKEN) throw new Error('Set APIFY_TOKEN (see .env.example), or use --from <file>');
        const { ApifyClient } = await import('apify-client'); // loaded here so --help and --from work before npm install
        const client = new ApifyClient({ token: process.env.APIFY_TOKEN });
        let datasetId = o.dataset;
        if (!datasetId) {
            const input = { maxItems: o.videos, maxComments: o.maxComments };
            if (o.urls.length) input.startUrls = o.urls;
            if (o.channels.length) input.channels = o.channels;
            if (o.searches.length) input.searchTerms = o.searches;
            // A plain search stops at YouTube's first two result pages (about 40 videos).
            if (o.searches.length && o.videos > 40) input.deepSearch = true;
            if (o.since) input.publishedAfter = o.since;
            if (o.sort) input.commentsSortBy = o.sort;
            if (o.replies > 0) input.maxReplies = o.replies;

            const est = estimateVideos(o);
            const fits = Math.floor((o.maxCharge - START_FEE) / PRICE_PER_VIDEO);
            if (est > fits) {
                console.error(`Note: this can return up to ${est} videos ($${(est * PRICE_PER_VIDEO).toFixed(2)}); `
                    + `--max-charge ${o.maxCharge} stops the run at about ${fits}. Raise --max-charge to get them all.`);
            }
            const what = [o.urls.length && `${o.urls.length} URL(s)`, o.channels.length && `${o.channels.length} channel(s)`,
                o.searches.length && `${o.searches.length} search(es)`].filter(Boolean).join(', ');
            console.error(`Running ${ACTOR} on ${what}, up to ${o.maxComments} comments per video ...`);
            // maxTotalChargeUsd caps what this run can cost you (pay-per-event Actors only).
            const run = await client.actor(ACTOR).call(input, { maxTotalChargeUsd: o.maxCharge });
            console.error(`Run ${run.id} finished: ${run.status}`);
            await printRunNotes(client, run);
            datasetId = run.defaultDatasetId;
        }
        items = [];
        for await (const it of client.dataset(datasetId).listItems()) items.push(it);
    }

    printVideoSummary(items);
    let rows = items.flatMap(toRows);
    let skipped = 0;
    if (o.append) {
        const seen = existingIds(o.out);
        const before = rows.length;
        rows = rows.filter((r) => !r.comment_id || !seen.has(r.comment_id));
        skipped = before - rows.length;
    }
    await writeOut(o, rows);

    const verb = o.append ? 'Appended' : 'Wrote';
    console.error(`${verb} ${rows.length} comments to ${o.out}${skipped ? ` (${skipped} already there)` : ''}`);
}

main().catch((err) => {
    console.error(err.message);
    process.exit(1);
});
