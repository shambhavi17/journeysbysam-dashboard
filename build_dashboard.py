#!/usr/bin/env python3
"""Analyze Instagram scraper data and build an interactive HTML analytics dashboard."""
import os, json, re, statistics
from datetime import datetime
from collections import defaultdict

_HERE = os.path.dirname(os.path.abspath(__file__))


def _pick(stable, fallback):
    """Use the auto-refreshed file in data/ if it exists, else the original export."""
    return stable if os.path.exists(stable) else fallback


# refresh_data.py writes the data/ files; until then we fall back to the manual exports.
SRC = _pick(os.path.join(_HERE, "data", "instagram_account.json"),
            "/Users/shambhaviadhikari/Downloads/dataset_instagram-scraper_2026-05-21_03-42-23-714.json")
COMP_SRC = _pick(os.path.join(_HERE, "data", "instagram_competitors.json"),
                 "/Users/shambhaviadhikari/Downloads/dataset_instagram-post-scraper_2026-05-21_04-12-48-951.json")
OUT = os.path.join(_HERE, "content_analytics_dashboard.html")

data = json.load(open(SRC))

def g(p, k):
    return p.get(k) or 0

def hook_of(p):
    c = (p.get("caption") or "").strip()
    return c.split("\n")[0].strip() if c else "(no caption)"

# ---- per-post enrichment ----
posts = []
for p in data:
    likes, comments = g(p, "likesCount"), g(p, "commentsCount")
    plays = g(p, "videoPlayCount")
    ts = p.get("timestamp")
    dt = datetime.fromisoformat(ts.replace("Z", "+00:00")) if ts else None
    eng = likes + comments
    eng_rate = round(eng / plays * 100, 2) if plays else None
    posts.append({
        "shortCode": p.get("shortCode"),
        "url": p.get("url"),
        "type": p.get("type"),
        "hook": hook_of(p),
        "caption": (p.get("caption") or "")[:280],
        "likes": likes,
        "comments": comments,
        "plays": plays,
        "engagement": eng,
        "engRate": eng_rate,
        "duration": round(g(p, "videoDuration"), 1) if p.get("type") == "Video" else None,
        "location": p.get("locationName") or "—",
        "hashtags": p.get("hashtags", []),
        "hashtagCount": len(p.get("hashtags", [])),
        "date": dt.strftime("%Y-%m-%d") if dt else None,
        "day": dt.strftime("%A") if dt else None,
        "music": (p.get("musicInfo") or {}).get("song_name") or "—",
    })

vids = [p for p in posts if p["type"] == "Video"]
cars = [p for p in posts if p["type"] == "Sidecar"]

# ---- summary ----
total_plays = sum(p["plays"] for p in posts)
total_eng = sum(p["engagement"] for p in posts)
total_likes = sum(p["likes"] for p in posts)
total_comments = sum(p["comments"] for p in posts)
vid_rates = [p["engRate"] for p in vids if p["engRate"] is not None]
avg_eng_rate = round(statistics.mean(vid_rates), 2) if vid_rates else 0

summary = {
    "creator": "Sam Adhikari — journeysbysam",
    "niche": "Bucketlist Travel + Chicago Lifestyle",
    "totalPosts": len(posts),
    "videos": len(vids),
    "carousels": len(cars),
    "totalPlays": total_plays,
    "totalLikes": total_likes,
    "totalComments": total_comments,
    "totalEngagement": total_eng,
    "avgEngRate": avg_eng_rate,
    "dateRange": f"{min(p['date'] for p in posts if p['date'])} to {max(p['date'] for p in posts if p['date'])}",
}

# ---- rankings ----
top_reach = sorted(vids, key=lambda p: p["plays"], reverse=True)[:10]
top_engrate = sorted([p for p in vids if p["engRate"] is not None],
                     key=lambda p: p["engRate"], reverse=True)[:10]
top_eng_abs = sorted(posts, key=lambda p: p["engagement"], reverse=True)[:10]

# ---- timeline (chronological reach) ----
timeline = sorted([p for p in vids if p["date"]], key=lambda p: p["date"])
timeline_data = [{"date": p["date"], "plays": p["plays"], "hook": p["hook"][:40]} for p in timeline]

# ---- content type comparison ----
def avg(lst, k):
    vals = [x[k] for x in lst if x[k] is not None]
    return round(statistics.mean(vals)) if vals else 0

type_compare = {
    "Video (Reels)": {"count": len(vids), "avgLikes": avg(vids, "likes"),
                       "avgComments": avg(vids, "comments"), "avgPlays": avg(vids, "plays")},
    "Sidecar (Carousels)": {"count": len(cars), "avgLikes": avg(cars, "likes"),
                            "avgComments": avg(cars, "comments"), "avgPlays": 0},
}

# ---- day-of-week ----
DOW = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
day_stats = {}
for d in DOW:
    dp = [p for p in vids if p["day"] == d]
    day_stats[d] = {"count": len(dp), "avgPlays": avg(dp, "plays")}

# ---- duration buckets ----
buckets = {"<10s": [], "10-20s": [], "20-40s": [], "40s+": []}
for p in vids:
    dur = p["duration"]
    if dur is None:
        continue
    k = "<10s" if dur < 10 else "10-20s" if dur < 20 else "20-40s" if dur < 40 else "40s+"
    buckets[k].append(p["plays"])
duration_data = {k: {"count": len(v), "avgPlays": round(statistics.mean(v)) if v else 0}
                 for k, v in buckets.items()}

# ---- hashtags ----
htag = defaultdict(lambda: {"n": 0, "eng": 0, "plays": 0})
for p in posts:
    for h in p["hashtags"]:
        key = h.lower()
        htag[key]["n"] += 1
        htag[key]["eng"] += p["engagement"]
        htag[key]["plays"] += p["plays"]
top_hashtags = sorted(htag.items(), key=lambda x: x[1]["n"], reverse=True)[:14]
hashtag_data = [{"tag": h, "uses": v["n"], "avgEng": round(v["eng"] / v["n"]),
                 "avgPlays": round(v["plays"] / v["n"])} for h, v in top_hashtags]

# ---- hook patterns ----
def classify(hook):
    h = hook.lower()
    cats = []
    if "pov" in h:
        cats.append("POV")
    if re.search(r"\b\d+ (hotel|things|ways|spots|reasons|mistakes|places|cafes|stops|day)", h) or re.match(r"^[^a-z]*\d", h):
        cats.append("Numbered list")
    if re.search(r"\b(this|these)\b", h):
        cats.append("'This/These' tease")
    if "?" in hook:
        cats.append("Question")
    if "everyone" in h or "nobody" in h or "stop " in h:
        cats.append("Contrarian")
    if hook and not hook[0].isalnum():
        cats.append("Emoji-led")
    return cats or ["Plain statement"]

pat_perf = defaultdict(lambda: {"n": 0, "plays": 0})
for p in vids:
    for c in classify(p["hook"]):
        pat_perf[c]["n"] += 1
        pat_perf[c]["plays"] += p["plays"]
hook_patterns = sorted(
    [{"pattern": k, "count": v["n"], "avgPlays": round(v["plays"] / v["n"])}
     for k, v in pat_perf.items()],
    key=lambda x: x["avgPlays"], reverse=True)

# best hooks = hooks from top reach posts
best_hooks = [{"hook": p["hook"], "plays": p["plays"], "engRate": p["engRate"],
               "url": p["url"]} for p in top_reach[:6]]

# ---- viral insights (data-derived) ----
best_day = max(day_stats.items(), key=lambda x: x[1]["avgPlays"])
best_dur = max(duration_data.items(), key=lambda x: x[1]["avgPlays"])
best_pattern = hook_patterns[0]
vid_avg_plays = avg(vids, "plays")
median_plays = round(statistics.median([p["plays"] for p in vids]))
viral_threshold = vid_avg_plays * 2
viral_posts = [p for p in vids if p["plays"] >= viral_threshold]

insights = [
    {"title": "Reels vastly outreach carousels",
     "detail": f"Your {len(vids)} Reels average {vid_avg_plays:,} plays. Carousels generate engagement but no comparable reach signal — Reels are your growth engine. Lead with video."},
    {"title": f"Short hooks win — {best_dur[0]} is your sweet spot",
     "detail": f"Videos in the {best_dur[0]} bucket average {best_dur[1]['avgPlays']:,} plays vs a {vid_avg_plays:,} overall average. Tighten edits; front-load the payoff."},
    {"title": f"'{best_pattern['pattern']}' hooks perform best",
     "detail": f"Posts opening with a {best_pattern['pattern']} hook average {best_pattern['avgPlays']:,} plays across {best_pattern['count']} posts. Make it a default opener."},
    {"title": f"{best_day[0]} is your strongest posting day",
     "detail": f"{best_day[0]} posts average {best_day[1]['avgPlays']:,} plays. Friday and Wednesday also overperform — front-load mid/late week."},
    {"title": f"{len(viral_posts)} posts broke 2x your average",
     "detail": f"Outliers above {viral_threshold:,} plays cluster around weather-driven Chicago content and seasonal travel — timely + locally relevant beats evergreen here."},
    {"title": "Comments outperform on niche-precise posts",
     "detail": f"Your highest comment counts (100+) come from specific 'how to' Kyoto/Chicago posts. Specificity drives saves and conversation, not just reach."},
]

# ---- content ideas (tailored) ----
locs = [p["location"] for p in posts if p["location"] != "—"]
loc_freq = defaultdict(int)
for l in locs:
    loc_freq[l] += 1

ideas = [
    {"title": "POV: Chicago's coldest day, captured cinematically",
     "why": "Your top post (15.2K plays) was a winter-warning Reel. Weather-timed Chicago content is proven viral fuel — keep a cold-snap Reel ready to post same-day.",
     "format": "Reel · <10s · POV hook", "tag": "Seasonal / Local"},
    {"title": "3 Hotel Booking Mistakes — Tokyo edition",
     "why": "The 'X Mistakes to Avoid' Paris Reel hit 11.3K. The numbered-list mistake format travels well; replicate it for Tokyo, Kyoto, and Big Sur.",
     "format": "Reel · 10-20s · Numbered hook", "tag": "Listicle / Travel"},
    {"title": "Hidden Chicago cafés — Part 2",
     "why": "Your café-discovery Reel earned your highest like count (427). #ChicagoHiddenGems is your top-engagement hashtag. Turn it into a recurring series.",
     "format": "Reel · <10s · 'This' tease hook", "tag": "Series / Local"},
    {"title": "How locals actually do Kyoto — temple edition",
     "why": "The 'Everyone visits Kyoto's temples… but THIS is how' Reel drew 108 comments. Contrarian + insider framing maximizes conversation.",
     "format": "Reel · 10-20s · Contrarian hook", "tag": "Insider / Travel"},
    {"title": "Big Sur road trip — the 5 stops you can't skip",
     "why": "Big Sur / Highway 1 hashtags appear 5× each with strong engagement. Bundle scattered coastal clips into one definitive numbered guide.",
     "format": "Reel · 20-40s · Numbered hook", "tag": "Guide / Travel"},
    {"title": "Carousel → Reel: repurpose your top photo dumps",
     "why": "19 carousels carry engagement but no reach. Re-cut your best carousel sets as Reels with a hook overlay to unlock the algorithm's reach surface.",
     "format": "Reel · convert from Sidecar", "tag": "Repurpose"},
]

# ---- competitor benchmark ----
comp_data = json.load(open(COMP_SRC))
by_owner = defaultdict(list)
for p in comp_data:
    by_owner[p.get("ownerUsername")].append(p)


def _ext_metrics(plist):
    """Deeper comparison metrics that work for any raw post list."""
    likes = [g(p, "likesCount") for p in plist]
    caps = [len(p.get("caption") or "") for p in plist]
    med = statistics.median(likes) if likes else 0
    breakouts = sum(1 for l in likes if med and l >= 2 * med)
    return {
        "avgCaption": round(statistics.mean(caps)) if caps else 0,
        "breakoutRate": round(breakouts / len(plist) * 100) if plist else 0,
    }


comp_accounts = []
for handle, plist in by_owner.items():
    likes = [g(p, "likesCount") for p in plist]
    cmts = [g(p, "commentsCount") for p in plist]
    dts = sorted(datetime.fromisoformat(p["timestamp"].replace("Z", "+00:00"))
                 for p in plist if p.get("timestamp"))
    span = ((dts[-1] - dts[0]).days or 1) if len(dts) > 1 else 1
    htc = [len(re.findall(r"#\w+", p.get("caption") or "")) for p in plist]
    tops = sorted(plist, key=lambda p: g(p, "likesCount"), reverse=True)[:3]
    comp_accounts.append({
        "handle": handle,
        "name": plist[0].get("ownerFullName") or handle,
        "posts": len(plist),
        "avgLikes": round(statistics.mean(likes)),
        "medLikes": round(statistics.median(likes)),
        "maxLikes": max(likes),
        "avgComments": round(statistics.mean(cmts)),
        "perWeek": round(len(plist) / span * 7, 1),
        "avgHashtags": round(statistics.mean(htc), 1),
        "topHooks": [hook_of(p) for p in tops],
        **_ext_metrics(plist),
    })
comp_accounts.sort(key=lambda c: c["avgLikes"], reverse=True)

# your recent (last 60d) cadence for a fair comparison
your_dts = sorted(datetime.fromisoformat(p["timestamp"].replace("Z", "+00:00"))
                  for p in data if p.get("timestamp"))
recent_n = sum(1 for x in your_dts if (your_dts[-1] - x).days <= 60)
your_per_week = round(recent_n / 60 * 7, 1)

you_bench = {
    "handle": "journeysbysam (You)",
    "name": "Sam Adhikari",
    "posts": len(posts),
    "avgLikes": round(statistics.mean([p["likes"] for p in posts])),
    "medLikes": round(statistics.median([p["likes"] for p in posts])),
    "maxLikes": max(p["likes"] for p in posts),
    "avgComments": round(statistics.mean([p["comments"] for p in posts])),
    "perWeek": your_per_week,
    "avgHashtags": round(statistics.mean([p["hashtagCount"] for p in posts]), 1),
    "topHooks": [p["hook"] for p in top_eng_abs[:3]],
    "isYou": True,
}
you_bench.update(_ext_metrics(data))
benchmark = [you_bench] + comp_accounts

# ---- deeper competitor insights ----
_big = comp_accounts[0]                                    # highest avg likes
_chatty = max(comp_accounts, key=lambda c: c["avgComments"])  # most comments
comp_deep = [
    {"title": "Comment-gating beats raw reach",
     "detail": f"@{_chatty['handle']} averages {_chatty['avgComments']} comments a post — even though its {_chatty['avgLikes']:,} avg likes sit far below @{_big['handle']}'s {_big['avgLikes']:,}. Its '‼️ comment for the location' CTA manufactures comments. You average {you_bench['avgComments']}. Comment-gate every Reel."},
    {"title": "Caption length & Instagram SEO",
     "detail": f"Your captions average {you_bench['avgCaption']} characters vs {_big['avgCaption']} for @{_big['handle']} and {_chatty['avgCaption']} for @{_chatty['handle']}. Instagram now searches caption text — a keyword-rich first line and real context keep posts findable for weeks."},
    {"title": "Breakout rate — how often posts swing big",
     "detail": f"{you_bench['breakoutRate']}% of your posts beat 2× your own median; @{_big['handle']} hits {_big['breakoutRate']}%. Even big accounts mostly post 'normal' content — a few deliberate big swings move the average. Plan 1–2 a month."},
    {"title": "Hashtags per post",
     "detail": f"You use {you_bench['avgHashtags']} hashtags/post; @{_big['handle']} uses {_big['avgHashtags']}, @{_chatty['handle']} {_chatty['avgHashtags']}. With Instagram's 5-tag cap, make each one count — 1 branded + 4 niche/destination tags."},
    {"title": "The reach ceiling",
     "detail": f"Your best post hit {you_bench['maxLikes']:,} likes; @{_big['handle']}'s best hit {_big['maxLikes']:,}. That gap is topic + hook, not talent — broadly shareable, aspirational posts are how you lift your own ceiling."},
]

gap_analysis = [
    {"title": "Your engagement ceiling is hooks, not effort",
     "detail": f"Your posts average {you_bench['avgLikes']} likes vs {comp_accounts[-1]['avgLikes']:,} and {comp_accounts[0]['avgLikes']:,} for your competitors. Your recent cadence ({your_per_week}/wk) already matches them — the gap is hook strength and topic reach, not output."},
    {"title": "Your hooks describe; theirs provoke",
     "detail": "Your openers narrate the scene ('POV: Chicago still looked cinematic'). Both competitors lead with a promise or pattern-interrupt — 'The 10 BEST things to do', '‼️ nobody was rushing here'. Curiosity and superlatives beat description."},
    {"title": "You're not gating comments",
     "detail": f"tripuntraveled averages {[c for c in comp_accounts if c['handle']=='tripuntraveled'][0]['avgComments']} comments — more than the bigger account — by hiding location/details behind '‼️location\U0001f447'. You average {you_bench['avgComments']}. A comment CTA is your fastest engagement lever."},
    {"title": "Local content caps your reach",
     "detail": "Your biggest posts are Chicago-local. peekinourjournal's 219K and 118K hits are broad bucketlist dreams — Hawaii, African safari, Iceland. Local builds loyalty; aspirational destinations unlock reach."},
    {"title": "Sustain the cadence you just found",
     "detail": f"You've ramped to {your_per_week}/wk recently, but your full-year average is 1.1/wk. Competitors hold 2–4/wk consistently. Gaps reset algorithmic momentum — protect the streak."},
]

what_to_add = [
    {"title": "Superlative title hooks",
     "detail": "Open with 'The BEST', 'Top 10', 'The #1' + an emoji. peekinourjournal's '10 BEST things to do in Hawaii' is a repeatable reach machine.",
     "tag": "Hook"},
    {"title": "Pattern-interrupt openers",
     "detail": "Lead with ‼️ or \U0001f92f and an unfinished promise: 'nobody tells you this about…', 'the last clip \U0001f92f'. Stop the scroll before the viewer decides.",
     "tag": "Hook"},
    {"title": "Comment-gated CTAs",
     "detail": "End every Reel with 'Comment LOCATION and I'll send details'. This is tripuntraveled's core growth loop — it converts passive viewers into comments the algorithm rewards.",
     "tag": "Engagement"},
    {"title": "Aspirational bucketlist destinations",
     "detail": "Mix big-dream locations (Hawaii, Iceland, safari, Japan) into your Chicago feed. Keep Chicago for your loyal niche; use marquee trips for reach spikes.",
     "tag": "Topic"},
    {"title": "'Save this' listicle Reels",
     "detail": "Numbered, save-worthy guides drive saves — a stronger growth signal than likes. Your Paris 'mistakes' Reel already proved this; systematize the format.",
     "tag": "Format"},
    {"title": "Deliberate big-swing posts",
     "detail": f"Even peekinourjournal's median is only {[c for c in comp_accounts if c['handle']=='peekinourjournal'][0]['medLikes']:,} — two viral hits carry their average. Design 1–2 'big swing' Reels a month, don't only play safe.",
     "tag": "Strategy"},
]

# ---- assemble ----
payload = {
    "summary": summary,
    "topReach": top_reach,
    "topEngRate": top_engrate,
    "topEngAbs": top_eng_abs,
    "timeline": timeline_data,
    "typeCompare": type_compare,
    "dayStats": day_stats,
    "durationData": duration_data,
    "hashtagData": hashtag_data,
    "hookPatterns": hook_patterns,
    "bestHooks": best_hooks,
    "insights": insights,
    "ideas": ideas,
    "benchmark": benchmark,
    "gapAnalysis": gap_analysis,
    "whatToAdd": what_to_add,
    "compDeep": comp_deep,
}

J = json.dumps(payload)

import growth_planner
PA = growth_planner.planner_assets()

HTML = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>journeysbysam Insights</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
__PLANNER_HEAD__
<style>
:root{
  --bg:#f7f1e9; --panel:#fffdfa; --panel2:#f3ece1; --line:#ebe0d0;
  --text:#4f463c; --muted:#a99c8d; --accent:#c0879a; --accent2:#9bb6c4;
  --gold:#caa86f; --pink:#d99fa6; --green:#9aae93;
}
*{margin:0;padding:0;box-sizing:border-box}
body{background:var(--bg);color:var(--text);font-family:'Jost',-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;line-height:1.55;padding:0 0 80px}
.wrap{max-width:1180px;margin:0 auto;padding:0 24px}
header{padding:48px 0 32px;border-bottom:1px solid var(--line);margin-bottom:36px;
  background:radial-gradient(800px 300px at 20% 0%,rgba(192,135,154,.18),transparent),
             radial-gradient(700px 300px at 90% 0%,rgba(155,182,196,.16),transparent)}
.eyebrow{color:var(--accent);font-size:13px;font-weight:600;letter-spacing:.12em;text-transform:uppercase}
h1{font-family:'Cormorant Garamond',Georgia,serif;font-size:42px;font-weight:600;margin:8px 0 6px;letter-spacing:.01em}
.sub{color:var(--muted);font-size:15px}
.section-title{font-family:'Cormorant Garamond',Georgia,serif;font-size:27px;font-weight:600;margin:46px 0 18px;
  display:flex;align-items:center;gap:10px}
.section-title::before{content:"";width:4px;height:20px;background:linear-gradient(var(--accent),var(--accent2));border-radius:3px}
.section-desc{color:var(--muted);font-size:14px;margin:-10px 0 18px}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:14px;margin-top:26px}
.kpi{background:var(--panel);border:1px solid var(--line);border-radius:14px;padding:18px}
.kpi .val{font-size:27px;font-weight:700;letter-spacing:-.02em}
.kpi .lbl{color:var(--muted);font-size:12.5px;margin-top:3px}
.kpi.a .val{color:var(--accent)}.kpi.b .val{color:var(--accent2)}
.kpi.c .val{color:var(--gold)}.kpi.d .val{color:var(--pink)}.kpi.e .val{color:var(--green)}
.grid2{display:grid;grid-template-columns:1fr 1fr;gap:18px}
.grid3{display:grid;grid-template-columns:repeat(3,1fr);gap:16px}
.card{background:var(--panel);border:1px solid var(--line);border-radius:14px;padding:20px}
.card h3{font-size:15px;font-weight:600;margin-bottom:14px}
.chart-box{position:relative;height:260px}
.chart-box.tall{height:300px}
table{width:100%;border-collapse:collapse;font-size:13.5px}
th{text-align:left;color:var(--muted);font-weight:600;font-size:11.5px;text-transform:uppercase;
  letter-spacing:.05em;padding:10px 12px;border-bottom:1px solid var(--line)}
td{padding:12px;border-bottom:1px solid var(--line);vertical-align:top}
tr:last-child td{border-bottom:none}
tr:hover td{background:var(--panel2)}
.rank{font-weight:700;color:var(--accent);font-size:15px;width:34px}
.hook-cell{max-width:420px}
.hook-cell a{color:var(--text);text-decoration:none;font-weight:500}
.hook-cell a:hover{color:var(--accent2)}
.meta{color:var(--muted);font-size:11.5px;margin-top:3px}
.num{font-variant-numeric:tabular-nums;font-weight:600}
.pill{display:inline-block;padding:2px 9px;border-radius:20px;font-size:11px;font-weight:600}
.pill.v{background:rgba(192,135,154,.18);color:var(--accent)}
.pill.s{background:rgba(155,182,196,.22);color:var(--accent2)}
.tabs{display:flex;gap:8px;margin-bottom:16px;flex-wrap:wrap}
.tab{background:var(--panel2);border:1px solid var(--line);color:var(--muted);
  padding:7px 15px;border-radius:9px;cursor:pointer;font-size:13px;font-weight:600;transition:.15s}
.tab.active{background:var(--accent);color:#fff;border-color:var(--accent)}
.tab:hover:not(.active){color:var(--text)}
.hooks-grid{display:grid;grid-template-columns:1fr 1fr;gap:14px}
.hook-card{background:var(--panel);border:1px solid var(--line);border-left:3px solid var(--gold);
  border-radius:11px;padding:16px}
.hook-card .q{font-size:14.5px;font-weight:600;margin-bottom:9px}
.hook-card .s{display:flex;gap:16px;color:var(--muted);font-size:12px}
.hook-card .s b{color:var(--green)}
.insight{background:var(--panel);border:1px solid var(--line);border-radius:12px;padding:18px;
  display:flex;gap:14px}
.insight .ic{flex-shrink:0;width:36px;height:36px;border-radius:10px;
  background:linear-gradient(135deg,var(--accent),var(--accent2));display:flex;
  align-items:center;justify-content:center;font-weight:700;color:#fff}
.insight h4{font-size:14.5px;font-weight:600;margin-bottom:4px}
.insight p{color:var(--muted);font-size:13px}
.idea{background:var(--panel);border:1px solid var(--line);border-radius:13px;padding:18px;
  transition:.15s;cursor:default}
.idea:hover{border-color:var(--accent);transform:translateY(-2px)}
.idea .tag{font-size:10.5px;font-weight:700;letter-spacing:.05em;text-transform:uppercase;
  color:var(--accent2)}
.idea h4{font-size:15px;font-weight:650;margin:7px 0 8px}
.idea p{color:var(--muted);font-size:12.8px;margin-bottom:12px}
.idea .fmt{font-size:11.5px;color:var(--gold);font-weight:600;border-top:1px solid var(--line);
  padding-top:10px}
footer{color:var(--muted);font-size:12px;text-align:center;margin-top:60px;
  border-top:1px solid var(--line);padding-top:24px}
@media(max-width:840px){.grid2,.grid3,.hooks-grid{grid-template-columns:1fr}h1{font-size:27px}}
/* ===== top-level tab navigation ===== */
body{transition:background .3s ease}
.topnav{position:sticky;top:0;z-index:60;backdrop-filter:blur(12px);
  -webkit-backdrop-filter:blur(12px);border-bottom:1px solid;
  background:rgba(247,241,233,.93);border-color:#ebe0d0}
.topnav-inner{max-width:1200px;margin:0 auto;padding:11px 26px;display:flex;
  align-items:center;gap:16px}
.topnav .brand{font-family:'Cormorant Garamond',Georgia,serif;font-size:21px;font-weight:600;
  color:#4f463c;letter-spacing:.01em}
.topnav-tabs{display:flex;gap:8px;margin-left:auto}
.topnav-tab{font-family:'Jost',sans-serif;font-size:13px;font-weight:500;cursor:pointer;
  padding:8px 16px;border-radius:10px;border:1px solid transparent;background:transparent;
  color:#7c7065;transition:.15s}
.topnav-tab:hover{color:#4f463c}
.topnav-tab.active{background:linear-gradient(135deg,#c0879a,#b3a3d6);color:#fff}
.topnav-refresh{font-family:'Jost',sans-serif;font-size:12.5px;font-weight:500;cursor:pointer;
  text-decoration:none;white-space:nowrap;padding:7px 13px;border-radius:10px;
  border:1px solid #c0879a;color:#a96b80;background:transparent;transition:.15s}
.topnav-refresh:hover{background:#f4dfe1;color:#4f463c}
.pane{display:none}
.pane.active{display:block}
#topnav-sync{margin-left:8px;display:inline-flex;align-items:center;gap:6px}
#topnav-sync .syncdot{font-size:10px;line-height:1;color:#bbb;transition:.2s}
#topnav-sync.sync-ok .syncdot{color:#5cae7e}
#topnav-sync.sync-err .syncdot{color:#d9776a}
#topnav-sync.sync-busy .syncdot{color:#e0b54a}
.jbs-modal-bg{position:fixed;inset:0;background:rgba(40,32,28,.5);z-index:9999;
  display:none;align-items:center;justify-content:center;padding:20px}
.jbs-modal{background:#fff;border-radius:16px;max-width:450px;width:100%;
  box-shadow:0 20px 60px rgba(0,0,0,.32);font-family:'Jost',sans-serif;
  max-height:90vh;overflow:auto}
.jbs-modal-h{display:flex;justify-content:space-between;align-items:center;
  padding:15px 20px;border-bottom:1px solid #eee;font-family:'Cormorant Garamond',Georgia,serif;
  font-size:21px;font-weight:600;color:#4f463c}
.jbs-modal-x{background:none;border:none;font-size:25px;cursor:pointer;color:#aaa;line-height:1}
.jbs-modal-b{padding:18px 20px}
.jbs-sync-intro{font-size:13px;color:#6b6258;line-height:1.55;margin:0 0 13px}
.sync-status{font-size:13px;padding:9px 12px;border-radius:8px;background:#f2efe9;
  color:#6b6258;margin-bottom:14px}
.sync-status.ok{background:#e7f3ec;color:#3d7d56}
.sync-status.err{background:#fbe9e6;color:#b5503f}
.jbs-sync-label{display:block;font-size:12px;font-weight:600;color:#4f463c;margin-bottom:5px}
.jbs-sync-hint{font-weight:400;color:#999}
.jbs-sync-input{width:100%;box-sizing:border-box;padding:9px 11px;border:1px solid #ddd;
  border-radius:8px;font-size:13px;font-family:inherit;margin-bottom:12px}
.jbs-sync-actions{display:flex;flex-wrap:wrap;gap:7px;margin-bottom:12px}
.jbs-sync-btn{padding:8px 13px;border:1px solid #ddd;background:#f7f4ef;border-radius:8px;
  font-size:12.5px;font-family:inherit;cursor:pointer;color:#4f463c}
.jbs-sync-btn.primary{background:linear-gradient(135deg,#c0879a,#b3a3d6);color:#fff;border:none}
.jbs-sync-btn.danger{color:#b5503f}
.jbs-sync-gist{font-size:12px;color:#6b6258;line-height:1.55;margin-bottom:10px}
.jbs-sync-gist a{color:#b3789a}
.jbs-sync-help{font-size:12.5px;color:#6b6258;border-top:1px solid #eee;padding-top:11px}
.jbs-sync-help summary{cursor:pointer;font-weight:600;color:#4f463c}
.jbs-sync-help ol{margin:9px 0 0;padding-left:20px;line-height:1.75}
.jbs-sync-help a{color:#b3789a}
__PLANNER_CSS__
</style>
</head>
<body class="theme-analytics">
<nav class="topnav"><div class="topnav-inner">
  <span class="brand">&#10022; journeysbysam Insights</span>
  <div class="topnav-tabs">
    <button class="topnav-tab active" data-pane="analytics">Content Analytics</button>
    <button class="topnav-tab" data-pane="planner">40K Growth Planner</button>
  </div>
  <a class="topnav-refresh" target="_blank" rel="noopener"
     href="https://github.com/shambhavi17/journeysbysam-dashboard/actions/workflows/refresh.yml"
     title="Opens GitHub — tap &quot;Run workflow&quot; to pull fresh Instagram data (~3 min)">&#8635; Refresh data</a>
  <button class="topnav-refresh sync-off" id="topnav-sync" onclick="jbsSyncOpen()"
     title="Sync your edits across all your devices">&#9729; <span class="syncdot">&#9679;</span> Sync</button>
</div></nav>
<div id="jbs-sync-modal" class="jbs-modal-bg">
  <div class="jbs-modal">
    <div class="jbs-modal-h"><span>&#9729; Device Sync</span>
      <button class="jbs-modal-x" onclick="jbsSyncClose()">&times;</button></div>
    <div class="jbs-modal-b">
      <p class="jbs-sync-intro">Keep your dashboard edits &mdash; calendar, performance log,
        ideas, hooks, everything &mdash; in sync across your laptop and iPad. Your data is stored
        in a <b>private GitHub Gist</b> that only you can see. Completely free.</p>
      <div id="syncStatus" class="sync-status">Not connected</div>
      <label class="jbs-sync-label">GitHub token
        <span class="jbs-sync-hint">(classic token, <b>gist</b> scope only)</span></label>
      <input type="password" id="syncToken" class="jbs-sync-input" placeholder="ghp_..." autocomplete="off">
      <div class="jbs-sync-actions">
        <button class="jbs-sync-btn primary" onclick="jbsSyncConnect()">Connect</button>
        <button class="jbs-sync-btn" onclick="jbsSyncPush()">Save now &#8593;</button>
        <button class="jbs-sync-btn" onclick="jbsSyncPull()">Pull latest &#8595;</button>
        <button class="jbs-sync-btn danger" onclick="jbsSyncDisconnect()">Disconnect</button>
      </div>
      <div id="syncGist" class="jbs-sync-gist"></div>
      <details class="jbs-sync-help">
        <summary>How do I get a token? (1-minute, one-time setup)</summary>
        <ol>
          <li>Open <a href="https://github.com/settings/tokens/new" target="_blank" rel="noopener">github.com/settings/tokens/new</a> (log in if asked)</li>
          <li>Note: type <b>journeysbysam dashboard sync</b></li>
          <li>Expiration: choose <b>No expiration</b> (or 1 year)</li>
          <li>Scopes: tick <b>only</b> the <b>gist</b> box &mdash; leave everything else off</li>
          <li>Click <b>Generate token</b>, copy it, paste it above, hit <b>Connect</b></li>
          <li>On your iPad: open this dashboard, tap Sync, paste the <b>same token</b>, Connect &mdash; it finds your data automatically</li>
        </ol>
        <p style="margin:9px 0 0">This token can only touch gists &mdash; not your repos or account. You can delete it anytime in GitHub settings.</p>
      </details>
    </div>
  </div>
</div>
<div id="pane-analytics" class="pane active">
<header>
  <div class="wrap">
    <div class="eyebrow">Content Analytics Dashboard</div>
    <h1 id="creator"></h1>
    <div class="sub" id="subline"></div>
    <div class="kpis" id="kpis"></div>
  </div>
</header>
<div class="wrap">

  <div class="section-title">Performance Over Time</div>
  <div class="section-desc">Reel reach (plays) in chronological order — spot what spikes.</div>
  <div class="card"><div class="chart-box tall"><canvas id="timelineChart"></canvas></div></div>

  <div class="section-title">What Drives Reach</div>
  <div class="grid2">
    <div class="card"><h3>Content Type — avg per post</h3><div class="chart-box"><canvas id="typeChart"></canvas></div></div>
    <div class="card"><h3>Video Length vs Avg Plays</h3><div class="chart-box"><canvas id="durChart"></canvas></div></div>
    <div class="card"><h3>Best Posting Day (Reels, avg plays)</h3><div class="chart-box"><canvas id="dayChart"></canvas></div></div>
    <div class="card"><h3>Hook Pattern vs Avg Plays</h3><div class="chart-box"><canvas id="hookChart"></canvas></div></div>
  </div>

  <div class="section-title">Top Hashtags</div>
  <div class="section-desc">Your 14 most-used tags, ranked by average engagement per post.</div>
  <div class="card"><div class="chart-box tall"><canvas id="hashtagChart"></canvas></div></div>

  <div class="section-title">Post Rankings</div>
  <div class="section-desc">Switch between reach, engagement rate, and total engagement.</div>
  <div class="tabs" id="rankTabs">
    <div class="tab active" data-k="topReach">By Reach (plays)</div>
    <div class="tab" data-k="topEngRate">By Engagement Rate</div>
    <div class="tab" data-k="topEngAbs">By Total Engagement</div>
  </div>
  <div class="card" style="padding:6px 6px"><table id="rankTable"></table></div>

  <div class="section-title">Your Best Hooks</div>
  <div class="section-desc">Opening lines from your six highest-reach Reels — your proven attention-grabbers.</div>
  <div class="hooks-grid" id="hooks"></div>

  <div class="section-title">Viral Patterns</div>
  <div class="section-desc">Data-derived signals about what makes your content take off.</div>
  <div class="grid2" id="insights"></div>

  <div class="section-title">Competitor Benchmark</div>
  <div class="section-desc">You vs <b>peekinourjournal</b> and <b>tripuntraveled</b> &mdash; on likes, comments, and cadence.</div>
  <div class="card" style="padding:6px 6px"><table id="benchTable"></table></div>
  <div class="grid2" style="margin-top:18px">
    <div class="card"><h3>Avg Likes per Post (log scale)</h3><div class="chart-box"><canvas id="compLikesChart"></canvas></div></div>
    <div class="card"><h3>Avg Comments &amp; Posts / Week</h3><div class="chart-box"><canvas id="compEngChart"></canvas></div></div>
    <div class="card"><h3>Avg Caption Length (characters)</h3><div class="chart-box"><canvas id="compCaptionChart"></canvas></div></div>
    <div class="card"><h3>Breakout Rate — posts beating 2&times; own median</h3><div class="chart-box"><canvas id="compBreakoutChart"></canvas></div></div>
  </div>

  <div class="section-title">Deeper Competitor Insights</div>
  <div class="section-desc">What the numbers reveal beyond likes — and how to close each gap.</div>
  <div class="grid2" id="compDeep"></div>

  <div class="section-title">What to Improve</div>
  <div class="section-desc">Where the gap is — and why closing it grows your account.</div>
  <div class="grid2" id="gapAnalysis"></div>

  <div class="section-title">What to Add to Your Content</div>
  <div class="section-desc">Tactics your competitors use that you currently don't.</div>
  <div class="grid3" id="whatToAdd"></div>

  <div class="section-title">Content Ideas — Tailored to Your Data</div>
  <div class="section-desc">Six concrete next posts, each backed by a pattern already working for you.</div>
  <div class="grid3" id="ideas"></div>

  <footer>
    Generated from Apify Instagram scraper export &middot; 50 posts analyzed &middot; TikTok data not included in this build.
  </footer>
</div>
</div><!-- /pane-analytics -->

<div id="pane-planner" class="pane">
__PLANNER_BODY__
</div>

<script>
const D = __DATA__;
const fmt = n => n>=1000 ? (n/1000).toFixed(n>=10000?0:1)+'k' : (n||0).toLocaleString();
const C = {accent:'#c0879a',accent2:'#9bb6c4',gold:'#caa86f',pink:'#d99fa6',green:'#9aae93',
  grid:'#ebe0d0',muted:'#a99c8d'};
Chart.defaults.color = C.muted;
Chart.defaults.font.family = "-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif";

// header
const s = D.summary;
document.getElementById('creator').textContent = s.creator;
document.getElementById('subline').textContent =
  s.niche + '  ·  ' + s.dateRange + '  ·  ' + s.totalPosts + ' posts analyzed';
const kpis = [
  ['a', fmt(s.totalPlays), 'Total Reel Plays'],
  ['b', s.avgEngRate + '%', 'Avg Engagement Rate'],
  ['c', fmt(s.totalLikes), 'Total Likes'],
  ['d', fmt(s.totalComments), 'Total Comments'],
  ['e', s.videos, 'Reels'],
  ['a', s.carousels, 'Carousels'],
];
document.getElementById('kpis').innerHTML = kpis.map(k =>
  `<div class="kpi ${k[0]}"><div class="val">${k[1]}</div><div class="lbl">${k[2]}</div></div>`).join('');

// timeline
new Chart(document.getElementById('timelineChart'), {
  type:'line',
  data:{labels:D.timeline.map(t=>t.date),
    datasets:[{label:'Plays',data:D.timeline.map(t=>t.plays),
      borderColor:C.accent,backgroundColor:'rgba(192,135,154,.15)',fill:true,
      tension:.35,pointRadius:3,pointHoverRadius:6,pointBackgroundColor:C.accent2}]},
  options:{maintainAspectRatio:false,plugins:{legend:{display:false},
    tooltip:{callbacks:{afterLabel:c=>D.timeline[c.dataIndex].hook}}},
    scales:{y:{grid:{color:C.grid},ticks:{callback:fmt}},x:{grid:{display:false}}}}
});

// type
const tc = D.typeCompare, tk = Object.keys(tc);
new Chart(document.getElementById('typeChart'),{
  type:'bar',
  data:{labels:tk,datasets:[
    {label:'Avg Likes',data:tk.map(k=>tc[k].avgLikes),backgroundColor:C.accent},
    {label:'Avg Comments',data:tk.map(k=>tc[k].avgComments),backgroundColor:C.accent2}]},
  options:{maintainAspectRatio:false,plugins:{legend:{position:'bottom'}},
    scales:{y:{grid:{color:C.grid}},x:{grid:{display:false}}}}
});

// duration
const dd = D.durationData, dk = Object.keys(dd);
new Chart(document.getElementById('durChart'),{
  type:'bar',
  data:{labels:dk,datasets:[{label:'Avg Plays',data:dk.map(k=>dd[k].avgPlays),
    backgroundColor:[C.green,C.accent,C.accent2,C.muted]}]},
  options:{maintainAspectRatio:false,plugins:{legend:{display:false},
    tooltip:{callbacks:{afterLabel:c=>dd[dk[c.dataIndex]].count+' posts'}}},
    scales:{y:{grid:{color:C.grid},ticks:{callback:fmt}},x:{grid:{display:false}}}}
});

// day
const ds = D.dayStats, dyk = Object.keys(ds);
new Chart(document.getElementById('dayChart'),{
  type:'bar',
  data:{labels:dyk.map(d=>d.slice(0,3)),datasets:[{label:'Avg Plays',
    data:dyk.map(k=>ds[k].avgPlays),backgroundColor:C.gold}]},
  options:{maintainAspectRatio:false,plugins:{legend:{display:false},
    tooltip:{callbacks:{afterLabel:c=>ds[dyk[c.dataIndex]].count+' posts'}}},
    scales:{y:{grid:{color:C.grid},ticks:{callback:fmt}},x:{grid:{display:false}}}}
});

// hook patterns
new Chart(document.getElementById('hookChart'),{
  type:'bar',
  data:{labels:D.hookPatterns.map(h=>h.pattern),
    datasets:[{label:'Avg Plays',data:D.hookPatterns.map(h=>h.avgPlays),
      backgroundColor:C.pink}]},
  options:{indexAxis:'y',maintainAspectRatio:false,plugins:{legend:{display:false},
    tooltip:{callbacks:{afterLabel:c=>D.hookPatterns[c.dataIndex].count+' posts'}}},
    scales:{x:{grid:{color:C.grid},ticks:{callback:fmt}},y:{grid:{display:false}}}}
});

// hashtags
new Chart(document.getElementById('hashtagChart'),{
  type:'bar',
  data:{labels:D.hashtagData.map(h=>'#'+h.tag),
    datasets:[
      {label:'Avg Engagement',data:D.hashtagData.map(h=>h.avgEng),backgroundColor:C.accent},
      {label:'Times Used',data:D.hashtagData.map(h=>h.uses),backgroundColor:C.accent2,
        yAxisID:'y1'}]},
  options:{maintainAspectRatio:false,plugins:{legend:{position:'bottom'}},
    scales:{y:{grid:{color:C.grid},title:{display:true,text:'Avg Engagement'}},
      y1:{position:'right',grid:{display:false},title:{display:true,text:'Uses'}},
      x:{grid:{display:false},ticks:{maxRotation:55,minRotation:55,font:{size:10}}}}}
});

// rankings table
function renderTable(key){
  const rows = D[key];
  const isRate = key==='topEngRate';
  let h = `<thead><tr><th>#</th><th>Hook / Post</th><th>Type</th>`+
    `<th>Plays</th><th>Likes</th><th>Comments</th>`+
    `<th>${isRate?'Eng Rate':'Engagement'}</th><th>Date</th></tr></thead><tbody>`;
  rows.forEach((p,i)=>{
    const pill = p.type==='Video'
      ? '<span class="pill v">Reel</span>' : '<span class="pill s">Carousel</span>';
    const last = isRate ? (p.engRate!=null?p.engRate+'%':'—') : fmt(p.engagement);
    h += `<tr><td class="rank">${i+1}</td>`+
      `<td class="hook-cell"><a href="${p.url}" target="_blank">${p.hook}</a>`+
      `<div class="meta">${p.location} &middot; ${p.hashtagCount} hashtags`+
      (p.duration?` &middot; ${p.duration}s`:'')+`</div></td>`+
      `<td>${pill}</td>`+
      `<td class="num">${p.plays?fmt(p.plays):'—'}</td>`+
      `<td class="num">${fmt(p.likes)}</td>`+
      `<td class="num">${fmt(p.comments)}</td>`+
      `<td class="num" style="color:var(--green)">${last}</td>`+
      `<td class="meta">${p.date||'—'}</td></tr>`;
  });
  document.getElementById('rankTable').innerHTML = h+'</tbody>';
}
renderTable('topReach');
document.querySelectorAll('#rankTabs .tab').forEach(t=>{
  t.onclick=()=>{
    document.querySelectorAll('#rankTabs .tab').forEach(x=>x.classList.remove('active'));
    t.classList.add('active');
    renderTable(t.dataset.k);
  };
});

// hooks
document.getElementById('hooks').innerHTML = D.bestHooks.map(h=>
  `<div class="hook-card"><div class="q">&ldquo;${h.hook}&rdquo;</div>`+
  `<div class="s"><span><b>${fmt(h.plays)}</b> plays</span>`+
  `<span>${h.engRate!=null?h.engRate+'% eng rate':''}</span>`+
  `<span><a href="${h.url}" target="_blank" style="color:var(--accent2);text-decoration:none">view &rarr;</a></span>`+
  `</div></div>`).join('');

// insights
document.getElementById('insights').innerHTML = D.insights.map((x,i)=>
  `<div class="insight"><div class="ic">${i+1}</div>`+
  `<div><h4>${x.title}</h4><p>${x.detail}</p></div></div>`).join('');

// ---- competitor benchmark ----
const B = D.benchmark;
let bh = `<thead><tr><th>Account</th><th>Posts/wk</th><th>Avg Likes</th>`+
  `<th>Median Likes</th><th>Best Post</th><th>Avg Comments</th><th>Signature Hook</th></tr></thead><tbody>`;
B.forEach(a=>{
  const you = a.isYou;
  bh += `<tr${you?' style="background:rgba(192,135,154,.10)"':''}>`+
    `<td><b style="color:${you?'var(--accent)':'var(--text)'}">@${a.handle.replace(' (You)','')}</b>`+
    (you?' <span class="pill v">You</span>':'')+
    `<div class="meta">${a.name}</div></td>`+
    `<td class="num">${a.perWeek}</td>`+
    `<td class="num" style="color:var(--gold)">${fmt(a.avgLikes)}</td>`+
    `<td class="num">${fmt(a.medLikes)}</td>`+
    `<td class="num">${fmt(a.maxLikes)}</td>`+
    `<td class="num" style="color:var(--green)">${fmt(a.avgComments)}</td>`+
    `<td class="hook-cell"><div class="meta" style="font-size:12px;color:#c7c9d6">`+
    `&ldquo;${(a.topHooks[0]||'').slice(0,60)}&rdquo;</div></td></tr>`;
});
document.getElementById('benchTable').innerHTML = bh+'</tbody>';

const compColors = B.map(a=>a.isYou?C.accent:C.muted);
new Chart(document.getElementById('compLikesChart'),{
  type:'bar',
  data:{labels:B.map(a=>'@'+a.handle.replace(' (You)','')),
    datasets:[{label:'Avg Likes',data:B.map(a=>a.avgLikes),backgroundColor:compColors}]},
  options:{maintainAspectRatio:false,plugins:{legend:{display:false}},
    scales:{y:{type:'logarithmic',grid:{color:C.grid},ticks:{callback:fmt}},
      x:{grid:{display:false}}}}
});
new Chart(document.getElementById('compEngChart'),{
  type:'bar',
  data:{labels:B.map(a=>'@'+a.handle.replace(' (You)','')),
    datasets:[
      {label:'Avg Comments',data:B.map(a=>a.avgComments),backgroundColor:C.green},
      {label:'Posts / Week',data:B.map(a=>a.perWeek),backgroundColor:C.pink,yAxisID:'y1'}]},
  options:{maintainAspectRatio:false,plugins:{legend:{position:'bottom'}},
    scales:{y:{grid:{color:C.grid},title:{display:true,text:'Avg Comments'}},
      y1:{position:'right',grid:{display:false},title:{display:true,text:'Posts/wk'}},
      x:{grid:{display:false}}}}
});

new Chart(document.getElementById('compCaptionChart'),{
  type:'bar',
  data:{labels:B.map(a=>'@'+a.handle.replace(' (You)','')),
    datasets:[{label:'Avg caption length',data:B.map(a=>a.avgCaption),backgroundColor:compColors}]},
  options:{maintainAspectRatio:false,plugins:{legend:{display:false}},
    scales:{y:{grid:{color:C.grid}},x:{grid:{display:false}}}}
});
new Chart(document.getElementById('compBreakoutChart'),{
  type:'bar',
  data:{labels:B.map(a=>'@'+a.handle.replace(' (You)','')),
    datasets:[{label:'Breakout rate',data:B.map(a=>a.breakoutRate),backgroundColor:compColors}]},
  options:{maintainAspectRatio:false,plugins:{legend:{display:false}},
    scales:{y:{grid:{color:C.grid},ticks:{callback:v=>v+'%'}},x:{grid:{display:false}}}}
});

// deeper competitor insights
document.getElementById('compDeep').innerHTML = D.compDeep.map((x,i)=>
  `<div class="insight"><div class="ic" style="background:linear-gradient(135deg,var(--accent),var(--accent2))">${i+1}</div>`+
  `<div><h4>${x.title}</h4><p>${x.detail}</p></div></div>`).join('');

// gap analysis
document.getElementById('gapAnalysis').innerHTML = D.gapAnalysis.map((x,i)=>
  `<div class="insight"><div class="ic" style="background:linear-gradient(135deg,var(--pink),var(--gold))">${i+1}</div>`+
  `<div><h4>${x.title}</h4><p>${x.detail}</p></div></div>`).join('');

// what to add
document.getElementById('whatToAdd').innerHTML = D.whatToAdd.map(x=>
  `<div class="idea"><div class="tag">${x.tag}</div><h4>${x.title}</h4>`+
  `<p>${x.detail}</p></div>`).join('');

// ideas
document.getElementById('ideas').innerHTML = D.ideas.map(x=>
  `<div class="idea"><div class="tag">${x.tag}</div><h4>${x.title}</h4>`+
  `<p>${x.why}</p><div class="fmt">${x.format}</div></div>`).join('');
</script>
<script>
__PLANNER_JS__
</script>
<script>
/* ===== Device Sync via private GitHub Gist ===== */
(function(){
  var SYNC_CFG='jbs_sync_cfg', GIST_FILE='jbs_dashboard.json',
      GIST_DESC='journeysbysam Insights - dashboard sync', lastSnap=null;
  function cfg(){try{return JSON.parse(localStorage.getItem(SYNC_CFG)||'{}');}catch(e){return {};}}
  function setCfg(c){localStorage.setItem(SYNC_CFG,JSON.stringify(c));}
  function syncable(k){return k&&k.indexOf('jbs_')===0&&k!==SYNC_CFG&&k!=='jbs_ai_cfg';}
  function snapshot(){var o={};for(var i=0;i<localStorage.length;i++){var k=localStorage.key(i);
    if(syncable(k))o[k]=localStorage.getItem(k);}return o;}
  function snapStr(){return JSON.stringify(snapshot());}
  function applyData(d){Object.keys(d||{}).forEach(function(k){
    if(syncable(k)){try{localStorage.setItem(k,d[k]);}catch(e){}}});}
  function hdr(c){return {'Authorization':'Bearer '+c.token,'Accept':'application/vnd.github+json'};}
  function badge(s){var b=document.getElementById('topnav-sync');
    if(b)b.className='topnav-refresh sync-'+s;}
  function stat(msg,cls){var e=document.getElementById('syncStatus');
    if(e){e.textContent=msg;e.className='sync-status '+(cls||'');}}
  function showGist(){var c=cfg(),e=document.getElementById('syncGist');if(!e)return;
    e.innerHTML=c.gistId?('Connected. Your data lives in a <a href="https://gist.github.com/'+c.gistId+
      '" target="_blank" rel="noopener">private gist</a>. Use the same token on your other devices.'):'';}
  function push(){
    var c=cfg();if(!c.token)return Promise.resolve(false);
    badge('busy');stat('Saving...','');
    var body={files:{}};
    body.files[GIST_FILE]={content:JSON.stringify({updatedAt:new Date().toISOString(),data:snapshot()})};
    var url='https://api.github.com/gists',method='POST';
    if(c.gistId){url+='/'+c.gistId;method='PATCH';}else{body.description=GIST_DESC;body.public=false;}
    return fetch(url,{method:method,headers:hdr(c),body:JSON.stringify(body)})
      .then(function(r){if(!r.ok)throw new Error('HTTP '+r.status);return r.json();})
      .then(function(j){if(!c.gistId){c.gistId=j.id;setCfg(c);showGist();}
        lastSnap=snapStr();badge('ok');stat('Saved to cloud ✓','ok');return true;})
      .catch(function(e){badge('err');stat('Save failed: '+e.message,'err');return false;});
  }
  function fetchGist(c){
    return fetch('https://api.github.com/gists/'+c.gistId,{headers:hdr(c)})
      .then(function(r){if(!r.ok)throw new Error('HTTP '+r.status);return r.json();})
      .then(function(j){var f=j.files&&j.files[GIST_FILE];if(!f)return null;
        if(f.truncated&&f.raw_url)return fetch(f.raw_url).then(function(r){return r.text();});
        return f.content;});
  }
  function pull(reload){
    var c=cfg();if(!c.token||!c.gistId)return Promise.resolve(false);
    badge('busy');stat('Loading...','');
    return fetchGist(c).then(function(content){
      if(!content){badge('ok');stat('Nothing stored yet — click Save now','');return false;}
      applyData((JSON.parse(content)||{}).data||{});
      lastSnap=snapStr();badge('ok');stat('Loaded from cloud ✓','ok');
      if(reload)setTimeout(function(){location.reload();},450);
      return true;
    }).catch(function(e){badge('err');stat('Load failed: '+e.message,'err');return false;});
  }
  window.jbsSyncOpen=function(){document.getElementById('jbs-sync-modal').style.display='flex';
    var c=cfg();if(c.token)document.getElementById('syncToken').value=c.token;showGist();
    if(!c.token)stat('Not connected','');
    else if(c.gistId)stat('Connected — syncing on','ok');
    else stat('Token saved — click Connect to finish','');};
  window.jbsSyncClose=function(){document.getElementById('jbs-sync-modal').style.display='none';};
  window.jbsSyncConnect=function(){
    var t=document.getElementById('syncToken').value.trim();
    if(!t){stat('Paste your token first','err');return;}
    var c={token:t};setCfg(c);badge('busy');stat('Connecting...','');
    fetch('https://api.github.com/gists',{headers:hdr(c)})
      .then(function(r){if(!r.ok)throw new Error('HTTP '+r.status+' — check the token & gist scope');return r.json();})
      .then(function(list){
        var found=(list||[]).filter(function(g){return g.description===GIST_DESC&&g.files&&g.files[GIST_FILE];})[0];
        if(found){c.gistId=found.id;setCfg(c);showGist();
          stat('Found your data — loading...','ok');pull(true);}
        else{stat('New setup — creating your cloud storage...','');
          push().then(function(ok){if(ok)stat('Connected & saved ✓','ok');});}
      })
      .catch(function(e){badge('err');stat('Connect failed: '+e.message,'err');});
  };
  window.jbsSyncPush=function(){push();};
  window.jbsSyncPull=function(){pull(true);};
  window.jbsSyncDisconnect=function(){
    if(!confirm('Disconnect sync on this device? Your edits stay here but stop syncing.'))return;
    localStorage.removeItem(SYNC_CFG);badge('off');stat('Disconnected','');
    document.getElementById('syncToken').value='';showGist();};
  function initSync(){
    lastSnap=snapStr();
    var c=cfg();
    if(!c.token){badge('off');return;}
    badge('ok');
    if(c.gistId&&!sessionStorage.getItem('jbs_sync_done')){
      sessionStorage.setItem('jbs_sync_done','1');
      var pre=snapStr();
      fetchGist(c).then(function(content){
        if(!content){badge('ok');return;}
        applyData((JSON.parse(content)||{}).data||{});
        var now=snapStr();lastSnap=now;
        if(now!==pre)location.reload();else badge('ok');
      }).catch(function(){badge('err');});
    }
    setInterval(function(){var cc=cfg();if(!cc.token)return;
      var now=snapStr();if(lastSnap!==null&&now!==lastSnap)push();},30000);
    document.addEventListener('visibilitychange',function(){
      if(document.hidden){var cc=cfg();if(cc.token&&snapStr()!==lastSnap)push();}});
  }
  if(document.readyState!=='loading')initSync();
  else document.addEventListener('DOMContentLoaded',initSync);
})();
</script>
<script>
(function(){
  var panes={analytics:document.getElementById('pane-analytics'),
             planner:document.getElementById('pane-planner')};
  document.querySelectorAll('.topnav-tab').forEach(function(t){
    t.addEventListener('click',function(){
      var p=t.dataset.pane;
      document.querySelectorAll('.topnav-tab').forEach(function(x){
        x.classList.toggle('active',x===t);});
      panes.analytics.classList.toggle('active',p==='analytics');
      panes.planner.classList.toggle('active',p==='planner');
      document.body.className='theme-'+p;
      if(p==='planner'&&window.__initPlanner)window.__initPlanner();
      window.scrollTo(0,0);
    });
  });
})();
</script>
</body>
</html>"""

HTML = HTML.replace("__PLANNER_HEAD__", PA["head"])
HTML = HTML.replace("__PLANNER_CSS__", PA["css"])
HTML = HTML.replace("__PLANNER_BODY__", PA["body"])
HTML = HTML.replace("__PLANNER_JS__", PA["js"])
HTML = HTML.replace("__DATA__", J)
open(OUT, "w").write(HTML)
print("written:", OUT, len(HTML), "bytes")
print("videos:", len(vids), "carousels:", len(cars), "total plays:", total_plays)
