#!/usr/bin/env python3
"""
Truth Social watcher.

Checks the trumpstruth.org RSS feed for newly published Trump posts and posts
each new one to a Microsoft Teams channel via a Workflows ("incoming webhook")
Adaptive Card.

We read trumpstruth.org's RSS feed rather than scraping Truth Social directly:
there is no official Truth Social API, and scraping it means fighting bot
detection / ToS. trumpstruth.org (Defending Democracy Together) does that work
and exposes a clean, standard RSS feed — the polite, first-party-ish source.

State is kept in state.json (last-seen publish date + recently-posted IDs) so
nothing is missed or double-posted even if a scheduled run is skipped.

Reads two secrets from the environment:
  TEAMS_WEBHOOK_URL - Teams Workflows webhook URL (the Truth Social channel)

Zero third-party dependencies (stdlib only).
"""

import html
import json
import os
import re
import sys
import urllib.request
import xml.etree.ElementTree as ET
from datetime import datetime, timezone
from email.utils import parsedate_to_datetime

# --- config -----------------------------------------------------------------

FEED_URL = "https://www.trumpstruth.org/feed"
STATE_FILE = os.path.join(os.path.dirname(__file__), "state.json")

# trumpstruth.org's custom namespace (for truth:originalId / truth:originalUrl).
TRUTH_NS = "https://truthsocial.com/ns"

# How many recently-posted IDs to remember (dedupe guard). He posts a lot, so
# keep a generous window.
SEEN_IDS_CAP = 500

# On the very first run (no state file), seed silently instead of flooding the
# channel with the entire current feed. Override by setting SEED_AND_POST=1.
SEED_AND_POST = os.environ.get("SEED_AND_POST") == "1"

# A browser-ish UA — some CDNs 403 the default urllib agent.
USER_AGENT = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) truth-social-alerts/1.0"


# --- helpers ----------------------------------------------------------------

def log(msg):
    print(msg, flush=True)


def load_state():
    try:
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        return None
    except (json.JSONDecodeError, OSError) as e:
        log(f"WARNING: could not read state file ({e}); treating as first run.")
        return None


def save_state(state):
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f, indent=2)
        f.write("\n")


def clean_text(value):
    """Neutralise markdown so untrusted post text can't render as links/markup
    in a Teams card. Teams TextBlocks render a subset of markdown, so a hostile
    string like '[click](http://evil)' would otherwise become a live link."""
    if not value:
        return ""
    out = str(value)
    for ch in "[]()`*_#>|":
        out = out.replace(ch, "\\" + ch)
    return out


def safe_https(value):
    """Return the URL only if it's a plain https:// link, else ''.
    Blocks javascript:, data:, http:, etc. from reaching a card button/image."""
    if not value:
        return ""
    v = str(value).strip()
    return v if v.lower().startswith("https://") else ""


def strip_html(raw):
    """Turn the description HTML into plain text: drop tags, unescape entities,
    collapse whitespace. Retweets come through as 'RT: <url>' which we keep."""
    if not raw:
        return ""
    # Preserve some structure: turn block/break tags into newlines first.
    text = re.sub(r"(?i)<\s*(br|/p|/div)\s*/?>", "\n", raw)
    text = re.sub(r"<[^>]+>", "", text)          # strip remaining tags
    text = html.unescape(text)                    # &amp; -> & etc.
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def parse_dt(value):
    """Parse a timestamp to an aware datetime, or None. Handles both the feed's
    RFC 822 pubDate ('Wed, 01 Jul 2026 01:31:48 +0000') and our own ISO 8601
    watermark ('2026-07-01T01:31:48Z') — the latter is what we save to state, so
    it must round-trip on reload."""
    if not value:
        return None
    # ISO 8601 first (our stored watermark).
    if "T" in value:
        try:
            return datetime.fromisoformat(value.replace("Z", "+00:00"))
        except (ValueError, AttributeError):
            pass
    # RFC 822 (feed pubDate).
    try:
        dt = parsedate_to_datetime(value)
        if dt is not None and dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        return dt
    except (TypeError, ValueError):
        return None


def _find_text(item, tag):
    el = item.find(tag)
    return el.text if el is not None and el.text else ""


def fetch_items():
    """Fetch and parse the RSS feed into a list of dicts (feed order: newest
    first)."""
    req = urllib.request.Request(
        FEED_URL, headers={"Accept": "application/rss+xml, application/xml",
                           "User-Agent": USER_AGENT},
    )
    with urllib.request.urlopen(req, timeout=30) as resp:
        body = resp.read().decode("utf-8")

    root = ET.fromstring(body)
    channel = root.find("channel")
    if channel is None:
        return []

    items = []
    for item in channel.findall("item"):
        original_id = _find_text(item, f"{{{TRUTH_NS}}}originalId")
        original_url = _find_text(item, f"{{{TRUTH_NS}}}originalUrl")
        # Fall back to guid/link if the custom id is ever missing.
        guid = _find_text(item, "guid")
        link = _find_text(item, "link")
        items.append({
            "id": original_id or guid or link,
            "title": _find_text(item, "title"),
            "text": strip_html(_find_text(item, "description")),
            "link": link,
            "original_url": original_url,
            "pubDate": _find_text(item, "pubDate"),
        })
    return items


def build_card(item):
    """Build the Teams Adaptive Card envelope for one Truth Social post."""
    text = clean_text(item.get("text"))
    # trumpstruth.org archive page (stable) as the primary link; original Truth
    # Social URL as a secondary if present.
    archive_url = safe_https(item.get("link"))
    truth_url = safe_https(item.get("original_url"))

    when = parse_dt(item.get("pubDate"))
    when_str = when.strftime("%d %b %Y %H:%M UTC") if when else ""

    body = [
        {"type": "TextBlock", "text": "🇺🇸 New Truth Social post — Donald Trump",
         "weight": "Bolder", "size": "Medium", "color": "Accent", "wrap": True},
    ]
    if when_str:
        body.append({"type": "TextBlock", "text": when_str, "isSubtle": True,
                     "spacing": "None", "size": "Small"})
    if text:
        body.append({"type": "TextBlock", "text": text, "wrap": True,
                     "spacing": "Medium"})
    else:
        body.append({"type": "TextBlock", "text": "_(no text — likely an image, "
                     "video, or repost; open to view)_", "wrap": True,
                     "isSubtle": True, "spacing": "Medium"})

    actions = []
    if archive_url:
        actions.append({"type": "Action.OpenUrl", "title": "📄 View post", "url": archive_url})
    if truth_url:
        actions.append({"type": "Action.OpenUrl", "title": "↗ On Truth Social", "url": truth_url})

    card = {
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "type": "AdaptiveCard",
        "version": "1.4",
        "body": body,
        "actions": actions,
    }
    return {
        "type": "message",
        "attachments": [{
            "contentType": "application/vnd.microsoft.card.adaptive",
            "content": card,
        }],
    }


def post_to_teams(webhook_url, card):
    payload = json.dumps(card).encode("utf-8")
    req = urllib.request.Request(
        webhook_url, data=payload,
        headers={"Content-Type": "application/json"}, method="POST",
    )
    with urllib.request.urlopen(req, timeout=30) as resp:
        return resp.status


# --- main -------------------------------------------------------------------

def main():
    webhook_url = os.environ.get("TEAMS_WEBHOOK_URL")
    if not webhook_url:
        log("ERROR: TEAMS_WEBHOOK_URL must be set.")
        return 1

    try:
        items = fetch_items()
    except Exception as e:
        log(f"ERROR fetching/parsing feed: {e}")
        return 1

    log(f"Fetched {len(items)} item(s) from feed.")
    state = load_state()
    first_run = state is None
    if first_run:
        state = {"last_published": None, "seen_ids": []}

    # Ordered list (oldest→newest) + set for O(1) lookup. Order matters so the
    # trim below keeps the genuinely most-recent IDs, not an arbitrary subset.
    seen_list = list(state.get("seen_ids") or [])
    seen_ids = set(seen_list)
    last_published = parse_dt(state.get("last_published"))

    # Feed is newest-first; reverse so the channel reads chronologically.
    items = list(reversed(items))

    new_items = []
    for item in items:
        item_id = item.get("id")
        if not item_id or item_id in seen_ids:
            continue
        pub = parse_dt(item.get("pubDate"))
        # If we have a watermark, only take strictly newer items.
        if last_published and pub and pub <= last_published:
            continue
        new_items.append(item)

    # Decide whether to actually post.
    posting = not (first_run and not SEED_AND_POST)
    if first_run and not SEED_AND_POST:
        log(f"First run: seeding state with {len(new_items)} item(s), posting none. "
            f"(Set SEED_AND_POST=1 to post on a manual run.)")

    posted = 0
    for item in new_items:
        if posting:
            try:
                status = post_to_teams(webhook_url, build_card(item))
                log(f"Posted: {item.get('id')} @ {item.get('pubDate')} (HTTP {status})")
                posted += 1
            except Exception as e:
                log(f"ERROR posting {item.get('id')}: {e} — will retry next run.")
                # Don't record as seen, so the next run tries again.
                continue
        iid = item.get("id")
        if iid not in seen_ids:
            seen_ids.add(iid)
            seen_list.append(iid)
        pub = parse_dt(item.get("pubDate"))
        if pub and (last_published is None or pub > last_published):
            last_published = pub

    # Persist trimmed state.
    state["last_published"] = last_published.strftime("%Y-%m-%dT%H:%M:%SZ") if last_published else None
    # Keep the most recent IDs only (seen_list is oldest→newest).
    state["seen_ids"] = seen_list[-SEEN_IDS_CAP:]
    save_state(state)

    log(f"Done. New: {len(new_items)}, posted: {posted}, "
        f"watermark: {state['last_published']}.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
