#!/usr/bin/env python
"""AGOL sharing audit: which maps, and which apps that embed them, their viewers cannot open.

Answers one question per layer reference: is every audience member of the MAP also an
audience member of the LAYER?  Audience containment, not an access "rank": 'org' and
'shared' are incomparable audiences, not equal ones (an org member outside the group
genuinely cannot load a group-shared layer).

--edges asks the same question one level up: is every audience member of the APP (dashboard,
web mapping application, StoryMap, Experience Builder app) also an audience member of the web
map it embeds?  --wcag runs a small accessibility rule pack over the same web maps.

HARD RULE. Never attribute-access an unknown field on an Item.  Item.__getattr__ does a
network round trip *before* raising AttributeError (measured 0.665s each); the original
script spent ~95% of its runtime inside two such probes for columns that were empty in
every row.  Dict access only: it["access"], it.get("numViews", 0).  The only permitted
attribute/method calls on an Item are .get_data() and .id.

Run:  python sightline.py [-o PATH] [--limit N] [--deep] [--edges] [--wcag] [--apply]
Test: python sightline.py --self-test      (--self-check is an alias; no network, no login)
"""

import os
import re
import ast
import io
import sys
import json
import time
import types
import shutil
import getpass
import argparse
import datetime
import tempfile
import warnings
import threading
import contextlib
import importlib.util
import http.server
import urllib.error
import urllib.parse
import urllib.request
import collections

# pandas, openpyxl and arcgis load only when a live run needs them (see load_report_libs and
# load_arcgis), so the decision core and --self-test run on a bare Python.
pd = None
get_column_letter = None
GIS = None
FeatureLayer = None

# openpyxl's own ILLEGAL_CHARACTERS_RE, copied so scrub() needs no import.  The self-test
# compares it with openpyxl's when openpyxl is installed.
ILLEGAL_CHARACTERS_RE = re.compile(r"[\000-\010]|[\013-\014]|[\016-\037]")


def load_report_libs():
    """Import what the Excel report needs.  Called by a live run, never by --self-test."""
    global pd, get_column_letter
    try:
        import pandas
        from openpyxl.utils import get_column_letter as column_letter
    except ModuleNotFoundError as exc:
        sys.exit("%s was not found. Install it: pip install arcgis pandas openpyxl\n"
                 "ArcGIS Pro's own Python already has all three." % exc.name)
    pd, get_column_letter = pandas, column_letter


def load_arcgis():
    """Import the ArcGIS API for Python.  Called when a sign-in or --deep needs it."""
    global GIS, FeatureLayer
    try:
        from arcgis.gis import GIS as gis_class
        from arcgis.features import FeatureLayer as layer_class
    except ModuleNotFoundError as exc:
        # These ship with ArcGIS Pro's bundled Python.  On a plain interpreter the
        # bare ImportError says nothing about where to find one that works.
        sys.exit(
            "%s was not found. Run this with the Python that ships with ArcGIS Pro:\n"
            r'  "C:\Program Files\ArcGIS\Pro\bin\Python\envs\arcgispro-py3\python.exe" '
            "sightline.py\n"
            "or install the dependencies yourself: pip install arcgis pandas openpyxl"
            % exc.name
        )
    GIS, FeatureLayer = gis_class, layer_class

# =============================================================================
#  CONFIGURATION. Everything you may want to change lives in this block.
#  Every value here is a default; the matching command-line flag wins when given.
# =============================================================================

# --- connection ---------------------------------------------------------------
# Portal URL.  ArcGIS Online for everyone except ArcGIS Enterprise sites, which
# use e.g. "https://gis.example.org/portal".
PORTAL_URL = "https://www.arcgis.com"

# Leave USERNAME blank to use, in order: $ARCGIS_USER -> the signed-in ArcGIS Pro
# session -> an interactive prompt.
USERNAME = ""

# PASSWORD: leave blank.  Set $ARCGIS_PASS instead, or let the script prompt.
# A password typed here is a password you will eventually commit to git.
PASSWORD = ""

# --- what to scan -------------------------------------------------------------
MAX_MAPS = 10000        # ceiling on web maps fetched
MAX_APPS = 10000        # ceiling on apps fetched, per item type (--edges)
DEFAULT_LIMIT = None    # None = scan everything; set an int to smoke-test (--limit)
DEEP_DEFAULT = False    # True = always collect FeatureCount + LastEditUTC (--deep)

# --- output -------------------------------------------------------------------
OUTPUT_DIR = os.path.join(os.path.expanduser("~"), "Downloads")
OUTPUT_PREFIX = "sightline"              # final name: <prefix>_YYYYmmdd_HHMMSS.xlsx
MAX_COL_WIDTH = 55                  # Excel column width cap
WIDTH_SAMPLE_ROWS = 400             # rows sampled when sizing columns

# --- service probing ----------------------------------------------------------
# The probe asks, anonymously, "can the public actually load this service?"
PROBE_TIMEOUT = 8       # seconds per service URL
MAX_HOST_FAILS = 3      # consecutive transport failures before a host is skipped.
                        # Without this, one unreachable on-prem box x 50 urls x 8s
                        # is a 400s hang inside a ~160s run.
UA = {"User-Agent": "Mozilla/5.0"}

# TLS verification stays ON.  A service whose certificate does not verify is one an
# anonymous browser would also refuse, and that is a finding (UNREACHABLE /
# SSLCertVerificationError), not something to paper over with an unverified context.

# ArcGIS REST error codes.  400 is not a typo: a request for a service that no longer
# exists answers "Invalid URL" with code 400, not 404 (verified against 12 dead urls --
# an org_admin session gets the identical 400, so it is absence, not permissions).
SECURED_CODES = {401, 403, 498, 499}
DEAD_CODES = {400, 404}

# --- batching (see README "Why it is fast") -----------------------------------
ITEM_CHUNK = 50         # ids per id:(a OR b ...) search; 18.8x faster than one-by-one
GROUP_CHUNK = 80        # ids per content/itemsgroups call

# --- report shape -------------------------------------------------------------
# Severity is driven by the MAP's audience: a broken layer in a public map matters
# more than the same layer in a private one.  Sorts correctly as text.  For --edges it is
# the APP's audience, because that is who opens the link.
SEV = {"public": "1 High", "org": "2 Medium", "shared": "3 Low", "private": "4 Info"}

FIND_COLS = ["Severity", "Verdict", "Reason", "MapTitle", "MapOwner", "MapAccess",
             "MapViews", "MapUrl", "LayerTitle", "LayerAccess", "LayerUrl"]
ALL_COLS = ["MapTitle", "MapId", "MapOwner", "MapAccess", "MapViews", "MapGroups", "MapUrl",
            "Section", "LayerTitle", "LayerType", "LayerItemId", "LayerUrl", "LayerHost",
            "LayerOwner", "LayerAccess", "LayerGroups", "LayerState", "Verdict", "Severity",
            "Reason", "Overexposed"]
VERDICTS = ("OK", "BROKEN", "DEAD", "UNKNOWN")

# --- app edges (--edges) ------------------------------------------------------
# Item type names as the ArcGIS REST "Items and item types" page spells them.
EDGE_TYPES = ("Dashboard", "Web Mapping Application", "StoryMap", "Web Experience")
# Keys that hold an item id in an app's JSON.  Esri documents none of these for the four
# item types (see README), so a candidate only becomes an edge when its id resolves to an
# item whose type is "Web Map".  A key missing from this list is a missed edge, not a wrong one.
MAP_KEYS = ("itemId", "webmap", "webmapId", "mapId")
ID_RE = re.compile(r"^[0-9a-fA-F]{32}$")
EDGE_COLS = ["Severity", "Verdict", "Reason", "AppType", "AppTitle", "AppId", "AppOwner",
             "AppAccess", "AppViews", "AppGroups", "AppUrl", "MapTitle", "MapId", "MapOwner",
             "MapAccess", "MapGroups", "MapUrl"]

# --- accessibility rule pack (--wcag) -----------------------------------------
# rule id -> (WCAG 2.2 success criterion, level, what the rule looks at).  A clean run is not
# a conformance claim: these are the checks a web map's JSON can answer without drawing it.
WCAG_RULES = collections.OrderedDict([
    ("map-title-placeholder", ("2.4.2", "A", "the item title is empty or a placeholder")),
    ("map-no-summary", ("1.1.1", "A", "the item has neither a summary nor a description")),
    ("layer-no-title", ("2.4.6", "AA", "a layer has no title for the layer list and legend")),
    ("layer-popup-off", ("1.1.1", "A", "a layer disables its pop-up")),
    ("renderer-color-only", ("1.4.1", "A", "unique-value classes differ by color alone")),
    ("label-halo-contrast", ("1.4.3", "AA", "label text and its halo are under the minimum ratio")),
])
# Lower-case, whitespace-collapsed titles that say nothing about the map.  A heuristic list.
PLACEHOLDER_TITLES = frozenset(["untitled", "untitled map", "untitled web map", "new map",
                                "map", "web map"])
WCAG_COLS = ["Severity", "Rule", "Criterion", "Level", "MapTitle", "MapId", "MapOwner",
             "MapAccess", "MapViews", "MapUrl", "Layer", "Detail"]
TEXT_RATIO = 4.5        # WCAG 1.4.3, normal text
LARGE_TEXT_RATIO = 3.0  # WCAG 1.4.3, text at least 18 point, or 14 point bold

# =============================================================================
#  END CONFIGURATION
# =============================================================================

FS_RE = re.compile(r"/featureserver/\d+/?$")          # strict: the service ROOT cannot be counted
# Query parameters that may carry a credential.  A layer url is copied into the report and
# its hyperlinks, so these never leave clean_url().  A heuristic on the NAME, not the value.
SECRET_PARAM_RE = re.compile(r"token|key|pass|pwd|secret|sig|auth|credential", re.I)
USERINFO_RE = re.compile(r"^([A-Za-z][A-Za-z0-9+.-]*://)[^/]*@")


# --------------------------------------------------------------------------- plumbing

def connect():
    """Resolve credentials, most explicit first, and sign in.

    CONFIG block -> environment -> signed-in ArcGIS Pro session -> interactive prompt.
    A password is never written to disk by this script.
    """
    if GIS is None:
        load_arcgis()
    # PORTAL_URL always holds a value, so $ARCGIS_URL overrides it; USERNAME and
    # PASSWORD default to blank, so a value typed in CONFIG overrides the environment.
    url = os.environ.get("ARCGIS_URL") or PORTAL_URL
    user = USERNAME or os.environ.get("ARCGIS_USER")
    pw = PASSWORD or os.environ.get("ARCGIS_PASS")
    if user and pw:
        return GIS(url, user, pw)
    if not user:
        try:
            return GIS("pro")
        except Exception as e:
            print("No active ArcGIS Pro session (%s); falling back to prompt." % type(e).__name__)
        user = input("Username: ")
    return GIS(url, user, pw or getpass.getpass("Password: "))


def iso(ms):
    """epoch MILLISECONDS -> 'YYYY-MM-DD HH:MM:SS+00:00' UTC.  '' when falsy."""
    if not ms:
        return ""
    dt = datetime.datetime.fromtimestamp(ms / 1000.0, datetime.timezone.utc)
    return str(dt.replace(microsecond=0))


def chunks(seq, n):
    seq = list(seq)
    for i in range(0, len(seq), n):
        yield seq[i:i + n]


def host(url):
    try:
        return urllib.parse.urlparse(url or "").netloc
    except ValueError:
        return ""


def clean_url(url):
    """url without userinfo, fragment and credential-named query parameters.

    Everything the report or a request shows of a layer url goes through here.  The other
    query parameters stay: a layer built from a .../query?where=... url needs them to load.
    """
    head, _, query = (url or "").split("#")[0].partition("?")
    head = USERINFO_RE.sub(r"\1", head)
    keep = [q for q in query.split("&") if q and not SECRET_PARAM_RE.search(q.split("=")[0])]
    return head + ("?" + "&".join(keep) if keep else "")


def item_url_base(gis):
    return gis.url.rstrip("/") + "/home/item.html?id="


def walk(data, out):
    """Append (section, layer_dict) for every real layer reference in a web map's JSON."""
    if not isinstance(data, dict):
        return
    for section in ("operationalLayers", "tables"):
        for lyr in (data.get(section) or []):
            _emit(section, lyr, out)
    for lyr in ((data.get("baseMap") or {}).get("baseMapLayers") or []):
        _emit("baseMap", lyr, out)


def _emit(section, lyr, out):
    if not isinstance(lyr, dict):
        return
    # ponytail: children are filtered to entries with a target, so MapServer sublayer
    # style-overrides (which also live under "layers") never leak in as phantom rows.
    kids = [c for c in (lyr.get("layers") or [])
            if isinstance(c, dict) and (c.get("url") or c.get("itemId"))]
    if lyr.get("url") or lyr.get("itemId") or not kids:
        out.append((section, lyr))          # a GroupLayer shell with children is not itself a ref
    for c in kids:
        _emit(section, c, out)


# --------------------------------------------------------------------------- phase 1-3

def collect(gis, maps, wcag_out=None):
    """One get_data per map (~0.09s).  A failed map yields an error tuple, never vanishes.

    With wcag_out (a list) the accessibility rule rows for every readable map are appended
    to it, from the JSON already fetched, so --wcag costs no extra request.
    """
    base = item_url_base(gis)
    refs, errors, t0, n = [], [], time.time(), len(maps)
    for i, m in enumerate(maps, 1):
        meta = {"MapTitle": m["title"], "MapId": m["id"], "MapOwner": m["owner"],
                "MapAccess": m["access"], "MapViews": m.get("numViews", 0),
                "MapUrl": base + m["id"]}
        found = []
        try:
            # walk() must stay inside the try: nested-corrupt map JSON (a string where
            # baseMap should be, an int under "tables") raises there, not in get_data,
            # and one such map would otherwise abort the whole run with no error row.
            walk(m.get_data(try_json=True), found)
        except Exception as e:
            errors.append((meta, "%s: %s" % (type(e).__name__, e)))
        else:
            for section, lyr in found:
                refs.append((meta, section, lyr))
            if wcag_out is not None:
                wcag_out.extend(wcag_check(meta, m.get("title"), m.get("snippet"),
                                           m.get("description"), found))
        # outside the try/else: a failing last map must not swallow the final progress line
        if i % 25 == 0 or i == n:
            el = time.time() - t0
            print("  %d/%d maps  %.0fs elapsed  ~%.0fs left" % (i, n, el, el / i * (n - i)))
    return refs, errors


def resolve_items(gis, ids):
    """Batched id lookup -> ({id: Item}, {ids whose lookup never completed}).

    18.8x faster than gis.content.get() one at a time, same results.  The second return
    value matters: an id missing because its chunk errored is NOT the same as an id
    missing because the item was deleted, and reporting a throttled request as a deleted
    layer would put up to ITEM_CHUNK false DEAD rows in the report.
    """
    out, unresolved = {}, set()
    for chunk in chunks(sorted(ids), ITEM_CHUNK):
        try:
            hits = gis.content.search(query="id:(" + " OR ".join(chunk) + ")",
                                      max_items=len(chunk) * 2, outside_org=True)
        except Exception as e:
            print("  item batch failed (%s: %s), %d ids unresolved" % (type(e).__name__, e, len(chunk)))
            unresolved.update(chunk)
            continue
        for it in hits:
            out[it.id] = it
    return out, unresolved


def fetch_groups(gis, ids):
    """itemId -> {groupId} and groupId -> title, 80 ids per round trip.

    Ids from a FAILED chunk stay absent so .get() returns None (UNKNOWN); ids from a
    successful chunk are pre-seeded with an empty set (genuinely shared to no group).
    """
    gids, titles = {}, {}
    url = gis._portal.resturl + "content/itemsgroups"
    for chunk in chunks(sorted(ids), GROUP_CHUNK):
        try:
            resp = gis._con.get(url, {"f": "json", "items": ",".join(chunk)}) or {}
        except Exception as e:
            print("  group batch failed (%s: %s), %d ids unknown" % (type(e).__name__, e, len(chunk)))
            continue
        for iid in chunk:
            gids.setdefault(iid, set())
        for gid, blk in resp.items():
            blk = blk or {}
            titles[gid] = blk.get("groupTitle") or gid
            for ent in (blk.get("items") or []):
                iid = (ent or {}).get("itemId")
                if iid:
                    gids.setdefault(iid, set()).add(gid)
    return gids, titles


# --------------------------------------------------------------------------- phase 4 probe

def code_state(code):
    if code in SECURED_CODES:
        return ("SECURED", "code %s" % code)
    # ponytail: ArcGIS REST answers a request for a service that no longer exists with
    # 400 "Invalid URL", not 404.  Verified against 12 distinct dead urls: the org_admin
    # session gets the identical 400, so it is absence, not a permission problem.
    if code in DEAD_CODES:
        return ("DEAD", "code %s" % code)
    return ("UNKNOWN", "code %s" % code)


def probe(url, cache, fails):
    """One anonymous GET url?f=json.  Answers: can the public actually load this service?"""
    key = clean_url(url)
    if key in cache:
        return cache[key]
    h = host(key)
    head, _, query = key.partition("?")
    ask = [q for q in query.split("&") if q and q.split("=")[0].lower() != "f"]
    # ponytail: per-host circuit breaker. Without it one unreachable on-prem box x 50 urls
    # x 8s timeout is a 400s hang. Reset on any completed round trip: a 499 means alive.
    if fails[h] >= MAX_HOST_FAILS:
        return ("UNREACHABLE", "host circuit-open")
    try:
        req = urllib.request.Request(head + "?" + "&".join(ask + ["f=json"]), headers=UA)
        with urllib.request.urlopen(req, timeout=PROBE_TIMEOUT) as r:
            raw = r.read()
    except urllib.error.HTTPError as e:
        fails[h] = 0
        res = code_state(e.code)
    except Exception as e:
        fails[h] += 1
        res = ("UNREACHABLE", type(e).__name__)
    else:
        fails[h] = 0
        try:
            body = json.loads(raw.decode("utf-8", "replace"))
        except ValueError:
            body = None
        if not isinstance(body, dict):
            res = ("UNKNOWN", "non-JSON response")
        else:
            err = body.get("error")
            # A third-party JSON endpoint may answer {"error": "text"} or a list as the code.
            code = err.get("code", 0) if isinstance(err, dict) else None
            if not err:
                res = ("PUBLIC", "")
            elif isinstance(code, (int, float, str)):
                res = code_state(code)
            else:
                res = ("UNKNOWN", "error reply of unknown shape")
    cache[key] = res
    return res


# --------------------------------------------------------------------------- phase 5 verdict

def effective(access, groups):
    """'shared' to zero known groups is an owner-only audience, the same as 'private'."""
    return "private" if access == "shared" and groups is not None and not groups else access


def covers(la, lg, lown, ma, mg, mown, titles=None):
    """Is every viewer of the map also a viewer of the layer?  (True / False / None=unknown)"""
    la, ma = effective(la, lg), effective(ma, mg)
    if la == "public":
        return True, "public layer covers every audience"
    if ma == "private":
        if la == "private" and lown != mown:
            return False, "private layer owned by %s" % (lown or "another user")
        return True, "private map; the audience is the owner only"
    if la == "org":
        if ma == "public":
            return False, "org-only layer in a public map; anonymous viewers cannot load it"
        if ma == "shared":
            # Esri lets a group take members from other organisations, who cannot load an
            # org-only item.  Membership is not read here, so this cannot be called OK.
            return None, ("org-only layer in a group-shared map; group members from another "
                          "organisation cannot load it, and membership is not read")
        return True, "org layer covers org and group audiences"
    if la == "shared":
        if ma != "shared":
            return False, "%s map, group-restricted layer" % ma
        # Both sides must be known.  Treating an unavailable MAP group set as "shared to no
        # groups" would silently score a never-performed check as OK, the one failure mode
        # this tool must not have.  A failed fetch_groups chunk is exactly how that happens.
        if lg is None or mg is None:
            return None, "group membership unavailable"
        missing = set(mg) - lg
        if missing:
            names = sorted((titles or {}).get(g, g) for g in missing)
            return False, "map shared to groups not on the layer: " + ", ".join(names)
        return True, "layer shared to every group the map is shared to"
    return False, "private layer in a %s map" % ma


def verdict(state, detail, la, lg, lown, ma, mg, mown, titles=None):
    if state == "MISSING_ITEM":
        return "DEAD", "layer item deleted or inaccessible, and no service url to probe"
    if state == "DEAD":
        return "DEAD", "service gone (%s)" % (detail or "no longer published")
    if state == "EMBEDDED":
        return "OK", "feature collection embedded in the map; nothing external to share"
    if state in ("UNREACHABLE", "UNKNOWN", "NO_TARGET"):
        return "UNKNOWN", detail or "layer reference has no item id and no service url"
    if state == "SECURED":
        if ma == "public":
            return "BROKEN", "secured service in a public map; anonymous viewers cannot load it"
        return "UNKNOWN", "secured service; on-prem credentials are not verifiable from AGOL"
    ok, why = covers(la, lg, lown, ma, mg, mown, titles)
    return ("UNKNOWN" if ok is None else "OK" if ok else "BROKEN"), why


def build(refs, items, gids, gtitle, probes, fails, org_id, unresolved=()):
    rows = []
    for meta, section, lyr in refs:
        # A url is probed and reported only after clean_url(): the report is a file people
        # share, and a token in a layer url would otherwise be written into it and its links.
        iid, url = lyr.get("itemId") or "", clean_url(lyr.get("url"))
        it = items.get(iid) if iid else None
        la = lown = None
        lg, detail = None, ""
        if it is not None:
            state, la, lown, lg = "item", it["access"], it["owner"], gids.get(it.id)
        elif url:
            # Probe the SERVICE, not just the item: a deleted item whose service is still
            # public means the layer still draws.  Short-circuiting on the batch miss
            # would call that DEAD, a whole false-positive class, removed by one branch.
            state, detail = probe(url, probes, fails)
            if state == "PUBLIC":
                la, lown, lg = "public", "", set()
        elif iid in unresolved:
            # its lookup errored, so we never learned anything. Do not claim it is deleted
            state, detail = "UNKNOWN", "item lookup failed; membership not verified"
        elif iid:
            state = "MISSING_ITEM"
        elif lyr.get("featureCollection"):
            state = "EMBEDDED"
        else:
            state = "NO_TARGET"
        mg = gids.get(meta["MapId"])
        v, why = verdict(state, detail, la, lg, lown, meta["MapAccess"], mg,
                         meta["MapOwner"], gtitle)
        # Overexposed is a DIFFERENT failure mode on a different axis; never in Findings.
        # The org-owned gate suppresses public Esri basemap noise.
        over = "Yes" if (la == "public" and meta["MapAccess"] != "public"
                         and it is not None and it.get("orgId") == org_id) else ""
        row = dict(meta)
        row.update({
            "Section": section, "LayerTitle": lyr.get("title") or "", "LayerType": lyr.get("layerType") or "",
            "LayerItemId": iid, "LayerUrl": url, "LayerHost": host(url), "LayerOwner": lown or "",
            "LayerAccess": la or "", "LayerState": state, "Verdict": v, "Reason": why,
            "Severity": "" if v == "OK" else SEV.get(meta["MapAccess"], "4 Info"),
            "Overexposed": over,
            "MapGroups": "; ".join(sorted(gtitle.get(g, g) for g in (mg or ()))),
            "LayerGroups": "; ".join(sorted(gtitle.get(g, g) for g in (lg or ()))),
        })
        rows.append(row)
    return rows


def error_rows(errors):
    """An error IS a finding. It lands in Findings automatically, so there is no Errors sheet."""
    out = []
    for meta, msg in errors:
        row = dict(meta)
        row.update({"Section": "<error>", "Verdict": "UNKNOWN", "Reason": msg,
                    "Severity": SEV.get(meta["MapAccess"], "4 Info"), "LayerState": "UNKNOWN"})
        out.append(row)
    return out


# --------------------------------------------------------------------------- app edges

def map_candidates(data):
    """Every item-id-shaped string under a MAP_KEYS key, anywhere in an app's JSON.

    Iterative, so a hostile deeply nested document cannot hit the recursion limit.  This is
    only a candidate list: build_edges keeps an id when it resolves to a Web Map item.
    """
    found, stack = set(), [data]
    while stack:
        node = stack.pop()
        if isinstance(node, dict):
            for key, val in node.items():
                if key in MAP_KEYS and isinstance(val, str) and ID_RE.match(val):
                    found.add(val.lower())
                if isinstance(val, (dict, list)):
                    stack.append(val)
        elif isinstance(node, list):
            stack.extend(v for v in node if isinstance(v, (dict, list)))
    return sorted(found)


def collect_apps(gis, limit):
    """Search each app type, read each app's JSON once -> ([(meta, [candidate ids])], errors)."""
    base = item_url_base(gis)
    cap = MAX_APPS if limit is None else limit
    apps, errors, seen = [], [], set()
    for kind in EDGE_TYPES:
        for app in gis.content.search("", item_type=kind, max_items=cap):
            if app["id"] in seen:
                continue
            seen.add(app["id"])
            meta = {"AppType": app["type"], "AppTitle": app["title"], "AppId": app["id"],
                    "AppOwner": app["owner"], "AppAccess": app["access"],
                    "AppViews": app.get("numViews", 0), "AppUrl": base + app["id"]}
            try:
                apps.append((meta, map_candidates(app.get_data(try_json=True))))
            except Exception as e:
                errors.append((meta, "%s: %s" % (type(e).__name__, e)))
    return apps, errors


def edge_verdict(aa, ag, aown, ma, mg, mown, titles=None):
    """Can every viewer of the APP (access aa, groups ag, owner aown) load the embedded MAP?

    The map plays the part the layer plays in covers(), and the app plays the map.
    """
    aa, ma = effective(aa, ag), effective(ma, mg)
    ok, why = covers(ma, mg, mown, aa, ag, aown, titles)
    if ok is None:
        if ma == "org":
            return "UNKNOWN", ("app is shared to groups and the web map is org-only; group "
                               "members from another organisation cannot load it, and "
                               "membership is not read")
        return "UNKNOWN", why
    if ok:
        return "OK", "every viewer of the app can load the web map"
    if aa == "shared" and ma == "shared":
        names = sorted((titles or {}).get(g, g) for g in set(ag) - mg)
        return "BROKEN", "app is shared to groups the web map is not: " + ", ".join(names)
    if aa == "private" and ma == "private":
        return "BROKEN", "web map is private to %s; the app owner cannot load it" % (mown or "another user")
    return "BROKEN", "app is shared %s but the web map is shared %s" % (aa, ma)


def build_edges(apps, errors, items, gids, gtitle, base, unresolved=()):
    """-> (edge rows, Counter of candidate ids that did not become an edge).

    A candidate that resolves to a layer, a scene or any non-map item is dropped, as is one
    that resolves to nothing; the counter says how many of each so the loss is visible.
    """
    rows, skipped = [], collections.Counter()
    for meta, cands in apps:
        for cid in cands:
            it = items.get(cid)
            if it is None:
                skipped["lookup failed" if cid in unresolved else "no item"] += 1
                continue
            if it["type"] != "Web Map":
                skipped["not a web map"] += 1
                continue
            ag, mg = gids.get(meta["AppId"]), gids.get(cid)
            v, why = edge_verdict(meta["AppAccess"], ag, meta["AppOwner"],
                                  it["access"], mg, it["owner"], gtitle)
            row = dict(meta)
            row.update({
                "Severity": "" if v == "OK" else SEV.get(meta["AppAccess"], "4 Info"),
                "Verdict": v, "Reason": why, "MapTitle": it["title"], "MapId": cid,
                "MapOwner": it["owner"], "MapAccess": it["access"], "MapUrl": base + cid,
                "AppGroups": "; ".join(sorted(gtitle.get(g, g) for g in (ag or ()))),
                "MapGroups": "; ".join(sorted(gtitle.get(g, g) for g in (mg or ()))),
            })
            rows.append(row)
    for meta, msg in errors:
        row = dict(meta)
        row.update({"Severity": SEV.get(meta["AppAccess"], "4 Info"), "Verdict": "UNKNOWN",
                    "Reason": msg})
        rows.append(row)
    return rows, skipped


# --------------------------------------------------------------------------- wcag rule pack

def dig(obj, *keys):
    """obj[k1][k2]... where any level that is not a dict yields None instead of raising."""
    for key in keys:
        if not isinstance(obj, dict):
            return None
        obj = obj.get(key)
    return obj


def as_list(value):
    return value if isinstance(value, list) else []


def rgb_of(color):
    """[r, g, b, 255] -> (r, g, b).  Anything else -> None.

    A color that is not fully opaque is drawn over the basemap, which this tool cannot see,
    so its contrast cannot be judged and the rule stays silent.
    """
    if (isinstance(color, (list, tuple)) and len(color) == 4
            and all(isinstance(v, (int, float)) and 0 <= v <= 255 for v in color)
            and color[3] == 255):
        return tuple(color[:3])
    return None


def luminance(rgb):
    """WCAG 2.2 relative luminance of an sRGB (r, g, b) triple of 0..255."""
    lin = []
    for c in rgb:
        s = c / 255.0
        lin.append(s / 12.92 if s <= 0.04045 else ((s + 0.055) / 1.055) ** 2.4)
    return 0.2126 * lin[0] + 0.7152 * lin[1] + 0.0722 * lin[2]


def contrast(a, b):
    """WCAG 2.2 contrast ratio of two (r, g, b) triples: (L1 + 0.05) / (L2 + 0.05)."""
    la, lb = luminance(a), luminance(b)
    return (max(la, lb) + 0.05) / (min(la, lb) + 0.05)


def label_issues(lyr):
    """[(ratio, needed)] for each label of a layer whose text/halo contrast is too low.

    Only labels that draw (showLabels true) and only a halo that exists (haloSize over 0)
    are judged, because the halo is then the background the text sits on.
    """
    if lyr.get("showLabels") is not True:
        return []
    out = []
    for info in as_list(dig(lyr, "layerDefinition", "drawingInfo", "labelingInfo")):
        sym = dig(info, "symbol") or {}
        text, halo = rgb_of(dig(sym, "color")), rgb_of(dig(sym, "haloColor"))
        size = dig(sym, "haloSize")
        if text is None or halo is None or not isinstance(size, (int, float)) or size <= 0:
            continue
        fsize, weight = dig(sym, "font", "size"), dig(sym, "font", "weight")
        fsize = fsize if isinstance(fsize, (int, float)) else 0
        large = fsize >= 18 or (weight == "bold" and fsize >= 14)
        need = LARGE_TEXT_RATIO if large else TEXT_RATIO
        ratio = contrast(text, halo)
        if ratio < need:
            out.append((ratio, need))
    return out


def symbol_shape(sym):
    """What a user could tell apart without color: marker/line/fill kind, style, size, outline."""
    if not isinstance(sym, dict):
        return None
    kind = sym.get("type")
    if kind == "esriSMS":
        shape = (sym.get("style"), sym.get("size"))
    elif kind == "esriSLS":
        shape = (sym.get("style"), sym.get("width"))
    elif kind == "esriSFS":
        shape = (sym.get("style"),)
    else:
        return None                         # picture and CIM symbols differ by more than color
    ol = sym.get("outline")
    return (kind, shape, (ol.get("style"), ol.get("width")) if isinstance(ol, dict) else None)


def color_only(lyr):
    """True when a unique-value renderer's classes differ by fill color and nothing else."""
    rend = dig(lyr, "layerDefinition", "drawingInfo", "renderer")
    if not isinstance(rend, dict) or rend.get("type") != "uniqueValue":
        return False
    # uniqueValueGroups wins over uniqueValueInfos when present, as the web map spec says.
    groups = as_list(rend.get("uniqueValueGroups"))
    if groups:
        classes = [c for g in groups for c in as_list(dig(g, "classes")) if isinstance(c, dict)]
    else:
        classes = [c for c in as_list(rend.get("uniqueValueInfos")) if isinstance(c, dict)]
    syms = [c.get("symbol") for c in classes]
    shapes = [symbol_shape(s) for s in syms]
    # repr() keeps a corrupt, unhashable value from raising inside the set.
    colors = set(repr(s["color"]) for s in syms
                 if isinstance(s, dict) and isinstance(s.get("color"), list))
    return (len(syms) >= 2 and None not in shapes and len(set(repr(x) for x in shapes)) == 1
            and len(colors) >= 2)


def wcag_check(meta, title, snippet, description, found):
    """Accessibility rule rows for one web map.  Pure: the JSON is already in hand.

    found is walk()'s (section, layer) list.  The base map is Esri's content, not the
    author's, so its layers are never judged.
    """
    hits = []
    norm = " ".join(str(title or "").lower().split())
    if not norm or norm in PLACEHOLDER_TITLES:
        hits.append(("map-title-placeholder", "", "the title is %s" % (repr(title) if norm else "empty")))
    # shortcut: tags are stripped with a regex, not parsed; an entity-only description
    # (&nbsp;) counts as text.  Upgrade if a real description trips it.
    plain = re.sub(r"<[^>]*>", " ", str(description or ""))
    if not str(snippet or "").strip() and not plain.strip():
        hits.append(("map-no-summary", "", "no summary and no description on the item"))
    for section, lyr in found:
        if section == "baseMap":
            continue
        name = str(lyr.get("title") or "").strip()
        label = name or lyr.get("itemId") or clean_url(lyr.get("url"))
        if not name:
            hits.append(("layer-no-title", label, "the layer has no title"))
        if lyr.get("disablePopup") is True:
            hits.append(("layer-popup-off", label, "disablePopup is true"))
        if color_only(lyr):
            hits.append(("renderer-color-only", label,
                         "the unique-value classes share one shape and differ only by color"))
        for ratio, need in label_issues(lyr):
            hits.append(("label-halo-contrast", label,
                         "label text on its halo is %.2f:1, needs %.1f:1" % (ratio, need)))
    rows = []
    for rule, label, detail in hits:
        crit, level, _ = WCAG_RULES[rule]
        rows.append({"Severity": SEV.get(meta["MapAccess"], "4 Info"), "Rule": rule,
                     "Criterion": crit, "Level": level, "MapTitle": meta["MapTitle"],
                     "MapId": meta["MapId"], "MapOwner": meta["MapOwner"],
                     "MapAccess": meta["MapAccess"], "MapViews": meta["MapViews"],
                     "MapUrl": meta["MapUrl"], "Layer": label, "Detail": detail})
    return rows


# --------------------------------------------------------------------------- phase 6-7

def deep_stats(gis, urls):
    """--deep only.  Cached BY URL: many web-map layers share one FeatureServer endpoint."""
    if FeatureLayer is None:
        load_arcgis()
    out = {}
    for u in sorted(urls):
        if not FS_RE.search(u.lower()):
            continue
        try:
            fl = FeatureLayer(u, gis)
            out[u] = (fl.query(return_count_only=True),
                      iso((fl.properties.get("editingInfo") or {}).get("lastEditDate")))
        except Exception as e:
            out[u] = ("", type(e).__name__)
    return out


def summary_rows(rows, meta, edges=None, wcag=None):
    """(Section, Metric, Value) triples.  Plain Python, so the counting is testable offline."""
    out = [("Run", k, v) for k, v in meta.items()]
    counts = collections.Counter(r.get("Verdict") for r in rows)
    for v in VERDICTS:
        out.append(("Verdicts", v, counts[v]))
    out.append(("Verdicts", "Overexposed (public layer, non-public map)",
                sum(1 for r in rows if r.get("Overexposed") == "Yes")))
    bad = [r for r in rows if r.get("Verdict") != "OK"]
    worst = collections.Counter((r.get("MapTitle"), r.get("MapAccess"), r.get("MapViews")) for r in bad)
    for (t, a, vw), n in sorted(worst.items(), key=lambda kv: (-kv[1], str(kv[0])))[:15]:
        out.append(("Worst maps", "%s (%s, %s views)" % (t, a, vw), n))
    owners = collections.Counter(r.get("MapOwner") for r in bad)
    for owner, n in sorted(owners.items(), key=lambda kv: (-kv[1], str(kv[0])))[:10]:
        out.append(("Top owners", owner, n))          # this is who you email
    if edges is not None:
        ec = collections.Counter(r.get("Verdict") for r in edges)
        for v in VERDICTS:
            out.append(("Edges", v, ec[v]))
    if wcag is not None:
        wc = collections.Counter(r.get("Rule") for r in wcag)
        for rule in WCAG_RULES:
            out.append(("WCAG", rule, wc[rule]))
    return out


def frame(rows, cols):
    """A DataFrame of rows with every missing cell as "".

    The blanks are filled in Python, before pandas sees them, so no NaN is ever created.
    fillna("") on a column that held only NaN is float64 in pandas 3 and warns that it will
    raise later, and fillna on object columns warns about downcasting in pandas 2.  Either
    would end a multi-minute run at the report step.
    """
    body = [["" if r.get(c) is None else r.get(c) for c in cols] for r in rows]
    return pd.DataFrame(body, columns=cols)


def summarize(rows, meta, edges=None, wcag=None):
    return pd.DataFrame(summary_rows(rows, meta, edges, wcag),
                        columns=["Section", "Metric", "Value"])


def save_report(path, findings, all_rows, summary, extra=()):
    """write_excel, but never lose a completed audit to a locked output file.

    Re-running with the previous report still open in Excel is the single most likely
    real-world failure for this audience, and it would otherwise discard a run that has
    already spent minutes of API calls.
    """
    try:
        write_excel(path, findings, all_rows, summary, extra)
        return path
    except PermissionError:
        stem, ext = os.path.splitext(path)
        alt = "%s_%s%s" % (stem, datetime.datetime.now().strftime("%H%M%S"), ext)
        print("  %s is locked (open in Excel?); writing %s instead" % (path, alt))
        write_excel(alt, findings, all_rows, summary, extra)
        return alt


def defuse(ws):
    """Stop a hostile item title from becoming a live formula in the report.

    A layer or map title is attacker-controllable. Any org member can name an item
    `=cmd|' /C calc'!A0`, and openpyxl faithfully types a leading =/+/-/@ as a formula.
    Forcing the cell's type to string keeps the text exactly as-is (an apostrophe prefix
    would corrupt the value) while making it inert.
    """
    for row in ws.iter_rows():
        for cell in row:
            if cell.data_type == "f":
                cell.data_type = "s"


def scrub(df):
    """Drop control characters openpyxl refuses to write.

    Unreachable through the AGOL web UI, reachable through the REST API, and it would
    otherwise raise IllegalCharacterError at the very end of a multi-minute run.
    """
    # ponytail: guard on the value, not the column dtype. pandas 3 types text columns as
    # StringDtype rather than object, so a `dtype == object` test silently scrubs nothing.
    out = df.copy()
    for col in out.columns:
        out[col] = out[col].map(
            lambda v: ILLEGAL_CHARACTERS_RE.sub("", v) if isinstance(v, str) else v)
    return out


def write_excel(path, findings, all_rows, summary, extra=()):
    """extra: (sheet name, DataFrame, hyperlink columns) for the optional --edges/--wcag sheets."""
    sheets = [("Findings", findings, ("MapUrl", "LayerUrl")),
              ("AllLayers", all_rows, ("MapUrl", "LayerUrl")),
              ("Summary", summary, ())] + list(extra)
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        for name, df, links in sheets:
            df = scrub(df)
            df.to_excel(xw, index=False, sheet_name=name)
            ws = xw.sheets[name]
            defuse(ws)
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for i, col in enumerate(df.columns, 1):
                widths = [len(str(col))] + [len(str(v)) for v in df[col].head(WIDTH_SAMPLE_ROWS)]
                ws.column_dimensions[get_column_letter(i)].width = min(max(widths) + 2, MAX_COL_WIDTH)
                if col not in links:
                    continue
                for r in range(2, len(df) + 2):
                    cell = ws.cell(row=r, column=i)
                    if cell.value:
                        cell.hyperlink = str(cell.value)
                        cell.style = "Hyperlink"


# --------------------------------------------------------------------------- self-test
# Everything below the decision core is test scaffolding.  Nothing here runs in a live audit.

_MISSING = object()


def _hid(n):
    """A 32-hex-digit item id, the shape of a real one, from a small number."""
    return "%032x" % n


@contextlib.contextmanager
def _patched(**kw):
    """Swap module globals for the length of a with block, then put every one back."""
    g = globals()
    old = dict((k, g.get(k, _MISSING)) for k in kw)
    g.update(kw)
    try:
        yield
    finally:
        for k, v in old.items():
            if v is _MISSING:
                g.pop(k, None)
            else:
                g[k] = v


@contextlib.contextmanager
def _setattr(obj, **kw):
    old = dict((k, getattr(obj, k)) for k in kw)
    for k, v in kw.items():
        setattr(obj, k, v)
    try:
        yield
    finally:
        for k, v in old.items():
            setattr(obj, k, v)


@contextlib.contextmanager
def _environ(**kw):
    """Set (or, for None, remove) environment variables, then restore them."""
    old = dict((k, os.environ.get(k)) for k in kw)
    for k, v in kw.items():
        if v is None:
            os.environ.pop(k, None)
        else:
            os.environ[k] = v
    try:
        yield
    finally:
        for k, v in old.items():
            if v is None:
                os.environ.pop(k, None)
            else:
                os.environ[k] = v


@contextlib.contextmanager
def _modules(**kw):
    """Plant entries in sys.modules; a None entry makes that import raise ModuleNotFoundError."""
    old = dict((k, sys.modules.get(k, _MISSING)) for k in kw)
    sys.modules.update(kw)
    try:
        yield
    finally:
        for k, v in old.items():
            if v is _MISSING:
                sys.modules.pop(k, None)
            else:
                sys.modules[k] = v


@contextlib.contextmanager
def _quiet():
    """Capture stdout and stderr; yields the (out, err) buffers."""
    out, err = io.StringIO(), io.StringIO()
    with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
        yield out, err


def _exits(fn):
    """The code or message of the SystemExit that fn raises; AssertionError if it does not exit."""
    try:
        fn()
    except SystemExit as e:
        return e.code
    raise AssertionError("expected SystemExit")


class _Item(dict):
    """The slice of an arcgis Item the audit touches: dict access, .id and .get_data()."""

    def __init__(self, iid, title="t", owner="u", access="private", kind="Web Map",
                 data=None, **extra):
        dict.__init__(self, id=iid, title=title, owner=owner, access=access, type=kind, **extra)
        self.id = iid
        self._data = data

    def get_data(self, try_json=None):
        # The real call returns parsed JSON only for try_json=True; otherwise raw text.
        if isinstance(self._data, Exception):
            raise self._data
        return self._data if try_json else json.dumps(self._data)

    def __getattr__(self, name):
        # The README HARD RULE: an unknown attribute on a real Item costs a round trip.
        raise AssertionError("attribute probe on an Item: %s" % name)


class _Content(object):
    def __init__(self, items, fail_ids=(), leak=None):
        self.items, self.fail_ids, self.leak = items, set(fail_ids), leak or {}

    def search(self, query="", item_type=None, max_items=100, outside_org=False):
        if query.startswith("id:("):
            ids = query[4:-1].split(" OR ")
            if self.fail_ids & set(ids):
                raise RuntimeError("throttled")
            hits = [self.items[i] for i in ids if i in self.items]
        else:
            hits = [it for it in self.items.values() if it["type"] == item_type]
            hits += [self.items[i] for i in self.leak.get(item_type, ())]
        return hits[:max_items]


class _Con(object):
    """gis._con.get over a real socket, so fetch_groups' request really leaves the process."""

    def get(self, url, params):
        with urllib.request.urlopen(url + "?" + urllib.parse.urlencode(params), timeout=5) as r:
            return json.loads(r.read().decode("utf-8"))


class _Gis(object):
    def __init__(self, rest, items, fail_ids=(), leak=None):
        self.url = "https://portal.test/"
        self._portal = types.SimpleNamespace(resturl=rest)
        self._con = _Con()
        self.content = _Content(items, fail_ids, leak)
        self.properties = {"name": "Test Org", "id": "org1"}
        self.users = types.SimpleNamespace(me=types.SimpleNamespace(username="admin"))


_MEMBERS = {"m1": ["g1", "g2"], "l1": ["g1"], "e1": [], "d1": ["g1"],
            _hid(1): ["g1", "g2"], _hid(10): ["g1"], _hid(20): ["g1"]}
_GTITLES = {"g1": "Alpha", "g2": "Beta"}
_ROUTES = {
    "/ok/FeatureServer/0": (200, '{"name": "x"}'),
    "/token/FeatureServer/0": (200, '{"error": {"code": 499, "message": "Token Required"}}'),
    "/forbidden/FeatureServer/0": (403, "{}"),
    "/gone/FeatureServer/0": (400, '{"error": {"code": 400}}'),
    "/missing/FeatureServer/0": (404, ""),
    "/boom/FeatureServer/0": (500, ""),
    "/html/FeatureServer/0": (200, "<html></html>"),
    "/list/FeatureServer/0": (200, "[1, 2]"),
    "/odd/FeatureServer/0": (200, '{"error": {"code": 418}}'),
    "/errstr/FeatureServer/0": (200, '{"error": "forbidden"}'),
    "/errlist/FeatureServer/0": (200, '{"error": ["x"]}'),
    "/errnum/FeatureServer/0": (200, '{"error": 5}'),
    "/errcode/FeatureServer/0": (200, '{"error": {"code": [1]}}'),
    "/errstrcode/FeatureServer/0": (200, '{"error": {"code": "498"}}'),
}


def _stub(log):
    """A local http.server that answers like the services and the portal.  Returns the server."""

    class Handler(http.server.BaseHTTPRequestHandler):
        def do_GET(self):
            parts = urllib.parse.urlparse(self.path)
            log.append(self.path)
            if parts.path.startswith("/drop"):
                self.close_connection = True        # no reply at all: a dead host
                return
            if parts.path == "/sharing/rest/content/itemsgroups":
                ids = urllib.parse.parse_qs(parts.query)["items"][0].split(",")
                if "boom" in ids:
                    status, body = 500, ""
                else:
                    blocks = {}
                    for gid, title in _GTITLES.items():
                        mine = [i for i in ids if gid in _MEMBERS.get(i, ())]
                        if mine:
                            blocks[gid] = {"groupTitle": title,
                                           "items": [{"itemId": i} for i in mine]}
                    status, body = 200, json.dumps(blocks)
            elif parts.path == "/needs/query":
                # answers only when the layer's own query string arrives
                has = "where" in urllib.parse.parse_qs(parts.query)
                status, body = (200, '{"features": []}') if has else (400, '{"error": {"code": 400}}')
            else:
                status, body = _ROUTES.get(parts.path, (404, ""))
            raw = body.encode("utf-8")
            self.send_response(status)
            self.send_header("Content-Length", str(len(raw)))
            self.end_headers()
            self.wfile.write(raw)

        def log_message(self, *args):
            pass

    srv = http.server.HTTPServer(("127.0.0.1", 0), Handler)
    thread = threading.Thread(target=srv.serve_forever)
    thread.daemon = True
    thread.start()
    return srv


class _T(object):
    """The check()/raises() harness: count what passed, remember what failed by label."""

    def __init__(self):
        self.passed = 0
        self.failed = []

    def check(self, cond, label):
        if cond:
            self.passed += 1
            print("PASS  %s" % label)
        else:
            self.failed.append(label)
            print("FAIL  %s" % label)

    def raises(self, fn, label, exc=ValueError):
        try:
            fn()
        except exc:
            self.check(True, label)
        except Exception as e:
            self.check(False, "%s (wrong exception %r)" % (label, e))
        else:
            self.check(False, "%s (no error raised)" % label)


def _footer(passed, failed, skipped=False):
    if skipped:
        print("SKIP  the Excel report assertions did not run, so this count is lower than "
              "the full suite")
    print("-" * 68)
    total = passed + len(failed)
    if failed:
        print("%d assertions, %d failed" % (total, len(failed)))
        for label in failed:
            print("  FAILED: %s" % label)
        return 1
    print("%d assertions, 0 failed" % total)
    return 0


def self_test():
    """Run every assertion in a temp directory that is removed even when one raises.

    No network beyond 127.0.0.1, no credential, no portal.  The only files written are the
    ones under the temp directory.
    """
    tmp = tempfile.mkdtemp(prefix="sightline-selftest-")
    log = []
    srv = _stub(log)
    # a proxy named in the environment must not capture a request to the local stub
    urllib.request.install_opener(urllib.request.build_opener(urllib.request.ProxyHandler({})))
    try:
        return _self_test(tmp, srv, log)
    finally:
        urllib.request.install_opener(None)
        srv.shutdown()
        srv.server_close()
        shutil.rmtree(tmp, ignore_errors=True)


def _have_report_libs():
    try:
        import pandas
        import openpyxl
    except ModuleNotFoundError:
        return False
    return bool(pandas and openpyxl)


def _self_test(tmp, srv, log):
    t = _T()
    base = "http://127.0.0.1:%d" % srv.server_address[1]
    rest = base + "/sharing/rest/"
    print("sightline self-test: 127.0.0.1 stub and temporary files only, no portal, no login")
    print("-" * 68)
    _t_harness(t)
    _t_helpers(t)
    _t_matrix(t)
    _t_core(t)
    _t_probe(t, base, log)
    _t_portal(t, rest)
    _t_edges(t, rest)
    _t_wcag(t)
    _t_summary(t)
    _t_login(t)
    _t_cli(t)
    _t_source(t)
    with _modules(pandas=None):
        t.check(_have_report_libs() is False, "a missing pandas is detected, so the Excel assertions would skip")
    probe_t = _T()
    with _quiet() as (out, _):
        _t_report_gate(probe_t, tmp, rest, False)
    t.check(probe_t.passed == 0 and "SKIP  report writer" in out.getvalue(),
            "without the report libraries the Excel assertions are skipped by name, not silently")
    _t_report_gate(t, tmp, rest, _have_report_libs())
    return _footer(t.passed, t.failed, not _have_report_libs())


def _t_report_gate(t, tmp, rest, have):
    if have:
        _t_report(t, tmp, rest)
    else:
        print("SKIP  report writer: pandas and openpyxl are not installed, so the Excel "
              "assertions did not run")


def _t_helpers(t):
    """The patch helpers restore what was there, or a test would leak into the next one."""
    check = t.check
    with _environ(SL_PROBE="outer"):
        with _environ(SL_PROBE="inner", SL_PROBE_GONE=None):
            inside = os.environ["SL_PROBE"]
        after = os.environ["SL_PROBE"]
    check(inside == "inner" and after == "outer" and "SL_PROBE" not in os.environ,
          "_environ() restores a variable that existed and removes one that did not")
    with _modules(sl_probe_mod=types.ModuleType("one")):
        with _modules(sl_probe_mod=types.ModuleType("two")):
            inside = sys.modules["sl_probe_mod"].__name__
        after = sys.modules["sl_probe_mod"].__name__
    check(inside == "two" and after == "one" and "sl_probe_mod" not in sys.modules,
          "_modules() restores an entry that existed and removes one that did not")
    with _patched(SL_PROBE_GLOBAL=1):
        inside = globals()["SL_PROBE_GLOBAL"]
    check(inside == 1 and "SL_PROBE_GLOBAL" not in globals() and iso(0) == "",
          "_patched() removes a global that did not exist")
    box = types.SimpleNamespace(x=1)
    with _setattr(box, x=2):
        inside = box.x
    check(inside == 2 and box.x == 1, "_setattr() restores the attribute")


def _t_harness(t):
    """The harness must be able to fail, or every PASS below it means nothing."""
    check, raises = t.check, t.raises
    probe_t = _T()
    with _quiet():
        probe_t.check(False, "probe: a false condition must be recorded as a failure")
        probe_t.raises(lambda: None, "probe: a call that raises nothing must fail")
        probe_t.raises(lambda: [][0], "probe: the wrong exception must fail")
        probe_t.raises(lambda: [][0], "probe: the right exception must pass", exc=IndexError)
    check(len(probe_t.failed) == 3 and probe_t.passed == 1,
          "check() and raises() really do record a failure  <-- pinned defect")
    raises(lambda: int("x"), "and raises() accepts the ValueError it is looking for")
    raises(lambda: _Item("x").title_of_nothing, "the fake Item refuses an attribute lookup, as the HARD RULE demands",
           exc=AssertionError)
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        red = _footer(1, ["probe one", "probe two"])
        green = _footer(2, [])
    out = buf.getvalue()
    check(red == 1 and "3 assertions, 2 failed" in out and "  FAILED: probe two" in out,
          "the footer reports failures by count and by name, and returns 1  <-- pinned defect")
    check(green == 0 and out.rstrip().endswith("2 assertions, 0 failed") and "SKIP" not in out,
          "a green footer is the last line and returns 0")
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        short = _footer(2, [], True)
    lines = buf.getvalue().splitlines()
    check(short == 0 and lines[0].startswith("SKIP  the Excel report assertions did not run")
          and lines[-1] == "2 assertions, 0 failed",
          "a run that skipped the Excel assertions says so above a footer that still ends the same way  <-- pinned defect")


def _t_matrix(t):
    """The sharing matrix of the original --self-check, one weight per audience pair."""
    check = t.check
    u = "u"
    g1, g2 = {"g1"}, {"g1", "g2"}
    titles = {"g1": "Alpha", "g2": "Beta"}

    def c(la, ma, lg=None, mg=None, lown=u, mown=u):
        return covers(la, lg, lown, ma, mg, mown, titles)[0]

    # All 18 audience pairs.  (layer access, map access, expected covers())  'MISSING' pairs
    # are decided by state, not covers, and are asserted via verdict below.  Each pair counts
    # once; the audit this table came from weighted them by measured counts, which stay out
    # of a public repo.
    matrix = [
        ("public", "public", True), ("public", "private", True),
        ("shared", "shared", True), ("MISSING", "private", None),
        ("private", "private", True), ("org", "org", True),
        ("shared", "private", True), ("org", "private", True),
        ("MISSING", "shared", None), ("private", "shared", False),
        ("public", "org", True), ("MISSING", "public", None),
        ("org", "public", False), ("private", "org", False),
        ("MISSING", "org", None), ("shared", "org", False),
        ("private", "public", False), ("org", "shared", None),
    ]
    check(len(matrix) == 18, "the matrix has all 18 audience pairs")
    broken = dead = 0
    for la, ma, want in matrix:
        state = "MISSING_ITEM" if la == "MISSING" else "item"
        lg = g1 if la == "shared" else None
        mg = g1 if ma == "shared" else set()
        if want is not None:
            check(c(la, ma, lg, mg) is want, "matrix row map %s, layer %s" % (ma, la))
        v = verdict(state, "", la, lg, u, ma, mg, u, titles)[0]
        if v == "BROKEN":
            broken += 1
        elif v == "DEAD":
            dead += 1
    # shared/shared is group-dependent and scores OK with equal groups, so it lands in neither
    # bucket.  Four pairs were broken under the old rank model; the fifth, org-map with a
    # group-shared layer, is deliberately reclassified: 'org' and 'shared' are incomparable,
    # not equal.
    check((broken, dead) == (5, 4), "the pairs total 5 broken and 4 dead, got %s" % ((broken, dead),))

    # group containment, both directions, and the reason names the missing title
    ok, why = covers("shared", g1, u, "shared", g2, u, titles)
    check(ok is False and "Beta" in why, "a map shared to a group the layer lacks is broken and names it")
    check(covers("shared", g2, u, "shared", g1, u, titles)[0] is True,
          "a layer shared to more groups than the map still covers it")
    # a map shared to zero groups normalizes to private.  Probed with an unknown-group layer:
    # without normalization this is UNKNOWN, with it the map is owner-only -> OK.
    check(verdict("item", "", "shared", None, "other", "shared", set(), u, titles)[0] == "OK",
          "a map shared to zero groups is owner-only")
    check(verdict("item", "", "shared", None, "other", "shared", g1, u, titles)[0] == "UNKNOWN",
          "unknown layer groups are unknown, not broken")
    # a private layer owned by someone else is invisible even inside a private map
    check(c("private", "private", lown="bob", mown="alice") is False,
          "a private layer owned by someone else is invisible in a private map")
    check(c("private", "private", lown="alice", mown="alice") is True,
          "a private layer owned by the map owner is visible in a private map")
    check(verdict("item", "", "shared", g1, u, "shared", None, u, titles)[0] == "UNKNOWN",
          "unknown map groups are unknown, never a silent OK")
    check(covers("shared", g1, u, "shared", None, u, titles)[0] is None,
          "covers() itself answers None for unknown map groups")
    unsure, why = covers("org", None, u, "shared", g1, u, titles)
    check(unsure is None and "another organisation" in why
          and verdict("item", "", "org", None, u, "shared", g1, u, titles)[0] == "UNKNOWN",
          "an org-only layer in a group-shared map is unknown, since a group can hold outside members  <-- pinned defect")
    check(c("org", "shared", None, set()) is True and c("org", "org") is True,
          "an org-only layer is fine in an org map, and in a map shared to no group")
    # a layer shared to zero groups is owner-only too, so the owner test applies to it
    check(verdict("item", "", "shared", set(), "bob", "private", None, "alice", titles)[0] == "BROKEN",
          "a layer shared to no group is invisible to a private map owned by someone else  <-- pinned defect")
    check(verdict("item", "", "shared", set(), "alice", "private", None, "alice", titles)[0] == "OK"
          and verdict("item", "", "shared", set(), "alice", "shared", set(), "alice", titles)[0] == "OK",
          "a layer shared to no group is fine for its own owner's private or no-group map")
    ok, why = covers("shared", set(), u, "shared", g1, u, titles)
    check(ok is False and why == "private layer in a shared map",
          "a layer shared to no group is owner-only inside a group-shared map")


def _t_core(t):
    check, raises = t.check, t.raises
    u = "u"
    # probe error-code map
    for code in (499, 403, 401, 498):
        check(code_state(code)[0] == "SECURED", "code %s means a token is needed" % code)
    check(code_state(404) == ("DEAD", "code 404"), "code 404 means the service is gone")
    # 400 "Invalid URL" is how ArcGIS REST reports a service that no longer exists.
    check(code_state(400) == ("DEAD", "code 400"), "code 400 means the service is gone")
    check(code_state(500)[0] == "UNKNOWN", "code 500 is unknown")
    check(verdict("SECURED", "", None, None, None, "public", None, u)[0] == "BROKEN",
          "a secured service in a public map is broken")
    check(verdict("SECURED", "", None, None, None, "org", None, u)[0] == "UNKNOWN",
          "a secured service in an org map is unknown")
    check(verdict("PUBLIC", "", "public", set(), "", "public", set(), u)[0] == "OK",
          "a public service in a public map is fine")
    check(verdict("EMBEDDED", "", None, None, None, "public", None, u)[0] == "OK",
          "an embedded feature collection is fine")
    check(verdict("NO_TARGET", "", None, None, None, "public", None, u)[1].startswith("layer reference has no"),
          "a layer with no target says so")
    check("400" in verdict("DEAD", "code 400", None, None, None, "public", None, u)[1],
          "a dead reason carries the real code, not a hardcoded 404")
    check("no longer published" in verdict("DEAD", "", None, None, None, "public", None, u)[1],
          "a dead reason with no code still reads")

    # circuit breaker returns before any request is attempted
    fails = collections.Counter({"dead.example.gov": MAX_HOST_FAILS})
    cache = {}
    check(probe("https://dead.example.gov/x/FeatureServer/0", cache, fails)
          == ("UNREACHABLE", "host circuit-open"), "an open circuit answers without a request")
    check(cache == {}, "circuit-open must not poison the url cache")

    # walk: GroupLayer children, tables, basemap; shells and junk excluded
    out = []
    walk({"operationalLayers": [
        {"title": "Grp", "layerType": "GroupLayer", "layers": [
            {"title": "Kid", "url": "https://x/FeatureServer/0", "itemId": "k1"}]},
        {"title": "Solo", "url": "https://x/FeatureServer/1"},
        {"title": "Svc", "url": "https://x/MapServer", "layers": [{"id": 0, "popupInfo": {}}]}],
        "tables": [{"title": "Tbl", "itemId": "t1"}],
        "baseMap": {"baseMapLayers": [{"title": "BM", "url": "https://b/MapServer"}]}}, out)
    got = set((s, lyr.get("title")) for s, lyr in out)
    check(("operationalLayers", "Kid") in got and ("operationalLayers", "Grp") not in got,
          "a group layer's child is a reference and the shell is not")
    check(("tables", "Tbl") in got and ("baseMap", "BM") in got, "tables and the base map are walked")
    check(len(out) == 5, "five references: Kid, Solo, Svc, Tbl, BM; the sublayer override is dropped")
    for junk in (None, {}, [], "x"):
        o = []
        walk(junk, o)
        check(o == [], "walk() of %r finds nothing" % (junk,))
    walk({"operationalLayers": ["junk", 5, {"title": "Only"}]}, out)
    check(len(out) == 6, "a non-dict layer entry is skipped and a bare titled one is kept")
    out = []
    walk({"operationalLayers": [{"title": "Grp", "layers": [
        {"title": "ByUrl", "url": "https://x/FeatureServer/0"}, {"title": "ById", "itemId": "k2"},
        {"title": "Override", "popupInfo": {}}]}]}, out)
    check(sorted(l["title"] for _, l in out) == ["ById", "ByUrl"],
          "a group child with only a url or only an item id is a reference; the shell is not  <-- pinned defect")

    # iso takes MILLISECONDS
    check(iso(0) == "" and iso(None) == "", "iso() of nothing is empty")
    check(iso(1700000000000).startswith("2023-11-"), "iso() reads milliseconds, got %s" % iso(1700000000000))
    check(sorted(["2 Medium", "1 High", "4 Info", "3 Low"]) == ["1 High", "2 Medium", "3 Low", "4 Info"],
          "the severity bands sort as text")
    check(sorted(SEV.values()) == ["1 High", "2 Medium", "3 Low", "4 Info"] and SEV["public"] == "1 High",
          "SEV maps public to the first band")
    # strict feature-layer regex: the service root cannot be counted
    check(FS_RE.search("https://x/arcgis/rest/services/a/featureserver/0") is not None,
          "a feature layer url is countable")
    check(FS_RE.search("https://x/arcgis/rest/services/a/featureserver") is None,
          "the feature service root is not countable")
    check(FS_RE.search("https://x/arcgis/rest/services/a/mapserver/0") is None,
          "a map service layer is not countable")
    check(list(chunks(range(5), 2)) == [[0, 1], [2, 3], [4]] and list(chunks([], 3)) == [],
          "chunks() splits and tolerates nothing")
    check(host("https://a.example/x?y=1") == "a.example" and host(None) == "" and host("http://[bad") == "",
          "host() reads the netloc and survives None and a malformed url")

    # a corrupt nested map shape raises out of walk so collect() can record it as an error
    # row, rather than being silently dropped
    for corrupt in ({"baseMap": "corrupt"}, {"tables": 5},
                    {"operationalLayers": [{"url": "x", "layers": 7}]}):
        raises(lambda c=corrupt: walk(c, []), "walk() raises on the corrupt shape %r" % (corrupt,),
               exc=(AttributeError, TypeError))

    # build(): every state of a layer reference, with a fake item
    org = "org1"
    items = {"i-pub": _Item("i-pub", access="public", owner="bob", orgId=org),
             "i-org": _Item("i-org", access="org", owner="bob", orgId=org),
             "i-ext": _Item("i-ext", access="public", owner="esri", orgId="esri-org")}
    meta = {"MapTitle": "M", "MapId": "m1", "MapOwner": "alice", "MapAccess": "org",
            "MapViews": 3, "MapUrl": "u"}
    refs = [(meta, "operationalLayers", {"itemId": "i-pub", "title": "A", "layerType": "ArcGISFeatureLayer"}),
            (meta, "operationalLayers", {"itemId": "i-org", "url": "https://h.example/a/FeatureServer/0"}),
            (meta, "operationalLayers", {"itemId": "i-ext"}),
            (meta, "operationalLayers", {"itemId": "gone"}),
            (meta, "operationalLayers", {"itemId": "late"}),
            (meta, "operationalLayers", {"featureCollection": {"layers": []}}),
            (meta, "operationalLayers", {}),
            (meta, "operationalLayers", {"url": "https://h.example/b/FeatureServer/0"})]
    fakeprobe = {"https://h.example/b/FeatureServer/0": ("PUBLIC", "")}
    rows = build(refs, items, {"m1": set()}, {}, fakeprobe, collections.Counter(), org, {"late"})
    check([r["Verdict"] for r in rows] == ["OK", "OK", "OK", "DEAD", "UNKNOWN", "OK", "UNKNOWN", "OK"],
          "build() scores each kind of layer reference, got %s" % [r["Verdict"] for r in rows])
    check([r["Overexposed"] for r in rows] == ["Yes", "", "", "", "", "", "", ""],
          "only an org-owned public layer in a non-public map is overexposed")
    check(rows[3]["Severity"] == "2 Medium" and rows[0]["Severity"] == "",
          "severity follows the map audience and is blank for OK")
    check(rows[4]["Reason"].startswith("item lookup failed"),
          "an item whose lookup failed is unknown, never dead")
    check((rows[0]["LayerOwner"], rows[0]["LayerAccess"], rows[0]["LayerType"])
          == ("bob", "public", "ArcGISFeatureLayer")
          and (rows[1]["LayerOwner"], rows[1]["LayerAccess"]) == ("bob", "org")
          and (rows[5]["LayerOwner"], rows[5]["LayerAccess"], rows[5]["LayerType"]) == ("", "", ""),
          "the layer owner, access and type columns carry what the item says  <-- pinned defect")
    dirty = "https://user:pw@h.example/a/FeatureServer/0?token=SECRET9&where=1%3D1&f=json#frag"
    tidy = "https://h.example/a/FeatureServer/0?where=1%3D1&f=json"
    rows2 = build([(meta, "operationalLayers", {"url": dirty})], {}, {}, {}, {tidy: ("PUBLIC", "")},
                  collections.Counter(), org)
    check(rows2[0]["LayerUrl"] == tidy and rows2[0]["LayerHost"] == "h.example"
          and "SECRET9" not in repr(rows2) and "pw" not in repr(rows2),
          "a token or password in a layer url never reaches a report row  <-- pinned defect")
    check(clean_url(None) == "" and clean_url("https://h/x?a=1&&b") == "https://h/x?a=1&b"
          and clean_url("https://h/x?access_token=1&Key=2&apikey=3&pwd=4&SIG=5") == "https://h/x"
          and clean_url("https://h/a@b/x") == "https://h/a@b/x" and clean_url("//h/x") == "//h/x"
          and clean_url("https://a:b@c/x") == "https://c/x" and clean_url("x?q=token") == "x?q=token",
          "clean_url() drops credential-named parameters and userinfo and nothing else")
    erows = error_rows([(meta, "KeyError: x")])
    check(erows[0]["Verdict"] == "UNKNOWN" and erows[0]["Section"] == "<error>"
          and erows[0]["Severity"] == "2 Medium", "an unreadable map becomes an UNKNOWN row")
    # Overexposed was deliberately kept out of Findings: it is a different axis
    check(rows[0]["Verdict"] == "OK" and rows[0]["Overexposed"] == "Yes",
          "an overexposed layer is OK in the verdict column")


def _t_probe(t, base, log):
    check = t.check
    cache, fails = {}, collections.Counter()

    def go(path):
        return probe(base + path, cache, fails)

    check(go("/ok/FeatureServer/0") == ("PUBLIC", ""), "a service that answers JSON is public")
    seen = len(log)
    check(go("/ok/FeatureServer/0") == ("PUBLIC", "") and len(log) == seen,
          "a second probe of the same url is served from the cache, not the network")
    check(go("/token/FeatureServer/0") == ("SECURED", "code 499"),
          "a 200 whose body is error 499 is secured")
    check(go("/forbidden/FeatureServer/0") == ("SECURED", "code 403"), "an HTTP 403 is secured")
    check(go("/gone/FeatureServer/0") == ("DEAD", "code 400"), "an HTTP 400 is a dead service")
    check(go("/missing/FeatureServer/0") == ("DEAD", "code 404"), "an HTTP 404 is a dead service")
    check(go("/boom/FeatureServer/0") == ("UNKNOWN", "code 500"), "an HTTP 500 is unknown")
    check(go("/html/FeatureServer/0") == ("UNKNOWN", "non-JSON response"), "an HTML reply is unknown")
    check(go("/list/FeatureServer/0") == ("UNKNOWN", "non-JSON response"), "a JSON array is unknown")
    check(go("/odd/FeatureServer/0") == ("UNKNOWN", "code 418"), "an unlisted error code is unknown")
    check(fails[host(base)] == 0, "completed round trips keep the failure count at zero")
    mark = len(log)
    check(probe(base + "/ok2/FeatureServer/0?token=SECRET123", cache, fails)[0] == "DEAD"
          and not any("SECRET123" in p for p in log[mark:]),
          "a token already in the layer url is not sent  <-- pinned defect")
    mark = len(log)
    check(go("/needs/query?where=1%3D1&f=geojson&token=SECRET55") == ("PUBLIC", "")
          and "/needs/query?where=1%3D1&f=json" in log[mark:]
          and not any("SECRET55" in q or "geojson" in q for q in log[mark:]),
          "a layer url keeps its own query, with f=json swapped in and the token dropped  <-- pinned defect")
    check(go("/needs/query") == ("DEAD", "code 400"), "the same service without its query is not the same layer")
    for path, want in (("/errstr", "forbidden"), ("/errlist", "list"), ("/errnum", "number"),
                       ("/errcode", "list code")):
        check(go(path + "/FeatureServer/0") == ("UNKNOWN", "error reply of unknown shape"),
              "a 200 whose error is a %s is unknown, not a crash  <-- pinned defect" % want)
    check(go("/errstrcode/FeatureServer/0") == ("UNKNOWN", "code 498"),
          "a string error code is not one of the integer codes, so it stays unknown")
    # a failure, a success, a failure ...: only consecutive failures open the circuit
    fresh, fresh_cache = collections.Counter(), {}
    for route in ("/drop10", "/ok/FeatureServer/0?a=1", "/drop11", "/ok/FeatureServer/0?a=2", "/drop12"):
        probe(base + route, fresh_cache, fresh)
    check(fresh[host(base)] == 1 and probe(base + "/ok/FeatureServer/0?a=3", fresh_cache, fresh) == ("PUBLIC", ""),
          "a success between failures resets the host count, so three separate failures do not open the circuit  <-- pinned defect")
    for n in (1, 2, 3):
        res = go("/drop%d" % n)
    check(res == ("UNREACHABLE", "RemoteDisconnected") and fails[host(base)] == MAX_HOST_FAILS,
          "a host that never answers is unreachable and counted")
    mark = len(log)
    check(go("/ok/other/FeatureServer/0") == ("UNREACHABLE", "host circuit-open") and len(log) == mark,
          "after three failures the host is skipped with no request")
    fails[host(base)] = 2
    go("/gone/FeatureServer/9")
    check(fails[host(base)] == 0, "any completed round trip resets the host failure count")


def _t_portal(t, rest):
    check = t.check
    items = {"a1": _Item("a1"), "a2": _Item("a2"), "a4": _Item("a4")}
    gis = _Gis(rest, items, fail_ids={"a3"})
    with _patched(ITEM_CHUNK=2), _quiet() as (out, _):
        got, unresolved = resolve_items(gis, {"a1", "a2", "a3", "a4"})
    check(sorted(got) == ["a1", "a2"] and unresolved == {"a3", "a4"},
          "a failed lookup chunk is unresolved, never reported as deleted  <-- pinned defect")
    check("item batch failed (RuntimeError: throttled), 2 ids unresolved" in out.getvalue(),
          "a failed lookup chunk is announced")
    got, unresolved = resolve_items(gis, {"a1", "zz"})
    check(sorted(got) == ["a1"] and unresolved == set(), "an id with no item is missing, not unresolved")
    asked = []

    class Spy(_Content):
        def search(self, query="", **kw):
            asked.append(query.count(" OR ") + 1)
            return []

    spy = _Gis(rest, {})
    spy.content = Spy({})
    resolve_items(spy, set(_hid(n) for n in range(120)))
    check(asked == [50, 50, 20] and ITEM_CHUNK == 50, "item ids go out to the portal 50 per search, as the README says")

    gids, titles = fetch_groups(gis, {"m1", "l1", "e1"})
    check(gids == {"m1": {"g1", "g2"}, "l1": {"g1"}, "e1": set()} and titles == {"g1": "Alpha", "g2": "Beta"},
          "group membership and titles are read from the portal reply")
    check(gids["e1"] == set(), "an item shared to no group is an empty set, not unknown")
    gids, titles = fetch_groups(gis, {"l1"})
    check(gids == {"l1": {"g1"}} and titles == {"g1": "Alpha"},
          "a group with no member among the asked ids is not listed")
    odd = _Gis(rest, {})
    odd._con = types.SimpleNamespace(get=lambda url, params: {
        "g9": {"groupTitle": None, "items": [None, {}, {"itemId": "x1"}]}, "g8": None})
    gids, titles = fetch_groups(odd, {"x1", "x2"})
    check(gids == {"x1": {"g9"}, "x2": set()} and titles == {"g9": "g9", "g8": "g8"},
          "an empty entry, a null group and a missing title in the reply do not crash it")
    odd._con = types.SimpleNamespace(get=lambda url, params: None)
    check(fetch_groups(odd, {"x1"}) == ({"x1": set()}, {}), "an empty reply means shared to no group")
    with _patched(GROUP_CHUNK=1), _quiet() as (out, _):
        gids, titles = fetch_groups(gis, {"boom", "m1"})
    check("boom" not in gids and gids["m1"] == {"g1", "g2"},
          "an item from a failed group chunk stays unknown  <-- pinned defect")
    check("group batch failed (HTTPError: HTTP Error 500" in out.getvalue(), "a failed group chunk is announced")

    # collect(): good map, corrupt map, raising map; wcag rows ride along
    good = {"operationalLayers": [{"title": "A", "itemId": "l1"}]}
    maps = [_Item("m1", title="Good", owner="alice", access="public", data=good, numViews=7),
            _Item("m2", title="Corrupt", data={"baseMap": "corrupt"}),
            _Item("m3", title="Boom", data=RuntimeError("no data"))]
    wcag = []
    with _quiet() as (out, _):
        refs, errors = collect(gis, maps, wcag)
    check(len(refs) == 1 and refs[0][0]["MapUrl"] == "https://portal.test/home/item.html?id=m1",
          "collect() returns the references of the readable map")
    check([e[0]["MapId"] for e in errors] == ["m2", "m3"] and errors[1][1] == "RuntimeError: no data",
          "a corrupt or failing map is an error tuple, never silently dropped  <-- pinned defect")
    check("3/3 maps" in out.getvalue(),
          "collect() prints its final progress line even when the last map fails  <-- pinned defect")
    check(sorted(set(r["Rule"] for r in wcag)) == ["map-no-summary"] and len(wcag) == 1,
          "collect() appends rule rows for the readable maps only")
    with _quiet():
        refs2, _ = collect(gis, maps[:1])
    check(len(refs2) == 1, "collect() without --wcag keeps working")

    # deep_stats
    class FL(object):
        def __init__(self, url, g):
            if url.endswith("/1"):
                raise OSError("down")
            self.properties = {"editingInfo": {"lastEditDate": 1700000000000}}

        def query(self, return_count_only):
            return 42 if return_count_only is True else {"features": []}

    urls = {"https://x/rest/a/FeatureServer/0", "https://x/rest/a/FeatureServer/1", "https://x/rest/a/MapServer/0"}
    with _patched(FeatureLayer=FL):
        stats = deep_stats(gis, urls)
    check(stats == {"https://x/rest/a/FeatureServer/0": (42, "2023-11-14 22:13:20+00:00"),
                    "https://x/rest/a/FeatureServer/1": ("", "OSError")},
          "deep_stats() counts feature layers, records a failure by type, skips other services, got %s" % stats)
    fake = types.ModuleType("arcgis")
    fake_features = types.ModuleType("arcgis.features")
    fake_features.FeatureLayer = FL
    fake_gis = types.ModuleType("arcgis.gis")
    fake_gis.GIS = object
    with _patched(FeatureLayer=None, GIS=None), _modules(arcgis=fake, **{"arcgis.features": fake_features,
                                                                         "arcgis.gis": fake_gis}):
        stats = deep_stats(gis, {"https://x/rest/a/FeatureServer/0"})
    check(list(stats.values()) == [(42, "2023-11-14 22:13:20+00:00")],
          "deep_stats() loads the ArcGIS library itself when it has not been loaded")
    with _patched(FeatureLayer=None, GIS=None), _modules(arcgis=None, **{"arcgis.gis": None, "arcgis.features": None}):
        raises_exit = _exits(load_arcgis)
    check(raises_exit.startswith("arcgis") and "ArcGIS Pro" in raises_exit,
          "a missing ArcGIS library explains where to find a Python that has it")
    with _patched(FeatureLayer=None, GIS=None), _modules(arcgis=fake, **{"arcgis.features": fake_features,
                                                                         "arcgis.gis": fake_gis}):
        load_arcgis()
        loaded = (GIS, FeatureLayer)
    check(loaded == (object, FL), "load_arcgis() puts the library classes where connect() looks")
    with _patched(pd=None, get_column_letter=None):
        pandas_stub = types.ModuleType("pandas")
        with _modules(pandas=pandas_stub, openpyxl=None, **{"openpyxl.utils": None}):
            msg = _exits(load_report_libs)
    check(msg.startswith("openpyxl") and "was not found" in msg and "pip install" in msg,
          "a missing report library says which one and how to install it")
    with _patched(pd=None, get_column_letter=None):
        with _modules(pandas=None):
            msg = _exits(load_report_libs)
    check(msg.startswith("pandas was not found"), "a missing pandas is named")
    t.raises(lambda: _exits(lambda: None), "_exits() fails loudly when the call does not exit", exc=AssertionError)


def _t_edges(t, rest):
    check = t.check
    mid = "a" * 32
    lid = "b" * 32
    # map_candidates: the four key spellings, nested lists, case, junk
    data = {"widgets": [{"type": "mapWidget", "itemId": mid}],
            "values": {"webmap": lid.upper(), "title": mid},
            "nodes": {"n1": {"type": "webmap", "data": {"itemId": mid}}},
            "other": [[{"webmapId": lid}], {"mapId": "not-an-id"}, {"itemId": 7}]}
    check(map_candidates(data) == [mid, lid],
          "candidates are found under every known key, lower-cased and de-duplicated")
    check(map_candidates({"title": mid, "id": mid}) == [],
          "an id under a key that is not a map key is not a candidate")
    for junk in (None, "x", 5, [], {}, [1, "a"]):
        check(map_candidates(junk) == [], "no candidates in %r" % (junk,))
    deep = cur = {}
    for _ in range(5000):
        cur["k"] = {}
        cur = cur["k"]
    cur["itemId"] = mid
    check(map_candidates(deep) == [mid], "a 5000-level document does not hit the recursion limit  <-- pinned defect")

    # edge_verdict: the audience matrix, app side against map side
    g1, g2 = {"g1"}, {"g1", "g2"}
    titles = {"g1": "Alpha", "g2": "Beta"}

    def ev(aa, ma, ag=None, mg=None, aown="alice", mown="alice"):
        return edge_verdict(aa, ag, aown, ma, mg, mown, titles)

    check(ev("public", "public")[0] == "OK", "a public app on a public map is fine")
    check(ev("public", "org")[0] == "BROKEN", "a public app on an org map is broken")
    check(ev("public", "private")[0] == "BROKEN", "a public app on a private map is broken")
    check(ev("public", "shared", None, g1)[0] == "BROKEN", "a public app on a group map is broken")
    check(ev("org", "shared", None, g1)[0] == "BROKEN", "an org app on a group map is broken")
    check(ev("org", "org")[0] == "OK" and ev("org", "public")[0] == "OK", "an org app on an org or public map is fine")
    verdict_, why = ev("shared", "org", g1)
    check(verdict_ == "UNKNOWN" and "membership is not read" in why and "web map is org-only" in why,
          "a group app on an org map is unknown: a group can hold members from another organisation  <-- pinned defect")
    check(ev("private", "private")[0] == "OK", "a private app on the same owner's private map is fine")
    verdict_, why = ev("private", "private", mown="bob")
    check(verdict_ == "BROKEN" and why == "web map is private to bob; the app owner cannot load it",
          "a private app on another user's private map is broken")
    check(ev("private", "org")[0] == "OK" and ev("private", "shared", None, g1)[0] == "OK",
          "a private app is seen by its owner only, and the owner can open an org or group map")
    verdict_, why = ev("shared", "shared", g2, g1)
    check(verdict_ == "BROKEN" and why == "app is shared to groups the web map is not: Beta",
          "a group app shared to a group the map is not names that group")
    check(ev("shared", "shared", g1, g2)[0] == "OK", "a group app whose groups are all on the map is fine")
    check(ev("shared", "shared", None, g1)[0] == "UNKNOWN" and ev("shared", "shared", g1, None)[0] == "UNKNOWN",
          "unknown groups on either side are unknown, never a silent OK")
    check(ev("shared", "shared", set(), g1)[0] == "OK",
          "an app shared to no group is owner-only, so any map its owner can open covers it")
    check(ev("shared", "shared", set(), None)[0] == "OK",
          "an app shared to no group is owner-only even when the map's groups are unknown  <-- pinned defect")
    verdict_, why = ev("public", "org")
    check(why == "app is shared public but the web map is shared org", "the reason names both audiences")
    # a map shared to no group is owner-only, whoever owns it
    verdict_, why = ev("private", "shared", None, set(), mown="bob")
    check(verdict_ == "BROKEN" and why == "web map is private to bob; the app owner cannot load it",
          "a private app on another user's map that is shared to no group is broken  <-- pinned defect")
    check(ev("private", "shared", None, set())[0] == "OK" and ev("public", "shared", None, set())[0] == "BROKEN",
          "a map shared to no group is fine for its own owner's app and broken for a public one")
    # the README table, cell by cell: rows are the map's audience, columns the app's
    table = {
        "public": {"public": "OK", "org": "OK", "shared": "OK", "private": "OK"},
        "org": {"public": "BROKEN", "org": "OK", "shared": "UNKNOWN", "private": "OK"},
        "shared": {"public": "BROKEN", "org": "BROKEN", "shared": "OK", "private": "OK"},
        "private": {"public": "BROKEN", "org": "BROKEN", "shared": "BROKEN", "private": "OK"},
    }
    for map_side, row in table.items():
        for app_side, want in row.items():
            got = ev(app_side, map_side, g1 if app_side == "shared" else None,
                     g1 if map_side == "shared" else None)[0]
            check(got == want, "app %s on a %s map is %s" % (app_side, map_side, want))
    verdict_, why = edge_verdict("shared", g2, "a", "shared", g1, "a")
    check(verdict_ == "BROKEN" and why.endswith("g2"), "with no titles dict a missing group is named by its id")

    # build_edges and collect_apps over a fake portal
    ids = dict((k, _hid(n)) for k, n in (("m-pub", 1), ("m-org", 2), ("m-grp", 3), ("lay", 4), ("ghost", 255)))
    specs = {"m-pub": ("Pub", "public", "Web Map"), "m-org": ("Org", "org", "Web Map"),
             "m-grp": ("Grp", "shared", "Web Map"), "lay": ("Lay", "public", "Feature Service")}
    maps = dict((ids[k], _Item(ids[k], title=v[0], access=v[1], owner="alice", kind=v[2]))
                for k, v in specs.items())
    apps_items = {
        "d1": _Item("d1", title="Dash", access="public", owner="alice", kind="Dashboard", numViews=9,
                    data={"widgets": [{"type": "mapWidget", "itemId": ids["m-org"]},
                                      {"type": "list", "itemId": ids["lay"]},
                                      {"type": "map", "itemId": ids["ghost"]}]}),
        "d2": _Item("d2", title="Second", access="private", owner="alice", kind="Dashboard", data={}),
        "x2": _Item("x2", title="Exp2", access="private", owner="alice", kind="Web Experience",
                    data={"dataSources": {"dataSource_1": {"type": "WEB_MAP", "itemId": ids["m-pub"]}}}),
        "s1": _Item("s1", title="Story", access="org", owner="alice", kind="StoryMap",
                    data={"nodes": {"n": {"data": {"itemId": ids["m-pub"]}}}}),
        "w1": _Item("w1", title="App", access="shared", owner="alice", kind="Web Mapping Application",
                    data={"values": {"webmap": ids["m-grp"]}}),
        "x1": _Item("x1", title="Exp", access="public", owner="alice", kind="Web Experience",
                    data=RuntimeError("no data")),
    }
    pool = dict(maps)
    pool.update(apps_items)
    gis = _Gis(rest, pool, leak={"Web Experience": ["s1"]})
    with _patched(MAX_APPS=50):
        apps, errors = collect_apps(gis, None)
    check(sorted(m["AppId"] for m, _ in apps) == ["d1", "d2", "s1", "w1", "x2"] and [e[0]["AppId"] for e in errors] == ["x1"],
          "collect_apps() reads each app once, even if two searches return it, and records a failure")
    check(apps[0][0]["AppUrl"] == "https://portal.test/home/item.html?id=d1", "an app row carries its item url")
    with _patched(MAX_APPS=50):
        one, _ = collect_apps(gis, 1)
    check(len(one) == 4 and "d2" not in [m["AppId"] for m, _ in one],
          "--limit caps each item type: one dashboard, story, app and experience; got %d" % len(one))
    cand = sorted(set(c for _, cs in apps for c in cs))
    gids = {"d1": set(), "s1": set(), "x2": set(), "w1": g2, ids["m-grp"]: g1, ids["m-org"]: set(), ids["m-pub"]: set()}
    rows, skipped = build_edges(apps, errors, maps, gids, titles, "https://portal.test/home/item.html?id=",
                                {ids["ghost"]})
    by_app = dict((r["AppId"], r) for r in rows)
    check(by_app["d1"]["Verdict"] == "BROKEN" and by_app["d1"]["Severity"] == "1 High",
          "a public dashboard on an org map is a High finding")
    check(by_app["s1"]["Verdict"] == "OK" and by_app["s1"]["Severity"] == "",
          "an org story on a public map is fine and has no severity")
    check(by_app["w1"]["Verdict"] == "BROKEN" and by_app["w1"]["Reason"].endswith("Beta"),
          "an app shared to a group the map lacks names the group")
    check(by_app["w1"]["AppGroups"] == "Alpha; Beta" and by_app["w1"]["MapGroups"] == "Alpha",
          "edge rows carry the group titles of both sides")
    check(by_app["x2"]["Verdict"] == "OK" and by_app["x2"]["MapTitle"] == "Pub",
          "an Experience Builder style data source of type WEB_MAP is followed to its map")
    check(by_app["x1"]["Verdict"] == "UNKNOWN" and by_app["x1"]["Reason"] == "RuntimeError: no data",
          "an app that could not be read is an UNKNOWN row")
    check(dict(skipped) == {"not a web map": 1, "lookup failed": 1},
          "candidates that were a layer or unresolved are counted, got %s" % dict(skipped))
    rows2, skipped2 = build_edges(apps, [], maps, gids, titles, "b", set())
    check(skipped2["no item"] == 1 and skipped2["lookup failed"] == 0,
          "an id that resolves to nothing, with no failed lookup, is counted as no item")
    check(rows2[0]["MapUrl"].startswith("b") and rows2[0]["MapTitle"] in ("Org", "Pub", "Grp"),
          "an edge row carries the map title and url")
    check(set(EDGE_COLS) <= set(rows[0]) and set(rows[0]) <= set(EDGE_COLS) | {"Verdict"},
          "every edge row fits the EDGE_COLS sheet layout")
    check(cand == sorted([ids["m-org"], ids["m-pub"], ids["m-grp"], ids["lay"], ids["ghost"]]),
          "collect_apps() surfaces every candidate id, including a layer and an unknown one")


def _t_wcag(t):
    check = t.check
    meta = {"MapTitle": "Parks", "MapId": "m1", "MapOwner": "u", "MapAccess": "public",
            "MapViews": 5, "MapUrl": "https://p/m1"}

    def rules(title="Parks", snippet="s", description="", layers=(), meta=meta):
        found = [("operationalLayers", lyr) for lyr in layers]
        return [r["Rule"] for r in wcag_check(meta, title, snippet, description, found)]

    # contrast: the numbers the thresholds rest on
    black, white = (0, 0, 0), (255, 255, 255)
    check(abs(contrast(black, white) - 21.0) < 1e-9, "black on white is 21:1")
    check(abs(contrast(white, white) - 1.0) < 1e-9, "a color on itself is 1:1")
    check(abs(contrast(white, black) - contrast(black, white)) < 1e-12, "contrast does not depend on argument order")
    check(contrast((118, 118, 118), white) >= TEXT_RATIO, "#767676 on white clears 4.5:1")
    check(contrast((119, 119, 119), white) < TEXT_RATIO, "#777777 on white misses 4.5:1")
    check(abs(luminance((255, 0, 0)) - 0.2126) < 1e-9 and abs(luminance((0, 0, 255)) - 0.0722) < 1e-9,
          "luminance weights red at 0.2126 and blue at 0.0722")
    check(rgb_of([0, 0, 0, 255]) == (0, 0, 0) and rgb_of([0, 255, 0, 255]) == (0, 255, 0)
          and rgb_of([-1, 2, 3, 255]) is None,
          "a channel of 0 is a valid color, so black can be judged  <-- pinned defect")
    low = 10 / 255.0
    check(abs(luminance((10, 10, 10)) - low / 12.92) < 1e-12
          and abs(luminance((10.2, 10.2, 10.2)) - 0.04 / 12.92) < 1e-12
          and abs(luminance((10.4, 10.4, 10.4)) - ((10.4 / 255.0 + 0.055) / 1.055) ** 2.4) < 1e-12
          and abs(luminance((30, 30, 30)) - ((30 / 255.0 + 0.055) / 1.055) ** 2.4) < 1e-12,
          "luminance uses s/12.92 up to 0.04045 and the 2.4 power above it  <-- pinned defect")
    check(rgb_of([1, 2, 3, 255]) == (1, 2, 3) and rgb_of([1, 2, 3, 128]) is None
          and rgb_of([1, 2, 3]) is None and rgb_of("x") is None and rgb_of([1, 2, 3, 256]) is None
          and rgb_of([1, 2, "3", 255]) is None and rgb_of(None) is None,
          "only a fully opaque, well-formed color can be judged")
    check(dig({"a": {"b": 1}}, "a", "b") == 1 and dig({"a": 5}, "a", "b") is None and dig(None, "a") is None,
          "dig() stops at the first level that is not a dict")

    # map-level rules
    check(rules() == [], "a titled map with a summary raises nothing")
    check(rules(title="Untitled Map") == ["map-title-placeholder"], "a placeholder title is flagged")
    check(rules(title="  UNTITLED   web map ") == ["map-title-placeholder"],
          "case and spacing do not hide a placeholder title")
    check(rules(title="") == ["map-title-placeholder"] and rules(title=None) == ["map-title-placeholder"],
          "an empty or missing title is flagged")
    check(rules(snippet="", description="") == ["map-no-summary"], "no summary and no description is flagged")
    check(rules(snippet=None, description=None) == ["map-no-summary"], "missing summary fields are flagged")
    check(rules(snippet="", description="<p> </p><br/>") == ["map-no-summary"],
          "a description that is only tags is empty")
    check(rules(snippet="", description="<p>Parks open in summer</p>") == [], "a description is enough")
    check(rules(snippet="  ", description="") == ["map-no-summary"], "a blank summary does not count")

    # layer-level rules
    check(rules(layers=[{"title": "Trails"}]) == [], "a titled layer raises nothing")
    check(rules(layers=[{"url": "u"}]) == ["layer-no-title"] and rules(layers=[{"title": "  "}]) == ["layer-no-title"],
          "a layer with no or a blank title is flagged")
    check(rules(layers=[{"title": "T", "disablePopup": True}]) == ["layer-popup-off"],
          "a layer with the pop-up disabled is flagged")
    check(rules(layers=[{"title": "T", "disablePopup": False}, {"title": "U"}]) == [],
          "a layer that leaves pop-ups on is not")
    check(wcag_check(meta, "Parks", "s", "", [("baseMap", {"url": "u"})]) == [],
          "base map layers are Esri's and are never judged")

    def uv(*symbols, **kw):
        classes = [{"value": str(i), "symbol": s} for i, s in enumerate(symbols)]
        rend = {"type": "uniqueValue", "uniqueValueInfos": classes}
        rend.update(kw)
        return {"title": "T", "layerDefinition": {"drawingInfo": {"renderer": rend}}}

    def sms(color, style="esriSMSCircle", size=8, **kw):
        s = {"type": "esriSMS", "style": style, "size": size, "color": color}
        s.update(kw)
        return s

    red, blue = [200, 0, 0, 255], [0, 0, 200, 255]
    check(rules(layers=[uv(sms(red), sms(blue))]) == ["renderer-color-only"],
          "two marker classes that differ only by color are flagged")
    check(rules(layers=[uv(sms(red), sms(blue, style="esriSMSSquare"))]) == [],
          "classes that differ by shape are fine")
    check(rules(layers=[uv(sms(red), sms(blue, size=14))]) == [], "classes that differ by size are fine")
    check(rules(layers=[uv(sms(red), sms(blue, outline={"style": "esriSLSDash", "width": 1}))]) == [],
          "classes that differ by outline are fine")
    check(rules(layers=[uv(sms(red), sms(red))]) == [], "classes with one color are not a color-only failure")
    check(rules(layers=[uv(sms(red))]) == [], "a single class has nothing to tell apart")
    sls = lambda col, style="esriSLSSolid", width=2: {"type": "esriSLS", "style": style, "width": width, "color": col}
    sfs = lambda col, style="esriSFSSolid": {"type": "esriSFS", "style": style, "color": col}
    check(rules(layers=[uv(sls(red), sls(blue))]) == ["renderer-color-only"]
          and rules(layers=[uv(sls(red), sls(blue, style="esriSLSDash"))]) == [],
          "line classes are judged by dash and width")
    check(rules(layers=[uv(sls(red, width=1), sls(blue, width=4))]) == [],
          "line classes that differ by width are not color-only  <-- pinned defect")
    check(rules(layers=[uv(sfs(red), sfs(blue))]) == ["renderer-color-only"]
          and rules(layers=[uv(sfs(red), sfs(blue, style="esriSFSCross"))]) == [],
          "fill classes are judged by pattern")
    check(rules(layers=[uv({"type": "esriPMS", "url": "a.png"}, {"type": "esriPMS", "url": "b.png"})]) == [],
          "picture marker classes differ by more than color")
    check(rules(layers=[uv(sms(red), "junk")]) == [] and rules(layers=[uv(sms(red), {"type": "esriSMS"})]) == [],
          "a class with a malformed or colorless symbol is not judged")
    check(rules(layers=[{"title": "T", "layerDefinition": {"drawingInfo": {"renderer": {"type": "simple"}}}}]) == []
          and rules(layers=[{"title": "T", "layerDefinition": "x"}]) == []
          and rules(layers=[{"title": "T", "layerDefinition": {"drawingInfo": {"renderer": {
              "type": "classBreaks", "classBreakInfos": []}}}}]) == [],
          "simple, class-break and malformed renderers are not judged")
    grouped = {"title": "T", "layerDefinition": {"drawingInfo": {"renderer": {
        "type": "uniqueValue",
        "uniqueValueInfos": [{"symbol": sms(red)}, {"symbol": sms(red, style="esriSMSX")}],
        "uniqueValueGroups": [{"heading": "h", "classes": [{"symbol": sms(red)}, {"symbol": sms(blue)}]},
                              {"heading": "i", "classes": "bad"}, "junk"]}}}}
    check(rules(layers=[grouped]) == ["renderer-color-only"],
          "uniqueValueGroups wins over uniqueValueInfos when both are present")
    bad_color = uv(sms([[1], [2]]), sms([[3], [4]]))
    check(rules(layers=[bad_color]) == ["renderer-color-only"],
          "an unhashable color value does not crash the rule")

    def lab(text, halo, size=1, font=None, show=True):
        sym = {"type": "esriTS", "color": text, "haloColor": halo, "haloSize": size}
        if font is not None:
            sym["font"] = font
        return {"title": "T", "showLabels": show,
                "layerDefinition": {"drawingInfo": {"labelingInfo": [{"symbol": sym}]}}}

    grey, pale = [119, 119, 119, 255], [255, 255, 255, 255]
    check(rules(layers=[lab(grey, pale)]) == ["label-halo-contrast"],
          "label text under 4.5:1 on its halo is flagged")
    check(rules(layers=[lab([118, 118, 118, 255], pale)]) == [], "label text at 4.54:1 on its halo is fine")
    check(rules(layers=[lab(grey, pale, font={"size": 18})]) == [] and
          rules(layers=[lab(grey, pale, font={"size": 14, "weight": "bold"})]) == [],
          "large text needs only 3:1: 18 point, or 14 point bold")
    check(rules(layers=[lab(grey, pale, font={"size": 14})]) == ["label-halo-contrast"]
          and rules(layers=[lab(grey, pale, font={"size": "big"})]) == ["label-halo-contrast"],
          "14 point regular and an unreadable size are normal text")
    check(rules(layers=[lab([0, 0, 0, 255], [0, 0, 0, 255])]) == ["label-halo-contrast"]
          and rules(layers=[lab([0, 0, 0, 255], [255, 255, 255, 255])]) == [],
          "black text on a black halo is flagged, on a white halo it is fine  <-- pinned defect")
    check(rules(layers=[lab(grey, pale, show=False)]) == [] and rules(layers=[lab(grey, pale, size=0)]) == [],
          "labels that do not draw, or sit on no halo, are not judged")
    check(rules(layers=[lab(grey, [255, 255, 255, 100])]) == [] and rules(layers=[lab(None, pale)]) == [],
          "a translucent or missing color cannot be judged")
    check(rules(layers=[{"title": "T", "showLabels": True,
                         "layerDefinition": {"drawingInfo": {"labelingInfo": "bad"}}}]) == []
          and rules(layers=[{"title": "T", "showLabels": True,
                             "layerDefinition": {"drawingInfo": {"labelingInfo": ["x", {"symbol": "y"}]}}}]) == [],
          "a malformed labelingInfo is skipped, not a crash")
    check(rules(layers=[lab(grey, pale)], meta=dict(meta, MapAccess="weird")) == ["label-halo-contrast"],
          "an access value outside the four known bands still produces a row")
    rows = wcag_check(meta, "Untitled", "", "", [("operationalLayers", {"title": "L", "itemId": "x",
                                                                          "disablePopup": True})])
    check([r["Rule"] for r in rows] == ["map-title-placeholder", "map-no-summary", "layer-popup-off"]
          and rows[0]["Severity"] == "1 High" and rows[2]["Layer"] == "L",
          "rule rows come in a fixed order with the map's severity and the layer name")
    check(all(set(r) == set(WCAG_COLS) for r in rows), "every rule row fits the WCAG_COLS sheet layout")
    check(wcag_check(meta, "Untitled", "", "", [("operationalLayers", {"itemId": "abc"})])[2]["Layer"] == "abc",
          "an untitled layer is named by its item id in the report")
    named = wcag_check(meta, "Parks", "s", "", [("operationalLayers", {"url": "https://u:p@h/x?token=TOP7&a=1"})])
    check(named[0]["Layer"] == "https://h/x?a=1" and "TOP7" not in repr(named),
          "an untitled layer named by its url shows no token or password  <-- pinned defect")
    check(set(r for r in WCAG_RULES) == set(["map-title-placeholder", "map-no-summary", "layer-no-title",
                                            "layer-popup-off", "renderer-color-only", "label-halo-contrast"])
          and all(len(v) == 3 and v[1] in ("A", "AA") for v in WCAG_RULES.values()),
          "the rule registry names six rules, each with a criterion and a level")


def _t_summary(t):
    check = t.check
    rows = [{"Verdict": "OK", "MapTitle": "A", "MapAccess": "public", "MapViews": 1, "MapOwner": "x"},
            {"Verdict": "BROKEN", "MapTitle": "B", "MapAccess": "org", "MapViews": 2, "MapOwner": "y",
             "Overexposed": "Yes"},
            {"Verdict": "BROKEN", "MapTitle": "B", "MapAccess": "org", "MapViews": 2, "MapOwner": "y"},
            {"Verdict": "DEAD", "MapTitle": "C", "MapAccess": "org", "MapViews": 3, "MapOwner": "z"},
            {"Section": "<error>", "MapTitle": "D", "MapAccess": "org", "MapViews": 0, "MapOwner": "z"}]
    out = summary_rows(rows, {"maps scanned": 4})
    d = dict(((s, m), v) for s, m, v in out)
    check(d[("Run", "maps scanned")] == 4 and d[("Verdicts", "OK")] == 1 and d[("Verdicts", "BROKEN")] == 2
          and d[("Verdicts", "DEAD")] == 1 and d[("Verdicts", "UNKNOWN")] == 0,
          "verdict counts add up from the rows")
    check(d[("Verdicts", "Overexposed (public layer, non-public map)")] == 1, "overexposed rows are counted")
    check([m for s, m, v in out if s == "Worst maps"] == ["B (org, 2 views)", "C (org, 3 views)", "D (org, 0 views)"]
          and d[("Worst maps", "B (org, 2 views)")] == 2, "worst maps are ranked by findings, ties by name")
    check([(m, v) for s, m, v in out if s == "Top owners"] == [("y", 2), ("z", 2)],
          "top owners are the people to email, ties by name")
    check(len(summary_rows([], {"maps scanned": 0})) > 0 and ("Edges" not in [s for s, _, _ in out]),
          "an org with no web maps still summarizes, and the optional sections are absent  <-- pinned defect")
    out = summary_rows(rows, {}, edges=[{"Verdict": "BROKEN"}, {"Verdict": "OK"}],
                       wcag=[{"Rule": "layer-no-title"}, {"Rule": "layer-no-title"}])
    d = dict(((s, m), v) for s, m, v in out)
    check(d[("Edges", "BROKEN")] == 1 and d[("Edges", "OK")] == 1 and d[("WCAG", "layer-no-title")] == 2
          and d[("WCAG", "map-no-summary")] == 0, "edge verdicts and rule counts get their own sections")
    big = [{"Verdict": "BROKEN", "MapTitle": "m%02d" % i, "MapAccess": "org", "MapViews": 0,
            "MapOwner": "o%02d" % i} for i in range(30)]
    out = summary_rows(big, {})
    check(len([1 for s, _, _ in out if s == "Worst maps"]) == 15 and len([1 for s, _, _ in out if s == "Top owners"]) == 10,
          "the worst-maps table holds 15 rows and the owners table 10")


def _t_login(t):
    check = t.check
    calls, pro_fails = [], []

    def fake_gis(*args):
        calls.append(args)
        if args == ("pro",) and pro_fails:
            raise RuntimeError("no pro session with the secret-pw")
        return "session"

    prompts = []
    clean = dict(ARCGIS_USER=None, ARCGIS_PASS=None, ARCGIS_URL=None)
    with _patched(GIS=fake_gis, USERNAME="", PASSWORD=""):
        with _environ(**dict(clean, ARCGIS_USER="eu", ARCGIS_PASS="ep")):
            check(connect() == "session" and calls[-1] == (PORTAL_URL, "eu", "ep"),
                  "both found in the environment: sign in with them")
        with _environ(**dict(clean, ARCGIS_USER="eu", ARCGIS_PASS="ep", ARCGIS_URL="https://e.example/portal")):
            connect()
            check(calls[-1][0] == "https://e.example/portal", "ARCGIS_URL overrides PORTAL_URL")
        with _patched(USERNAME="cu", PASSWORD="cp"), _environ(**dict(clean, ARCGIS_USER="eu", ARCGIS_PASS="ep")):
            connect()
            check(calls[-1][1:] == ("cu", "cp"), "a value typed in CONFIG overrides the environment")
        with _environ(**clean):
            connect()
            check(calls[-1] == ("pro",), "no username anywhere: use the signed-in ArcGIS Pro session")
        pro_fails.append(True)
        with _environ(**clean), _patched(input=lambda p: prompts.append(p) or "bob"), \
                _setattr(getpass, getpass=lambda p: prompts.append(p) or "typed"), _quiet() as (out, _):
            connect()
        check(calls[-1] == (PORTAL_URL, "bob", "typed") and prompts == ["Username: ", "Password: "],
              "no Pro session: prompt for the username and then the password")
        check("RuntimeError" in out.getvalue() and "secret-pw" not in out.getvalue(),
              "the fallback notice names the error type and never its message  <-- pinned defect")
        with _environ(**dict(clean, ARCGIS_USER="eu")), _setattr(getpass, getpass=lambda p: "typed2"):
            connect()
        check(calls[-1] == (PORTAL_URL, "eu", "typed2"), "a username with no password prompts for the password only")
    fake = types.ModuleType("arcgis")
    fake_gis_mod = types.ModuleType("arcgis.gis")
    fake_gis_mod.GIS = lambda *a: "loaded"
    fake_features = types.ModuleType("arcgis.features")
    fake_features.FeatureLayer = object
    with _patched(GIS=None, USERNAME="", PASSWORD=""), _environ(ARCGIS_USER="u", ARCGIS_PASS="p", ARCGIS_URL=None), \
            _modules(arcgis=fake, **{"arcgis.gis": fake_gis_mod, "arcgis.features": fake_features}):
        check(connect() == "loaded", "connect() loads the ArcGIS library when it is not loaded yet")


def _t_cli(t):
    check = t.check
    p = build_parser()

    def refuses(argv):
        with _quiet():
            return _exits(lambda: p.parse_args(argv)) == 2

    check(p.parse_args(["--self-test"]).self_test and p.parse_args(["--self-check"]).self_test,
          "--self-check is an alias of --self-test")
    check(p.parse_args([]).apply is False and p.parse_args(["--apply"]).apply is True, "--apply is off unless given")
    check(refuses(["--ap"]), "a prefix of the write flag, --ap, is refused  <-- pinned defect")
    check(refuses(["--appl"]), "--appl is refused too")
    check(refuses(["--self-ch"]), "a unique prefix of the longest option must be refused, not expanded  <-- pinned defect")
    check(refuses(["--self-te"]) and refuses(["--ed"]) and refuses(["--wc"]) and refuses(["--li", "3"])
          and refuses(["--de"]), "no prefix of any flag is expanded")
    a = p.parse_args(["--edges", "--wcag", "--limit", "3", "--deep", "-o", "x.xlsx"])
    check(a.edges and a.wcag and a.limit == 3 and a.deep and a.out == "x.xlsx", "every flag parses")
    check(p.parse_args(["--max", "4"]).limit == 4, "--max still means --limit")
    check(p.allow_abbrev is False, "the parser sets allow_abbrev=False")
    with _patched(self_test=lambda: 1):
        red = main(["--self-test"])
    with _patched(self_test=lambda: 0):
        green = main(["--self-check"])
    check(red == 1 and green == 0,
          "main returns the exit status of --self-test, so a red self-test fails CI  <-- pinned defect")
    check(p.parse_args([]).edges is False and p.parse_args([]).wcag is False,
          "--edges and --wcag are off unless given, so a plain run is unchanged")
    check(build_parser().parse_args([]).out != "" and OUTPUT_PREFIX in build_parser().parse_args([]).out,
          "the default output name uses the configured prefix")


def _t_source(t):
    check = t.check
    path = os.path.abspath(__file__)
    with open(path, "rb") as fh:
        raw = fh.read()
    check(all(b < 128 for b in raw), "the source is pure ASCII")
    ast.parse(raw.decode("ascii"), feature_version=(3, 9))
    check(True, "the source parses as Python 3.9")
    spec = importlib.util.spec_from_file_location("sightline_imported", path)
    imported = importlib.util.module_from_spec(spec)
    buf = io.StringIO()
    # No bytecode cache, so the self-test writes nothing outside its own temporary directory.
    before = sys.dont_write_bytecode
    sys.dont_write_bytecode = True
    try:
        with contextlib.redirect_stdout(buf):
            spec.loader.exec_module(imported)
    finally:
        sys.dont_write_bytecode = before
    check(buf.getvalue() == "" and imported.covers("public", None, "u", "org", None, "u")[0] is True,
          "importing the module prints nothing and exposes the core")
    check(imported.pd is None and imported.GIS is None,
          "importing the module loads neither pandas nor the ArcGIS library  <-- pinned defect")


def _t_report(t, tmp, rest):
    """The Excel writer and the whole run.  Needs pandas and openpyxl, which a live run needs too."""
    check = t.check
    import openpyxl
    base = rest.split("/sharing/")[0]
    with _patched(pd=None, get_column_letter=None):
        load_report_libs()
        check(pd is not None and get_column_letter(3) == "C",
              "load_report_libs() puts pandas and openpyxl's column helper where the writer looks")
        from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE as theirs
        check(ILLEGAL_CHARACTERS_RE.pattern == theirs.pattern,
              "the copied illegal-character pattern equals openpyxl's  <-- pinned defect")
        # a hostile item title must not survive as a live Excel formula
        ws = openpyxl.Workbook().active
        ws["A1"] = '=cmd|" /C calc"!A0'
        check(ws["A1"].data_type == "f", "openpyxl still types a leading = as a formula; revisit defuse() if not")
        defuse(ws)
        check(ws["A1"].data_type == "s" and ws["A1"].value == '=cmd|" /C calc"!A0',
              "defuse() makes the title inert and leaves its text alone")
        check(scrub(pd.DataFrame({"a": ["ok\x02bad"]}))["a"][0] == "okbad",
              "control characters are stripped rather than crashing the write")
        check(len(summarize([], {"maps scanned": 0})) > 0,
              "an org with no web maps summarizes into a DataFrame, not a crash")
        with warnings.catch_warnings():
            warnings.simplefilter("error")
            only_error = frame(error_rows([({"MapTitle": "t", "MapAccess": "org"}, "x")]), ALL_COLS)
            nothing = frame([], EDGE_COLS)
        check(len(only_error) == 1 and only_error["LayerTitle"][0] == "" and len(nothing) == 0,
              "a column that no row filled is blank, with no pandas deprecation warning  <-- pinned defect")

        # write_excel / save_report on a real file
        meta = {"MapTitle": "=evil()", "MapId": "m1", "MapOwner": "u", "MapAccess": "org",
                "MapViews": 1, "MapUrl": "https://p/m1"}
        rows = build([(meta, "operationalLayers", {"title": "L\x02", "itemId": "x", "url": "https://h/a"})],
                     {}, {}, {}, {"https://h/a": ("DEAD", "code 400")}, collections.Counter(), "o")
        df = frame(rows, ALL_COLS)
        findings = df[df.Verdict != "OK"][FIND_COLS]
        xlsx = os.path.join(tmp, "direct.xlsx")
        edf = frame([{"AppUrl": "https://p/a", "MapUrl": "https://p/m"}], EDGE_COLS)
        write_excel(xlsx, findings, df, summarize(rows, {"k": "v"}), [("Edges", edf, ("AppUrl", "MapUrl"))])
        wb = openpyxl.load_workbook(xlsx)
        check(wb.sheetnames == ["Findings", "AllLayers", "Summary", "Edges"], "the workbook has the three sheets and the extra one")
        f = wb["Findings"]
        check(f["A1"].value == "Severity" and f.freeze_panes == "A2" and f["D2"].value == "=evil()"
              and f["D2"].data_type == "s", "headers are frozen and a formula-shaped title is stored as text")
        check(f["H2"].hyperlink is not None and f["I2"].value == "L", "map links are live and control characters are gone")
        check(wb["Edges"]["K2"].hyperlink is not None, "extra sheets get live links on the columns they name")
        real = write_excel
        state = []

        def locked(path, *args):
            state.append(path)
            if len(state) == 1:
                raise PermissionError("locked")
            return real(path, *args)

        target = os.path.join(tmp, "locked.xlsx")
        with _patched(write_excel=locked), _quiet() as (out, _):
            written = save_report(target, findings, df, summarize(rows, {}))
        check(written != target and os.path.exists(written) and "is locked" in out.getvalue(),
              "a locked report is written under a second name, never lost  <-- pinned defect")
        check(save_report(os.path.join(tmp, "free.xlsx"), findings, df, summarize(rows, {})).endswith("free.xlsx"),
              "an unlocked report is written where it was asked to be")

        # the whole run, against the stub and a fake portal
        org = "org1"
        m1, m2, m3, lay, dash, story = (_hid(n) for n in (1, 2, 3, 10, 20, 21))
        pub_map = {"operationalLayers": [
            {"title": "Org layer", "itemId": lay},
            {"title": "Open service", "url": base + "/ok/FeatureServer/0"},
            {"title": "Token service", "url": base + "/token/FeatureServer/0"},
            {"title": "Deleted", "itemId": "zz"},
            {"title": "Drawn here", "featureCollection": {"layers": []}}]}
        bad_map = {"operationalLayers": [{"url": base + "/ok/FeatureServer/0", "disablePopup": True}]}
        pool = dict((i.id, i) for i in (
            _Item(m1, title="Parks", owner="alice", access="public", data=pub_map, snippet="Park finder"),
            _Item(m2, title="Untitled Map", owner="alice", access="org", data=bad_map),
            _Item(m3, title="Staff map", owner="alice", access="org", data={}, snippet="s"),
            _Item(lay, title="Org layer", owner="bob", access="org", kind="Feature Service", orgId=org),
            _Item(dash, title="Public board", owner="alice", access="public", kind="Dashboard",
                  data={"widgets": [{"itemId": m3}, {"itemId": lay}]}),
            _Item(story, title="Org story", owner="alice", access="org", kind="StoryMap",
                  data={"nodes": {"a": {"data": {"itemId": m1}}}})))
        gis = _Gis(rest, pool)
        secret = "sentinel-pw-7431"
        out_path = os.path.join(tmp, "run.xlsx")

        class FL(object):
            def __init__(self, url, g):
                self.properties = {}

            def query(self, return_count_only):
                return 7 if return_count_only is True else {"features": []}

        def run(argv):
            with _quiet() as (out, err):
                code = main(argv)
            return code, out.getvalue(), err.getvalue()

        def put(path, data):
            with open(path, "wb") as fh:
                fh.write(data)

        def get(path, n=-1):
            with open(path, "rb") as fh:
                return fh.read(n)

        with _patched(connect=lambda: gis):
            code, out, err = run(["-o", out_path, "--edges", "--wcag"])
        check(code == 0 and os.path.exists(out_path) and "findings of" in out, "a full run exits 0 and writes the report")
        wb = openpyxl.load_workbook(out_path)
        check(wb.sheetnames == ["Findings", "AllLayers", "Summary", "Edges", "WCAG"],
              "--edges and --wcag add their own sheets, got %s" % wb.sheetnames)
        verdicts = [r[1] for r in wb["Findings"].iter_rows(min_row=2, values_only=True)]
        check(sorted(verdicts) == ["BROKEN", "BROKEN", "DEAD"],
              "the org layer, the token service and the deleted layer are the findings, got %s" % verdicts)
        edge_rows = [r for r in wb["Edges"].iter_rows(min_row=2, values_only=True)]
        check([(r[1], r[3]) for r in edge_rows] == [("BROKEN", "Dashboard"), ("OK", "StoryMap")],
              "the public dashboard on the org map is the edge finding, listed before the clean one, got %s"
              % [(r[1], r[3]) for r in edge_rows])
        wrows = [r[1] for r in wb["WCAG"].iter_rows(min_row=2, values_only=True)]
        check(sorted(wrows) == ["layer-no-title", "layer-popup-off", "map-no-summary", "map-title-placeholder"],
              "the untitled, summary-less map yields four rule rows, got %s" % wrows)
        summary = dict(((r[0], r[1]), r[2]) for r in wb["Summary"].iter_rows(min_row=2, values_only=True))
        check(summary[("Edges", "BROKEN")] == 1 and summary[("WCAG", "layer-no-title")] == 1
              and summary[("Run", "apps scanned")] == 2
              and summary[("Run", "candidate ids that were not web maps")] == "not a web map 1",
              "the summary counts the edges, the rules and the candidates that were not maps")
        check("Traceback" not in err, "a clean run prints no traceback")
        check([c.value for c in wb["AllLayers"][1]] == ALL_COLS and "INCOMPLETE" not in out,
              "without --deep the AllLayers sheet has exactly the standard columns  <-- pinned defect")

        put(out_path, b"old")
        with _patched(connect=lambda: gis):
            refused = _exits(lambda: run(["-o", out_path]))
        check(refused == 2 and get(out_path) == b"old",
              "an existing report is not replaced without --apply  <-- pinned defect")
        with _patched(connect=lambda: gis):
            code, _out, _err = run(["-o", out_path, "--apply"])
        check(code == 0 and get(out_path, 2) == b"PK", "--apply replaces it")

        empty = os.path.join(tmp, "empty.xlsx")
        with _patched(connect=lambda: gis):
            code, out, _err = run(["-o", empty, "--limit", "0", "--edges", "--wcag"])
        check(code == 0 and "Scanning 0 web maps" in out and os.path.exists(empty),
              "--limit 0 means zero, and an empty audit still writes its report  <-- pinned defect")

        # the report must not carry a credential that sat in a layer url
        leaky = {"operationalLayers": [
            {"url": base + "/ok/FeatureServer/0?token=TOKEN-LEAK-42&where=1%3D1"},
            {"title": "Named", "url": "https://user:PW-LEAK-42@" + base.split("//")[1] + "/ok/FeatureServer/0"}]}
        gis_leak = _Gis(rest, dict((i.id, i) for i in (
            _Item(_hid(40), title="Leaky", owner="alice", access="public", data=leaky, snippet="s"),)))
        leak_path = os.path.join(tmp, "leak.xlsx")
        with _patched(connect=lambda: gis_leak):
            code, out, err = run(["-o", leak_path, "--wcag"])
        wb = openpyxl.load_workbook(leak_path)
        cells = [c.value for c in wb["AllLayers"]["L"][1:]]
        blob = out + err + "".join(_zip_text(leak_path))
        check(code == 0 and "TOKEN-LEAK-42" not in blob and "PW-LEAK-42" not in blob
              and cells[0] == base + "/ok/FeatureServer/0?where=1%3D1"
              and wb["WCAG"]["K2"].value == base + "/ok/FeatureServer/0?where=1%3D1",
              "no token or password from a layer url reaches the cells, the links, or the console  <-- pinned defect")

        # --limit must not hide the groups of a map that only an app embeds
        pool_lim = dict((i.id, i) for i in (
            _Item(_hid(30), title="Front", owner="alice", access="public", data={}, snippet="s"),
            _Item(_hid(1), title="Embedded", owner="alice", access="shared", data={}, snippet="s"),
            _Item(_hid(20), title="Board", owner="alice", access="shared", kind="Dashboard",
                  data={"widgets": [{"itemId": _hid(1)}]})))
        lim_path = os.path.join(tmp, "lim.xlsx")
        with _patched(connect=lambda: _Gis(rest, pool_lim)):
            code, out, err = run(["-o", lim_path, "--edges", "--limit", "1"])
        edges = list(openpyxl.load_workbook(lim_path)["Edges"].iter_rows(min_row=2, values_only=True))
        check(code == 0 and [(r[1], r[3]) for r in edges] == [("OK", "Dashboard")],
              "the groups of a map that is embedded but not scanned are still read  <-- pinned defect")

        # a run that read nothing, or could not read something, must not look complete
        for label, items_, argv in (
                ("no web maps", {}, ["-o", os.path.join(tmp, "none.xlsx")]),
                ("a map that cannot be read", {_hid(50): _Item(_hid(50), data=RuntimeError("x"))},
                 ["-o", os.path.join(tmp, "bad.xlsx")]),
                ("an app that cannot be read",
                 {_hid(51): _Item(_hid(51), data={}),
                  _hid(52): _Item(_hid(52), kind="Dashboard", data=RuntimeError("x"))},
                 ["-o", os.path.join(tmp, "app.xlsx"), "--edges"])):
            with _patched(connect=lambda items_=items_: _Gis(rest, items_)):
                code, out, err = run(argv)
            check(code == 3 and "INCOMPLETE" in out and os.path.exists(argv[1]),
                  "%s: the report is written and the exit code is 3, not 0  <-- pinned defect" % label)

        # a folder that does not exist is refused before any portal call
        gone = os.path.join(tmp, "no_such_folder", "r.xlsx")

        contacted = []
        with _patched(connect=lambda: contacted.append(1)), _quiet() as (_o, err):
            code = _exits(lambda: main(["-o", gone]))
        check(code == 2 and "does not exist" in err.getvalue() and not contacted
              and not os.path.exists(os.path.dirname(gone)),
              "an output folder that does not exist is refused before the audit starts  <-- pinned defect")

        deep = os.path.join(tmp, "deep.xlsx")
        with _patched(GIS=lambda *a: gis, FeatureLayer=FL, USERNAME="", PASSWORD=""), \
                _environ(ARCGIS_USER="svc", ARCGIS_PASS=secret, ARCGIS_URL=None):
            code, out, err = run(["-o", deep, "--deep"])
        wb = openpyxl.load_workbook(deep)
        head = [c.value for c in wb["AllLayers"][1]]
        check(code == 0 and head[-2:] == ["FeatureCount", "LastEditUTC"], "--deep adds the count and last-edit columns")
        blob = out + err + "".join(_zip_text(deep))
        check(secret not in blob, "the password reaches neither the console nor the report")
        check(wb.sheetnames == ["Findings", "AllLayers", "Summary"], "without --edges and --wcag the workbook is unchanged")


def _zip_text(path):
    import zipfile
    with zipfile.ZipFile(path) as z:
        for name in z.namelist():
            yield z.read(name).decode("utf-8", "replace")


# --------------------------------------------------------------------------- main

def edge_order(rows):
    """Findings first, worst audience first, busiest app first; clean edges last."""
    return sorted(rows, key=lambda r: (r["Verdict"] == "OK", r["Severity"],
                                       -int(r.get("AppViews") or 0), r["AppTitle"], r["MapTitle"] if "MapTitle" in r else ""))


def wcag_order(rows):
    return sorted(rows, key=lambda r: (r["Severity"], -int(r["MapViews"] or 0), r["MapTitle"], r["Rule"], r["Layer"]))


def build_parser():
    # allow_abbrev=False: a prefix such as --self-ch or --ap must not silently expand to a flag
    p = argparse.ArgumentParser(description="Audit AGOL web maps, and the apps that embed them, "
                                            "for viewers who cannot open them.",
                                allow_abbrev=False)
    default = os.path.join(OUTPUT_DIR, "%s_%s.xlsx"
                           % (OUTPUT_PREFIX, datetime.datetime.now().strftime("%Y%m%d_%H%M%S")))
    p.add_argument("-o", "--out", default=default, help="output .xlsx (default: %(default)s)")
    p.add_argument("--limit", "--max", type=int, dest="limit", metavar="N", default=DEFAULT_LIMIT,
                   help="cap maps scanned, and apps per type; smoke-test before the full run")
    p.add_argument("--deep", action="store_true", default=DEEP_DEFAULT,
                   help="add FeatureCount + LastEditUTC (~30s)")
    p.add_argument("--edges", action="store_true",
                   help="also flag apps shared wider than the web map they embed")
    p.add_argument("--wcag", action="store_true",
                   help="also run the accessibility rule pack over every web map")
    p.add_argument("--apply", action="store_true",
                   help="replace the output file if it already exists")
    p.add_argument("--self-test", "--self-check", dest="self_test", action="store_true",
                   help="run the offline assertions and exit")
    return p


def main(argv=None):
    p = build_parser()
    a = p.parse_args(argv)
    if a.self_test:
        return self_test()
    # Refuse before the multi-minute audit, not after it.  A new file needs no flag.
    folder = os.path.dirname(os.path.abspath(a.out))
    if not os.path.isdir(folder):
        p.error("the folder %s does not exist; create it or choose another -o" % folder)
    if os.path.exists(a.out) and not a.apply:
        p.error("%s already exists; add --apply to replace it" % a.out)
    load_report_libs()

    t0 = time.time()
    gis = connect()
    org = gis.properties["name"]
    org_id = gis.properties["id"]
    print("Connected to %s as %s" % (org, gis.users.me.username))

    # `a.limit is None` rather than a truthiness test: --limit 0 must mean zero, not "all".
    maps = gis.content.search("", item_type="Web Map",
                              max_items=MAX_MAPS if a.limit is None else a.limit)
    print("Scanning %d web maps..." % len(maps))
    wcag_rows = [] if a.wcag else None
    refs, errors = collect(gis, maps, wcag_rows)
    apps, app_errors = collect_apps(gis, a.limit) if a.edges else ([], [])
    if a.edges:
        print("Read %d apps (%d unreadable)" % (len(apps) + len(app_errors), len(app_errors)))

    ids = set(l.get("itemId") for _, _, l in refs if l.get("itemId"))
    cands = set(c for _, cs in apps for c in cs)
    print("%d layer refs, %d distinct layer items" % (len(refs), len(ids)))
    items, unresolved = resolve_items(gis, ids | cands)
    embedded = set(c for c in cands if c in items and items[c]["type"] == "Web Map")
    app_ids = set(m["AppId"] for m, _ in apps)
    gids, gtitle = fetch_groups(gis, ids | set(m["id"] for m in maps) | embedded | app_ids)

    probes, fails = {}, collections.Counter()
    rows = build(refs, items, gids, gtitle, probes, fails, org_id, unresolved) + error_rows(errors)
    df = frame(rows, ALL_COLS)

    if a.deep:
        stats = deep_stats(gis, set(df.LayerUrl) - {""})
        df["FeatureCount"] = [stats.get(u, ("", ""))[0] for u in df.LayerUrl]
        df["LastEditUTC"] = [stats.get(u, ("", ""))[1] for u in df.LayerUrl]

    extra, edge_rows = [], None
    if a.edges:
        edge_rows, skipped = build_edges(apps, app_errors, items, gids, gtitle, item_url_base(gis), unresolved)
        edge_rows = edge_order(edge_rows)
        extra.append(("Edges", frame(edge_rows, EDGE_COLS), ("AppUrl", "MapUrl")))
    if a.wcag:
        wcag_rows = wcag_order(wcag_rows)
        extra.append(("WCAG", frame(wcag_rows, WCAG_COLS), ("MapUrl",)))

    meta = {
        "organization": org, "signed in as": gis.users.me.username,
        "started": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "maps scanned": len(maps), "maps that failed to read": len(errors),
        "layer refs": len(refs), "distinct layer items": len(ids),
        "distinct probed urls": len(probes),
        "feature counts": "included (--deep)" if a.deep else "skipped, re-run with --deep",
    }
    if a.edges:
        meta["apps scanned"] = len(apps) + len(app_errors)
        meta["candidate ids that were not web maps"] = ", ".join(
            "%s %d" % kv for kv in sorted(skipped.items())) or "none"
    findings = df[df.Verdict != "OK"].sort_values(
        ["Severity", "MapViews", "MapTitle", "LayerTitle"],
        ascending=[True, False, True, True])[FIND_COLS]
    meta["elapsed s"] = round(time.time() - t0, 1)
    written = save_report(a.out, findings, df, summarize(rows, meta, edge_rows, wcag_rows), extra)

    print("\n%d findings of %d layer refs in %.0fs -> %s"
          % (len(findings), len(df), time.time() - t0, written))
    for v, n in df.Verdict.value_counts().items():
        print("  %-8s %d" % (v, n))
    if a.edges:
        print("  %d app edges, %d not openable by every viewer"
              % (len(edge_rows), sum(1 for r in edge_rows if r["Verdict"] != "OK")))
    if a.wcag:
        print("  %d accessibility findings" % len(wcag_rows))
    # A report was written, but a run that read nothing, or could not read some items, did not
    # audit everything.  Exit 3 lets a scheduled task tell that from a complete run.
    if errors or app_errors or (not maps and a.limit != 0):
        print("  INCOMPLETE: %d maps and %d apps could not be read, %d maps were scanned; exit 3"
              % (len(errors), len(app_errors), len(maps)))
        return 3
    return 0


if __name__ == "__main__":
    sys.exit(main())
