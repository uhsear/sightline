# sightline

Audit ArcGIS Online web maps, and the apps that embed them, for viewers who cannot open them.

A team publishes a storm dashboard for the public. The dashboard is shared with everyone. The
web map inside it is shared with the organization only, because the map was built first, for
staff, and nobody shared it again. Every person in the building opens the dashboard and sees
the map, because everyone in the building is signed in. The public opens the same link and sees
the dashboard frame around a grey rectangle. No error is raised. The dashboard loads, the map
widget does not, and nobody who could fix it ever sees the problem.

A sharing audit that reads each item on its own passes this. The map is consistent with its
layers. The dashboard is consistent with itself. The fault is on the edge between them: the app
reaches a wider audience than the map it embeds.

sightline checks every edge, from the layer up to the app. For every layer in every web map, it
answers one question:

> Is every viewer of the **map** also a viewer of the **layer**?

With `--edges`, it asks the same question one level up:

> Is every viewer of the **app** also a viewer of the **web map** it embeds?

With `--wcag`, it also runs six accessibility rules over the same web maps.

This is a run of `main()` against the synthetic portal that the self-test builds. The portal
object is a fake and the services are a local stub, so no portal and no login were involved:

```
Connected to Test Org as admin
Scanning 3 web maps...
  3/3 maps  0s elapsed  ~0s left
Read 2 apps (0 unreadable)
6 layer refs, 2 distinct layer items

3 findings of 6 layer refs in 0s -> demo.xlsx
  OK       3
  BROKEN   2
  DEAD     1
  2 app edges, 1 not openable by every viewer
  4 accessibility findings
```

The `Edges` sheet of that run held two rows, worst first:

```
1 High  BROKEN  Dashboard  Public board -> Staff map   app is shared public but the web map is shared org
-       OK      StoryMap   Org story -> Parks           every viewer of the app can load the web map
```

## What already exists

Esri's item pages show the relationship between an app and its map. Esri's ArcGIS blog post
["Discover how your layers, scenes, maps and apps connect"](https://www.esri.com/arcgis-blog/products/arcgis-online/announcements/discover-how-your-layers-scenes-maps-and-apps-connect)
describes "Items used" and "Used by" sections on item pages. Use them to see what depends on
what. I found no sharing comparison between the two ends in them, but I could not open the
post itself, so check the item page in your own organization.

The ArcGIS API for Python has an
[`itemgraph`](https://developers.arcgis.com/python/latest/guide/itemgraph-and-remap-data/)
module. It builds a directed graph of the items an item requires and the items that require
it. Its examples include a web map inside a StoryMap. Its documentation page does not describe
a sharing comparison. If you need the full graph, use it.

GitHub has community tools in the same area. A repository search on 2026-10-22 found these,
and I read only their descriptions. I did not run them.

- [Austin-AECEomnis/ArcGIS-Online-Sharing-Audit-Toolkit](https://github.com/Austin-AECEomnis/ArcGIS-Online-Sharing-Audit-Toolkit)
  audits item sharing across an organization and applies sharing updates.
- [valdesrosier/arcGIS-item-dependency-finder](https://github.com/valdesrosier/arcGIS-item-dependency-finder)
  and [phaakma/arcgis_item_graph](https://github.com/phaakma/arcgis_item_graph) trace and graph
  the dependencies between items.
- [dymaptic/accessible-map-agent](https://github.com/dymaptic/accessible-map-agent) writes
  accessible text descriptions for web maps with an AI model. sightline does not write
  descriptions. It only tells you which maps have none.

The search showed no tool that compares an app's audience with the audience of the map it
embeds. That shows only what the search found.

For the accessibility rules, the
[Massachusetts GIS accessibility guidance for web maps](https://www.mass.gov/info-details/gis-accessibility-guidance-web-maps)
gives guidance for people who build web maps. A browser tool such as axe checks a rendered page. sightline checks
what is in the web map's JSON, before anything is drawn.

## Install

`--self-test` needs Python 3.9 or newer and nothing else. A live run needs `arcgis`, `pandas`
and `openpyxl`. ArcGIS Pro ships all three, so you can use its interpreter and install nothing:

```powershell
& "C:\Program Files\ArcGIS\Pro\bin\Python\envs\arcgispro-py3\python.exe" sightline.py --self-test
```

Otherwise:

```bash
pip install arcgis pandas openpyxl
python sightline.py --self-test
```

`--self-test` runs the assertion suite with no portal and no credentials. It starts a small
web server on 127.0.0.1 and talks to that. If the last line reads `0 failed` and no `SKIP` line
sits above it, the install is good. A `SKIP` line means pandas or openpyxl is missing, so the
Excel report was not tested. `--self-check` is the old name for the same flag and still works.

## Use

```bash
python sightline.py --limit 10     # smoke-test first
python sightline.py                # the real run
python sightline.py --deep         # also collect feature counts and last-edit dates
python sightline.py --edges        # also flag apps shared wider than their web map
python sightline.py --wcag         # also run the accessibility rule pack
python sightline.py -o report.xlsx # choose the output path
python sightline.py -o report.xlsx --apply   # replace report.xlsx if it exists
```

| Flag | Default | What it does |
|---|---|---|
| `-o`, `--out` | `~/Downloads/sightline_<time>.xlsx` | The report path. |
| `--limit N` (`--max`) | all | Cap the web maps scanned, and the apps scanned per item type. |
| `--deep` | off | Add `FeatureCount` and `LastEditUTC` to the `AllLayers` sheet. |
| `--edges` | off | Add the `Edges` sheet: apps shared wider than the web map they embed. |
| `--wcag` | off | Add the `WCAG` sheet: the accessibility rule pack. |
| `--apply` | off | Replace the output file when it already exists. |
| `--self-test` (`--self-check`) | off | Run the offline assertions and exit. |

Every flag must be spelled in full. The parser sets `allow_abbrev=False`, so `--ap` is refused
with exit code 2 instead of being read as `--apply`, and `--self-ch` is refused instead of
being read as `--self-check`.

sightline is read-only toward your organization. The one thing it writes is the report. A new
report path needs no flag. An existing file is never replaced without `--apply`. The check for
an existing file, and the check that the output folder exists, run before the audit starts, so
a long run cannot end in a refusal.

Exit codes:

| Code | Meaning |
|---|---|
| `0` | A complete run, or a green self-test. Findings do not change the code. |
| `1` | A red self-test, or a missing library. |
| `2` | A refused flag, an output file that exists, or an output folder that does not. |
| `3` | The report was written, but the run is incomplete: no web maps were found, or a web map or app could not be read. The console prints `INCOMPLETE`. |

A scheduled task can test for `3` to tell a run that audited nothing from a clean one.

Credentials resolve as a pair, and none of them is a file sightline writes. Username comes
from `USERNAME` in the config block, or `ARCGIS_USER` if that is blank; password from
`PASSWORD`, or `ARCGIS_PASS`. Then:

1. Both found, it signs in with them.
2. No username anywhere, it uses your signed-in ArcGIS Pro session, and prompts if there is none.
3. Username but no password, it prompts for the password.

`ARCGIS_URL` overrides `PORTAL_URL`, because `PORTAL_URL` always holds a value.

Run it as an org administrator. A standard user only sees their own content, so the audit
covers a fraction of the org and everything else reads as clean.

**Scheduled tasks need a Python without ArcPy.** sightline imports the `arcgis` package when
it signs in. That package probes for ArcPy when it loads. Under a non-interactive scheduled
task, the ArcGIS Pro licence may not be initialised. The import then fails with "The Product
License has not been initialized", before sightline reaches your portal. The ArcGIS Pro Python
has ArcPy, so it has this problem. For a scheduled run, use a Python that has `arcgis`
installed but not `arcpy`. Set `ARCGIS_USER` and `ARCGIS_PASS` for that run too. Without ArcPy
there is no Pro session to sign in with, and a task cannot answer a prompt.

If the task must use the ArcGIS Pro Python, block ArcPy before `arcgis` loads. An import hook
that raises `ImportError` for `arcpy` makes the package use an engine without ArcPy.
sightline does not install this hook itself. Save this wrapper next to `sightline.py` as
`run_sightline.py`, and point the task at the wrapper:

```python
import builtins, os, runpy

os.environ.setdefault("ARCGIS_IO_ENGINE", "shapefile")
_real_import = builtins.__import__

def _no_arcpy(name, *args, **kwargs):
    if name.split(".")[0] in ("arcpy", "arcgisscripting"):
        raise ImportError("%s blocked for a scheduled run" % name)
    return _real_import(name, *args, **kwargs)

builtins.__import__ = _no_arcpy
here = os.path.dirname(os.path.abspath(__file__))
runpy.run_path(os.path.join(here, "sightline.py"), run_name="__main__")
```

`ARCGIS_IO_ENGINE=shapefile` stops the package from choosing ArcPy for file work. The hook
also removes Pro sign-in, so `ARCGIS_USER` and `ARCGIS_PASS` are still required.

## Configure

Everything tunable sits in one `CONFIG` block at the top of the script. Command-line flags
override the matching value.

| Setting | Default | What it does |
|---|---|---|
| `PORTAL_URL` | `https://www.arcgis.com` | Set to your portal URL for ArcGIS Enterprise |
| `USERNAME` / `PASSWORD` | blank | Leave blank, prefer env vars or the prompt |
| `MAX_MAPS` | `10000` | Ceiling on web maps fetched |
| `MAX_APPS` | `10000` | Ceiling on apps fetched, per item type (`--edges`) |
| `DEFAULT_LIMIT` | `None` | Cap every run without passing `--limit` |
| `DEEP_DEFAULT` | `False` | Always collect feature counts |
| `OUTPUT_DIR` | `~/Downloads` | Where the report goes |
| `OUTPUT_PREFIX` | `sightline` | Report filename is `<prefix>_YYYYmmdd_HHMMSS.xlsx` |
| `MAX_COL_WIDTH` | `55` | Excel column width cap |
| `WIDTH_SAMPLE_ROWS` | `400` | Rows sampled when sizing columns |
| `PROBE_TIMEOUT` | `8` | Seconds per service URL |
| `MAX_HOST_FAILS` | `3` | Failures before a host is skipped |
| `SECURED_CODES` | `{401,403,498,499}` | REST codes meaning "needs a token" |
| `DEAD_CODES` | `{400,404}` | REST codes meaning "not there" |
| `ITEM_CHUNK` | `50` | Item ids per batched lookup |
| `GROUP_CHUNK` | `80` | Item ids per `content/itemsgroups` call |
| `SEV` | public = High | Severity band per audience |
| `EDGE_TYPES` | four item types | Item types searched for `--edges` |
| `MAP_KEYS` | four JSON keys | Keys that may hold an embedded web map id |
| `PLACEHOLDER_TITLES` | six titles | Titles the `map-title-placeholder` rule flags |

Do not put a real password in `PASSWORD`. A password typed into a script is a password you
will eventually commit.

## The report

Three sheets, plus one for each of `--edges` and `--wcag`:

**Findings** is every layer reference that is not OK, sorted worst first. This is the work queue.

**AllLayers** is every layer reference examined, including healthy ones, for pivoting.

**Summary** carries run metadata, verdict totals, the worst maps, and the owners with the
most problems. That last table is who you email. With `--edges` it also counts edge verdicts.
With `--wcag` it counts findings per rule.

**Edges** (`--edges`) is every app-to-map edge, findings first. Severity follows the app's
audience, because the app is the link people open.

**WCAG** (`--wcag`) is one row per rule hit, with the rule, the WCAG criterion, the layer
and the reason. Severity follows the map's audience.

### Verdicts

| Verdict | Meaning |
|---|---|
| `OK` | Every map viewer can load the layer |
| `BROKEN` | Some map viewers cannot load the layer |
| `DEAD` | The layer's service is gone, or its item was not found by the signed-in account. The item may be deleted, or the account may not be able to see it. The `Reason` column says which |
| `UNKNOWN` | Could not be determined. The `Reason` column says why |

Severity follows the map's audience, because that is who is affected:
`1 High` public, `2 Medium` org, `3 Low` shared, `4 Info` private. Findings sorts by
severity then view count, so the most-visited broken public map is the first row.

The `Overexposed` column tracks a separate problem: a public layer inside a non-public map.
Nothing is broken, but the data is shared more widely than its map implies.

## How it decides

Sharing is audience containment, not a severity rank. The common shortcut
(`private < org = shared < public`) hides real breakage, because `org` and `shared` are
incomparable audiences. An org member outside the group cannot load a group-shared layer,
so an org-wide map with a group-shared layer is broken for most of its audience.

| Layer | Map | Result |
|---|---|---|
| `public` | anything | OK |
| anything | `private` | OK |
| `private` owned by someone else | `private` | BROKEN |
| `org` | `public` | BROKEN |
| `org` | `org` | OK |
| `org` | `shared` | UNKNOWN |
| `shared` | `org` or `public` | BROKEN |
| `shared` | `shared` | Compare group sets |
| `shared` | `shared`, groups unknown | UNKNOWN |

An item shared to zero groups normalizes to `private`, since its real audience is the owner.
This holds for the map and for the layer. A layer shared to no group, owned by someone else,
is therefore BROKEN in a private map.

An `org` layer in a group-shared map is `UNKNOWN`, not OK. Esri lets a group take members from
other organizations ([Share items](https://doc.arcgis.com/en/arcgis-online/share-maps/share-items.htm)),
and those members cannot load an org-only item. I took this from an audit
note and could not re-open the page when I wrote it, so check the page. sightline does not read group members, so it
cannot say that every member belongs to your organization.

Group sets are compared by group id, never by title. Two groups can have the same title, for
example two "GIS Staff" groups made by different owners. A comparison by title would call
them the same audience. Titles appear in the report only as labels.

### Layers that are not items

Many layer references are not ArcGIS Online items at all. They are registered service URLs,
on-premises server layers and basemaps. Tools that only compare item permissions mark all of
them fine.

sightline fetches each distinct service URL anonymously, with no token, and reads what the
public would get:

| Response | Meaning |
|---|---|
| Normal | The layer really is public |
| `401` `403` `498` `499` | Needs a token. In a public map this is BROKEN |
| `400` `404` | The service is gone |
| A 200 reply whose `error` is not an object with a code | UNKNOWN, reason `error reply of unknown shape` |
| Connection failure | UNKNOWN, with the reason recorded |

This catches a map titled "Internal" that is shared publicly and points at token-secured
services. No data leaks, but the map is broken for every visitor.

The probe sends the layer's own query string, because a layer built from a `.../query?where=...`
url needs it. It replaces any `f=` parameter with `f=json`. It drops every parameter whose name
contains `token`, `key`, `pass`, `pwd`, `secret`, `sig`, `auth` or `credential`, and it drops
any `user:password@` part, so the probe sends no credential. The report shows the same cleaned
url in the `LayerUrl` column, in its hyperlink and in the `WCAG` sheet. The name test is a
heuristic. A secret in a parameter with another name is still copied.

TLS verification stays on. A certificate a browser would reject is a finding, not something
to suppress.

## App edges (`--edges`)

`--edges` searches four item types, named as the Esri
[Items and item types](https://developers.arcgis.com/rest/users-groups-and-items/items-and-item-types/)
page spells them: `Dashboard`, `Web Mapping Application`, `StoryMap` and `Web Experience`
(Experience Builder). It reads each app's JSON once.

The app is treated as the "map" in the table above, and the embedded web map is treated as the
"layer". Read the table below as: the row is the web map's audience, the column is the app's.

| Web map is | App `public` | App `org` | App `shared` | App `private` |
|---|---|---|---|---|
| `public` | OK | OK | OK | OK |
| `org` | BROKEN | OK | UNKNOWN | OK |
| `shared` | BROKEN | BROKEN | Compare group sets | OK |
| `private` | BROKEN | BROKEN | BROKEN | OK if the owner is the same |

The same rules apply as for layers. An item shared to zero groups is `private`, on either
side. Unknown group membership on either side is `UNKNOWN`, never a silent OK. An `org` web map
under a group-shared app is `UNKNOWN`, for the reason given in the layer table.

**How an embedded map is found.** Esri documents the data of these item types only in part:
the item-type page above lists no JSON shape for any of the four. Esri's
[configurable templates](https://enterprise.arcgis.com/en/portal/10.5/use/configurable-templates.htm)
page (the 10.5 version) describes a `values` object that holds settings such as a web map id. I found no official
shape for Dashboard, StoryMap or Experience Builder data. So sightline does not parse one shape
per type. It collects every 32-digit hex string held under the key `itemId`, `webmap`,
`webmapId` or `mapId` anywhere in the app's JSON. It then looks those ids up. It keeps an id
only when it resolves to an item of type `Web Map`. A layer id, a scene id and a stale id all
drop out. The summary counts the ids it dropped.

This has one consequence you must know. A key I did not list is a missed edge, not a wrong
one. The edges sightline reports are real. The edges it misses are unknown. An app whose map
was deleted is not reported either, because a deleted id cannot be told from a deleted layer.

## Accessibility rules (`--wcag`)

The rules read the web map's JSON and its item summary. Nothing is drawn. Each rule names the
[WCAG 2.2](https://www.w3.org/TR/WCAG22/) success criterion it serves. A clean run is not a
conformance claim. These are the checks a web map's JSON can answer.

| Rule | WCAG | Level | Fires when |
|---|---|---|---|
| `map-title-placeholder` | 2.4.2 Page Titled | A | The item title is empty, or one of `PLACEHOLDER_TITLES` such as "Untitled Map". |
| `map-no-summary` | 1.1.1 Non-text Content | A | The item has no summary and no description. |
| `layer-no-title` | 2.4.6 Headings and Labels | AA | A layer has an empty or blank `title`. |
| `layer-popup-off` | 1.1.1 Non-text Content | A | A layer sets `disablePopup` to `true`. |
| `renderer-color-only` | 1.4.1 Use of Color | A | A `uniqueValue` renderer has two or more classes whose symbols share one shape, size and outline, and differ only by color. |
| `label-halo-contrast` | 1.4.3 Contrast (Minimum) | AA | A drawn label has a halo, and its text and halo are under 4.5 to 1. Large text needs 3 to 1. |

What each rule rests on, checked against the sources:

- WCAG 2.2 defines 1.4.1, 1.4.3, 1.1.1, 2.4.2 and 2.4.6 as sightline uses them. The contrast
  ratio is `(L1 + 0.05) / (L2 + 0.05)`, from relative luminance with the 0.04045 sRGB
  threshold. The W3C
  [Contrast (Minimum)](https://www.w3.org/WAI/WCAG22/Understanding/contrast-minimum.html)
  page gives 4.5 to 1 for normal text, 3 to 1 for large text, and defines large as at least
  18 point, or 14 point bold.
- The Esri web map specification says a layer `title` is "a user-friendly string title for the
  layer that can be used in a table of contents", that `disablePopup` indicates "whether to
  allow a client to ignore popups", that labels display when `showLabels` is `true`, and that
  `uniqueValueGroups` should be used instead of `uniqueValueInfos` when present. sightline
  reads `uniqueValueGroups` first for that reason. Colors are four-element arrays of red,
  green, blue and alpha, 0 to 255. Label `font.size` is in points. See the
  [layer](https://developers.arcgis.com/web-map-specification/objects/layer/),
  [feature layer](https://developers.arcgis.com/web-map-specification/objects/featureLayer/),
  [unique value renderer](https://developers.arcgis.com/web-map-specification/objects/uniqueValueRenderer/)
  and [text symbol](https://developers.arcgis.com/web-map-specification/objects/esriTS_symbol/)
  pages.
- The item `snippet` (summary) and `description` are item parameters in the REST
  [common parameters](https://developers.arcgis.com/rest/users-groups-and-items/common-parameters/)
  page.

What the rules choose not to do:

- The base map is Esri's content, so its layers are never judged.
- A color that is not fully opaque is skipped. It is drawn over a base map that sightline
  cannot see, so its contrast has no answer.
- `classBreaks` renderers are skipped. They show an ordered value as a ramp, which is a
  different design question from telling nominal classes apart.
- Picture and CIM symbols are skipped. They differ by more than color.
- A layer with no `popupInfo` is not flagged. The specification lets a client use the popup
  that the service defines, so a missing `popupInfo` does not mean popups are off.
- `layer-popup-off` is advisory. A reference layer may have good reason to hide its popup.

## Why it is fast

The naive version of this audit took about 19 minutes against an organization of a few
hundred web maps. This one takes under three. I measured the numbers below against a live
organization while building the tool. The self-test cannot reproduce them, because it has no
portal. Three things account for the difference:

**Never touch an unknown attribute on an `Item`.** `Item.__getattr__` performs a network
round trip before raising `AttributeError`. One `hasattr()` call costs 0.665 seconds. The
tool sightline replaced called it twice per layer to fill two columns that were empty in
every row, spending roughly 95% of its runtime in dead code.

**Resolve items in batches.** One `id:(a OR b OR ...)` search resolves 50 ids per round trip
instead of one, measured 18.8 times faster with byte-identical results.

**Do not reach for threads.** On a shared `GIS` session, 4 workers gave 1.3x, 8 gave 1.1x
and 16 gave 0.7x, slower than sequential. The connection serialises, so sightline stays
sequential and readable.

The self-test pins the rule and the batch size, not the times. Its fake `Item` raises an
`AssertionError` on any attribute that is not `id`, `get_data` or a dict key, so a full run
fails if the code ever probes one. A chunk test checks that ids go out 50 at a time.

`--wcag` adds no request. It reads the JSON that the layer audit already fetched. `--edges`
adds one search per item type, and one read per app.

## Self-test

```
$ python sightline.py --self-test
sightline self-test: 127.0.0.1 stub and temporary files only, no portal, no login
--------------------------------------------------------------------
PASS  check() and raises() really do record a failure  <-- pinned defect
...
PASS  an org-only layer in a group-shared map is unknown, since a group can hold outside members  <-- pinned defect
...
PASS  a layer shared to no group is invisible to a private map owned by someone else  <-- pinned defect
...
PASS  a token already in the layer url is not sent  <-- pinned defect
PASS  a layer url keeps its own query, with f=json swapped in and the token dropped  <-- pinned defect
...
PASS  a 200 whose error is a list is unknown, not a crash  <-- pinned defect
...
PASS  a failed lookup chunk is unresolved, never reported as deleted  <-- pinned defect
...
PASS  collect() prints its final progress line even when the last map fails  <-- pinned defect
...
PASS  a public app on an org map is broken
...
PASS  a group app on an org map is unknown: a group can hold members from another organisation  <-- pinned defect
...
PASS  an app shared to a group the map lacks names the group
...
PASS  an Experience Builder style data source of type WEB_MAP is followed to its map
...
PASS  two marker classes that differ only by color are flagged
...
PASS  label text under 4.5:1 on its halo is flagged
...
PASS  black text on a black halo is flagged, on a white halo it is fine  <-- pinned defect
...
PASS  a prefix of the write flag, --ap, is refused  <-- pinned defect
...
PASS  main returns the exit status of --self-test, so a red self-test fails CI  <-- pinned defect
...
PASS  an existing report is not replaced without --apply  <-- pinned defect
...
PASS  --limit 0 means zero, and an empty audit still writes its report  <-- pinned defect
PASS  no token or password from a layer url reaches the cells, the links, or the console  <-- pinned defect
PASS  the groups of a map that is embedded but not scanned are still read  <-- pinned defect
PASS  no web maps: the report is written and the exit code is 3, not 0  <-- pinned defect
...
PASS  an output folder that does not exist is refused before the audit starts  <-- pinned defect
...
--------------------------------------------------------------------
306 assertions, 0 failed
```

The full run prints all 306 assertions. Each `...` line is where this block is cut.
Lines marked `<-- pinned defect` are cases where a real defect was found, or could return, and
the assertion keeps it from coming back.

The count is the same on every host that has `pandas` and `openpyxl`:

| Host | Python | Count |
|---|---|---|
| Windows, ArcGIS Pro conda environment (has `arcgis`, `arcpy`, pandas, openpyxl) | 3.13.7 | 306 |
| Windows, virtual environment | 3.9.25, pandas 2.3.3, openpyxl 3.1.5 | 306 |
| Linux, virtual environment | 3.12.3, pandas 3.1.0 | 306 |
| Windows, plain `C:\Python313`, no pandas | 3.13.2 | 273 |
| Linux, system `python3`, no pandas | 3.12.3 | 273 |

The self-test never imports `arcgis` or `arcpy`, so they change nothing. On a Python without
`pandas` or `openpyxl`, 33 Excel assertions do not run. The run prints `SKIP  report writer`
where they would be, and a second `SKIP` line above the footer. The count drops to 273. The
run still ends `0 failed`, so read the `SKIP` line before you trust the footer. The two hosts
without pandas give the same count, 273, and the three with pandas give 306.

`coverage run --branch sightline.py --self-check` reports 100 percent of 1466 statements and
344 branches on the Windows ArcGIS Pro environment and on the Linux virtual environment.

## Limits

- **Never run against a live portal in this work.** The portal calls (`content.search`,
  `get_data`) are covered by a fake object with the same shape. The real `arcgis.gis.GIS`
  class insists on HTTPS (an `http://` portal URL raised an `SSLError` on the first request),
  so a stub for it needs a throwaway certificate, which I did not make. The login, the search
  and the item reads are therefore unproven against a real portal. The anonymous service probe
  and the group lookup ran against a real local web server.
- **App JSON is undocumented.** See "How an embedded map is found". The fixtures are synthetic
  and written from the rules above, not copied from a real app. A real Dashboard, StoryMap or
  Experience Builder app may keep its map id under a key that is not in `MAP_KEYS`.
- **`content/itemsgroups` is not in the REST reference.** I found no page for it. The group
  lookup reads the reply shape that the earlier versions of sightline read. The stub proves
  the parser reads that shape, not that your portal sends it.
- **The 498 and 499 codes are not in an Esri page I could find.** The code table was measured
  against live services when sightline was first built. The stub proves the table is applied,
  not that it is complete.
- **Web maps and four app types.** Web scenes, Hub sites and other item types are not scanned.
  An app that embeds a layer directly, with no web map, has no edge in this tool.
- **Group membership is item-level.** sightline compares which groups a map and layer are
  shared to. It does not enumerate group members, so it cannot tell you a group is empty.
- **UNKNOWN is deliberate.** An unreachable host may be a firewall rule rather than a defect.
  Check the `Reason` column before acting.
- **A secured service in a non-public map stays UNKNOWN.** Whether your org members hold
  credentials for an on-premises service cannot be determined from ArcGIS Online.
- **An embedded feature collection that is an empty object counts as no target.** The earlier
  versions did the same. It shows as `UNKNOWN`.
- **The `--wcag` description check strips tags with a pattern, not a parser.** A description
  made only of entities such as `&nbsp;` counts as text.
- **The probe fetches any url a web map holds.** A layer url whose scheme is not http or https
  goes to `urlopen` as given, redirects are followed, and private-network addresses are
  fetched too. A hostile web map can make the machine that runs sightline request an
  arbitrary url. Only a status is kept. Run it from a host whose network reach you accept.
- **The probe reads any JSON object with no `error` key as PUBLIC.** That includes a reply
  from a service that is not an ArcGIS service.
- **A string error code is UNKNOWN.** A reply such as `{"error": {"code": "498"}}` is not
  classed SECURED, because the code table holds integers only.
- **Malformed layer urls trip the host circuit breaker.** Three urls with a space, a control
  character or a non-ASCII character fail as transport errors. Later good urls on the same
  host then read `UNREACHABLE` instead of being probed.
- **A layer `itemId`, `url` or title that is not a string stops the run.** A number or an
  object there raises a `TypeError` or an `AttributeError` before any report is written.
- **An upper-case layer `itemId` reads as DEAD.** It is looked up as given, and the search
  returns lower-case ids, so a live layer shows as a false DEAD.
- **`Overexposed` only fires for org-owned layer items.** A public service found by the
  probe, or a public item from another organization, is never marked, which is narrower than
  "a public layer inside a non-public map".
- **Run as a non-admin, some apps get no edge row.** An app whose web map is private to
  someone else, or deleted, is missing from the item lookup. Only the `no item` count in the
  Summary sheet shows it.
- **An app that holds its map id under a key outside `MAP_KEYS` has no row in `Edges`.**
  Absence from the sheet is not a clean result.
- **After a failed item lookup, read the Summary sheet before you trust the edge count.** A
  candidate id whose lookup failed is dropped from `Edges` and counted in the Summary row
  `candidate ids that were not web maps` as `lookup failed N`. The console line and the exit
  code do not show it.
- **Owner names compare as exact strings.** `Alice` and `alice` are two owners.
- **A private map owned by A, with a layer shared to groups owned by B, reads OK.** sightline
  does not check that A is in those groups.
- **The edge reason for a public app on a group-shared map reads "shared shared".** The words
  are awkward and the verdict is right.
- **The report path check runs once, before the audit.** A second run that writes the same
  path during the audit replaces the file without `--apply`. The locked-file fallback also
  writes a new `<name>_HHMMSS.xlsx` without `--apply`, and its name has one-second
  resolution, so two locked runs in the same second share a name. The default output name has
  the same resolution.
- **Error text is copied into the report.** The text of a failed request or read goes into
  the `Reason` column and the console as the exception gives it. An exception that echoes its
  request url could carry a token from a legacy token-in-query login. `clean_url()` cleans the
  layer url columns only.
- **An item title with a lone surrogate stops the Excel write.** It is reachable only through
  the REST API, and it raises `UnicodeEncodeError` at the end of the run.
- **A title longer than 32767 characters is cut.** openpyxl cuts it in the workbook without a
  message.
- **Non-ASCII text can stop the console.** Printing an organization name or an error with
  characters outside the console code page, such as cp1252 in a redirected scheduled run,
  raises `UnicodeEncodeError`.
- **Hyperlink cells keep the url scheme.** A `javascript:`, `file:` or UNC url from a hostile
  web map becomes a clickable link. A url over about 2079 characters may make Excel offer to
  repair the file. Do not click links in a report from an organization you do not trust.
- **`fetch_groups()` expects the reply shape it knows.** A group block that is not an object
  raises an `AttributeError`. This needs a malformed reply.
- **`connect()` ignores `ARCGIS_PASS` when a Pro session is found.** That happens when no
  username is set.
- **`USERNAME` and `PASSWORD` in the config block win over the environment.** The safe
  path, an environment variable or the prompt, is a convention. Nothing enforces it.
- **The Mass.gov link was not opened by the author.** The page returned HTTP 403 to a
  scripted fetch, and a search did not confirm the exact address. Check that it resolves.
- **The "Product License has not been initialized" failure is reported, not reproduced.** It
  comes from scheduled-task experience. The self-test does not trigger it. The wrapper above
  ran in the self-test environment, and `arcgis.gis` imported with `arcpy` blocked.
- **Coverage on a bare interpreter is 84 percent.** Without pandas and openpyxl the Excel
  writer and the whole-run tests cannot execute, so 233 statements are missed. Those are
  the lines in `write_excel()`, `save_report()` and `_t_report()`. With both libraries it is
  100 percent.

## Contributing

Bug reports and feature requests go in
[GitHub issues](https://github.com/uhsear/sightline/issues).

Run `python sightline.py --self-test` before opening a pull request. The suite asserts the
full classification matrix, including the rule that `org` and `shared` are incomparable, so
reverting to a rank comparison fails loudly instead of quietly under-reporting.

## Author

Built by [Asir Khan](https://www.linkedin.com/in/asir-khan-310317264/).

## License

MIT. See [LICENSE](LICENSE).

## Related

Other single-file tools in this portfolio that pair with this one:

- [sharewatch](https://github.com/uhsear/sharewatch) - the same platform from the other side: what CHANGED last night, rather than what a viewer sees today
- [ghostsvc](https://github.com/uhsear/ghostsvc) - the anonymous services a viewer-side audit never reaches
- [whobreaks](https://github.com/uhsear/whobreaks) - what breaks if you delete one of the items it flags
- [deadwidget](https://github.com/uhsear/deadwidget) - the Experience Builder side: widgets bound to a layer or sublayer that no longer exists
