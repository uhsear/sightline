# sightline

Audit ArcGIS Online web maps for layers their viewers cannot see.

For every layer in every web map, sightline answers one question:

> Is every viewer of the **map** also a viewer of the **layer**?

When the answer is no, the map is broken for somebody. A public map whose data is org-only
shows the public an empty canvas, and nobody gets told.

📖 **[Documentation](https://uhsear.github.io/sightline/)**

```
336 web maps, 3,583 layer references, 161 seconds

  OK       2941
  BROKEN    142   viewers of the map cannot load the layer
  DEAD      137   the layer's item or service no longer exists
  UNKNOWN   363   could not be determined
```

## Install

Needs `arcgis`, `pandas` and `openpyxl`. ArcGIS Pro ships all three, so you can use its
interpreter and install nothing:

```powershell
& "C:\Program Files\ArcGIS\Pro\bin\Python\envs\arcgispro-py3\python.exe" sightline.py --self-check
```

Otherwise:

```bash
pip install arcgis pandas openpyxl
python sightline.py --self-check
```

`--self-check` runs the assertion suite with no network and no credentials. If it prints
`self-check passed`, the install is good.

## Use

```bash
python sightline.py --limit 10     # smoke-test first
python sightline.py                # the real run
python sightline.py --deep         # also collect feature counts and last-edit dates
python sightline.py -o report.xlsx # choose the output path
```

Credentials resolve in this order, and none of them is a file sightline writes:

1. `USERNAME` / `PASSWORD` in the config block
2. `ARCGIS_USER` / `ARCGIS_PASS` environment variables
3. Your signed-in ArcGIS Pro session
4. An interactive prompt

Run it as an org administrator. A standard user only sees their own content, so the audit
covers a fraction of the org and everything else reads as clean.

## Configure

Everything tunable sits in one `CONFIG` block at the top of the script. Command-line flags
override the matching value.

| Setting | Default | What it does |
|---|---|---|
| `PORTAL_URL` | `https://www.arcgis.com` | Set to your portal URL for ArcGIS Enterprise |
| `USERNAME` / `PASSWORD` | blank | Leave blank, prefer env vars or the prompt |
| `MAX_MAPS` | `10000` | Ceiling on web maps fetched |
| `DEFAULT_LIMIT` | `None` | Cap every run without passing `--limit` |
| `DEEP_DEFAULT` | `False` | Always collect feature counts |
| `OUTPUT_DIR` | `~/Downloads` | Where the report goes |
| `PROBE_TIMEOUT` | `8` | Seconds per service URL |
| `MAX_HOST_FAILS` | `3` | Failures before a host is skipped |
| `SECURED_CODES` | `{401,403,498,499}` | REST codes meaning "needs a token" |
| `DEAD_CODES` | `{400,404}` | REST codes meaning "not there" |
| `ITEM_CHUNK` | `50` | Item ids per batched lookup |
| `SEV` | public → High | Severity band per map audience |

Do not put a real password in `PASSWORD`. A password typed into a script is a password you
will eventually commit.

## The report

Three sheets:

**Findings** is everything that is not OK, sorted worst first. This is the work queue.

**AllLayers** is every layer reference examined, including healthy ones, for pivoting.

**Summary** carries run metadata, verdict totals, the worst maps, and the owners with the
most problems. That last table is who you email.

### Verdicts

| Verdict | Meaning |
|---|---|
| `OK` | Every map viewer can load the layer |
| `BROKEN` | Some map viewers cannot load the layer |
| `DEAD` | The layer's item or service no longer exists |
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
| `org` | `org` or `shared` | OK |
| `shared` | `org` or `public` | BROKEN |
| `shared` | `shared` | Compare group sets |
| `shared` | `shared`, groups unknown | UNKNOWN |

A map shared to zero groups normalizes to `private`, since its real audience is the owner.

### Layers that are not items

In the org this was built against, 57% of layer references are not ArcGIS Online items at
all. They are registered service URLs, on-premises server layers and basemaps. Tools that
only compare item permissions mark all of them fine.

sightline fetches each distinct service URL anonymously, with no token, and reads what the
public would get:

| Response | Meaning |
|---|---|
| Normal | The layer really is public |
| `401` `403` `498` `499` | Needs a token. In a public map this is BROKEN |
| `400` `404` | The service is gone |
| Connection failure | UNKNOWN, with the reason recorded |

This catches a map titled "Internal" that is shared publicly and points at token-secured
services. No data leaks, but the map is broken for every visitor.

TLS verification stays on. A certificate a browser would reject is a finding, not something
to suppress.

## Why it is fast

The naive version of this audit takes about 19 minutes against a 336-map org. This one takes
under three. Three things account for the difference, each measured:

**Never touch an unknown attribute on an `Item`.** `Item.__getattr__` performs a network
round trip before raising `AttributeError`. One `hasattr()` call costs 0.665 seconds. The
tool sightline replaced called it twice per layer to fill two columns that were empty in 171
of 171 rows, spending roughly 95% of its runtime in dead code.

**Resolve items in batches.** One `id:(a OR b OR ...)` search resolves 50 ids per round trip
instead of one, measured 18.8 times faster with byte-identical results.

**Do not reach for threads.** On a shared `GIS` session, 4 workers gave 1.3x, 8 gave 1.1x
and 16 gave 0.7x, slower than sequential. The connection serialises, so sightline stays
sequential and readable.

## Limitations

**Web maps only.** Dashboards, Experience Builder apps, StoryMaps and Web Mapping
Applications are not scanned. Those usually outnumber web maps about two to one.

**Group membership is item-level.** sightline compares which groups a map and layer are
shared to. It does not enumerate group members, so it cannot tell you a group is empty.

**UNKNOWN is deliberate.** An unreachable host may be a firewall rule rather than a defect.
Check the `Reason` column before acting.

**A secured service in a non-public map stays UNKNOWN.** Whether your org members hold
credentials for an on-premises service cannot be determined from ArcGIS Online.

sightline is read-only. It changes nothing in your organisation.

## Contributing

Bug reports and feature requests go in
[GitHub issues](https://github.com/uhsear/sightline/issues).

Run `python sightline.py --self-check` before opening a pull request. The suite asserts the
full classification matrix, including the rule that `org` and `shared` are incomparable, so
reverting to a rank comparison fails loudly instead of quietly under-reporting.

## Author

Built by [Asir Khan](https://www.linkedin.com/in/asir-khan-310317264/).

## License

MIT. See [LICENSE](LICENSE).
