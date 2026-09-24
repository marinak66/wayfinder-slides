# Confluence Sync

## Prerequisites

`.env` with Confluence credentials (`CONFLUENCE_BASE_URL`, `CONFLUENCE_USERNAME`, `CONFLUENCE_API_TOKEN`, `CONFLUENCE_ROOT_PAGE_ID`).

## Step 1 - Add links to stage.md

Ensure one or both sections to `stage.md`:

**Best practices**:

```
## Best Practices Links
https://dt-rnd.atlassian.net/wiki/spaces/d1coe/pages/1247150978/1.+Assess+Design
```

**Public docs**:

```
## Official Doc Links
https://docs.dynatrace.com/docs/manage/segments/upgrade-guide-segments
https://docs.dynatrace.com/docs/ingest-from/setup-on-k8s/guides/metadata-automation/k8s-metadata-telemetry-enrichment
```

## Step 2 - Pull content (external sources → cache)

Run this before generating workshop materials to refresh the local cache.
Fetched content lands under `AI-automation/wayfinder-post-process/` (gitignored)
so it stays out of the curated `Wayfinder-content/` tree. Delete the stage's
cache first — the script only writes and never removes, so stale files from
old links would otherwise persist.

```bash
cd AI-automation && STAGE="../Wayfinder-content/source/wayfinder/001_PREPARE/stages-library/001:PLF - Assess & Design"
CACHE="wayfinder-post-process/001_PREPARE/stages-library/001:PLF - Assess & Design"
rm -rf "$CACHE/best-practices" "$CACHE/public-docs"
.venv/bin/python scripts/sync/pull_best_practices.py "$STAGE"
```

Result (cache layout, per stage):

```
AI-automation/wayfinder-post-process/<phase>/stages-library/<stage>/
├── best-practices/
│   ├── assets/
│   └── 1-assess-design.md
└── public-docs/
    └── upgrade-guide-segments.md
```

The curated tree stays clean:

```
Wayfinder-content/source/wayfinder/<phase>/stages-library/<stage>/
└── stage.md          ← the only manually authored file per stage
```

## Push Wayfinder to Confluence (repo → confluence)

**This repo is the source of truth for `stage.md` and `wayfinder.md`.** The
fetched best-practices/public-docs cache mirrors Confluence — edits to those
fetched files can still be pushed back to Confluence, but they will not be
detected by `--changed` (the cache is gitignored). Pass explicit paths
instead.

## Pushing local changes to Confluence

```bash
# Push all curated files changed since last commit (stage.md / wayfinder.md only)
cd AI-automation && python scripts/sync/push_to_confluence.py --changed

# Push a specific fetched best-practice (explicit path required — cache is gitignored)
python scripts/sync/push_to_confluence.py wayfinder-post-process/001_PREPARE/stages-library/.../best-practices/my-doc.md

# Push with per-file confirmation
python scripts/sync/push_to_confluence.py --changed --confirm
```
