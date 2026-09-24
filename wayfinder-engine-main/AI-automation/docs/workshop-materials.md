# Generate Workshop Materials

## Step 2 — Check Existing Workshop Materials on SharePoint

Before generating, check which stages already have workshop files uploaded:

```bash
cd AI-automation && python scripts/sync/list_sharepoint_workshop_status.py \
  --stage "../Wayfinder-content/source/wayfinder/001_PREPARE/stages-library/001:PLF - Assess & Design/stage.md"
```

## Step 3 — Generate or iterate

### New stage (no existing content)

Ask the AI to generate workshop slides from scratch:

| Tool        | Command                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                            |
| ----------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| Claude Code | `@instructions/shared/01-safeguards.md @instructions/slides-general/01-dynatrace-styleguide.md @instructions/slides-general/02-visual-emphasis.md @instructions/slides-general/03-slide-layouts.md @instructions/slides-general/04-image-usage.md @instructions/slides-general/05-json-format.md @instructions/slides-workshop/01-technical.md Generate workshop for stage ../Wayfinder-content/source/wayfinder/001_PREPARE/stages-library/001:PLF - Assess & Design/stage.md`                                    |
| Copilot     | `#file:instructions/shared/01-safeguards.md #file:instructions/slides-general/01-dynatrace-styleguide.md #file:instructions/slides-general/02-visual-emphasis.md #file:instructions/slides-general/03-slide-layouts.md #file:instructions/slides-general/04-image-usage.md #file:instructions/slides-general/05-json-format.md #file:instructions/slides-workshop/01-technical.md Generate workshop for stage ../Wayfinder-content/source/wayfinder/001_PREPARE/stages-library/001:PLF - Assess & Design/stage.md` |

The AI reads `stage.md` and the stage's cached `best-practices/` folder under
`AI-automation/wayfinder-post-process/<phase>/stages-library/<stage>/`, produces a
`.response.json`, and runs the build to generate the `.pptx`. Upload to SharePoint
is a separate step (see Step 4).

> Prerequisite: run [`scripts/sync/pull_best_practices.py`](../scripts/sync/pull_best_practices.py)
> against the stage first if the cache is empty. See [confluence-sync.md](confluence-sync.md).

### Existing stage (iterate)

Ask the AI to review an existing workshop:

```
Update workshop for stage `../Wayfinder-content/source/wayfinder/001_PREPARE/stages-library/001:PLF - Assess & Design/stage.md`
```

When a `.response.json` already exists, the AI asks to **Iterate** or **Abort** — it
never silently overwrites. Choose **Iterate** to review the deck without changing content.

**What iterate does:**

1. Downloads the living PPTX from SharePoint (the `.response.json` is NOT read or rebuilt)
2. Extracts slide text from the downloaded PPTX
3. Reads the current `best-practices/` and `public-docs/` for the stage from `AI-automation/wayfinder-post-process/<phase>/stages-library/<stage>/`
4. Identifies what is outdated, missing, or violates the style guide
5. Injects native PowerPoint review comments on every flagged slide
6. Appends new slides if a topic is missing entirely
7. Uploads back to SharePoint with eTag conflict detection

Comments appear in PowerPoint's **Review pane** under the author "Wayfinder Engine".
Read each comment, make the change directly in PowerPoint if you agree, and delete
the comment when done. You can also ask the agent to apply a batch of changes if preferred.

#### Comment format

Each comment is self-contained with three sections:

```
Old: <exact current text>

New: <exact replacement>

Why: <one sentence reason>
```

#### Iterate pipeline commands

The AI runs these steps automatically during an iterate session. You can also run them manually:

```bash
cd AI-automation && STAGE="Segment Data"
DECK="outputs/stages/$STAGE/workshop/segment-data-workshop-45min.pptx"
REMOTE="Wayfinder/stages/$STAGE/workshop/segment-data-workshop-45min.pptx"

# 1. Download the living PPTX from SharePoint + save eTag
python scripts/sync/fetch_pptx_for_review.py "$REMOTE" --out "$DECK"

# 2. Inject review comments (one --comment per issue)
python scripts/content-creation/inject_pptx_comments.py "$DECK" \
  --comment $'Slide Title:Old: "exact current text"\n\nNew: "replacement"\n\nWhy: one sentence reason'

# 3. (Optional) Append new slides
python scripts/content-creation/append_pptx_slides.py "$DECK" \
  --section-title "New section" \
  --slides '[{"title":"New slide","layout":"standard","bullets":["Bullet"]}]'

# 4. Upload back with conflict detection
python scripts/sync/upload_to_sharepoint.py "$DECK" \
  --folder "Wayfinder/stages/$STAGE/workshop" \
  --force --if-match
```

`--if-match` reads the `.etag` file saved by `fetch_pptx_for_review.py` and aborts
if someone else edited the file on SharePoint between your fetch and upload.

---

## JSON lifecycle

`.response.json` files are **initial-build artifacts only** — not the ongoing source of truth.

| Phase                             | Authoritative source                       |
| --------------------------------- | ------------------------------------------ |
| After `Generate workshop`         | `.response.json` → build → upload          |
| After first upload                | SharePoint PPTX is the living document     |
| All future `Update workshop` runs | Download from SharePoint; JSON is not read |

Every `.response.json` includes a `_meta` block that records its archived status and links to the SharePoint document:

```json
{
  "_meta": {
    "status": "initial-build-only",
    "warning": "This file generated the initial PPTX and is now archived. Do not rebuild or iterate from this file. The living document is on SharePoint.",
    "sharepoint_url": "https://dynatrace.sharepoint.com/...",
    "generated_at": "YYYY-MM-DD"
  }
}
```

## Step 4 — Push to SharePoint

Upload the generated PPTX to SharePoint and record the URL in the mapping:

```bash
cd AI-automation && STAGE="Post-ingest Enrichment"

find "outputs/stages/$STAGE" \
  \( -name '*.pptx' -o -name '*.xlsx' \) ! -name '~$*' | while read f; do
  REL="${f#outputs/stages/$STAGE/}"
  SUBFOLDER="$(dirname "$REL")"
  if [ "$SUBFOLDER" = "." ]; then
    FOLDER="Wayfinder/stages/$STAGE"
  else
    FOLDER="Wayfinder/stages/$STAGE/$SUBFOLDER"
  fi
  python scripts/sync/upload_to_sharepoint.py "$f" \
    --folder "$FOLDER" \
    --force \
    --update-mapping "$STAGE"
done
```

> **Initial upload only** — use `--force` here because this is the first time the file is pushed.
> For iterate uploads (after reviewing an existing deck), use `--if-match` instead of `--force`
> to detect concurrent edits (see iterate pipeline commands above).

## Step 5 — Sync SharePoint URLs into Mapping

Fetch the stage folder URLs from SharePoint and write them into `../Wayfinder-content/source/mermaid/wayfinder-v3-mapping.json`:

```bash
cd AI-automation && python scripts/sync/fetch_sharepoint_stage_urls.py
```

> Note: these links are used for the project planning.

---

## Manual comment injection

You can inject comments into any existing PPTX without running the full iterate flow:

```bash
cd AI-automation && python scripts/content-creation/inject_pptx_comments.py \
  "outputs/stages/Segment Data/workshop/segment-data-workshop-45min.pptx" \
  --comment $'Slide Title:Old: "current text"\n\nNew: "replacement"\n\nWhy: reason'
```

Slides are matched by title text. The comment author is "Wayfinder Engine".
