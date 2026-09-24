# Push to SharePoint

Upload generated deliverables (PPTX, XLSX, MMD, MD) to the Post-Sales SharePoint
site so they can be shared with delivery teams and customers. The upload
script ([scripts/sync/upload_to_sharepoint.py](../scripts/sync/upload_to_sharepoint.py)) handles a single file at a time;
the snippets below loop it across a folder of outputs.

## SharePoint layout

The Dynatrace SharePoint site has a `Wayfinder/` root folder with this
structure:

```
Wayfinder/
├── projects/
│   └── <project-name>/        ← customer-facing delivery plans
│       ├── mapping.md
│       ├── mapping.mmd
│       ├── mapping-simplified.mmd
│       ├── mapping-simplified-practitioner.mmd
│       ├── project-plan-<name>.xlsx
│       └── slides-<name>.pptx
└── stages/
    └── <stage>/               ← stage-specific workshop content
        └── workshop/
            └── *.pptx
```

`projects/` mirrors the local `outputs/projects/` folder. Each customer
project gets its own subfolder containing the full set of deliverables.

`stages/` mirrors stage-level training output (workshops, exercises,
slides) used by the delivery team.

## Auth

The script uses Microsoft Graph and supports two auth flows, picked
automatically from `.env`:

- **App-only (preferred for Sites.Selected)** — set `TENANT_ID`,
  `CLIENT_ID` and `CLIENT_SECRET`.
- **Device-code fallback** — when `CLIENT_SECRET` is not set, the script
  prompts an interactive login on first run and caches the token.

## Push projects (`outputs/projects/` → `Wayfinder/projects/`)

Run this after generating the project deliverables (mapping, .mmd
diagrams, Excel plan, PPTX slides) to keep SharePoint in sync.

```bash
# Upload one project
cd AI-automation && PROJECT="classic-to-3rd-gen-upgrade"

find "outputs/projects/$PROJECT" -type f \
  \( -name '*.md' -o -name '*.mmd' -o -name '*.xlsx' -o -name '*.pptx' \) \
  ! -name '~$*' | while read f; do
  python scripts/sync/upload_to_sharepoint.py "$f" \
    --folder "Wayfinder/projects/$PROJECT"
done
```

```bash
# Upload every project under outputs/projects/
cd AI-automation && for d in outputs/projects/*/; do
  PROJECT="$(basename "$d")"
  find "$d" -type f \
    \( -name '*.md' -o -name '*.mmd' -o -name '*.xlsx' -o -name '*.pptx' \) \
    ! -name '~$*' | while read f; do
    python scripts/sync/upload_to_sharepoint.py "$f" \
      --folder "Wayfinder/projects/$PROJECT"
  done
done
```

Project pushes do **not** touch `wayfinder-v3-mapping.json` — only stage
pushes use `--update-mapping`.

## Push stages (`outputs/stages/` → `Wayfinder/stages/`)

Upload the generated workshop materials for a single stage and record
the resulting SharePoint URL in `../Wayfinder-content/source/mermaid/wayfinder-v3-mapping.json`:

```bash
cd AI-automation && STAGE="Assess & Design"

find "outputs/stages/$STAGE" \
  \( -name '*.pptx' -o -name '*.xlsx' \) ! -name '~$*' | while IFS= read -r f; do
  REL="${f#outputs/stages/$STAGE/}"
  SUBFOLDER="$(dirname "$REL")"
  if [ "$SUBFOLDER" = "." ]; then
    FOLDER="Wayfinder/stages/$STAGE"
  else
    FOLDER="Wayfinder/stages/$STAGE/$SUBFOLDER"
  fi
  python scripts/sync/upload_to_sharepoint.py "$f" \
    --folder "$FOLDER" \
    --update-mapping "$STAGE"
done
```

This does two things:

1. Uploads every `.pptx` and `.xlsx` in the stage folder (and any
   nested subfolders) to `Shared Documents / Wayfinder / stages / <stage> / <subfolder>`.
   For example, `outputs/stages/Assess & Design/workshop/assess-design-workshop-45min.pptx`
   is uploaded to `Wayfinder/stages/Assess & Design/workshop/`.
2. Writes the SharePoint folder URL into
   `../Wayfinder-content/source/mermaid/wayfinder-v3-mapping.json` as a `sharepointUrl` field
   on the matching node entry.

The URL is preserved across mapping regenerations —
[map_stages.py](../scripts/map_stages.py) carries forward any extra
fields (like `sharepointUrl`) by matching on `nodeLabel`.

> [!NOTE]
> The `--update-mapping` flag expects the exact `nodeLabel` from the
> mermaid diagram (e.g. `"Assess & Design"`, not the `stage.md` filename).
> Only the last upload in the loop actually needs the flag, but running
> it on each file is harmless — it overwrites the same URL.

## Single-file upload

For ad-hoc uploads, point at any file and any target folder:

```bash
cd AI-automation && python scripts/sync/upload_to_sharepoint.py path/to/file.pptx \
  --folder "Wayfinder/projects/my-project"
```

The script logs the resulting SharePoint web URL on success.

---

## Workshop iterate: fetch → review → upload

For existing workshop decks, use the iterate workflow instead of a plain upload.
This downloads the living PPTX from SharePoint, lets you review and comment on it,
then uploads back with conflict detection.

```bash
cd AI-automation && STAGE="Segment Data"
DECK="outputs/stages/$STAGE/workshop/segment-data-workshop-45min.pptx"
REMOTE="Wayfinder/stages/$STAGE/workshop/segment-data-workshop-45min.pptx"

# 1. Download + save eTag (prints all slide text to stdout for AI review)
python scripts/sync/fetch_pptx_for_review.py "$REMOTE" --out "$DECK"

# 2. Inject review comments or append slides (see docs/workshop-materials.md)

# 3. Upload with conflict detection
python scripts/sync/upload_to_sharepoint.py "$DECK" \
  --folder "Wayfinder/stages/$STAGE/workshop" \
  --force --if-match
```

`--if-match` reads the `.etag` file written by `fetch_pptx_for_review.py`.
If the file on SharePoint was edited between your fetch and upload, the script
aborts with a clear message instead of overwriting the concurrent changes.

> **Do not use `--force` alone for iterate uploads.** `--force` skips the
> overwrite prompt but does not check for concurrent edits. Use `--if-match`
> (which also passes `--force`) so the eTag guard is always active.
