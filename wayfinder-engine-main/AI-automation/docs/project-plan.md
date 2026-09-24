# Generate Project Plan

## Workflow

```
[pathway.mmd](docs/pathway.md) → Project Plan (Excel) → PPTX
```

## Step 2 — Sync SharePoint URLs into Mapping

Fetch the stage folder URLs from SharePoint and write them into `../Wayfinder-content/source/mermaid/wayfinder-v3-mapping.json`:

```bash
cd AI-automation && python3 scripts/sync/fetch_sharepoint_stage_urls.py
```

## Step 3 — Build Project Plan

Ask AI to build the project plan:

| Tool        | Command                                                                                                                                                                                                                     |
| ----------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Claude Code | `@instructions/shared/01-safeguards.md @instructions/project-plan/02-effort-estimation.md @instructions/project-plan/03-project-plan.md Build project plan for outputs/projects/classic-to-3rd-gen-upgrade/`                |
| Copilot     | `#file:instructions/shared/01-safeguards.md #file:instructions/project-plan/02-effort-estimation.md #file:instructions/project-plan/03-project-plan.md Build project plan for outputs/projects/classic-to-3rd-gen-upgrade/` |

### Output

Each project plan produces four files in `outputs/<name>/`:

| File | Content |
|------|---------||
| `project-plan-<name>.xlsx` | Formatted Excel workbook ready for delivery |

## Step 4 — Build Presentation

Ask AI to build the positioning/business slides:

| Tool        | Command                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                            |
| ----------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Claude Code | `@instructions/shared/01-safeguards.md @instructions/project-plan/02-effort-estimation.md @instructions/project-plan/03-project-plan.md @instructions/slides-general/01-dynatrace-styleguide.md @instructions/slides-general/02-visual-emphasis.md @instructions/slides-general/03-slide-layouts.md @instructions/slides-general/04-image-usage.md @instructions/slides-general/05-json-format.md @instructions/slides-business/01-positioning.md Build positioning/business slides for outputs/projects/classic-to-3rd-gen-upgrade/`                                              |
| Copilot     | `#file:instructions/shared/01-safeguards.md #file:instructions/project-plan/02-effort-estimation.md #file:instructions/project-plan/03-project-plan.md #file:instructions/slides-general/01-dynatrace-styleguide.md #file:instructions/slides-general/02-visual-emphasis.md #file:instructions/slides-general/03-slide-layouts.md #file:instructions/slides-general/04-image-usage.md #file:instructions/slides-general/05-json-format.md #file:instructions/slides-business/01-positioning.md Build positioning/business slides for outputs/projects/classic-to-3rd-gen-upgrade/` |

## (Optional) Push to SharePoint

Upload all generated files (excluding `.md`) to the Post-Sales SharePoint site (`Shared Documents / Wayfinder / projects / <name>` )

```bash
cd AI-automation && NAME=classic-to-3rd-gen-upgrade
find "outputs/projects/$NAME" -maxdepth 1 \
  \( -name '*.xlsx' -o -name '*.pptx' -o -name '*.mmd' \) | while read f; do
  python scripts/sync/upload_to_sharepoint.py "$f" \
    --folder "Wayfinder/projects/$NAME"
done
```
