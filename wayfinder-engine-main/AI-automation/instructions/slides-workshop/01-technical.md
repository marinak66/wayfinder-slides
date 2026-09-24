# Slides Workshop

## Trigger

**New stage** (no existing `.response.json`):

```
Generate workshop for stage `../Wayfinder-content/source/wayfinder/<phase>/<stage>/stage.md`
```

**Existing stage** (iterate on top of existing content):

```
Update workshop for stage `../Wayfinder-content/source/wayfinder/<phase>/<stage>/stage.md`
```

## Workshop Philosophy

The goal is to get the customer working as fast as possible. Every slide exists to enable a task — not to educate.

**Content hierarchy:**

1. Read `stage.md` — the tasks listed there are the primary output goal
2. Use the cached `best-practices/` under `AI-automation/wayfinder-post-process/<phase>/stages-library/<stage>/` as the primary explanation material; use `public-docs/` in the same cache to fill any gaps or add depth where best practices are insufficient (run `scripts/sync/pull_best_practices.py` first if the cache is empty)
3. Consider both sources together — best practices take priority, public docs supplement

**Design principle:** present only what the customer needs to know _right before_ doing each task. No background, no theory, no fluff. Jump straight to work.

## Slide Structure

| Section         | Purpose                                             |
| --------------- | --------------------------------------------------- |
| **Stage goal**  | One sentence — what we're achieving today           |
| **Task blocks** | One block per task: minimal context → hands-on step |
| **Wrap-up**     | What was done, what comes next in the wayfinder     |

- Max 45 minutes total
- Fewer slides is better
- Each task block should be actionable in < 10 min

## Pipeline

### New stage (no existing `.response.json`)

```
stage.md + wayfinder-post-process/<stage>/best-practices/ → .prompt.md → AI → .response.json → build → .pptx → upload
```

| Step              | What happens                                                     | AI?     |
| ----------------- | ---------------------------------------------------------------- | ------- |
| `prepare`         | Reads stage, gathers content, writes `.prompt.md` into `.ai/`    | No      |
| Agent writes JSON | Reads `.prompt.md`, writes `.response.json` to `outputs/stages/` | **Yes** |
| `build`           | Assembles branded `.pptx` from `.response.json`                  | No      |
| `upload`          | Pushes `.pptx` to SharePoint                                     | No      |

### Existing stage (iterate — review only, no content changes)

```
SharePoint .pptx (living artifact) → extract text → AI review → inject comments → optional new slides → conflict-safe upload
```

| Step                | What happens                                              | AI?     |
| ------------------- | --------------------------------------------------------- | ------- |
| Fetch PPTX          | Download living PPTX from SharePoint; save eTag           | No      |
| Extract text        | Pull slide titles, bullets, speaker notes from the PPTX   | No      |
| Compare with docs   | Identify stale, missing, or style-guide-violating content | **Yes** |
| Inject comments     | Add Old/New/Why comments on flagged slides                | **Yes** |
| Append slides       | Add new slides if a topic is missing entirely             | **Yes** |
| `upload --if-match` | eTag conflict-safe upload to SharePoint                   | No      |

The `.response.json` is **not read and not rebuilt** during iterate — it is an archived initial-build record. The living document is the PPTX on SharePoint.

The human acts on comments directly in PowerPoint. If they want the agent to apply a batch of changes, they can ask — but that is optional.

**Overwrite protection:** If `.response.json` already exists, always stop and ask **Iterate** or **Abort** — never silently overwrite. Regenerating from scratch is not an option.

### JSON lifecycle

`.response.json` is an initial-build artifact. After the PPTX is uploaded to SharePoint it becomes an archived record — never the source of truth for future iterate runs.

Every generated `.response.json` must include a `_meta` block:

```json
{
  "_meta": {
    "status": "initial-build-only",
    "warning": "This file generated the initial PPTX and is now archived. Do not rebuild or iterate from this file. The living document is on SharePoint.",
    "sharepoint_url": "https://dynatrace.sharepoint.com/sites/Post-Sales/Shared%20Documents/Wayfinder/stages/<Stage Name>",
    "generated_at": "YYYY-MM-DD"
  }
}
```

The build tool ignores `_meta` — it reads only `title` and `sections`.

### Where to write `.response.json`

Derive the Stage Name from the stage folder (e.g. `001 - 001:PLF - Allocate Costs` → `Allocate Costs`) and write to:

```
outputs/stages/<Stage Name>/workshop/<slug>-workshop-<duration>.response.json
```

Example: `outputs/stages/Allocate Costs/workshop/allocate-costs-workshop-45min.response.json`

Once `.response.json` is written, **immediately run the build** to produce the PPTX — do not wait for the user to ask:

```bash
python scripts/content-creation/generate_training.py build
```

The build scans `outputs/stages/*/workshop/` and produces PPTX for every `.response.json` it finds. No stage argument is needed.

### Injecting review comments

After building (from the unchanged JSON), run `inject_pptx_comments.py` using the Old/New/Why format so the human can act on each comment directly in PowerPoint:

```bash
python scripts/content-creation/inject_pptx_comments.py \
  "outputs/stages/Segment Data/workshop/segment-data-workshop-45min.pptx" \
  --comment $'Slide Title:Old: "exact current text or value"\n\nNew: "exact replacement"\n\nWhy: one sentence reason'
```

Slides are matched by title text (case-insensitive substring if exact match fails). Comments appear under author "Wayfinder Engine" in the PowerPoint Review pane.
