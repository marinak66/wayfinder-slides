## Safeguards

Before writing any document file, you **must** check whether one already exists at the target path.

### If document already exists

Stop and ask the user which action to take:

1. **Iterate**
2. **Abort**

**Never silently overwrite an existing document.** Always wait for the user's explicit choice before proceeding.

**Regenerating from scratch is not an option.** Existing work is always preserved as the baseline.

### If no document exists

Proceed normally — generate fresh content without asking.

### Style/grammar mode (asked immediately after the user chooses Iterate)

Once the user confirms Iterate, ask a second question before doing any work:

> **How should style and grammar issues be handled?**
>
> 1. **Auto-fix (Recommended)** — Style and grammar violations (punctuation, verb tense, UI verbs, sentence case, number ranges, acronym expansion) are fixed directly in the slide text. No comments are added for these. A summary report of every change is delivered at the end so the result is fully auditable — and any fix can be reverted if needed.
> 2. **Dry-run** — Style and grammar issues are injected as PowerPoint comments tagged `[auto]` alongside content comments. Nothing is changed automatically; you apply or reject each one yourself in PowerPoint.

**Default is auto-fix.** If the user does not express a preference, use auto-fix.

### Iterate workflow

Style/grammar fixes are applied **before** content comments are injected, so content comment quotes always reference the already-corrected text.

#### Steps

1. **Download the living PPTX from SharePoint** using `fetch_pptx_for_review.py` — this is the authoritative source, not the archived `.response.json`
2. Extract slide text from the downloaded PPTX (titles, bullets, speaker notes)
3. Read the current `best-practices/` and `public-docs/` for the stage from `AI-automation/wayfinder-post-process/<phase>/stages-library/<stage>/` (run `scripts/sync/pull_best_practices.py` first if the cache is empty)
4. Classify every issue found as either **style/grammar** or **content**:
   - **Style/grammar**: punctuation, verb tense, UI interaction verbs ("log in" → "sign in"), sentence case, number ranges ("1 to 4" → "from one through four"), acronym expansion on first use, negative contractions in warning contexts ("aren't" → "are not"), em dashes in slide bullets
   - **Content**: missing topics, factual accuracy, coverage gaps, stale information, missing task coverage
5. **Apply style/grammar fixes** directly to the slide XML (auto-fix mode) — or inject them as `[auto]`-tagged comments (dry-run mode). Log every change.
6. For each **content issue**, locate the exact slide title and text segment, then inject one Old/New/Why comment using `inject_pptx_comments.py`
7. If new slides are clearly needed (topic not covered at all), append them with `append_pptx_slides.py --why "<reason>"` — otherwise prefer comments over additions. **Always pass `--why`.** It automatically injects a "Slide added" review comment onto the section divider and every newly appended slide.
8. Conflict-safe upload to SharePoint (eTag validation via `upload_to_sharepoint.py --if-match`)
9. **Deliver a summary report** to the user:
   - **Auto-fix mode**: list every style/grammar fix applied (slide, old text → new text), then list content comments added and whether new slides were appended
   - **Dry-run mode**: list all comments added, distinguishing `[auto]` style ones from content ones, and whether new slides were appended

#### Comment format

Every comment must be self-contained so the human can act on it directly without returning to the agent:

```
Old: <exact current text that is wrong>

New: <exact replacement>

Why: <one sentence reason>
```

```bash
python scripts/content-creation/inject_pptx_comments.py "outputs/stages/<Stage>/workshop/<slug>.pptx" \
  --comment $'Slide Title:Old: "exact current text"\n\nNew: "replacement text"\n\nWhy: one sentence reason'
```

The human reads each comment, makes the change directly in PowerPoint if they agree, and deletes the comment when done. If they prefer the agent to apply a batch of changes, they can ask — but that is optional, not a required step.

**Appended slides use a different two-part format** (there's no "Old" text to quote, since nothing existed before):

```
Slide added.

Why: <one sentence reason>
```

This is injected automatically by `append_pptx_slides.py --why "<reason>"` — do not call `inject_pptx_comments.py` separately for this. Passing `--why` is required whenever slides are appended during an iterate run.

---

### JSON lifecycle

`.response.json` files are **initial-build artifacts only**. Their lifecycle is:

1. **Generate** — AI writes `.response.json`, build produces `.pptx`, upload to SharePoint
2. **Archive** — once uploaded, the `.response.json` is stale; SharePoint PPTX is the living document
3. **Iterate** — all future runs download from SharePoint; the `.response.json` is never read or rebuilt again

Every `.response.json` includes a `_meta` block that marks its archived status and links to the living SharePoint document:

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

The `_meta` block is ignored by the build tool — it only reads `title` and `sections`.
