# Step 4 — Build project plan PPTX

## Trigger

Prompts like:

> `Build positioning/business slides for outputs/projects/platform-foundations/`

## Inputs

1. **`outputs/<name>/project-plan-<name>.xlsx`** — the generated project
   plan. The script reads:
   - Project objective from the Overview & TOC sheet
   - Phase names and summaries from the Overview & TOC sheet
   - Project name from the sheet titles

2. **`dt-template/project-plan.pptx`** — the Dynatrace branded
   PowerPoint template. Contains 6 slides and 57 slide layouts. The
   build script uses slide 1 (cover) and slide 5 (objective + outcomes)
   as pre-built canvases — it keeps their shapes and replaces text.

3. **`outputs/<name>/mapping.mmd`** — the resolved wayfinder path and
   corresponding `stage.md` files under `../Wayfinder-content/source/wayfinder/`. The AI
   reads stage outcomes and phase structure to compose the three
   business outcomes.

---

## How to

- Use the slide number 5 of dt-template/project-plan.pptx as your strucure, as your template

- You need to replace the default objective body with the business objective derived from `input.md` (Goals section). Keep it short and punchy — one sentence, max 15 words. Use personal "your" language (e.g. "upgrade your team to Dynatrace 3rd Generation and scale the migration across your organization"). The word "Objective:" must stay bold and separate — it is the first run in the shape; only replace the body text in the subsequent runs. After replacing text, remove any `<a:br>` line-break elements and trailing empty `<a:r>` runs from the paragraph XML — the template has a stray line break that must be stripped. Do not add a trailing newline.

- Use the pathway output .mmd (e.g. outputs/projects/classic-to-3rd-gen-upgrade/mapping.mmd), to fill each of the titles. Instead of Ingest, Observe, Analyze as the template has, in this case, it should have the subgraphs, e.g. Prepare, Deploy, Setup. If parallel subgraphs happen, convey into 1, it has to be a streamline process, not a map

- Replace the bullets e.g. "Define scope and requirement, Design high-level architecture, Ingest and standardize" with the main tasks on each subgraph. **Maximum 3 bullets per phase box.** If the subgraph has more than 3 tasks, merge related ones into a single concise bullet.

- Replace the business outcome of that step with the brief **Stage Outcome** from the stage definition. Keep it short, customer-facing, and outcome-oriented — this is not the timeline or implementation detail.

- It should have 5 boxes or less. If you find 6 or more subgraphs, merge related phases until you have at most 5.
- Pass **only the phases present in the `.mmd`** — never pad with empty or invented phases. The build script automatically removes unused column boxes from the template and redistributes the remaining ones so they are horizontally centered on the slide. A 3-phase project shows 3 centered boxes; a 5-phase project shows all 5.

- The output PPTX must have exactly 4 slides in this order:
  1. **Cover slide** — add a new slide using layout index 0 ("Cover slide") from the template. Set the title placeholder ("Title 1") to the project name and the subtitle placeholder ("Subtitle 2") to "Dynatrace Delivery Plan". Remove slides 1–4 from the template before saving.
  2. **Positioning slide** — the customised slide 5 from the template (objective + phases + outcomes), as described above.
  3. **Effort table slide** — add a new slide using layout index 6 ("Standard slide"). Set the title to "Estimated Effort". Build a table using `add_table` with columns: **Stage**, **Description**, **Effort (days)**. Read stage rows (rows where `Tasks` is None and `Effort (days)` is not None) from the `Activities & Milestones` sheet of `outputs/<name>/project-plan-<name>.xlsx`. Use the brief stage outcome from the Description column. Add a **Total** row at the bottom summing all effort days. Below the table, add a text box with assumptions from `pathways-generation/default-pathways/<name>/input.md`: number of teams, number of applications, and change management maturity, formatted as: `Assumptions: X team · Y applications · Z change maturity`.
  4. **Delivery timeline slide** — add a new slide using layout index 6 ("Standard slide"). Set the title to "Delivery timeline". Pass the `--timeline` JSON (keys: `label`, `context`, `total_weeks`, `total`, `phases`) to render a Gantt-style week grid.

     **How to calculate the timeline from the customer profile:**
     The customer profile in `pathways-generation/default-pathways/<name>/input.md` drives where within each stage's effort range you land, and therefore how long each stage takes in calendar weeks.
     1. **Determine the complexity multiplier** from the three profile fields:
        - `# Hosts`: ≤100 → Low, 101–1000 → Medium, >1000 → High
        - `# Apps`: ≤10 → Low, 11–30 → Medium, >30 → High
        - `Customer Complexity`: read directly (Low / Medium / High)
          Score: count how many of the three are High. 0–1 → use the **low end** of each range; 2 → use the **midpoint**; 3 → use the **high end**.

     2. **Read each stage's `## Effort & Timeline` section** from `../Wayfinder-content/source/wayfinder/` and apply the multiplier to pick a point estimate for both effort (days) and timeline (weeks per stage).

     3. **Respect the dependency order from the `.mmd` file.** Edges in the mmd (e.g. `A --> B`) mean B cannot start until A finishes. Stages with no edge between them within the same phase can run in parallel — assign them the same `start` week.

     4. **Build the Gantt at stage level**, not phase level. Each row in the `phases` array represents one stage. Parallel stages occupy the same week columns and appear as overlapping bars. This makes parallelism visible to the customer.

     5. **Calculate each stage's `start` week** by walking the dependency graph from the mmd:
        - A stage starts the week after all its prerequisites finish.
        - Parallel stages (no dependency between them) share the same start week.

     6. **Set `total_weeks`** to the end week of the last stage to finish.

     7. **Set `context`** to a short string summarising the profile, e.g. `"SaaS · 50 apps · High complexity"`.
     - Build a table with rows = phases + 1 header row + 1 total row, and columns = 1 label column + one column per week (W1…Wn).
     - Header row: "Phase" cell (dark `#3D4B5A` fill, white text) + one cell per week showing "W1", "W2", … (same dark fill, white text, centered).
     - Phase rows: left cell = phase label; each week cell is filled DT cyan `#00B4E0` if the week falls within `[start, start + duration - 1]`, otherwise light gray `#F5F5F5`. Overlap between phases is shown naturally — two phase bars can share the same week column.
     - Total row: left cell = `"Total: <total>"` (bold, `#EBEBEB` fill); all week cells also `#EBEBEB`.
     - Below the table, add an italic text box: `"<label>  ·  <context>"`.
     - Below that, add an italic 9pt footer: `"Note: Timeline includes scheduling gaps, approval gates, and stakeholder availability — not just active work. Shaded weeks show when each phase is active; overlapping phases reduce total calendar time below the sum of individual phase durations."`
     - Pass the timeline via `--timeline '<json>'` on the CLI. If omitted, skip this slide.
