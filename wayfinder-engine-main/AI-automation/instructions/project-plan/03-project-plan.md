# Step 3 — Build Project Plan

## Inputs

You need three things:

1. **`outputs/<name>/mapping.mmd`** — the resolved wayfinder path (from Step 2).
   This tells you the exact stages and their dependency order.

2. **Stage definitions** — for each stage in the `.mmd`, read the corresponding
   `stage.md` from `../Wayfinder-content/source/wayfinder/<phase>/stages-library/`. These
   contain tasks, effort ranges, prerequisites, and risks.

3. **`pathways-generation/default-pathways/<name>/input.md`** — the customer profile and scope decisions.
   Use complexity drivers (teams, apps, technology, change management maturity)
   to position effort estimates within the stage's low–high range.

---

## Steps

1. **Read the `.mmd` file** and extract the ordered list of stages.
   The dependency order is encoded in the edges. Topologically sort the
   stages so prerequisites come first.

2. **Read each stage's `stage.md`** to get tasks, effort, and prerequisites.

3. **Read the customer input** to evaluate complexity drivers and position
   effort estimates (use `02-effort-estimation.md` logic).

4. **Read the Excel template** from `EPM-project-plan-template/TEMPLATE.xlsx`.
   This is a multi-sheet workbook with 17 tabs. The key sheet for this step is
   **"Activities & Milestones"** — its header row is already set (columns:
   WBS, Stage, Tasks, Description, Resources, Effort (days), Owner, Start Date,
   End Date, Notes, Status). The Resources column contains hyperlinks to
   the stage's SharePoint folder (from `wayfinder-v3-mapping.json`).
   The data rows are empty and must be populated per customer.

5. **Build the project plan** by running:

   ```bash
   python3 scripts/build_project_plan.py outputs/<name>/mapping.mmd -i pathways-generation/default-pathways/<name>/input.md --objective "<AI-generated objective>" --phases '<JSON phases>'
   ```

   This script:
   - Copies the generic `TEMPLATE.xlsx` to `outputs/<name>/project-plan-<name>.xlsx`
   - Reads the `.mmd` to extract the ordered stage list and subgraph phases
   - Reads each stage's `stage.md` for tasks, KPIs, risks, and stakeholders
   - Reads the customer `input.md` to position effort estimates
   - Populates six sheets dynamically:
     - **Activities & Milestones** — WBS-numbered tasks with stage-level effort
       and a Resources column linking each stage to its SharePoint folder
     - **Overview & TOC** — "Delivery Approach at a Glance" phase table
       derived from the mmd subgraph labels (e.g. Prepare, Setup, Upgrade)
       with AI-generated business summaries for each phase
     - **Assumptions & Prereqs** — customer profile fields (from Customer
       Profile and Scope sections) as assumptions with ID, Impact, Owner
     - **Project KPIs** — acceptance criteria extracted from each stage's
       `Acceptance / KPIs` section
     - **RACI** — responsibility matrix built from project stages with role
       assignments inferred from each stage's `Participating Stakeholders`
       section, plus three standard project management rows
     - **Risks** — risk register populated from each stage's
       `Risk consideration` section with sequential R-001 numbering
   - Replaces `[Project Name]` across all sheets
   - Inserts the AI-generated project objective via `--objective`
   - Inserts the AI-generated phase summaries via `--phases`
   - Saves the result

6. **Generate the project objective** (non-deterministic, AI step).
   Before running the script, read the customer's goals from `input.md` and
   the expected outcomes of every stage in the resolved path. Then compose
   a concise project objective (2–3 sentences) that:
   - States what the engagement aims to achieve (from goals)
   - References the key delivery outcomes across phases
   - Is written in active voice, present tense

   Pass this text via the `--objective` flag. If omitted, the script falls
   back to a basic sentence derived from the Goals section bullets.

7. **Generate the phase summaries** (non-deterministic, AI step).
   For each phase (subgraph) in the `.mmd`, read the Stage Outcome sections
   of all stages within that phase. Then compose a compact business summary
   (one to two sentences) that:
   - Explains what the phase delivers in business terms
   - Avoids listing individual stage names or raw outcome bullets
   - Is written for a project stakeholder, not a technical audience

   Pass these as a JSON string via the `--phases` flag:

   ```bash
   --phases '{"Prepare": "Define the observability strategy...", "Setup": "Configure data access..."}'
   ```

   Keys must match the subgraph labels in the `.mmd` exactly.
   If omitted, the script falls back to concatenated Stage Outcome text.

---

## Output

One file in `outputs/<name>/`:

| File                       | Content                                            |
| -------------------------- | -------------------------------------------------- |
| `project-plan-<name>.xlsx` | Formatted Excel workbook with all sheets populated |

---

## Template Sheets Reference

The template (`TEMPLATE.xlsx`) contains 17 sheets. The `build_project_plan.py`
script populates six sheets automatically. The remaining sheets provide the
project governance structure and are filled manually during delivery:

| Sheet                       | Purpose                                   |                   Auto-populated?                   |
| --------------------------- | ----------------------------------------- | :-------------------------------------------------: |
| Overview & TOC              | Project intro & navigation                | **Yes** — phases + AI summaries + project objective |
| Executive Summary           | Context, objectives, success criteria     |                         No                          |
| Migration Scope             | Environment, sizing metrics               |                         No                          |
| Assumptions & Prereqs       | Contractual & delivery guardrails         |      **Yes** — customer profile as assumptions      |
| Dependencies                | Internal & customer dependencies          |                         No                          |
| Technical Infrastructure    | High-level infra overview                 |                         No                          |
| Contacts                    | Stakeholders and teams                    |                         No                          |
| RACI                        | Responsibility matrix                     |            **Yes** — stages + RACI roles            |
| Risks                       | Risk assessment & management              |         **Yes** — from stage risk sections          |
| **Activities & Milestones** | **WBS with tasks, owners, dates, effort** |                       **Yes**                       |
| Issues                      | Issue log                                 |                         No                          |
| Decisions                   | Architecture & delivery decisions         |                         No                          |
| Quality Assurance           | Quality gates & acceptance                |                         No                          |
| Pre-Sales Checklists        | Pre-sales readiness items                 |                         No                          |
| Post-Sales Checklists       | Post-sales readiness items                |                         No                          |
| Appendices & Change Log     | Versioning & notes                        |                         No                          |
| DATA                        | Dropdown reference data                   |                         No                          |

Additionally, the **Project KPIs** sheet is populated with acceptance criteria
from each stage's `Acceptance / KPIs` section.

---

## Rules

- **Use the `.mmd` as the source of truth** for which stages to include and
  their order. Do not add stages that are not in the resolved path.
- **Every task from the stage.md must appear** — do not summarize or merge tasks.
- **Effort estimates go on the stage row, not on task rows.**
  The WBS row like `1.1` or `2.3` carries the total positioned effort for the
  stage in the Effort column. Individual task rows (`1.1.1`, `1.1.2`, etc.)
  have no effort value — they exist for tracking, not estimation.
- **Effort estimates must reflect the customer profile.** A 1-team, 3-app PoC
  should trend toward the low end. A 10-team enterprise should trend high.
- **Do NOT invent tasks** beyond what the stage.md defines.
- **Task text is split into Tasks and Description columns.**
  The Tasks column holds a short action-oriented name (e.g., "Define
  observability requirements"). The Description column holds the detail
  that follows the em-dash (e.g., "covering data access (IAM), partitioning
  (retention), segmentation (filters), and cost allocation (chargeback)").
- **WBS numbering** follows: `{phase}.{stage}.{task}` (e.g., `2.1.3`).
  Phases map to the wayfinder Steps (Understanding, Ingest, Observe, Analyze, Act).
  Stages map to wayfinder stages within each phase.
