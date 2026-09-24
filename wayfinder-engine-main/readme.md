# Wayfinder Engine

The repo is split into two top-level workspaces:

- **`Wayfinder-content/`** — owned and managed entirely by the **D1 CoE Team**. Contains all curated assets and resources under CoE ownership: the master Mermaid graph and a structured hierarchy of phase folders, each with a `stages-library/` holding one dedicated folder per stage, where each stage folder contains its `stage.md`. Manually authored by the CoE Team members.
- **`AI-automation/`** — automation workspace: scripts, docs,
  instructions, templates, inputs, outputs, plus a gitignored
  `wayfinder-post-process/` cache that holds best-practices and public-docs
  fetched on demand from Confluence and docs.dynatrace.com.

## Get Started

1. Clone this repo:

   ```bash
   git clone https://github.com/dynatrace-ace/wayfinder-engine.git
   ```

   <details>
   <summary><strong>Repo structure</strong></summary>

   ```
   wayfinder-engine/
   ├── AI-automation/               ← automation workspace (venv, scripts, docs, instructions)
   │   ├── docs/                    ← end-to-end workflow guides (run from AI-automation/)
   │   ├── pathways-generation/     ← per-project pathway inputs (default + custom examples)
   │   ├── outputs/                 ← generated project and stage deliverables
   │   │   ├── projects/
   │   │   └── stages/
   │   ├── scripts/                 ← generation, sync, and utility scripts
   │   │   ├── content-creation/
   │   │   ├── miro-to-mermaid/
   │   │   └── sync/
   │   ├── instructions/            ← reusable prompt modules per workflow
   │   │   ├── shared/
   │   │   ├── slides-general/
   │   │   ├── slides-workshop/
   │   │   └── …
   │   ├── wayfinder-post-process/  ← fetched best-practices + public-docs cache (gitignored)
   │   ├── dt-template/             ← Dynatrace-branded PowerPoint templates
   │   ├── EPM-project-plan-template/  ← Excel project-plan template
   │   ├── requirements.txt
   │   └── status.txt
   ├── Wayfinder-content/           ← curated content source-of-truth (mermaid + stage.md only)
   │   └── source/
   │       ├── mermaid/             ← master wayfinder-v3 graph + mapping JSON
   │       └── wayfinder/           ← phase folders with stage.md (no fetched content)
   └── docs/                        ← top-level guides
   ```

   > Commands in the guides assume you run them from `AI-automation/` (where the
   > `.venv`, `.env`, and `requirements.txt` live). Paths into content are then
   > `../Wayfinder-content/source/...`.

   </details>

&nbsp;

2. **Set up the automation workspace** (one-time, from `AI-automation/`):

   ```bash
   cd AI-automation
   python3 -m venv .venv
   source .venv/bin/activate
   pip install -r requirements.txt
   ```

   Then create `AI-automation/.env` with the credentials the scripts need
   (only the ones you'll use):

   ```
   # Confluence sync (pull_best_practices.py, push_to_confluence.py)
   CONFLUENCE_BASE_URL=https://dt-rnd.atlassian.net
   CONFLUENCE_USERNAME=you@dynatrace.com
   CONFLUENCE_API_TOKEN=<token>
   CONFLUENCE_ROOT_PAGE_ID=<id>

   # SharePoint sync (upload_to_sharepoint.py, fetch_pptx_for_review.py, …)
   TENANT_ID=<tenant-id>
   CLIENT_ID=<client-id>
   CLIENT_SECRET=<client-secret>   # omit for device-code login

   # Dynatrace dashboard sync (dt_dashboard.py)
   DT_ENVIRONMENT=https://<tenant>.apps.dynatrace.com
   DT_CLIENT_ID=<id>
   DT_CLIENT_SECRET=<secret>
   DT_URN=<urn>
   ```

   All subsequent commands assume `AI-automation/` as the working directory.

&nbsp;

3. **Pick what you want to do** and follow the corresponding guide:

| #   | Guide                                                          | What it does                                                                                                                                                                                                           |
| --- | -------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| 1   | [Add a Stage](docs/add-stage.md)                               | Add a new stage (and phase if needed) to the Wayfinder Engine framework: from deciding where it belongs, through diagram, SVG, MMD, and mapping, to a merged and validated `stage.md`. See [Diagram Structure](docs/wayfinder-structure.md) for the node types (phases, grouping containers, stage boxes, decision diamonds) and the rules for where a stage belongs. |
| 2   | [Pathway](AI-automation/docs/pathway.md)                       | Generate a customer pathway: map goals to relevant wayfinder stages.                                                                                                                                                   |
| 3   | [Project Plan](AI-automation/docs/project-plan.md)             | Generate a full customer delivery package: Excel project plan with effort estimates + a 3-slide PPTX (cover, positioning, effort table).                                                                               |
| 4   | [Pull Best Practices](AI-automation/docs/confluence-sync.md)   | Fetch the latest best-practice + public-doc content for a stage from Confluence and docs.dynatrace.com into `AI-automation/wayfinder-post-process/` (gitignored cache). Run this before generating workshop materials. |
| 5   | [Workshop Materials](AI-automation/docs/workshop-materials.md) | Generate or iterate on hands-on workshop decks for each stage. Reads the stage's cached best-practices; run step 4 first if the cache for that stage is empty.                                                         |
| 6   | [Push to SharePoint](AI-automation/docs/push-to-sharepoint.md) | Publish generated assets to SharePoint for distribution and stakeholder access.                                                                                                                                        |

&nbsp;

## Scripts and where they write

`Wayfinder-content/` is the curated source-of-truth. Most scripts only read
from it; a handful modify it. This table lists every script that touches the
filesystem, so you always know what a command will change before running it.

| Script                                                                                                                                                                                                                                                                                                                             |                   Reads `Wayfinder-content/`                    | Writes `Wayfinder-content/`                                                                                       | Writes `AI-automation/`                                                                   |
| ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | :-------------------------------------------------------------: | ----------------------------------------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------- |
| [scripts/sync/pull_best_practices.py](AI-automation/scripts/sync/pull_best_practices.py)                                                                                                                                                                                                                                           |                  `source/wayfinder/*/stage.md`                  | —                                                                                                                 | `wayfinder-post-process/<phase>/<stage>/{best-practices,public-docs}/` (gitignored cache) |
| [scripts/sync/push_to_confluence.py](AI-automation/scripts/sync/push_to_confluence.py)                                                                                                                                                                                                                                             |          `source/wayfinder/*/stage.md`, `wayfinder.md`          | **`stage.md` / `wayfinder.md`** — rewrites `lastSynced` frontmatter after a successful push                       | —                                                                                         |
| [scripts/map_stages.py](AI-automation/scripts/map_stages.py)                                                                                                                                                                                                                                                                       |              `source/wayfinder/*/stage.md` (scan)               | **`source/mermaid/wayfinder-v3.mmd`** (in-place ID normalisation), **`source/mermaid/wayfinder-v3-mapping.json`** | —                                                                                         |
| [scripts/sync/fetch_sharepoint_stage_urls.py](AI-automation/scripts/sync/fetch_sharepoint_stage_urls.py)                                                                                                                                                                                                                           |                                —                                | **`source/mermaid/wayfinder-v3-mapping.json`** (adds `sharepointUrl` fields)                                      | —                                                                                         |
| [scripts/sync/upload_to_sharepoint.py](AI-automation/scripts/sync/upload_to_sharepoint.py) `--update-mapping`                                                                                                                                                                                                                      |                                —                                | **`source/mermaid/wayfinder-v3-mapping.json`** (writes the folder URL for one stage)                              | —                                                                                         |
| [scripts/miro-to-mermaid/cli.py](AI-automation/scripts/miro-to-mermaid/cli.py)                                                                                                                                                                                                                                                     |                                —                                | Only the output path you pass on the CLI (typically `source/mermaid/…`)                                           | —                                                                                         |
| [scripts/build_wayfinder.py](AI-automation/scripts/build_wayfinder.py)                                                                                                                                                                                                                                                             |                       `source/mermaid/…`                        | —                                                                                                                 | `outputs/projects/<name>/mapping.mmd`                                                     |
| [scripts/build_project_plan.py](AI-automation/scripts/build_project_plan.py), [scripts/build_positioning_slides.py](AI-automation/scripts/build_positioning_slides.py)                                                                                                                                                             |                  `source/wayfinder/*/stage.md`                  | —                                                                                                                 | `outputs/projects/<name>/*`                                                               |
| [scripts/content-creation/generate_training.py](AI-automation/scripts/content-creation/generate_training.py) (`prepare`, `build`)                                                                                                                                                                                                  | `source/wayfinder/*/stage.md`, `wayfinder-post-process/*` cache | —                                                                                                                 | `outputs/stages/<name>/workshop/*`, `.ai/prompts/`, `.ai/responses/`                      |
| [scripts/sync/upload_to_sharepoint.py](AI-automation/scripts/sync/upload_to_sharepoint.py) (no `--update-mapping`), [scripts/sync/fetch_pptx_for_review.py](AI-automation/scripts/sync/fetch_pptx_for_review.py), [scripts/sync/list_sharepoint_workshop_status.py](AI-automation/scripts/sync/list_sharepoint_workshop_status.py) |                                —                                | —                                                                                                                 | `outputs/stages/*` (downloads only), `.msal_token_cache.json`                             |
| [scripts/dt_dashboard.py](AI-automation/scripts/dt_dashboard.py)                                                                                                                                                                                                                                                                   |                                —                                | —                                                                                                                 | Only the file paths you pass on the CLI                                                   |

> **`push_to_confluence.py` frontmatter bump:** on every successful page push, the
> script rewrites the local `stage.md` (or `wayfinder.md`) to update the
> `lastSynced: "…"` field. The end-of-run log tells you to `git add -u && git commit`
> those timestamps.
>
> **`Wayfinder-content/source/mermaid/`** is tracked but is modified by the three
> mapping-related scripts above. Review the diff and commit intentionally.
>
> **`Wayfinder-content/source/wayfinder/<phase>/…/{best-practices,public-docs}/`** is
> _never_ written to by any script after the refactor — fetched content only ever
> lands in the gitignored cache.
