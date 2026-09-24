# Generate Pathway

## Workflow

```
Input file → Maps goals to stages → Script traces full path
```

## Step 1 — Map Goals with Wayfinder Stages

Reads the customer input file and maps each goal to the most relevant wayfinder stages with relevance percentages.

| Tool        | Command                                                                                                                                                                             |
| ----------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Claude Code | `@instructions/shared/01-safeguards.md @instructions/pathway/01-pathway.md Generate pathway for pathways-generation/default-pathways/classic-to-3rd-gen-upgrade/input.md`           |
| Copilot     | `#file:instructions/shared/01-safeguards.md #file:instructions/pathway/01-pathway.md Generate pathway for pathways-generation/default-pathways/classic-to-3rd-gen-upgrade/input.md` |

### Output

This produces `outputs/projects/<name>/mapping.md` containing a single table:

```markdown
| Stage           | Relevance | Why |
| --------------- | :-------: | --- |
| Assess & Design |  **95%**  | ... |
```

The agent reads the customer's goals and scope, cross-references them against each stage's `Stage Outcome` section, and produces one row per relevant stage. This table is the input for Step 2.

### Review & adjust mapping (optional)

Before running the script, open `outputs/projects/<name>/mapping.md` and edit it manually if needed:

- **Add a stage** — insert a new row with the exact `nodeLabel` from `../Wayfinder-content/source/mermaid/wayfinder-v3-mapping.json`
- **Remove a stage** — delete its row

The script in Step 2 reads only the Stage column, so relevance and why values don't matter for path resolution.

## Step 2 — Generate Wayfinder

Once you know which stages are relevant, a Python script traces all downstream dependencies in the mermaid graph and walks back to the very beginning to build the complete path:

```bash
cd AI-automation && python3 scripts/build_wayfinder.py outputs/projects/.../mapping.md
```

This produces `outputs/projects/<name>/mapping.mmd` — a standalone mermaid diagram containing only the relevant stages with all prerequisite paths resolved. No AI involved, purely graph traversal.

The script uses as well the input.md file for the decision points

### Output

Each project plan produces two files in `outputs/projects/<name>/`:

| File | Content |
|------|---------||
| `mapping.md` | Goal-to-stage mapping with relevance percentages |
| `mapping.mmd` | Standalone mermaid file (resolved wayfinder path) |

## Step 3 - Validate

Copy the `.mmd` into a mermaid reader (e.g. [mermaid.live](https://mermaid.live)) to visually validate the customer's path:

![Validate wayfinder](../assets/validate-wayfinder.png)
