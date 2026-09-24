# Effort & Timeline Estimation Logic

## How to estimate effort for a stage

1. **Read the input file** (e.g., `pathways-generation/default-pathways/classic-to-3rd-gen-upgrade/input.md`).
   It tells you the customer's conditions: number of teams, number of applications, and change management maturity.

2. **Read the stage file** (e.g., `../Wayfinder-content/source/wayfinder/001_PREPARE/stages-library/001:PLF - Assess & Design/stage.md`).
   The `## Effort & Timeline` section contains the range and a rationale explaining what drives variation within it.

3. **Apply the customer's conditions to the rationale** and pick a specific estimate within the range.
   The rationale tells you what pushes effort toward the low or high end — map the customer's input directly to it.

4. **Output** effort (days) and timeline (weeks) for that stage, with one sentence explaining why you landed where you did.

## General principles

- **Effort ≠ Timeline.** Effort is active working days. Timeline includes scheduling gaps, approval gates, and stakeholder availability.
  - Timeline should consider amount of hosts, apps, and customer complexity and give an estimation on number of weeks needed to implement the solution as a whole
- **Always give a point estimate with a brief reason**, not just the range. The range is in the stage file; your job is to narrow it.
- **Effort scales with the customer's conditions**, not with the number of stages.
- When stages run in parallel, the timeline does not simply add up — account for overlaps.

## Calculating total project effort

1. Estimate each required stage individually using the steps above.
2. Identify parallelism opportunities (stages that can overlap once prerequisites are met).
3. Sum effort across all stages (working days).
4. Calculate total timeline (calendar weeks), accounting for parallelism and scheduling gaps.
5. Present as: `Total effort: X days | Timeline: Y weeks` with a one-line rationale.
