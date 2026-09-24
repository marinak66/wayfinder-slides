## Content Rules

### Focus Areas — your knowledge scope

The `# Focus Areas` section in each definition file is the **sole source of content scope** for the agent. Every focus area entry is either:

- **A CoE best-practices path** (e.g., `AI-automation/wayfinder-post-process/001_PREPARE/stages-library/001:PLF - Assess & Design/best-practices/index.md`) — the orchestrator reads matching `.md` files from the `wayfinder-post-process/` cache (populated by `scripts/sync/pull_best_practices.py`) and injects their titles and paths into the prompt as CoE knowledge. Images found in those directories are also included in the image catalog.
- **A Dynatrace public documentation URL** (e.g., `https://docs.dynatrace.com/docs/manage/identity-access-management`) — the agent should use this as a reference source. The content at that URL is publicly available and can be cited, summarized, or used to inform slide content.

**If the focus areas are empty, the agent receives no scoped knowledge and must rely only on its general Dynatrace knowledge.** Always prioritize information from the provided focus areas over general knowledge.

### Voice and perspective

- The audience is **the customer**. All slide-visible content (titles, bullets) must use **2nd person** — write "your environment", "your organization", "your technology landscape".
- Do NOT use 3rd person to refer to the audience on slides ("the customer's environment", "conversations with the customer"). Rephrase to address them directly.
- Speaker notes guide the presenter and may occasionally reference "the customer" when giving presenter-specific instructions, but keep slide content consistently in 2nd person.

### General rules

- Use ONLY publicly available Dynatrace documentation and the CoE knowledge provided.
- NEVER include product roadmap, release timelines, or future features.
- NEVER include customer-specific data, account names, or deal details.
- Focus on current GA (Generally Available) features only.
- NEVER include pricing, commercial terms, or internal ticket references.
- Write in a professional, engaging training style.
- Include practical examples and best practices from the CoE knowledge base.
- NEVER use em dashes (—) in any content (bullets, titles, speaker notes). Use regular hyphens (-) or rewrite the sentence instead.
- **Always check the "Additional Notes for AI" section** at the bottom of the prompt before generating content. If it contains specific instructions, feedback, or corrections, follow them precisely. These notes take priority over general guidelines when they conflict.
- **Avoid content repetition.** Do not restate the same concept across multiple slides. Each slide must add new information. If two slides cover similar ground, merge them or cut one. The audience should never feel they are seeing the same message twice.
