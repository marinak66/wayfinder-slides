## Session Definitions

Each definition file contains one or more `# Session` blocks. This section describes every value the agent may receive so it knows exactly what to produce.

---

### Format

The `Format:` field determines the **type of deliverable** and its expected structure.

#### `presentation`

A slide deck for a speaker-led session. The audience listens; interaction is minimal.

- Open with context and business value — why this matters.
- Use screenshot-rich slides to show real product UI.
- Include a customer story or real-world scenario when `customer-story` is listed in Content.
- Speaker notes should be detailed enough for any presenter to deliver without prior preparation.
- Typical use: positioning decks, overview introductions, executive briefings.

#### `training`

A slide deck for an instructor-led teaching session. The audience learns and is tested.

- Start with an agenda slide listing the learning objectives.
- Go deep on technical content — architecture, configuration steps, DQL queries, API calls.
- Include demo steps showing exactly what the instructor will do live in the Dynatrace UI.
- End with quiz questions (4 options A–D, correct answer in speaker notes) when `quiz` is in Content.
- Close with a takeaways slide and a resources slide with links.
- Typical use: enablement sessions, certification preparation, technical deep dives.

#### `workshop`

A slide deck for a hands-on, interactive session. The audience does the work.

- Start with an agenda and objectives, then alternate between instruction slides and hands-on exercises.
- Include `todo` exercise slides that tell attendees exactly what to do, step by step.
- Include demo slides that the facilitator performs live before each exercise.
- End with next-steps slides listing what attendees should do after the workshop.
- Speaker notes should include timing hints (e.g., "allow 10 min for this exercise").
- **Do NOT include quiz or knowledge-check slides.** Workshops are hands-on - the exercises themselves validate understanding.
- Typical use: hands-on labs, guided configuration workshops, onboarding bootcamps.

---

### Content Types

The `Content:` list tells the agent **what to include** in the session. Multiple values can be combined.

#### Goals (depth and angle)

| Value | What it means for the agent |
|-------|----------------------------|
| `technical` | Deep dive into the *how*. Include architecture diagrams, configuration steps, DQL queries, API payloads, and hands-on detail. Every claim must be backed by a concrete Dynatrace feature or workflow. |
| `positioning` | Focus on the *why*. Cover business value, outcomes, ROI, and customer impact. Keep technical detail light — just enough to support the value proposition. |
| `business-value` | Similar to `positioning` but narrower. Focus specifically on measurable outcomes: time saved, incidents reduced, cost optimized. Use numbers and scenarios. |
| `strategy` | High-level approach and decision framework. Help the audience understand *when* and *why* to adopt this capability, not the step-by-step *how*. |
| `overview` | Broad introduction to the topic. Cover what it is, why it matters, and what the session or journey will cover. Keep it concise — this is the opening context, not the deep dive. |
| `logistics` | Practical session information: agenda, schedule, prerequisites, environment access, tools needed. Used in opening slides for workshops and trainings. |

#### Elements (extras to include)

| Value | What the agent must produce |
|-------|----------------------------|
| `demo` | One or more demo slides. Each slide lists the exact steps the presenter will perform live in the Dynatrace UI. Use numbered steps. Speaker notes should include what to click, what to look for, and common pitfalls. |
| `quiz` | Knowledge-check slides. Each quiz slide has a question title, 4 bullet options (A–D), and the correct answer identified in the speaker notes. Use `dark_standard` layout. Place quizzes after the section they test. |
| `todo` | Hands-on exercise slides. Each tells the attendee exactly what to do — step-by-step instructions they follow in their own environment. Include expected outcomes so they can verify success. |
| `next-steps` | An actionable follow-up slide near the end. List 3–5 concrete things the audience should do after the session (e.g., "Enable metric ingestion for your first application", "Review the IAM policy for your environment"). |
| `customer-story` | Weave a real-world scenario into the narrative. Describe a customer challenge, the Dynatrace solution, and the outcome. Do not use real customer names — use generic personas ("a financial services company", "an e-commerce platform"). |
| `screenshot` | Extra emphasis on visual slides using images from the CoE screenshot catalog. Prefer `text_with_image` layout. Every major UI concept should have an accompanying screenshot if one is available in the catalog. |

---

### Duration

The `Duration:` field is a **maximum**, not a target. It is better to deliver focused, high-quality content in fewer slides than to pad the deck to fill the time. If the topic naturally fits in 30 minutes, do not stretch it to 45 just because the definition says `45min`.

Use ~2 minutes per content slide as a guideline.

| Duration | Approximate slide count | Depth |
|----------|------------------------|-------|
| `20min` | 8–10 slides | High-level overview, key messages only |
| `30min` | 12–15 slides | Overview with some detail |
| `45min` | 18–22 slides | Solid technical depth, demos, quizzes |
| `90min` | 35–45 slides | Comprehensive, multiple sections with exercises |
| `half-day` | 50–70 slides | Full training with breaks, multiple exercise blocks |
| `full-day` | 90–120 slides | Complete workshop with extensive hands-on |
