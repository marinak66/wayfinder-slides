## Slide Layouts

Choose the best layout for each slide:

| `layout` value | When to use |
|---|---|
| `standard` | Bullet-point or numbered-step slides (default). Keep to 3-5 items max. Use numbered steps (`"1. ..."`) for sequential content; the build tool auto-converts them to native PowerPoint numbering. |
| `text_with_image` | Bullets or numbered steps on the left + a CoE screenshot on the right. Set `image` to a path from the Available Images list. Use plain bullets for concept explanations. For step-by-step paths or sequences (e.g., creating a boundary, configuring a group), use numbered steps instead: `"1. Go to Policy Management"`, `"2. Go to the Boundaries tab"`, etc. The build tool auto-detects numbered items and applies native PowerPoint numbering (1. 2. 3.) instead of bullet characters. |
| `two_column` | Side-by-side comparison. Use `columns` array with 2 objects: `{header, body:[]}`. |
| `three_column` | Three items side by side. Use `columns` array with 3 objects. |
| `point_series` | **Multi-step content** where each step carries enough substance for its own slide. **Each step generates its own slide** with the active step highlighted. Use `steps` array (up to 4 objects, each with `label`, `bullets`, `speaker_notes`). You can add screenshot within the steps if needed, align it with text |
| `dark_standard` | Dark-themed bullet slide. Auto-applied for quiz, but you can also pick it for emphasis or key takeaways. |

> **IMPORTANT:** Do NOT use `full_image`. It has been removed. If you want to show a screenshot, always use `text_with_image` so the slide has context (bullets) alongside the image.

### text_with_image - text plus screenshot

Use `text_with_image` when a concept benefits from a visual. The left side holds bullets or numbered steps and the right side shows a screenshot.

**Image placement rules:**
- The build tool uses **contain mode** - the entire image is displayed at correct aspect ratio with no cropping. Never assume images will be cropped or clipped.
- **Prefer landscape-oriented screenshots** (wider than tall). The image placeholder is landscape-shaped (~7in wide x 6.5in tall), so wide screenshots fill the space best with minimal blank area.
- **Avoid very wide/thin banner screenshots** (e.g., navigation bars, single-row tables). These scale down to a tiny strip. If the screenshot is extremely wide and short, pair it with a standard layout bullet instead, or pick a different screenshot that shows more vertical context.
- **Avoid very tall/narrow screenshots** (e.g., long scrolling pages cropped to a narrow column). These leave large blank areas on both sides. Prefer screenshots that capture the full UI width.
- The ideal screenshot aspect ratio is roughly 4:3 or 16:9 - a typical browser window or Dynatrace UI view.

**Bullets vs numbered steps (applies to all layouts that use the `bullets` array):**
- Use **plain bullets** when the items are independent concepts, features, or tips (no inherent order)
- Use **numbered steps** (`"1. Go to..."`, `"2. Select..."`) when there's a sequential path or procedure
- The build tool auto-detects items starting with a digit+period and converts them to **native PowerPoint numbering** (`1.  2.  3.`) - the number prefix is stripped from the text and the paragraph gets a proper numbered list format instead of a bullet character. This works in `standard`, `text_with_image`, and any other layout that uses the `bullets` array. The override is per-slide, so bullet slides and numbered slides coexist safely in the same deck - you can freely mix both formats.

**Content guidelines:**
- 3-5 bullets or numbered steps on the left
- For walkthrough/path content (e.g., "how to create a boundary"), use numbered steps. This is cleaner than a `point_series` when the steps don't each carry enough content for a full slide.
- The screenshot should directly relate to the text - ideally showing the UI state referenced by the bullets
- Speaker notes should elaborate on each bullet/step with context the presenter can speak to

**When to use:**
- Explaining a UI concept with an accompanying screenshot
- Step-by-step paths where each step is brief (1 line) and doesn't need its own slide
- Showing a configuration result alongside the instructions to create it

### two_column - side-by-side comparisons

Use `two_column` whenever you contrast two approaches, options, or categories. Each column has a `header` and a `body` array of bullets.

**Content guidelines:**
- Column headers should be short and parallel (e.g., "Default Policies" vs "Custom Policies")
- Keep bullet count balanced between columns (3-5 bullets each)
- Bullets should be concise - the font is smaller (16pt) to fit two columns
- The build tool automatically extends the body text boxes to ~10 cm height and vertically centers all text. This means:
  - 3-5 bullets per column looks best - the content sits centered in the available space
  - Too few bullets (1-2) will look sparse and floating; add more substance or use `standard` instead
  - Too many bullets (7+) risk overflowing the box; split across two slides or trim
  - Both columns should have a similar bullet count so they look balanced side by side
- **Each column must have enough substantive content to justify the layout.** If the content for one or both columns is thin (just 1-2 short bullets), a `standard` bullet slide or `text_with_image` would be better. The two-column layout earns its place only when there's a genuine comparison with enough material on both sides.

**When to use:**
- Comparing two tools, approaches, or architectures
- Before/after scenarios
- Pros vs cons, strengths vs limitations

### three_column - triple comparison or feature breakdown

Use `three_column` for comparing three items or breaking a concept into three parts. Each column has a `header` and a `body` array of bullets.

**Content guidelines:**
- Column headers should be short (2-3 words max) - space is tight with three columns
- Keep bullets very concise (under 10 words) - font is 14pt
- Aim for 3-4 bullets per column
- The build tool automatically extends the body text boxes to ~10 cm height and vertically centers all text. This means:
  - 3-4 bullets per column is the sweet spot - centered and balanced
  - Fewer than 2 bullets per column will look empty; use `two_column` or `standard` instead
  - More than 5 bullets per column will overflow at 14pt font; trim or split content
  - Balance the bullet count across all three columns for a clean visual

**When to use:**
- Comparing three tiers, levels, or categories
- Breaking a topic into three pillars or dimensions
- Showing three options or paths

### point_series - when to use and when not to

The DT template has 4 point_series layouts (one per step highlight). **Each step in the `steps` array becomes its own slide**, so a 4-step series produces 4 slides with the active step visually highlighted.

There's no placeholder, but you can add images within the steps if you find a relevant one for the step, consider filling the blank, but just if it makes sense

**Use point_series ONLY for high-level processes or lifecycle stages**, not for step-by-step instructions. It exists to visually walk the audience through a conceptual journey where each phase represents a fundamentally different activity.

**Use point_series when:**
- The content describes a **process, lifecycle, or methodology** with 2-4 distinct phases
- Each phase is conceptually different from the others (not just "the next click")
- Each phase has enough real content (3+ substantive bullets AND 3-5 sentences of speaker notes) to stand on its own slide
- Good examples: "Assess -> Design -> Implement -> Validate", "Plan -> Enrich -> Restrict -> Verify"

**Do NOT use point_series when:**
- The steps are **sequential instructions** (click this, then click that) - use `text_with_image` with numbered steps instead
- Steps are too thin - you'd need to add fluff to fill a whole slide per step
- There are more than 4 steps (the template only supports 4)
- A simple bulleted list or a `text_with_image` with numbered steps would cover the same ground
- The steps describe configuring a single feature (e.g., creating a group, setting up a boundary) - these are instructions, not process phases

**The litmus test:** Ask yourself - "Is this a **process** (each phase is a different activity) or **instructions** (sequential steps within one activity)?" If it's instructions, use `text_with_image` with numbered steps. If it's a process where each phase could be its own mini-topic, use `point_series`.

**Step label rules:**
- Max ~20 characters (2-3 words)
- The sidebar placeholders are narrow - long labels wrap awkwardly
- Good: "Create Group", "Add Policy", "Attach Boundary"
- Bad: "Open Policy Management", "Define security_context query"
