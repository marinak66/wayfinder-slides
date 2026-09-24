## JSON Response Format

Return ONLY valid JSON in this schema:

```json
{
  "_meta": {
    "status": "initial-build-only",
    "warning": "This file generated the initial PPTX and is now archived. Do not rebuild or iterate from this file. The living document is on SharePoint.",
    "sharepoint_url": "https://dynatrace.sharepoint.com/sites/Post-Sales/Shared%20Documents/Wayfinder/stages/<Stage Name>",
    "generated_at": "YYYY-MM-DD"
  },
  "title": "Presentation title",
  "sections": [
    {
      "section_title": "Section name",
      "slides": [
        {
          "title": "Slide title",
          "layout": "standard",
          "bullets": ["bullet 1", "bullet 2", "bullet 3"],
          "speaker_notes": "Detailed speaker notes (3-5 sentences).",
          "slide_type": "content",
          "image": "content/.../assets/screenshot.png",
          "columns": [],
          "steps": []
        },
        {
          "title": "Lifecycle Overview",
          "layout": "point_series",
          "steps": [
            {
              "label": "Assess",
              "bullets": ["Evaluate current state", "Identify gaps"],
              "speaker_notes": "Start by assessing where the customer is today."
            },
            {
              "label": "Design",
              "bullets": ["Define target architecture", "Plan rollout"],
              "speaker_notes": "Design the target state based on assessment findings."
            }
          ]
        }
      ]
    }
  ]
}
```

Field reference:
- `layout`: `standard` | `text_with_image` | `two_column` | `three_column` | `point_series` | `dark_standard`
- `slide_type`: `content` | `technical` | `quiz` | `demo`. This is metadata only and does NOT appear in the presentation. Do not use any other values.
- `image`: repo-relative path from the image catalog (for `text_with_image`)
- `columns`: array of `{header, body:[]}` objects (for `two_column` / `three_column`)
- `steps`: array of up to 4 step objects for `point_series`. Each object: `{label, bullets:[], speaker_notes}`. **Each step becomes its own slide** with the active step highlighted. The layout does not support images.

Omit fields you don't use - they have sensible defaults.
