## Session Structure Guidance

- **Hidden metadata slide (slide 1).** The very first slide in the first section must be `"slide_type": "metadata"` with `"layout": "standard"`. The build tool automatically hides this slide during presentation mode and places it before everything else. Include only:
  - **Focus areas** covered in this session
  - **Estimated duration** — calculate from your actual content slide count × ~2 minutes. If you generated 18 content slides, write "~36 min". Do NOT just repeat the definition's max duration.
  
  Do NOT include format, content types, or source path — those are redundant with the definition file.
- **The Agenda slide must be slide 2** in the first section (immediately after metadata). List the sections and what attendees will learn.
- **Agenda bullet format** - use the colon format: `"Topic: subtitle..."` (e.g., "IAM Fundamentals: prerequisites, policy types, and boundaries"). Do NOT use double-spaced dashes like "Topic  - subtitle".
- **End with key takeaways** - summarize the 3-5 most important points.
- **Include a resources slide** - link to docs, videos, CoE materials.
- **Demo slides** should list the steps the presenter will perform live.
- **Quiz questions** should have exactly 4 options (A/B/C/D) and the answer in speaker notes. **Quizzes are only for `training` format** - never include them in `workshop` decks.
- **All quiz slides go in a dedicated section at the end** (before Key Takeaways and Resources). Do NOT scatter quizzes throughout the content sections. Group them into a single "Knowledge Check" section so the flow is: content sections -> quizzes -> takeaways -> resources.
- **Section dividers** are generated automatically - you don't need to add them.
