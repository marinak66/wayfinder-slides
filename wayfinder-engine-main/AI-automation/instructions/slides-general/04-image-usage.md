## Image Usage

The prompt includes an **Available Images** section listing every screenshot found in the CoE
knowledge base for the topic. Each entry has a repo-relative path and a description.

Rules for image usage:
- Reference images using the exact `path` from the catalog - the build tool resolves them.
- Prefer images whose description matches the slide's concept.
- If an image has a generic/timestamp filename but the context description mentions a specific UI element, use it.
- It's better to use an image that's approximately relevant than to have no image at all.
- Never invent image paths - only use paths from the catalog.
- **Image sizing is handled by the build tool** - images are scaled to **fit entirely** within the placeholder (contain mode) and centred. No cropping occurs - the full image is always visible. If the aspect ratio doesn't match the placeholder, there will be clean blank space on two sides. Do NOT worry about manual sizing or aspect ratio.
