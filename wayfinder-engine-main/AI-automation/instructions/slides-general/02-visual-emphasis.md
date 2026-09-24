## Visual Emphasis

Use **bold** to make slides easier to scan.

### Bold formatting

Wrap key terms in `**double asterisks**` inside bullet text. The build tool renders them as native PowerPoint bold runs. Use bold to highlight:
- The single most important term or concept per bullet (not the whole sentence)
- Names of Dynatrace features, policy names, or attribute names on first mention (e.g., `**dt.security_context**`, `**All Grail data read access**`)
- Action verbs in instructional steps (e.g., `"1. **Go to** Policy Management"`)
- Do NOT bold entire bullets or more than 2-3 words per bullet - it loses impact
- Do NOT use bold in slide titles or column headers (those are already styled by the template)

### Code formatting

Wrap inline code (attribute names, values, format strings, commands) in backticks. The build tool renders them in a monospace font. Use backticks for:
- Attribute names and values: `dt.security_context`, `restricted`
- Format strings and examples: `plat:<platform>/comp:<component>/bu:<bu>/app:<app>`
- Commands or paths: `python generate_training.py build`

Do NOT put backtick content inside bold (`**`). Use one or the other — bold for key terms, code for literal strings.

### Emoji usage

Emojis are **restricted to column layouts only** (`two_column` and `three_column`) where they visually distinguish sides (e.g., ✅ Recommended vs ❌ Avoid). Do NOT use emojis anywhere else.

| Emoji | Use for |
|---|---|
| ✅ | Recommended, correct, or best practice |
| ❌ | Avoid, incorrect, or anti-pattern |
| ⚠️ | Caution, limitation, or watch-out |

**Rules:**
- Emojis are allowed ONLY in `two_column` and `three_column` bullet bodies
- Use them when columns represent opposing sentiments (good vs bad, do vs don't, before vs after)
- Do NOT put emojis in standard bullets, text_with_image bullets, point_series steps, slide titles, column headers, speaker notes, or quiz options
- Maximum one emoji per bullet, and only at the start of the bullet
- If the columns don't represent a comparison or contrast, skip emojis entirely
- Most slides in a deck should have ZERO emojis - they are the exception, not the rule
