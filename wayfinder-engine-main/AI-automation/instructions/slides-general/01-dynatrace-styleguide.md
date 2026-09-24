## Dynatrace style guide

All agent-generated content must follow the
[Dynatrace Content Style Guide](https://styleguide.dynatrace.com/docs/category/style-guidelines/).
The rules below are the subset most relevant to slide decks, project plans,
and other deliverables produced by these agents.

---

### Voice and tone

The Dynatrace voice is **confident, approachable, and empowering**.

- Be clear, concise, and decisive. Describe with precision and without hyperbole.
- Be pragmatic. When there's an issue, explain why and share known solutions.
- Use plain English. Stay away from filler words, unnecessary parentheticals,
  and internal technical jargon.
- Use contractions to keep the tone conversational (it's, you're, we're, don't).
  Spell out negative contractions in warnings or limitations for clarity
  (for example, "do not" instead of "don't" when the consequence is serious).
- Address users as "you" (second person). Refer to Dynatrace as "we."
- Be inclusive and respectful of diverse audiences.
- Avoid jargon and internal terminology. Use only approved platform terminology.

---

### Active voice

Write in the active voice. Active statements convey authority, transparency,
and are easier to read. Passive voice often sounds defensive.

| Do | Don't |
|---|---|
| The monitoring system identified the problem. | The problem was identified by the monitoring system. |
| All three formats support the same request payloads. | The same request payloads are supported in all three formats. |

Use passive voice only when the actor is not important (for example, in
warnings or glossary definitions): "Deletion cannot be reversed."

---

### Verb tenses

Use only simple tenses (past, present, future). Avoid progressive tenses
("is running", "was deploying"). Present tense is best for describing
features and general system behavior.

| Do | Don't |
|---|---|
| Davis AI evaluates billions of transactions per second. | Davis AI is evaluating billions of transactions per second. |
| The agent collected host metrics every minute. | The agent was collecting host metrics every minute. |
| The dashboard shows real-time data. | The dashboard is showing real-time data. |

**Scope:** This rule applies to body text and speaker notes that describe features, system behavior, or steps. It does not apply to conversational slide title framing such as "What we're achieving today" — those are not describing system behavior and read naturally as-is.

For if-then statements describing standard behavior, use present tense in
both clauses. Use future tense in the main clause only when the outcome is
not guaranteed.

---

### Sentence-case capitalization

Capitalize only the first word and proper nouns in titles, headings, column
headers, slide titles, and labels.

| Do | Don't |
|---|---|
| The DPS license model allows broad adoption | The DPS License Model Allows Broad Adoption |
| How Dynatrace boosts production resilience | How Dynatrace Boosts Production Resilience |

---

### Titles and headings

- Use sentence case (see above).
- Do not end titles or headings with a period or other closing punctuation.
- Avoid gerund (-ing) verb forms **when the -ing word is the title's main verb**. Prefer active verb forms.
  This rule targets titles where the gerund drives the action — i.e., you could replace it with a bare imperative and produce a cleaner title. It does **not** apply to -ing words that function as nouns or modifiers within a compound noun phrase.

  | Pattern | Example | Action |
  |---|---|---|
  | Gerund leads the title | "Creating an anomaly detector" | Fix → "Create an anomaly detector" |
  | -ing word is a noun or modifier | "Connector mapping", "Alert suppression", "Cost alerting" | Leave it — it's not a verb |

  Test: can you cleanly replace the -ing word with a bare imperative verb? If yes, do it. If the -ing word is embedded in a noun compound and removing it distorts the meaning, leave it.
- Progressive constructions in slide titles ("What we're achieving today", "What we're doing today") are acceptable. They are conversational framing, not descriptions of system behavior, and read naturally. Do not flag or change them.
- Use a colon (`:`) followed by a space to break a long title into a title
  and subtitle. Capitalize the first word of the subtitle.
- Write concise headings front-loaded with keywords.

---

### Numbers

- Spell out zero through nine in sentences. Use numerals for 10 and above.
- Don't begin a sentence with a numeral. Rewrite or spell it out.
- In tables, labels, and constrained spaces, use numerals regardless of value.
- Use numerals for measurements, dimensions, metrics, and currency.
- Insert a space between a numeral and its unit (100 ms, 5 GB) except for
  percent signs (50%), currency symbols ($100), and temperatures (75°F).
- Number ranges: use "from X through Y" in prose; use an en dash (–) in
  tables or constrained spaces. Do not use hyphens for ranges.
- Percentages: spell out "percent" in prose (50 percent). Use % in tables.
- Ordinal numbers: spell them out (first, fifth). Don't use 1st, 3rd.
- Use thousands commas for numbers with four or more digits (1,000; 25,000).

---

### Acronyms and abbreviations

- On first use, write the full term followed by the acronym in parentheses.
- Don't capitalize all initial letters just to match an acronym. Capitalize
  only proper nouns.
  - Do: software-defined networking (SDN)
  - Don't: Software-Defined Networking (SDN)
- Don't define acronyms in headings. Introduce them in the paragraph body.

---

### Punctuation

**Serial comma:** Always use a comma before the conjunction in a list of
three or more items ("Infrastructure Monitoring, Log Monitoring, and
Application Security").

**Em dashes:** No spaces on either side of an em dash. Use em dashes to add
emphasis or replace parentheses. In Markdown, write `---`.
**In slide bullets: do not use em dashes (`---`) or double hyphens (`--`) as separators.** The build tool does not render em dashes correctly, and `--` is not a recognized punctuation mark. Rewrite the bullet to avoid the separator entirely.
- Do: "AWS multi-region + AKS, no serverless or on-prem"
- Don't: "AWS multi-region + AKS -- no serverless, no on-prem"
For non-slide content (project plans, docs), em dashes are acceptable per the Dynatrace style guide as long as there are no surrounding spaces.

**Compound predicates:** Don't insert a comma between two verbs that share
the same subject.
- Do: "Davis AI evaluates billions of transactions per second and alerts you when anomalies are detected."
- Don't: "Davis AI evaluates billions of transactions per second, and alerts you when anomalies are detected."

**Semicolons:** Use a semicolon to join independent clauses without a
conjunction: "Rules are evaluated from top to bottom; the first matching
rule applies."

---

### Lists

**Bulleted lists:**
- If any list item completes an introductory phrase, add a colon after the
  phrase and a period after each item.
- Don't use periods if all items are three or fewer words.
- If items are complete sentences on their own, use a period.

**Numbered lists:** Use for sequential steps only. Each item is a complete
sentence. No colon after the introductory phrase.

---

### UI interaction verbs

Use the correct verbs when describing interactions with Dynatrace:

| Action | Verb | Don't use |
|---|---|---|
| Click/tap a control | **Select** | Click, Tap |
| Navigate to a page/tab | **Go to** | Navigate to, Open (for pages) |
| Display a file/dashboard | **Open** | Launch |
| Dismiss a dialog/tab | **Close** | Exit, Quit |
| Type into a field | **Enter** | Type, Paste |
| Enable/disable a toggle | **Turn on** / **Turn off** | Toggle, Switch on |
| Make a preference choice | **Choose** | Pick |
| Remove a checkbox selection | **Clear** | Uncheck, Deselect |
| Authenticate | **Sign in** / **Sign in to** | Log in, Log in to |

Use `>` (with spaces) to chain UI navigation: **Settings** > **Preferences** > **Management zones**.

---

### Emoji

Emoji are inappropriate for enterprise Dynatrace content:

- **Don't** use emoji in documentation, marketing copy, or customer-facing
  deliverables.
- **Exception for slide decks:** in internal presentations, emoji (✅ ❌ ⚠️)
  are permitted **only** in `two_column` and `three_column` bullet bodies
  where columns represent opposing sentiments (good vs bad, do vs don't).
  This exception is defined in `03-visual-emphasis.md` and overrides the
  general "don't use emoji" rule for that specific context.
  Use words instead of emoji wherever possible.
- If a customer uses emoji first, you may reciprocate in casual communication.

---

### Plural nouns as adjectives

When a noun acts as an adjective, use the singular form:
- Do: metric browser, management-zone configuration, root cause analysis
- Don't: metrics browser, management-zones configuration

Exceptions: UI labels that exist in plural form (Services page, Problems list).

---

### Link text

Write meaningful link text that conveys the purpose of the link. Never use
"click here" or bare URLs as link text.
- Do: "For details, see [Identity and access management](https://docs.dynatrace.com/...)."
- Don't: "Click [here](https://docs.dynatrace.com/...) for details."