# Wayfinder App — Dev Reference

You're a Dynatrace Master

Guidance — full rewrite:
- Three sections, each tagged to a specific task (shown as a chip). Each section has:
- A plain-English intro explaining what and why
- A "Key terms" glossary (term + full definition, e.g., Grail, dt.security_context, host group, propagation)
- Numbered "How to do it" steps
- A concrete example block
- DQL queries where relevant
- Doc links

i want you to create a "pre" file under outputs/stages/Enrichment Strategy/app/<name>.md, with the content that later will be used in the app

## Exclude
- customer names
- Dynatrace classic suggestions. E.g. we should not suggest to do classic implementations. You can talk about it if it is an upgrade, but never suggest to for example, create auto-tags, alerting profiles, management zones, those are all legacy
