---
name: eia-document-router
description: Route and diagnose EIA Word document chunks before specialized extraction. Use when an EIA agent needs to classify monitoring-plan/report tables, decide whether highway/waterway/railway/airport content should go to noise, surface-water, ecology, air, vibration, or other rules, or diagnose failed table/header/point matching without letting the LLM extract final monitoring values or compliance conclusions.
---

# EIA Document Router

Use this skill as an on-demand enhancement layer, not as the main extraction path.

Core workflow:

1. Run existing deterministic rules first for known highway noise and surface-water templates.
2. Trigger routing only when a table is unclassified, low confidence, missing required fields, or manually selected for enhanced recognition.
3. Ask the LLM to classify the chunk by project type, document role, section type, table role, and candidate rule.
4. Keep the LLM output as routing or diagnosis metadata. Do not let it extract final monitoring records, calculate standard indices, or write compliance conclusions.
5. Let existing rules perform point matching, numeric extraction, standard calculation, and compliance judgment.
6. If a rule fails, ask the LLM to diagnose likely failure causes and candidate header mappings; accept only headers that exist in the source table.
7. Send unsupported waterway, railway, and airport content to review until dedicated rules are implemented.

For route schemas, supported domains, and failure-diagnosis rules, read `references/router-rules.md`.
