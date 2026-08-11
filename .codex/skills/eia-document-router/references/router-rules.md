# Router Rules

## Route Schema

- `project_type`: `highway`, `waterway`, `railway`, `airport`, or `unknown`.
- `document_role`: `plan`, `report`, `reference`, or `unknown`.
- `section_type`: `noise`, `surface_water`, `regional_status`, `ecology`, `air`, `vibration`, `sediment`, or `unknown`.
- `table_role`: `plan_points`, `monitor_results`, `standard`, `frequency`, `method`, `metadata`, or `unknown`.
- `candidate_rule`: implemented rule name, currently `noise_plan`, `noise_result`, `surface_water_plan`, or `surface_water_result`.
- `confidence`: 0 to 1.
- `needs_review`: true for unsupported, low-confidence, or ambiguous routes.

## Trigger Rules

- Do not call routing for standard templates that existing rules parse successfully.
- Call routing when headers are non-standard, required fields are missing, table titles are atypical, point matching fails, or a user enables enhanced recognition.
- Call failure diagnosis only after a rule reports missing fields, duplicated header mapping, unclassified monitoring-like table, or plan/report mismatch.

## Domain Expansion

- Highway: keep existing noise and surface-water rules as primary supported routes.
- Waterway: route water quality, sediment, ecology, aquatic biology, and channel-noise material to review until rules exist.
- Railway: route railway boundary noise, vibration, and sound-barrier material to review until rules exist.
- Airport: route aircraft noise, contours, sensitive points, and weighted indicators to review until rules exist.

## Guardrails

- Never use LLM output as final measured values, standards, point codes, or compliance conclusions.
- Candidate header mappings must refer to headers that exist in the source table.
- Unsupported domains must not be forced into highway rules.
- Low-confidence routes and diagnoses must be written as review items.
