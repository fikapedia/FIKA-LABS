# FIKA Labs — Climate Work

Operating guidance for producing climate consulting deliverables with Claude Code,
with sources pinned to verifiable vintages rather than model memory.

## Contents

| Path | What it is |
|---|---|
| `docs/claude-code-climate-work-operating-guide.docx` | Operating guide: workspace setup, source register, run loop, verification gates |
| `scripts/build_guide.js` | Reproducible build for the guide |

## Rebuilding the guide

```bash
npm install docx
node scripts/build_guide.js
```

Output lands in `docs/`.

## A note on currency

The source register in Section 3 of the guide is dated **18 September 2026**. Several
entries are moving targets — the EPA eGRID and Emission Factors Hub release cadence,
the GHG Protocol Scope 2 revision, CARB's SB 253 and SB 261 deadlines, and the SEC
rescission proceeding. Re-check them at the start of each engagement and update
`scripts/build_guide.js` rather than editing the `.docx` by hand.
