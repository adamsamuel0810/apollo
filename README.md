# ACME Brand Compliance Checker

Upload a PowerPoint (`.pptx`) and review brand-guideline violations slide by slide.
Each slide is rendered in the browser with clickable, numbered highlight boxes over
the flagged elements, and a per-slide panel lets a reviewer **accept or reject** each
flag and export the decisions.

The engine combines a **deterministic OOXML rule engine** (precise, code-checkable
rules) with an **AI orchestration layer** (Anthropic Claude) for the semantic checks
that are hard to codify. It is tuned for **low noise**: against the annotated sample
deck it achieves **100% recall on the deterministic-addressable issues and 0 false
positives in the default ("primary") review tier**.

---

## Why this architecture

Brand compliance splits cleanly into two classes of problem, so the system uses two
layers and merges the results.

### 1. Deterministic rule engine (precision-first)
Most brand rules are objective: a font that isn't Calibri, a title that isn't 24 pt,
a missing confidentiality statement, a statistics row that isn't shaded yellow,
"Percentile" written out instead of "%ile". These are checked in code against a
**normalized deck model** built directly from the PPTX OOXML.

The hard part is **style inheritance**. In PowerPoint a run rarely states its own
font size; it inherits through `run → paragraph → placeholder → slide layout →
slide master txStyles → theme`. Without resolving that chain, size/font rules produce
constant false positives. `lib/pptx/resolve.ts` implements the full cascade, which is
why typography checks are reliable.

Deterministic findings are deliberately conservative and carry a `severity` and a
numeric `confidence`. Subjective or geometry-sensitive checks (e.g. "title moved from
the layout default", "table title not centered") are emitted at low confidence / `info`
severity so they never pollute the default view — the ground-truth annotations are
themselves inconsistent on exactly these, which is the tell that they should not be
hard errors.

### 2. AI orchestration layer (recall for semantics)
Some issues require judgment: are bullets grammatically parallel? Does the headline
state a *takeaway* rather than a topic? Is the client's name used inconsistently across
the deck? Is a numeric column missing its `$MM` scale? Is "C.R Bard" missing a period?

`lib/ai/orchestrator.ts` sends one compact, structured payload per slide to Claude
(with deck-level context such as the dominant client name and the legend labels seen
across the deck). The model must return **strict JSON validated with Zod**, every
finding must quote **exact evidence text that actually exists on the slide** (a
hallucination guard drops anything that doesn't), and a confidence floor + dedupe keep
noise down. Calls run concurrently with a cap to stay within serverless limits.

### Merge + noise control
`lib/findings/merge.ts` collapses duplicate deterministic findings, drops AI findings
that duplicate a deterministic one (deterministic is authoritative), and sorts by
slide → severity → confidence. The UI shows **errors/warnings at or above a confidence
threshold** by default; a single toggle reveals the low-confidence / `info` tier. This
is the core of the "tight control over noise" requirement.

### Rendering without LibreOffice
Vercel's serverless runtime can't run LibreOffice/PowerPoint, so slides are
**re-rendered from the parsed geometry**: absolutely-positioned shapes, native HTML
tables with real cell fills/borders, embedded images as data URLs, and labeled
placeholders for native charts. Because the highlight overlays use the *same* EMU
geometry as the rules, the boxes map exactly onto the flagged elements.

---

## Tech stack

| Concern        | Choice                                             |
| -------------- | -------------------------------------------------- |
| Framework      | Next.js (App Router) + TypeScript, React           |
| Styling        | Tailwind CSS                                        |
| PPTX parsing   | `jszip` + `fast-xml-parser` (order-preserving)     |
| Validation     | `zod` (AI output schema)                            |
| AI             | `@anthropic-ai/sdk` (Claude)                        |
| Auth           | Custom password gate: middleware + signed HMAC cookie |
| Deployment     | Vercel                                              |

---

## Project layout

```
app/                     # Next.js routes (login, API: login/logout/analyze) + review UI
components/               # SlideCanvas, FindingsPanel, SlideRail, Uploader
lib/
  pptx/                  # parser.ts, resolve.ts (inheritance), render.ts, types.ts, xml.ts
  rules/                 # deterministic rules + engine + shared types/util
  deck/aggregate.ts      # deck-level aggregates (client name, footer norm, legends)
  ai/                    # Claude client + orchestrator
  findings/merge.ts      # dedupe / merge / threshold
  brand/                 # palette.ts, guidelines.ts (constants distilled from the guidelines)
  analyze/               # end-to-end pipeline + API result types
eval/                    # ground-truth harness (precision/recall) + smoke test
middleware.ts            # password gate
```

---

## Local setup

```bash
npm install
cp .env.example .env.local   # then edit values (see below)
npm run dev                  # http://localhost:3000
```

`.env.local`:

```
APP_PASSWORD=<the password reviewers type>
AUTH_SECRET=<long random string>
ANTHROPIC_API_KEY=<your Anthropic key>
ANTHROPIC_MODEL=claude-3-5-sonnet-latest   # optional
ENABLE_AI=1                                 # set 0 to run deterministic-only
```

> Without `ANTHROPIC_API_KEY` (or with `ENABLE_AI=0`) the app runs in
> **deterministic-only** mode — fully functional, just without the semantic checks.

---

## Evaluation harness

The annotated sample deck has issues flagged in its speaker notes. `eval/` reverse-
engineers those notes into ground-truth tags and measures the engine against them.

```bash
npm run eval
```

It reports recall per layer (deterministic / AI / out-of-scope) and the **false-positive
count in the primary tier**, and smoke-tests the second (unannotated) deck to confirm
the rules generalize and aren't tuned to one file.

Latest run (deterministic-only):

```
Deterministic recall: 17/17 (100%)
Primary-tier precision: 12/12 (100%)  — no false positives in the default view
```

---

## Deployment (Vercel)

1. Push this repo to GitHub (see below).
2. Import the repo in Vercel → **New Project**.
3. Add Environment Variables: `APP_PASSWORD`, `AUTH_SECRET`, `ANTHROPIC_API_KEY`
   (and optionally `ANTHROPIC_MODEL`, `ENABLE_AI`).
4. Deploy. The app is gated by `APP_PASSWORD` on every route.

---

## How findings are controlled (summary)

- **Severity** (`error | warning | info`) + numeric **confidence** on every finding.
- Default view = `error`/`warning` with confidence ≥ 0.6; everything else behind a toggle.
- Deterministic rules fire only on unambiguous violations; AI findings must quote
  on-slide evidence and clear a confidence floor.
- Cross-source dedupe + per-slide grouping; deck-level aggregates so consistency
  checks flag the *minority* variant rather than every instance.
```
