import { z } from "zod";
import { DeckModel, Slide, Shape, Rect } from "../pptx/types";
import { Finding, DeckAggregates, FindingCategory, Severity, makeFindingId } from "../rules/types";
import { getAnthropic, getModel } from "./client";
import { cellRect } from "../pptx/render";
import { normalizeText } from "../rules/util";

// ---------------------------------------------------------------------------
// Output schema (validated with zod to reject malformed model output).
// ---------------------------------------------------------------------------

const AiFindingSchema = z.object({
  category: z.enum([
    "consistency",
    "bullets",
    "title",
    "terminology",
    "table",
    "content",
    "footnote",
  ]),
  severity: z.enum(["error", "warning", "info"]),
  confidence: z.number().min(0).max(1),
  title: z.string().min(3).max(120),
  detail: z.string().min(3).max(400),
  evidence: z.string().min(1).max(200),
  suggestion: z.string().max(300).optional().default(""),
  ruleId: z.string().max(60).optional().default("ai-finding"),
});

const AiResponseSchema = z.object({ findings: z.array(AiFindingSchema).max(12) });

// Minimum confidence we will surface from the AI at all.
const AI_CONFIDENCE_FLOOR = 0.6;

// ---------------------------------------------------------------------------
// Prompt construction
// ---------------------------------------------------------------------------

const SYSTEM_PROMPT = `You are a meticulous brand-compliance reviewer for ACME, an executive compensation consulting firm that produces board-ready PowerPoint decks.

You review ONE slide at a time (with deck-wide context provided) and report only HIGH-VALUE, HIGH-CONFIDENCE brand/style issues that a senior reviewer would genuinely flag. You are the semantic layer; deterministic code already checks fonts, sizes, colors, the confidentiality statement, spelled-out "Percentile", and table shading. DO NOT report those.

Focus ONLY on these checks:
1. Bullet parallelism: bullets in the same list must share parallel grammatical structure. Flag a list that mixes structures (e.g., some bullets are noun phrases, others full sentences).
2. Headline quality: the slide title should state the takeaway/insight, not just the topic. Flag topic-only titles. (Headline line count >3 is handled by deterministic code — do NOT flag length.)
3. Client-name consistency: the deck's standard client name is provided. Flag any place that uses a different variant of the client's name (e.g., a longer/alternate form).
4. Proper-noun punctuation: flag clear punctuation errors in names (e.g., a missing period in an abbreviation like "C.R Bard" which should be "C.R. Bard").
5. Missing units/scale: flag numeric table columns (e.g., revenue, market cap) that lack a units/scale indicator (such as $MM or ($000s)) when the magnitude implies one is needed.
6. Label/terminology consistency: flag inconsistent labels for the same concept across the deck (e.g., calling a legend "Key" on some slides and "Legend" on others; "Company Name" vs "Company").

STRICT RULES:
- Only flag issues you can directly justify from the provided slide text. Every finding MUST include "evidence" that is an EXACT substring copied from the slide content.
- Be conservative. When in doubt, do NOT flag. Prefer missing a borderline issue over creating noise. Aim for precision.
- Do NOT flag anything about fonts, font sizes, colors, palette, the confidentiality/privacy statement, "Percentile" vs "%ile", or table row shading — those are handled elsewhere.
- Do NOT invent content. Do NOT restate the same issue multiple times.
- confidence reflects how sure you are it is a real, reportable issue (0.0-1.0).
- Respond with ONLY a JSON object: {"findings": [ ... ]}. If there are no issues, return {"findings": []}.`;

interface SlidePayload {
  slideIndex: number;
  title: string | null;
  textBlocks: { name: string; text: string }[];
  bulletLists: { name: string; bullets: { level: number; text: string }[] }[];
  tables: { headers: string[]; sampleRow: string[] }[];
}

function buildSlidePayload(slide: Slide): SlidePayload {
  const textBlocks: { name: string; text: string }[] = [];
  const bulletLists: SlidePayload["bulletLists"] = [];
  const tables: SlidePayload["tables"] = [];

  for (const s of slide.shapes) {
    if (s.placeholderType === "ftr" || s.placeholderType === "sldNum") continue;
    if (s.kind === "table" && s.table) {
      const headerRows = s.table.rows.slice(0, 2);
      const headers = headerRows
        .flatMap((r) => r.map((c) => c.text.trim()))
        .filter(Boolean);
      const firstBody = s.table.rows.find(
        (r, i) => i >= 2 && r.some((c) => c.text.trim())
      );
      tables.push({
        headers: dedupeStrings(headers).slice(0, 16),
        sampleRow: (firstBody || []).map((c) => c.text.trim()).filter(Boolean).slice(0, 16),
      });
      continue;
    }
    if (s.kind !== "text") continue;
    if (s.placeholderType === "title" || s.placeholderType === "ctrTitle") continue;
    const bulletParas = s.paragraphs.filter(
      (p) => (p.bulletType === "char" || p.bulletType === "auto") && p.text.trim()
    );
    if (bulletParas.length >= 2) {
      bulletLists.push({
        name: s.name,
        bullets: bulletParas.map((p) => ({ level: p.level, text: p.text.trim() })),
      });
    } else if (s.text.trim()) {
      textBlocks.push({ name: s.name, text: s.text.trim().slice(0, 600) });
    }
  }

  return {
    slideIndex: slide.index,
    title: slide.titleShape?.text.trim() || null,
    textBlocks: textBlocks.slice(0, 8),
    bulletLists: bulletLists.slice(0, 6),
    tables: tables.slice(0, 4),
  };
}

function dedupeStrings(arr: string[]): string[] {
  return [...new Set(arr)];
}

function buildUserMessage(payload: SlidePayload, agg: DeckAggregates): string {
  const deckContext = {
    standardClientName: agg.clientName.dominant,
    clientNameVariantsSeen: Object.keys(agg.clientName.variants),
    legendLabelsSeen: dedupeStrings(agg.legendOrders.map((l) => l.label)),
  };
  return [
    "DECK CONTEXT (for consistency checks):",
    JSON.stringify(deckContext),
    "",
    `SLIDE ${payload.slideIndex} CONTENT:`,
    JSON.stringify(
      {
        title: payload.title,
        textBlocks: payload.textBlocks,
        bulletLists: payload.bulletLists,
        tables: payload.tables,
      },
      null,
      0
    ),
    "",
    'Return {"findings": [...]} with only high-confidence brand issues for THIS slide.',
  ].join("\n");
}

// ---------------------------------------------------------------------------
// Evidence grounding + highlight location
// ---------------------------------------------------------------------------

function locateEvidence(
  slide: Slide,
  evidence: string
): { shapeId?: string; rect: Rect | null; cell?: { row: number; col: number } } {
  const needle = normalizeText(evidence).toLowerCase();
  if (!needle) return { rect: null };
  for (const s of slide.shapes) {
    if (s.kind === "table" && s.table) {
      for (let ri = 0; ri < s.table.rows.length; ri++) {
        for (let ci = 0; ci < s.table.rows[ri].length; ci++) {
          const t = normalizeText(s.table.rows[ri][ci].text).toLowerCase();
          if (t && (t.includes(needle) || needle.includes(t)) && t.length > 1) {
            return {
              shapeId: s.id,
              rect: cellRect(s.rect, s.table, ri, ci) || s.rect,
              cell: { row: ri, col: ci },
            };
          }
        }
      }
    }
    const st = normalizeText(s.text).toLowerCase();
    if (st && st.includes(needle)) {
      return { shapeId: s.id, rect: s.rect };
    }
  }
  return { rect: null };
}

function slideContainsEvidence(slide: Slide, evidence: string): boolean {
  const needle = normalizeText(evidence).toLowerCase();
  if (needle.length < 2) return false;
  const hay = normalizeText(
    slide.shapes
      .map((s) => (s.kind === "table" && s.table ? s.text : s.text))
      .join(" ")
  ).toLowerCase();
  return hay.includes(needle);
}

// ---------------------------------------------------------------------------
// Orchestration
// ---------------------------------------------------------------------------

async function mapLimit<T, R>(items: T[], limit: number, fn: (item: T) => Promise<R>): Promise<R[]> {
  const results: R[] = new Array(items.length);
  let i = 0;
  async function worker() {
    while (i < items.length) {
      const idx = i++;
      results[idx] = await fn(items[idx]);
    }
  }
  await Promise.all(Array.from({ length: Math.min(limit, items.length) }, worker));
  return results;
}

function parseJsonLoose(text: string): unknown {
  const start = text.indexOf("{");
  const end = text.lastIndexOf("}");
  if (start === -1 || end === -1 || end < start) return null;
  try {
    return JSON.parse(text.slice(start, end + 1));
  } catch {
    return null;
  }
}

async function analyzeSlide(
  slide: Slide,
  agg: DeckAggregates,
  model: string
): Promise<Finding[]> {
  const anthropic = getAnthropic();
  if (!anthropic) return [];

  const payload = buildSlidePayload(slide);
  // Skip slides with essentially no analyzable content.
  if (!payload.title && payload.textBlocks.length === 0 && payload.bulletLists.length === 0 && payload.tables.length === 0) {
    return [];
  }

  let raw: string;
  try {
    const resp = await anthropic.messages.create({
      model,
      max_tokens: 1024,
      temperature: 0,
      system: SYSTEM_PROMPT,
      messages: [{ role: "user", content: buildUserMessage(payload, agg) }],
    });
    raw = resp.content
      .filter((b): b is { type: "text"; text: string } => b.type === "text")
      .map((b) => b.text)
      .join("\n");
  } catch (err) {
    console.error(`AI call failed for slide ${slide.index}:`, err);
    return [];
  }

  const parsed = parseJsonLoose(raw);
  const validated = AiResponseSchema.safeParse(parsed);
  if (!validated.success) {
    console.error(`AI output invalid for slide ${slide.index}`);
    return [];
  }

  const out: Finding[] = [];
  const seen = new Set<string>();
  for (const f of validated.data.findings) {
    if (f.confidence < AI_CONFIDENCE_FLOOR) continue;
    // Hallucination guard: evidence must appear on the slide.
    if (!slideContainsEvidence(slide, f.evidence)) continue;
    const dedupeKey = `${f.category}:${f.evidence.toLowerCase().slice(0, 40)}`;
    if (seen.has(dedupeKey)) continue;
    seen.add(dedupeKey);

    const loc = locateEvidence(slide, f.evidence);
    out.push({
      id: makeFindingId(f.ruleId || "ai", slide.index),
      ruleId: f.ruleId ? `ai-${f.ruleId}` : "ai-finding",
      category: f.category as FindingCategory,
      severity: f.severity as Severity,
      confidence: f.confidence,
      slideIndex: slide.index,
      title: f.title,
      detail: f.detail,
      evidence: f.evidence,
      suggestion: f.suggestion || undefined,
      guideline: "AI semantic review (ACME brand guidelines)",
      source: "ai",
      shapeId: loc.shapeId,
      rect: loc.rect,
      cell: loc.cell,
    });
  }
  return out;
}

export async function runAiChecks(deck: DeckModel, agg: DeckAggregates): Promise<Finding[]> {
  if (!getAnthropic()) return [];
  const model = getModel();
  const perSlide = await mapLimit(deck.slides, 5, (slide) => analyzeSlide(slide, agg, model));
  return perSlide.flat();
}
