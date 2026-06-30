import { z } from "zod";
import { getAnthropic, getModel, aiEnabled } from "./client";
import { DeckModel } from "../pptx/types";
import { Finding, FixInstruction, DeckAggregates } from "../rules/types";
import { parseDeck } from "../pptx/parser";
import { titleText, estimateLines, emuToIn, shapeText } from "../rules/util";
import { TITLE } from "../brand/guidelines";
import { resolveAiFix, resolveAiShapeFix } from "./autofix";
import { inferFix } from "../fixes/inferFix";
import { computeAggregates } from "../deck/aggregate";

const AiFixItemSchema = z.object({
  findingId: z.string(),
  ruleId: z.string().optional(),
  slideIndex: z.number().optional(),
  shapeId: z.string().optional(),
  kind: z.enum([
    "text-replace",
    "title-rewrite",
    "deck-replace",
    "ai-rewrite",
    "font-family-set",
    "font-size-set",
    "font-bold-set",
    "color-set",
    "shape-reposition",
    "row-fill-set",
    "border-remove",
    "ensure-footer",
  ]),
  search: z.string().optional(),
  replace: z.string().optional(),
  newText: z.string().optional(),
  deckWide: z.boolean().optional(),
  caseInsensitive: z.boolean().optional(),
  regex: z.boolean().optional(),
  cellRow: z.number().optional(),
  cellCol: z.number().optional(),
  fontSizePt: z.number().optional(),
  fromColor: z.string().optional(),
  toColor: z.string().optional(),
  colorTarget: z.enum(["fill", "text", "both"]).optional(),
  rowIndex: z.number().optional(),
  fillColor: z.string().optional(),
});

const AiFixPlanSchema = z.object({
  fixes: z.array(AiFixItemSchema).max(200),
});

export type FixItem = {
  fix: FixInstruction;
  findingId: string;
  ruleId: string;
  title: string;
  source: "deterministic" | "ai";
};

const CHUNK_SIZE = 35;

function summarizeDeck(deck: DeckModel) {
  return deck.slides.map((s) => ({
    index: s.index,
    kind: s.kind,
    title: titleText(s)?.slice(0, 200) || "",
    shapes: s.shapes.slice(0, 30).map((sh) => ({
      id: sh.id,
      name: sh.name,
      placeholderType: sh.placeholderType,
      text: shapeText(sh).slice(0, 500),
    })),
  }));
}

function findingKey(f: Finding): string {
  return `${f.ruleId}|${f.slideIndex}|${f.shapeId || ""}|${f.cell?.row ?? ""}|${f.cell?.col ?? ""}|${f.id}`;
}

function toFixInstruction(item: z.infer<typeof AiFixItemSchema>): FixInstruction {
  const fix: FixInstruction = {
    kind: item.kind,
    slideIndex: item.slideIndex ?? 1,
    shapeId: item.shapeId,
    search: item.search,
    replace: item.replace,
    newText: item.newText,
    deckWide: item.deckWide,
    caseInsensitive: item.caseInsensitive,
    regex: item.regex,
    fontSizePt: item.fontSizePt,
    fromColor: item.fromColor,
    toColor: item.toColor,
    colorTarget: item.colorTarget,
    rowIndex: item.rowIndex,
    fillColor: item.fillColor,
    borderScope: item.kind === "border-remove" ? "table-sides" : undefined,
  };
  if (item.cellRow != null && item.cellCol != null) {
    fix.cell = { row: item.cellRow, col: item.cellCol };
  }
  return fix;
}

async function callAiFixPlanChunk(
  deck: DeckModel,
  findings: Finding[],
  aggregates: DeckAggregates
): Promise<Map<string, FixInstruction>> {
  const anthropic = getAnthropic();
  const out = new Map<string, FixInstruction>();
  if (!anthropic || findings.length === 0) return out;

  const slides = summarizeDeck(deck);
  const payload = findings.map((f) => ({
    findingId: f.id,
    ruleId: f.ruleId,
    slideIndex: f.slideIndex,
    shapeId: f.shapeId,
    cell: f.cell,
    title: f.title,
    detail: f.detail,
    evidence: f.evidence,
    suggestion: f.suggestion,
    category: f.category,
    severity: f.severity,
  }));

  const prompt = `You are an ACME brand-compliance auto-fix engine. Produce fixes for EVERY finding below.

You MUST emit exactly one fix entry per findingId listed. Do not skip any finding.

FIX CAPABILITIES (use the appropriate kind):
- text-replace / deck-replace: terminology, punctuation, designators, label consistency
- title-rewrite: rewrite full title or footer text (newText field)
- font-family-set: change non-Calibri text to Calibri (shapeId required)
- font-size-set: set font size in pt (fontSizePt, shapeId, optional cellRow/cellCol)
- font-bold-set: make title bold
- color-set: replace off-palette color (fromColor, toColor, colorTarget fill|text)
- shape-reposition: move footer to standard position (when confidentiality-moved)
- row-fill-set: shade statistics row yellow (rowIndex, fillColor FFFFDB)
- border-remove: remove table left/right borders (borderScope implied)

STANDARD VALUES:
- Confidentiality text: "${aggregates.confidentialityText || ""}"
- Client name: "${aggregates.clientName.dominant || ""}"
- Percentile -> %ile (deckWide)
- TGT -> Target (deckWide)
- Incumbent -> Executive
- Company Name -> Company
- Stats row fill: FFFFDB
- Title max 3 lines, 24pt, no trailing punctuation

For bullet parallelism or semantic issues: rewrite the affected text minimally using title-rewrite with newText.

SLIDES:
${JSON.stringify(slides)}

FINDINGS (${findings.length}) — fix ALL of them:
${JSON.stringify(payload)}

Return ONLY valid JSON:
{"fixes":[{"findingId":"<id>","kind":"text-replace","slideIndex":1,"shapeId":"...","search":"...","replace":"..."}]}`;

  try {
    const resp = await anthropic.messages.create({
      model: getModel(),
      max_tokens: 8192,
      temperature: 0,
      messages: [{ role: "user", content: prompt }],
    });

    const raw = resp.content
      .filter((b): b is { type: "text"; text: string } => b.type === "text")
      .map((b) => b.text)
      .join("");

    const jsonMatch = raw.match(/\{[\s\S]*\}/);
    if (!jsonMatch) return out;

    const parsed = AiFixPlanSchema.parse(JSON.parse(jsonMatch[0]));
    const byId = new Map(findings.map((f) => [f.id, f]));

    for (const item of parsed.fixes) {
      const match = byId.get(item.findingId);
      if (!match) continue;
      const fix = toFixInstruction({ ...item, slideIndex: item.slideIndex ?? match.slideIndex });
      if (!fix.shapeId && match.shapeId) fix.shapeId = match.shapeId;
      if (!fix.slideIndex) fix.slideIndex = match.slideIndex;
      out.set(match.id, fix);
    }
  } catch (err) {
    console.error("AI batch fix plan chunk failed:", err);
  }

  return out;
}

async function callAiFixPlan(
  deck: DeckModel,
  findings: Finding[],
  aggregates: DeckAggregates
): Promise<Map<string, FixInstruction>> {
  const merged = new Map<string, FixInstruction>();
  for (let i = 0; i < findings.length; i += CHUNK_SIZE) {
    const chunk = findings.slice(i, i + CHUNK_SIZE);
    const part = await callAiFixPlanChunk(deck, chunk, aggregates);
    for (const [k, v] of part) merged.set(k, v);
  }
  return merged;
}

function seedKnownFixes(
  findings: Finding[],
  deck: DeckModel,
  aggregates: DeckAggregates
): Map<string, FixInstruction> {
  const out = new Map<string, FixInstruction>();
  for (const f of findings) {
    if (f.fix) {
      out.set(f.id, { ...f.fix });
      continue;
    }
    const inferred = inferFix(f, deck, aggregates);
    if (inferred) out.set(f.id, inferred);
  }
  return out;
}

async function resolveFix(
  fix: FixInstruction,
  finding: Finding,
  deck: DeckModel
): Promise<FixInstruction | null> {
  if (fix.kind !== "ai-rewrite" && !(fix.kind === "title-rewrite" && !fix.newText)) {
    return fix;
  }

  const slide = deck.slides.find((s) => s.index === finding.slideIndex);
  if (!slide) return fix;

  const shape = fix.shapeId
    ? slide.shapes.find((s) => s.id === fix.shapeId)
    : slide.titleShape;
  const text = shape ? shapeText(shape) : finding.evidence || "";

  const isTitle =
    finding.ruleId === "title-lines" ||
    shape?.placeholderType === "title" ||
    shape?.placeholderType === "ctrTitle";

  if (isTitle) {
    const title = slide.titleShape;
    const tText = titleText(slide) || text;
    const w = title?.rect ? emuToIn(title.rect.cx) : 8.5;
    const lines = tText ? estimateLines(tText, TITLE.sizePt, w) : TITLE.maxLines + 1;
    return resolveAiFix(fix, tText, lines, w);
  }

  return resolveAiShapeFix(fix, text);
}

/** Build the complete fix list covering every error/warning finding. */
export async function buildBatchFixPlan(
  pptx: Buffer,
  findings: Finding[]
): Promise<FixItem[]> {
  const deck = await parseDeck(pptx);
  const aggregates = computeAggregates(deck);
  const actionable = findings.filter(
    (f) => f.severity === "error" || f.severity === "warning"
  );

  const seeded = seedKnownFixes(actionable, deck, aggregates);

  if (aiEnabled() && actionable.length > 0) {
    const aiPlan = await callAiFixPlan(deck, actionable, aggregates);
    for (const [id, fix] of aiPlan) {
      if (!seeded.has(id)) seeded.set(id, fix);
    }
  }

  // Final pass: infer any still-missing fixes deterministically
  for (const f of actionable) {
    if (!seeded.has(f.id)) {
      const inferred = inferFix(f, deck, aggregates);
      if (inferred) seeded.set(f.id, inferred);
    }
  }

  const items: FixItem[] = [];

  for (const f of actionable) {
    let fix = seeded.get(f.id);
    if (!fix) {
      console.warn(`No fix available for ${f.ruleId} on slide ${f.slideIndex}`);
      continue;
    }

    const resolved = await resolveFix(fix, f, deck);
    if (!resolved) continue;
    fix = resolved;

    items.push({
      fix,
      findingId: f.id,
      ruleId: f.ruleId,
      title: f.title,
      source: f.source === "ai" ? "ai" : "deterministic",
    });
  }

  return items;
}

export { findingKey };
