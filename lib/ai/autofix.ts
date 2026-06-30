import { getAnthropic, getModel } from "./client";
import { TITLE } from "../brand/guidelines";
import { FixInstruction } from "../rules/types";
import { deterministicShortenTitle } from "../fixes/shortenTitle";

/**
 * Use Claude to rewrite a slide title so it fits within the ACME 3-line limit
 * while preserving the takeaway meaning. Falls back to deterministic shortening.
 */
export async function aiRewriteTitle(
  currentTitle: string,
  estimatedLines: number,
  boxWidthIn: number
): Promise<string> {
  const anthropic = getAnthropic();
  if (!anthropic) {
    return deterministicShortenTitle(currentTitle, boxWidthIn);
  }

  const charsPerLine = Math.max(
    20,
    Math.floor((boxWidthIn - 0.2) * 72 / (TITLE.sizePt * 0.48))
  );
  const maxChars = charsPerLine * TITLE.maxLines;

  const prompt = `You are an ACME executive-compensation consultant editing a PowerPoint slide title.

CURRENT TITLE (${estimatedLines} visual lines — must be ≤ ${TITLE.maxLines}):
"""${currentTitle}"""

RULES:
- Rewrite to fit within ~${maxChars} characters total (≈${TITLE.maxLines} lines at 24pt in a ${boxWidthIn.toFixed(1)}" box).
- Keep the key takeaway/insight, not just the topic.
- Sentence case. No ending punctuation.
- Do NOT add quotes or markdown.
- Return ONLY the rewritten title text, nothing else.`;

  try {
    const resp = await anthropic.messages.create({
      model: getModel(),
      max_tokens: 256,
      temperature: 0,
      messages: [{ role: "user", content: prompt }],
    });
    const text = resp.content
      .filter((b): b is { type: "text"; text: string } => b.type === "text")
      .map((b) => b.text)
      .join("")
      .trim()
      .replace(/^["']|["']$/g, "")
      .trim();
    if (text && text.length <= maxChars * 1.6) return text;
  } catch (err) {
    console.error("AI title rewrite failed, using deterministic fallback:", err);
  }

  return deterministicShortenTitle(currentTitle, boxWidthIn);
}

/** Resolve an ai-rewrite fix instruction (AI with deterministic fallback). */
export async function resolveAiFix(
  fix: FixInstruction,
  currentText: string,
  estimatedLines?: number,
  boxWidthIn?: number
): Promise<FixInstruction | null> {
  if (fix.kind !== "ai-rewrite") return fix;
  if (fix.newText) return fix;

  const w = boxWidthIn ?? 8.5;
  const rewritten = await aiRewriteTitle(
    currentText,
    estimatedLines ?? TITLE.maxLines + 1,
    w
  );

  return { ...fix, kind: "title-rewrite", newText: rewritten };
}

/** AI rewrite for arbitrary shape text (bullets, body, semantic fixes). */
export async function resolveAiShapeFix(
  fix: FixInstruction,
  currentText: string
): Promise<FixInstruction | null> {
  if (fix.newText) return { ...fix, kind: "title-rewrite" };
  const anthropic = getAnthropic();
  if (!anthropic) return null;

  const prompt = `You are an ACME brand-compliance editor. Rewrite the following text to fix the issue described.

CURRENT TEXT:
"""${currentText}"""

ISSUE / INSTRUCTIONS:
${fix.aiContext || "Fix the brand compliance issue while preserving meaning."}

RULES:
- Return ONLY the corrected text, no quotes or markdown.
- Keep changes minimal and targeted.
- Preserve numbers, names, and data unless the issue requires changing them.`;

  try {
    const resp = await anthropic.messages.create({
      model: getModel(),
      max_tokens: 512,
      temperature: 0,
      messages: [{ role: "user", content: prompt }],
    });
    const text = resp.content
      .filter((b): b is { type: "text"; text: string } => b.type === "text")
      .map((b) => b.text)
      .join("")
      .trim()
      .replace(/^["']|["']$/g, "")
      .trim();
    if (text) return { ...fix, kind: "title-rewrite", newText: text };
  } catch (err) {
    console.error("AI shape rewrite failed:", err);
  }
  return null;
}

/** Generate AI rewrites for all ai-rewrite fixes in parallel (capped). */
export async function resolveAiFixes(
  items: {
    fix: FixInstruction;
    currentText: string;
    estimatedLines?: number;
    boxWidthIn?: number;
  }[]
): Promise<FixInstruction[]> {
  const out: FixInstruction[] = [];
  for (const item of items) {
    const resolved = await resolveAiFix(
      item.fix,
      item.currentText,
      item.estimatedLines,
      item.boxWidthIn
    );
    if (resolved) out.push(resolved);
  }
  return out;
}
