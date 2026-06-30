import { AnalyzeResult } from "../analyze/types";
import { Finding } from "../rules/types";
import { Shape } from "../pptx/types";
import { deterministicShortenTitle, titleMaxChars } from "../fixes/shortenTitle";
import { TITLE } from "../brand/guidelines";
import { emuToIn } from "../rules/util";

function buildReplaceRegex(fix: NonNullable<Finding["fix"]>): RegExp {
  if (fix.regex) {
    return new RegExp(fix.search!, fix.caseInsensitive ? "gi" : "g");
  }
  return new RegExp(
    fix.search!.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"),
    fix.caseInsensitive ? "gi" : "g"
  );
}

/** Predict the post-fix text for instant preview (no network). */
export function predictFixText(finding: Finding, slide: AnalyzeResult["slides"][0]): string | null {
  const fix = finding.fix;
  if (!fix) return null;

  if (fix.kind === "title-rewrite" && fix.newText) return fix.newText;

  if (fix.kind === "ai-rewrite" || fix.kind === "title-rewrite") {
    const title = slide.shapes.find((s) => s.id === fix.shapeId) || slide.shapes.find(
      (s) => s.placeholderType === "title" || s.placeholderType === "ctrTitle"
    );
    const current =
      title?.text.trim() ||
      finding.evidence ||
      "";
    const w = title?.rect ? emuToIn(title.rect.cx) : 8.5;
    return deterministicShortenTitle(current, w);
  }

  if (
    (fix.kind === "text-replace" || fix.kind === "deck-replace") &&
    fix.search &&
    fix.replace != null
  ) {
    const re = buildReplaceRegex(fix);
    const applyReplace = (text: string) => {
      const next = text.replace(re, fix.replace!);
      return fix.regex ? next.replace(/\s{2,}/g, " ").trim() : next;
    };
    if (fix.cell && fix.shapeId) {
      const shape = slide.shapes.find((s) => s.id === fix.shapeId);
      const table = shape?.table;
      if (table) {
        const cell = table.rows[fix.cell.row]?.[fix.cell.col];
        if (cell) return applyReplace(cell.text);
      }
    }
    const shape = fix.shapeId
      ? slide.shapes.find((s) => s.id === fix.shapeId)
      : slide.shapes.find((s) => s.text.includes(fix.search!));
    if (shape) return applyReplace(shape.text);
    return applyReplace(finding.evidence || "");
  }

  return null;
}

function setShapeText(shape: Shape, newText: string) {
  shape.text = newText;
  if (shape.paragraphs.length === 0) {
    shape.paragraphs.push({
      level: 0,
      runs: [
        {
          text: newText,
          font: "Calibri",
          sizePt: 10,
          bold: false,
          italic: false,
          underline: false,
          strike: false,
          color: null,
          explicit: { size: false, bold: false, font: false, color: false },
        },
      ],
      text: newText,
      bulletChar: null,
      bulletType: null,
      bulletFont: null,
      align: null,
      marginLeftIn: null,
      firstLineIndentIn: null,
      indentIn: null,
      spaceBeforePt: null,
      spaceAfterPt: null,
      lineSpacingPct: null,
    });
    return;
  }
  shape.paragraphs[0].text = newText;
  if (shape.paragraphs[0].runs.length > 0) {
    shape.paragraphs[0].runs[0].text = newText;
    for (let i = 1; i < shape.paragraphs[0].runs.length; i++) {
      shape.paragraphs[0].runs[i].text = "";
    }
  }
  for (let i = 1; i < shape.paragraphs.length; i++) {
    shape.paragraphs[i].text = "";
    for (const r of shape.paragraphs[i].runs) r.text = "";
  }
}

function replaceInShape(shape: Shape, re: RegExp, replace: string, regex?: boolean) {
  const applyReplace = (text: string) => {
    const next = text.replace(re, replace);
    return regex ? next.replace(/\s{2,}/g, " ").trim() : next;
  };
  if (shape.kind === "table" && shape.table) {
    for (const row of shape.table.rows) {
      for (const cell of row) {
        if (re.test(cell.text)) {
          cell.text = applyReplace(cell.text);
          if (cell.runs[0]) cell.runs[0].text = cell.text;
        }
      }
    }
    shape.text = shape.table.rows.map((r) => r.map((c) => c.text).join("\t")).join("\n");
    return;
  }
  setShapeText(shape, applyReplace(shape.text));
}

/** Apply a fix to the analyze result for immediate on-screen preview. */
export function applyPreviewPatch(
  result: AnalyzeResult,
  finding: Finding
): AnalyzeResult {
  const fix = finding.fix;
  if (!fix) return result;

  const next: AnalyzeResult = structuredClone(result);
  const slide = next.slides.find((s) => s.index === finding.slideIndex);
  if (!slide) return result;

  // Force HTML preview (drop static ConvertAPI image for this slide).
  slide.imageUrl = null;

  if (fix.kind === "ensure-footer" && fix.newText) {
    let shape =
      slide.shapes.find((s) => s.id === fix.shapeId) ||
      slide.shapes.find((s) => s.placeholderType === "ftr");
    if (!shape && fix.targetRect) {
      shape = {
        id: fix.shapeId || `ftr-${slide.index}`,
        name: "Footer Placeholder",
        kind: "text",
        placeholderType: "ftr",
        placeholderIdx: null,
        rect: fix.targetRect,
        paragraphs: [],
        text: "",
        bodyPr: {
          anchor: "b",
          lIns: 91440,
          tIns: 45720,
          rIns: 91440,
          bIns: 45720,
          wrap: "square",
        },
        table: null,
        imageDataUrl: null,
        fill: null,
      };
      slide.shapes.push(shape);
    }
    if (shape) setShapeText(shape, fix.newText);
    slide.findings = slide.findings.filter((f) => f.id !== finding.id);
    return next;
  }

  if (fix.kind === "ai-rewrite" || fix.kind === "title-rewrite") {
    const newText =
      fix.newText || predictFixText(finding, slide) || finding.evidence || "";
    const shape =
      slide.shapes.find((s) => s.id === fix.shapeId) ||
      slide.shapes.find((s) => s.placeholderType === "ftr") ||
      slide.shapes.find(
        (s) => s.placeholderType === "title" || s.placeholderType === "ctrTitle"
      );
    if (shape) setShapeText(shape, newText);
    slide.findings = slide.findings.filter((f) => f.id !== finding.id);
    return next;
  }

  if (fix.kind === "shape-reposition" && fix.targetRect && fix.shapeId) {
    const shape = slide.shapes.find((s) => s.id === fix.shapeId);
    if (shape?.rect) {
      Object.assign(shape.rect, fix.targetRect);
    }
    slide.findings = slide.findings.filter((f) => f.id !== finding.id);
    return next;
  }

  if (
    fix.kind === "font-family-set" ||
    fix.kind === "font-size-set" ||
    fix.kind === "font-bold-set" ||
    fix.kind === "color-set" ||
    fix.kind === "row-fill-set" ||
    fix.kind === "border-remove"
  ) {
    slide.findings = slide.findings.filter((f) => f.id !== finding.id);
    return next;
  }

  if (
    (fix.kind === "text-replace" || fix.kind === "deck-replace") &&
    fix.search &&
    fix.replace != null
  ) {
    const re = buildReplaceRegex(fix);

    if (fix.deckWide || fix.kind === "deck-replace") {
      for (const s of next.slides) {
        s.imageUrl = null;
        for (const shape of s.shapes) replaceInShape(shape, re, fix.replace, fix.regex);
        s.findings = s.findings.filter((f) => f.ruleId !== finding.ruleId);
      }
      return next;
    }

    const shape = fix.shapeId
      ? slide.shapes.find((s) => s.id === fix.shapeId)
      : undefined;
    if (shape && fix.cell && shape.table) {
      const cell = shape.table.rows[fix.cell.row]?.[fix.cell.col];
      if (cell) {
        const nextText = fix.regex
          ? cell.text.replace(re, fix.replace).replace(/\s{2,}/g, " ").trim()
          : cell.text.replace(re, fix.replace);
        cell.text = nextText;
        if (cell.runs[0]) cell.runs[0].text = cell.text;
      }
    } else if (shape) {
      replaceInShape(shape, re, fix.replace, fix.regex);
    }
  }

  slide.findings = slide.findings.filter((f) => f.id !== finding.id);

  return next;
}

/** Remove a finding from the preview (e.g. after successful title fix). */
export function removeFindingFromPreview(
  result: AnalyzeResult,
  ruleId: string,
  slideIndex: number
): AnalyzeResult {
  const next = structuredClone(result);
  const slide = next.slides.find((s) => s.index === slideIndex);
  if (!slide) return next;
  slide.findings = slide.findings.filter((f) => f.ruleId !== ruleId);
  return next;
}

export { titleMaxChars, deterministicShortenTitle };
