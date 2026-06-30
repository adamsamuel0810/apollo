import { AnalyzeSlide } from "../analyze/types";
import { Finding, FixInstruction } from "../rules/types";
import { CONFIDENTIALITY } from "../brand/guidelines";

/** Attach a fix instruction on the client for instant preview when the rule omitted it. */
export function enrichFindingForPreview(
  finding: Finding,
  slide: AnalyzeSlide
): Finding {
  if (finding.fix) return finding;

  const inferred = inferClientFix(finding, slide);
  if (!inferred) return finding;
  return { ...finding, fix: inferred };
}

function inferClientFix(
  finding: Finding,
  slide: AnalyzeSlide
): FixInstruction | null {
  const quoted = extractQuoted(finding.suggestion);

  switch (finding.ruleId) {
    case "confidentiality-missing":
    case "confidentiality-modified": {
      const text = quoted || CONFIDENTIALITY.canonical;
      const ftr =
        slide.shapes.find((s) => s.id === finding.shapeId) ||
        slide.shapes.find((s) => s.placeholderType === "ftr") ||
        slide.shapes.find((s) => s.name.toLowerCase().includes("footer"));
      if (finding.ruleId === "confidentiality-missing" && finding.rect) {
        return {
          kind: "ensure-footer",
          slideIndex: finding.slideIndex,
          shapeId: ftr?.id,
          newText: text,
          targetRect: finding.rect,
        };
      }
      if (!ftr) return null;
      return {
        kind: "title-rewrite",
        slideIndex: finding.slideIndex,
        shapeId: ftr.id,
        newText: text,
      };
    }

    case "draft-designator":
      if (!finding.shapeId) return null;
      return {
        kind: "text-replace",
        slideIndex: finding.slideIndex,
        shapeId: finding.shapeId,
        search: "\\b(DRAFT|WIP|INTERNAL ONLY|DO NOT DISTRIBUTE|PRELIMINARY)\\b\\s*",
        replace: "",
        regex: true,
        caseInsensitive: true,
      };

    case "term-percentile":
      return {
        kind: "deck-replace",
        slideIndex: finding.slideIndex,
        search: "Percentile",
        replace: "%ile",
        caseInsensitive: true,
        deckWide: true,
      };

    case "term-tgt":
      return {
        kind: "deck-replace",
        slideIndex: finding.slideIndex,
        search: "TGT",
        replace: "Target",
        deckWide: true,
      };

    case "term-incumbent":
      if (!finding.shapeId) return null;
      return {
        kind: "text-replace",
        slideIndex: finding.slideIndex,
        shapeId: finding.shapeId,
        cell: finding.cell,
        search: "Incumbent",
        replace: "Executive",
        caseInsensitive: true,
      };

    case "term-company-name":
      if (!finding.shapeId) return null;
      return {
        kind: "text-replace",
        slideIndex: finding.slideIndex,
        shapeId: finding.shapeId,
        cell: finding.cell,
        search: "Company Name",
        replace: "Company",
        caseInsensitive: true,
      };

    case "title-punct": {
      const title = slide.shapes.find(
        (s) => s.placeholderType === "title" || s.placeholderType === "ctrTitle"
      );
      const tText = title?.text.trim() || finding.evidence || "";
      if (!title || !tText) return null;
      return {
        kind: "title-rewrite",
        slideIndex: finding.slideIndex,
        shapeId: title.id,
        newText: tText.replace(/[.,;]+$/, ""),
      };
    }

    default:
      if (quoted && finding.shapeId) {
        return {
          kind: "text-replace",
          slideIndex: finding.slideIndex,
          shapeId: finding.shapeId,
          cell: finding.cell,
          search: finding.evidence || "",
          replace: quoted,
        };
      }
      return null;
  }
}

function extractQuoted(text?: string): string | null {
  if (!text) return null;
  const m = text.match(/"([^"]+)"/);
  return m?.[1] ?? null;
}
