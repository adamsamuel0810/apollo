import { DeckModel } from "../pptx/types";
import { Finding, FixInstruction, DeckAggregates } from "../rules/types";
import { nearestPaletteColor } from "../brand/palette";
import { TITLE, TYPEFACE } from "../brand/guidelines";
import { titleText, estimateLines, emuToIn } from "../rules/util";

/** Infer a fix for any error/warning finding not already covered. */
export function inferFix(
  finding: Finding,
  deck: DeckModel,
  aggregates: DeckAggregates
): FixInstruction | null {
  const slide = deck.slides.find((s) => s.index === finding.slideIndex);
  if (!slide) return null;

  switch (finding.ruleId) {
    case "draft-designator":
      if (!finding.shapeId || !finding.evidence) return null;
      return {
        kind: "text-replace",
        slideIndex: finding.slideIndex,
        shapeId: finding.shapeId,
        search: "\\b(DRAFT|WIP|INTERNAL ONLY|DO NOT DISTRIBUTE|PRELIMINARY)\\b\\s*",
        replace: "",
        regex: true,
        caseInsensitive: true,
      };

    case "confidentiality-modified":
      if (!finding.shapeId || !aggregates.confidentialityText) return null;
      return {
        kind: "title-rewrite",
        slideIndex: finding.slideIndex,
        shapeId: finding.shapeId,
        newText: aggregates.confidentialityText,
      };

    case "confidentiality-missing": {
      if (!aggregates.confidentialityText || !aggregates.footerRect) return null;
      const ftr =
        slide.shapes.find((s) => s.placeholderType === "ftr") ||
        slide.shapes.find((s) => s.name.toLowerCase().includes("footer"));
      return {
        kind: "ensure-footer",
        slideIndex: finding.slideIndex,
        shapeId: ftr?.id,
        newText: aggregates.confidentialityText,
        targetRect: aggregates.footerRect,
      };
    }

    case "confidentiality-moved":
      if (!finding.shapeId || !aggregates.footerRect) return null;
      return {
        kind: "shape-reposition",
        slideIndex: finding.slideIndex,
        shapeId: finding.shapeId,
        targetRect: aggregates.footerRect,
      };

    case "font-family":
      if (!finding.shapeId) return null;
      return {
        kind: "font-family-set",
        slideIndex: finding.slideIndex,
        shapeId: finding.shapeId,
        fontFamily: TYPEFACE,
      };

    case "title-size":
      if (!finding.shapeId) return null;
      return {
        kind: "font-size-set",
        slideIndex: finding.slideIndex,
        shapeId: finding.shapeId,
        fontSizePt: TITLE.sizePt,
      };

    case "title-bold":
      if (!finding.shapeId) return null;
      return {
        kind: "font-bold-set",
        slideIndex: finding.slideIndex,
        shapeId: finding.shapeId,
      };

    case "title-punct": {
      const title = slide.titleShape;
      const tText = titleText(slide);
      if (!title || !tText) return null;
      return {
        kind: "title-rewrite",
        slideIndex: finding.slideIndex,
        shapeId: title.id,
        newText: tText.trim().replace(/[.,;]+$/, ""),
      };
    }

    case "title-lines": {
      const title = slide.titleShape;
      const tText = titleText(slide);
      if (!title || !tText) return null;
      const w = title.rect ? emuToIn(title.rect.cx) : 8.5;
      const lines = estimateLines(tText, TITLE.sizePt, w);
      return {
        kind: "ai-rewrite",
        slideIndex: finding.slideIndex,
        shapeId: title.id,
        aiContext: `Shorten to ${TITLE.maxLines} lines at 24pt in ${w.toFixed(1)}" box. Currently ~${lines} lines.`,
      };
    }

    case "off-palette-color": {
      if (!finding.shapeId || !finding.evidence) return null;
      const hex = finding.evidence.replace(/^#/, "").toUpperCase();
      const near = nearestPaletteColor(hex);
      if (!near) return null;
      const isFill = finding.detail?.toLowerCase().includes("fill");
      return {
        kind: "color-set",
        slideIndex: finding.slideIndex,
        shapeId: finding.shapeId,
        fromColor: hex,
        toColor: near.color,
        colorTarget: isFill ? "fill" : "text",
        cell: finding.cell,
      };
    }

    case "table-font-consistency": {
      if (!finding.shapeId || !finding.cell) return null;
      const ptMatch = finding.suggestion?.match(/(\d+(?:\.\d+)?)\s*pt/i);
      const dominant = ptMatch ? parseFloat(ptMatch[1]) : null;
      if (!dominant) return null;
      return {
        kind: "font-size-set",
        slideIndex: finding.slideIndex,
        shapeId: finding.shapeId,
        cell: finding.cell,
        fontSizePt: dominant,
      };
    }

    case "table-side-border":
      if (!finding.shapeId) return null;
      return {
        kind: "border-remove",
        slideIndex: finding.slideIndex,
        shapeId: finding.shapeId,
        borderScope: "table-sides",
      };

    case "table-statsrow-color":
      if (!finding.shapeId || finding.cell == null) return null;
      return {
        kind: "row-fill-set",
        slideIndex: finding.slideIndex,
        shapeId: finding.shapeId,
        rowIndex: finding.cell.row,
        fillColor: "FFFFDB",
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

    case "bullet-punct":
      if (!finding.shapeId) return null;
      return {
        kind: "text-replace",
        slideIndex: finding.slideIndex,
        shapeId: finding.shapeId,
        search: "[.;]\\s*$",
        replace: "",
        regex: true,
      };

    default:
      // AI semantic findings: attempt text fix from evidence + suggestion
      if (finding.source === "ai" && finding.evidence && finding.suggestion) {
        return inferAiTextFix(finding, aggregates);
      }
      return null;
  }
}

function inferAiTextFix(
  finding: Finding,
  aggregates: DeckAggregates
): FixInstruction | null {
  const replaceMatch = finding.suggestion?.match(
    /(?:replace|use|rename to|change to|correct to)\s+["']([^"']+)["']/i
  );
  const searchMatch = finding.suggestion?.match(
    /(?:replace|change)\s+["']([^"']+)["']/i
  );

  if (replaceMatch && searchMatch && finding.shapeId) {
    return {
      kind: "text-replace",
      slideIndex: finding.slideIndex,
      shapeId: finding.shapeId,
      cell: finding.cell,
      search: searchMatch[1],
      replace: replaceMatch[1],
    };
  }

  // Client name consistency
  if (
    finding.ruleId.includes("client") &&
    aggregates.clientName.dominant &&
    finding.evidence &&
    finding.shapeId
  ) {
    const dominant = aggregates.clientName.dominant;
    if (finding.evidence.includes(dominant)) return null;
    return {
      kind: "text-replace",
      slideIndex: finding.slideIndex,
      shapeId: finding.shapeId,
      cell: finding.cell,
      search: finding.evidence.slice(0, 80),
      replace: dominant,
    };
  }

  // Proper noun punctuation (e.g. C.R Bard -> C.R. Bard)
  if (finding.evidence && finding.suggestion && finding.shapeId) {
    const quoted = finding.suggestion.match(/["']([^"']+)["']/);
    if (quoted && quoted[1] !== finding.evidence) {
      return {
        kind: "text-replace",
        slideIndex: finding.slideIndex,
        shapeId: finding.shapeId,
        cell: finding.cell,
        search: finding.evidence,
        replace: quoted[1],
      };
    }
  }

  return {
    kind: "ai-rewrite",
    slideIndex: finding.slideIndex,
    shapeId: finding.shapeId,
    aiContext: `${finding.detail}. Suggestion: ${finding.suggestion || ""}. Evidence: "${finding.evidence}"`,
  };
}
