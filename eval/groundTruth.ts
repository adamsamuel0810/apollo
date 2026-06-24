// Ground truth reverse-engineered from the speaker notes in
// "Goofy Corp Peer Group Assessment - ANNOTATED.pptx".
//
// Each tag represents a distinct issue a reviewer flagged on that slide.
// Tags are split by which layer is expected to catch them:
//  - deterministic: precise, code-checkable issues
//  - ai: semantic/subjective issues for the AI layer
//  - hard: issues that are out of scope for an automated v1 (documented)

export type Layer = "deterministic" | "ai" | "hard";

export interface GtTag {
  tag: string;
  layer: Layer;
}

export const GROUND_TRUTH: Record<number, GtTag[]> = {
  1: [{ tag: "draft", layer: "deterministic" }],
  2: [{ tag: "bullets-parallel", layer: "ai" }],
  3: [
    { tag: "percentile", layer: "deterministic" },
    { tag: "table-title-centered", layer: "deterministic" },
    { tag: "title-moved", layer: "deterministic" },
    { tag: "client-name", layer: "ai" },
    { tag: "bullets-parallel", layer: "ai" },
    { tag: "legend-spacing", layer: "hard" },
    { tag: "legend-box", layer: "hard" },
    { tag: "first-col-margin", layer: "hard" },
    { tag: "title-lines", layer: "ai" },
  ],
  4: [
    { tag: "font-consistency", layer: "deterministic" },
    { tag: "title-moved", layer: "deterministic" },
  ],
  5: [
    { tag: "percentile", layer: "deterministic" },
    { tag: "legend-order", layer: "ai" },
    { tag: "first-col-margin", layer: "hard" },
  ],
  6: [
    { tag: "merged-cells", layer: "deterministic" },
    { tag: "first-col-margin", layer: "hard" },
  ],
  9: [{ tag: "title-lines", layer: "ai" }],
  10: [
    { tag: "confidentiality-missing", layer: "deterministic" },
    { tag: "legend-naming", layer: "ai" },
    { tag: "legend-spacing", layer: "hard" },
  ],
  11: [
    { tag: "legend-order", layer: "ai" },
    { tag: "missing-units", layer: "ai" },
    { tag: "propernoun-punct", layer: "ai" },
  ],
  12: [
    { tag: "confidentiality-overlap", layer: "deterministic" },
    { tag: "company-name", layer: "deterministic" },
  ],
  13: [
    { tag: "confidentiality-missing", layer: "deterministic" },
    { tag: "statsrow-color", layer: "deterministic" },
    { tag: "legend-naming", layer: "ai" },
  ],
  14: [
    { tag: "confidentiality-missing", layer: "deterministic" },
    { tag: "statsrow-color", layer: "deterministic" },
    { tag: "legend-naming", layer: "ai" },
    { tag: "missing-labels", layer: "hard" },
  ],
  23: [{ tag: "confidentiality-missing", layer: "deterministic" }],
  24: [{ tag: "confidentiality-moved", layer: "deterministic" }],
};

// Map a finding's ruleId to a ground-truth tag.
export const RULE_TO_TAG: Record<string, string> = {
  "draft-designator": "draft",
  "term-percentile": "percentile",
  "term-company-name": "company-name",
  "term-incumbent": "incumbent",
  "table-font-consistency": "font-consistency",
  "table-statsrow-color": "statsrow-color",
  "table-merged-body": "merged-cells",
  "table-title-centered": "table-title-centered",
  "confidentiality-missing": "confidentiality-missing",
  "confidentiality-modified": "confidentiality-modified",
  "confidentiality-moved": "confidentiality-moved",
  "confidentiality-overlap": "confidentiality-overlap",
  "title-moved": "title-moved",
};

// AI findings map by category to a representative tag for credit in eval.
export const AI_CATEGORY_TO_TAGS: Record<string, string[]> = {
  bullets: ["bullets-parallel"],
  title: ["title-lines"],
  consistency: ["client-name", "legend-order", "legend-naming"],
  terminology: ["company-name", "legend-naming"],
  table: ["missing-units"],
  content: ["propernoun-punct", "missing-units"],
  footnote: ["propernoun-punct"],
};
