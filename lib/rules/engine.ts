import { DeckModel, Slide } from "../pptx/types";
import { Finding, RuleContext, SlideRule } from "./types";
import { computeAggregates } from "../deck/aggregate";

import { typographyRules } from "./typography";
import { draftDesignatorRule, titleSlideMetaRule } from "./titleSlide";
import { confidentialityRule } from "./confidentiality";
import { terminologyRules } from "./terminology";
import { tableRules } from "./tables";
import { colorRules } from "./color";
import { titlePositionRule } from "./titlePosition";
import { bulletRules } from "./bullets";

const SLIDE_RULES: SlideRule[] = [
  typographyRules,
  draftDesignatorRule,
  titleSlideMetaRule,
  confidentialityRule,
  terminologyRules,
  tableRules,
  colorRules,
  titlePositionRule,
  bulletRules,
];

export function runDeterministicRules(deck: DeckModel): {
  findings: Finding[];
  ctx: RuleContext;
} {
  const aggregates = computeAggregates(deck);
  const ctx: RuleContext = { deck, aggregates };

  const findings: Finding[] = [];
  for (const slide of deck.slides) {
    for (const rule of SLIDE_RULES) {
      try {
        findings.push(...rule(slide, ctx));
      } catch (err) {
        // A single failing rule should never break the whole analysis.
        console.error(`Rule failed on slide ${slide.index}:`, err);
      }
    }
  }

  return { findings, ctx };
}
