import { Shape, SlideKind } from "../pptx/types";
import { Finding, Severity } from "../rules/types";

export interface AnalyzeSlide {
  index: number;
  kind: SlideKind;
  size: { cx: number; cy: number };
  shapes: Shape[];
  findings: Finding[];
}

export interface AnalyzeSummary {
  total: number;
  bySeverity: Record<Severity, number>;
  slidesWithFindings: number;
}

export interface AnalyzeResult {
  fileName: string;
  slideCount: number;
  generatedAt: string;
  aiEnabled: boolean;
  summary: AnalyzeSummary;
  slides: AnalyzeSlide[];
}
