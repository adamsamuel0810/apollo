import { Shape, SlideKind } from "../pptx/types";
import { Finding, Severity, FixRecord } from "../rules/types";

export interface AnalyzeSlide {
  index: number;
  kind: SlideKind;
  size: { cx: number; cy: number };
  shapes: Shape[];
  findings: Finding[];
  /**
   * Pixel-identical slide image (data URL) from the cloud renderer, when
   * configured. When present the UI shows this instead of the HTML render.
   */
  imageUrl?: string | null;
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
  /** Fixes applied during this session (accumulated client-side). */
  fixRecords?: FixRecord[];
}
