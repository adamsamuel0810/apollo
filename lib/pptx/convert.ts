/**
 * High-fidelity slide rendering via a cloud conversion API (ConvertAPI).
 *
 * PowerPoint/LibreOffice cannot run on Vercel's serverless runtime, so to show
 * slides that are pixel-identical to the original deck we offload rasterization
 * to ConvertAPI: PPTX -> one PNG per slide. The images are returned as data
 * URLs and used as the slide background; highlight overlays are positioned on
 * top using the EMU geometry we already parse.
 *
 * If no API secret is configured the functions return null and the app falls
 * back to the built-in HTML renderer (free, private, but not pixel-perfect).
 */

export interface SlideImage {
  /** 1-based slide index. */
  index: number;
  /** PNG as a data URL, ready to drop into an <img src>. */
  dataUrl: string;
}

interface ConvertApiFile {
  FileName: string;
  FileExt: string;
  FileData: string; // base64
}
interface ConvertApiResponse {
  Files?: ConvertApiFile[];
}

export function imageRenderingEnabled(): boolean {
  return !!(process.env.CONVERTAPI_TOKEN || process.env.CONVERTAPI_SECRET);
}

/** Target width in px for crisp (retina) slide images; height auto-scales. */
const IMAGE_WIDTH = 1600;

/**
 * Convert a PPTX into one PNG per slide. Returns null when image rendering is
 * not configured or the conversion fails, so callers can fall back gracefully.
 *
 * Auth accepts either a modern API token (CONVERTAPI_TOKEN, sent as a Bearer
 * header) or the legacy secret (CONVERTAPI_SECRET, sent as a query param).
 */
export async function convertPptxToImages(
  data: Buffer | ArrayBuffer | Uint8Array,
  fileName: string
): Promise<SlideImage[] | null> {
  const token = process.env.CONVERTAPI_TOKEN;
  const secret = process.env.CONVERTAPI_SECRET;
  if (!token && !secret) return null;

  const bytes =
    data instanceof Buffer
      ? data
      : Buffer.from(data instanceof Uint8Array ? data : new Uint8Array(data));
  const base64 = bytes.toString("base64");
  const safeName = fileName.toLowerCase().endsWith(".pptx")
    ? fileName
    : `${fileName}.pptx`;

  const endpoint = token
    ? "https://v2.convertapi.com/convert/pptx/to/png"
    : `https://v2.convertapi.com/convert/pptx/to/png?Secret=${encodeURIComponent(
        secret!
      )}`;

  const headers: Record<string, string> = {
    "Content-Type": "application/json",
  };
  if (token) headers["Authorization"] = `Bearer ${token}`;

  const body = {
    Parameters: [
      { Name: "File", FileValue: { Name: safeName, Data: base64 } },
      { Name: "ScaleImage", Value: true },
      { Name: "ScaleProportions", Value: true },
      { Name: "ImageWidth", Value: IMAGE_WIDTH },
      { Name: "ImageInterpolation", Value: true },
      { Name: "StoreFile", Value: false },
    ],
  };

  // Conversion of a multi-slide deck can take a while; cap it well under the
  // route's maxDuration so we fail over to HTML instead of timing out hard.
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 45_000);

  try {
    const res = await fetch(endpoint, {
      method: "POST",
      headers,
      body: JSON.stringify(body),
      signal: controller.signal,
    });

    if (!res.ok) {
      console.error(
        `ConvertAPI returned ${res.status}: ${await res.text().catch(() => "")}`
      );
      return null;
    }

    const json = (await res.json()) as ConvertApiResponse;
    const files = json.Files;
    if (!files || files.length === 0) {
      console.error("ConvertAPI returned no files");
      return null;
    }

    // Files come back in slide order; map each to a 1-based slide index.
    return files
      .filter((f) => f.FileData)
      .map((f, i) => ({
        index: i + 1,
        dataUrl: `data:image/png;base64,${f.FileData}`,
      }));
  } catch (err) {
    console.error("ConvertAPI conversion failed:", err);
    return null;
  } finally {
    clearTimeout(timeout);
  }
}
