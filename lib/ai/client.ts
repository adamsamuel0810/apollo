import Anthropic from "@anthropic-ai/sdk";

let client: Anthropic | null = null;

export function getAnthropic(): Anthropic | null {
  if (process.env.ENABLE_AI === "0") return null;
  const key = process.env.ANTHROPIC_API_KEY;
  if (!key) return null;
  if (!client) client = new Anthropic({ apiKey: key });
  return client;
}

export function getModel(): string {
  return process.env.ANTHROPIC_MODEL || "claude-3-5-sonnet-latest";
}

export function aiEnabled(): boolean {
  return getAnthropic() != null;
}
