import { beforeEach, describe, expect, it, vi } from "vitest";
import { NextRequest } from "next/server";

const createCompletion = vi.fn();

vi.mock("openai", () => ({
  default: class {
    chat = { completions: { create: createCompletion } };
  },
}));

const ALLOWED_ORIGIN = "https://kontorcompaniet-ai.vercel.app";

function request(options: { origin?: string | null; body?: string; ip?: string } = {}) {
  const headers = new Headers({ "content-type": "application/json" });
  if (options.origin !== null) headers.set("origin", options.origin ?? ALLOWED_ORIGIN);
  headers.set("x-forwarded-for", options.ip ?? "10.0.0.1");
  return new NextRequest("https://kontorcompaniet-ai.vercel.app/api/ai/reply", {
    method: "POST",
    headers,
    body: options.body ?? JSON.stringify({ context: "Hei, kan dere levere i morgen?" }),
  });
}

async function loadRoute() {
  vi.resetModules();
  return import("@/app/api/ai/reply/route");
}

describe("POST /api/ai/reply", () => {
  beforeEach(() => {
    process.env.OPENAI_API_KEY = "test-key";
    delete process.env.KPILOT_ALLOWED_ORIGINS;
    createCompletion.mockReset();
    createCompletion.mockResolvedValue({ choices: [{ message: { content: "Utkast" } }] });
  });

  it("avviser forespørsler uten Origin-header (f.eks. direkte curl-kall)", async () => {
    const { POST } = await loadRoute();
    const response = await POST(request({ origin: null }));
    expect(response.status).toBe(403);
    expect(createCompletion).not.toHaveBeenCalled();
  });

  it("avviser forespørsler fra et annet nettsted", async () => {
    const { POST } = await loadRoute();
    const response = await POST(request({ origin: "https://evil.example" }));
    expect(response.status).toBe(403);
    expect(createCompletion).not.toHaveBeenCalled();
  });

  it("lar taskpane-en på eget domene generere et svar", async () => {
    const { POST } = await loadRoute();
    const response = await POST(request());
    expect(response.status).toBe(200);
    expect((await response.json()).reply).toBe("Utkast");
    expect(createCompletion).toHaveBeenCalledTimes(1);
  });

  it("respekterer KPILOT_ALLOWED_ORIGINS", async () => {
    process.env.KPILOT_ALLOWED_ORIGINS = "https://kontorcompaniet.no";
    const { POST } = await loadRoute();
    expect((await POST(request())).status).toBe(403);
    expect((await POST(request({ origin: "https://kontorcompaniet.no" }))).status).toBe(200);
  });

  it("rate-limiter etter 20 kall per minutt fra samme IP", async () => {
    const { POST } = await loadRoute();
    for (let i = 0; i < 20; i += 1) {
      expect((await POST(request({ ip: "10.0.0.9" }))).status).toBe(200);
    }
    const limited = await POST(request({ ip: "10.0.0.9" }));
    expect(limited.status).toBe(429);
    expect((await POST(request({ ip: "10.0.0.10" }))).status).toBe(200);
  });

  it("avviser for store forespørsler uten å kalle OpenAI", async () => {
    const { POST } = await loadRoute();
    const response = await POST(request({ body: JSON.stringify({ context: "x".repeat(70_000) }) }));
    expect(response.status).toBe(413);
    expect(createCompletion).not.toHaveBeenCalled();
  });

  it("avviser ugyldig JSON", async () => {
    const { POST } = await loadRoute();
    expect((await POST(request({ body: "{ikke json" }))).status).toBe(400);
  });

  it("lekker ikke interne feilmeldinger til klienten", async () => {
    createCompletion.mockRejectedValue(new Error("Incorrect API key provided: sk-hemmelig"));
    vi.spyOn(console, "error").mockImplementation(() => {});
    const { POST } = await loadRoute();
    const response = await POST(request());
    expect(response.status).toBe(500);
    expect(JSON.stringify(await response.json())).not.toContain("sk-hemmelig");
  });

  it("svarer 503 når OPENAI_API_KEY mangler", async () => {
    delete process.env.OPENAI_API_KEY;
    vi.spyOn(console, "error").mockImplementation(() => {});
    const { POST } = await loadRoute();
    expect((await POST(request())).status).toBe(503);
  });
});
