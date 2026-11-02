import { describe, expect, it, vi } from "vitest";
import { GET, POST } from "@/app/api/ai/reply/route";

describe("/api/ai/reply (faset ut)", () => {
  it("svarer 410 på POST uten noe nettverkskall (OpenAI)", async () => {
    const fetchSpy = vi.spyOn(globalThis, "fetch");
    const response = await POST();
    expect(response.status).toBe(410);
    expect((await response.json()).reply).toMatch(/faset ut/);
    expect(fetchSpy).not.toHaveBeenCalled();
  });

  it("svarer 410 på GET", async () => {
    expect((await GET()).status).toBe(410);
  });
});
