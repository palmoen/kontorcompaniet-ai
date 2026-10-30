import { describe, expect, it } from "vitest";
import { FixedWindowRateLimiter, clientKey, isAllowedOrigin } from "@/lib/requestGuard";

describe("isAllowedOrigin", () => {
  it("krever eksakt treff", () => {
    expect(isAllowedOrigin("https://a.no", ["https://a.no"])).toBe(true);
    expect(isAllowedOrigin("https://a.no.evil.example", ["https://a.no"])).toBe(false);
    expect(isAllowedOrigin(null, ["https://a.no"])).toBe(false);
  });
});

describe("FixedWindowRateLimiter", () => {
  it("nullstiller etter vinduet", () => {
    let now = 0;
    const limiter = new FixedWindowRateLimiter(2, 1000, () => now);
    expect(limiter.take("a")).toBe(true);
    expect(limiter.take("a")).toBe(true);
    expect(limiter.take("a")).toBe(false);
    now = 1001;
    expect(limiter.take("a")).toBe(true);
  });
});

describe("clientKey", () => {
  it("bruker første IP i x-forwarded-for", () => {
    expect(clientKey(new Headers({ "x-forwarded-for": "1.2.3.4, 5.6.7.8" }))).toBe("1.2.3.4");
  });
});
