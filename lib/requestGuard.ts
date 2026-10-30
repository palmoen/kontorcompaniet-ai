/**
 * Beskyttelse for /api/ai/reply inntil ekte innlogging (Office SSO) er på plass.
 *
 * VIKTIG: Dette er IKKE autentisering. Origin-sjekken stopper nettlesere på andre
 * nettsteder fra å bruke endepunktet, og rate limiting/størrelsesgrense begrenser hvor
 * mye OpenAI-kostnad én klient kan påføre — men en angriper som kaller endepunktet
 * direkte (curl e.l.) kan fortsatt forfalske Origin-headeren. Se docs/SIKKERHET.md for
 * forslaget om Office SSO, som er den egentlige løsningen.
 */

const DEFAULT_ALLOWED_ORIGINS = ["https://kontorcompaniet-ai.vercel.app"];

/** Maks antall tegn i hele forespørselskroppen — en lang e-posttråd er godt under dette. */
export const MAX_BODY_CHARS = 60_000;

export function allowedOrigins(): string[] {
  const fromEnv = process.env.KPILOT_ALLOWED_ORIGINS;
  if (!fromEnv) return DEFAULT_ALLOWED_ORIGINS;
  return fromEnv
    .split(",")
    .map((origin) => origin.trim())
    .filter(Boolean);
}

/**
 * Taskpane-en kaller endepunktet med en relativ URL fra samme domene, og nettleseren
 * (WebView2/Edge, Safari/WKWebView, nettleser-Outlook) sender da alltid en Origin-header
 * på en POST. Mangler den, eller peker den et annet sted, avvises forespørselen.
 */
export function isAllowedOrigin(origin: string | null, origins: string[] = allowedOrigins()): boolean {
  if (!origin) return false;
  return origins.includes(origin);
}

interface Window {
  count: number;
  resetAt: number;
}

/**
 * Enkel fast-vindu-teller i minnet. På Vercel har hver serverinstans sin egen teller, så
 * grensen er "per instans" — godt nok til å stoppe en enkelt løpsk klient eller et
 * naivt misbruk, ikke en garanti. En delt teller (f.eks. Upstash/Redis) er neste steg.
 */
export class FixedWindowRateLimiter {
  private readonly windows = new Map<string, Window>();

  constructor(
    private readonly limit: number,
    private readonly windowMs: number,
    private readonly now: () => number = Date.now,
  ) {}

  /** Returnerer `true` hvis forespørselen er innenfor grensen (og teller den). */
  take(key: string): boolean {
    const now = this.now();
    const current = this.windows.get(key);
    if (!current || current.resetAt <= now) {
      this.windows.set(key, { count: 1, resetAt: now + this.windowMs });
      this.pruneExpired(now);
      return true;
    }
    if (current.count >= this.limit) return false;
    current.count += 1;
    return true;
  }

  private pruneExpired(now: number): void {
    if (this.windows.size < 1000) return;
    for (const [key, window] of this.windows) {
      if (window.resetAt <= now) this.windows.delete(key);
    }
  }
}

/** Første IP i x-forwarded-for (satt av Vercel), ellers en felles nøkkel. */
export function clientKey(headers: Headers): string {
  const forwarded = headers.get("x-forwarded-for");
  const first = forwarded?.split(",")[0]?.trim();
  return first || headers.get("x-real-ip") || "unknown";
}
