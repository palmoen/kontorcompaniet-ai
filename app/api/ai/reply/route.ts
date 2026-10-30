import { NextRequest, NextResponse } from "next/server";
import OpenAI from "openai";
import {
  FixedWindowRateLimiter,
  MAX_BODY_CHARS,
  clientKey,
  isAllowedOrigin,
} from "@/lib/requestGuard";

export const runtime = "nodejs";

// Opprettes først ved første kall — ellers feiler `next build` (og hele ruten) når
// OPENAI_API_KEY mangler, i stedet for å svare 503 under.
let openai: OpenAI | undefined;
function getOpenAI(): OpenAI {
  openai ??= new OpenAI({ apiKey: process.env.OPENAI_API_KEY });
  return openai;
}

/** 20 AI-kall per minutt per klient-IP (per serverinstans, se lib/requestGuard.ts). */
const rateLimiter = new FixedWindowRateLimiter(20, 60_000);

type Intent =
  | "reply"
  | "summary"
  | "tasks"
  | "complaint"
  | "new_email"
  | "tools";

type Tone =
  | "standard"
  | "kort"
  | "mer_formell"
  | "mer_uformell"
  | "mer_salgsrettet";

function normalizeIntent(value: unknown): Intent {
  const safe = String(value || "reply").trim().toLowerCase();

  if (
    safe === "reply" ||
    safe === "summary" ||
    safe === "tasks" ||
    safe === "complaint" ||
    safe === "new_email" ||
    safe === "tools"
  ) {
    return safe as Intent;
  }

  return "reply";
}

function normalizeTone(value: unknown): Tone {
  const safe = String(value || "standard").trim().toLowerCase();

  if (
    safe === "standard" ||
    safe === "kort" ||
    safe === "mer_formell" ||
    safe === "mer_uformell" ||
    safe === "mer_salgsrettet"
  ) {
    return safe as Tone;
  }

  return "standard";
}

function getToneInstruction(tone: Tone) {
  switch (tone) {
    case "kort":
      return "Hold svaret ekstra kort og effektivt.";
    case "mer_formell":
      return "Bruk en mer formell og profesjonell tone.";
    case "mer_uformell":
      return "Bruk en mer uformell, varm og naturlig tone.";
    case "mer_salgsrettet":
      return "Bruk en mer salgsrettet tone, men fortsatt profesjonell og troverdig.";
    case "standard":
    default:
      return "Bruk en profesjonell, men uformell tone.";
  }
}

function buildLegacyContext(payload: {
  subject?: string;
  body?: string;
  from?: string;
}) {
  return `
Fra: ${payload.from || ""}
Emne: ${payload.subject || ""}
Innhold:
${payload.body || ""}
  `.trim();
}

function buildPrompt({
  intent,
  tone,
  context,
  instruction,
  previousDraft,
}: {
  intent: Intent;
  tone: Tone;
  context: string;
  instruction?: string;
  previousDraft?: string;
}) {
  const toneInstruction = getToneInstruction(tone);

  const baseRules = `
Du er KontorPilot AI for Kontorcompaniet.

Skriv alltid på norsk.

Felles krav:
- Profesjonell, men uformell stil, med mindre tonevalget sier noe annet
- Konkret og ryddig
- Maks 6–8 linjer når output er e-posttekst
- Foreslå neste steg når det er naturlig
- Relevans for B2B, kontormøbler, interiør, leveranser, tilbud og reklamasjoner
- Ingen unødvendig fluff
- Ikke bruk emoji
- Ikke skriv forklaringer om hva du gjør
- Returner kun selve resultatet
- Tonevalg: ${toneInstruction}
  `.trim();

  const intentRulesMap: Record<Intent, string> = {
    reply: `
Oppgave:
- Lag et konkret svarutkast til e-posten
- Svaret skal kunne limes rett inn i Outlook
- Hvis previousDraft finnes og instruction ber om justering, skriv en forbedret versjon av draftet
    `.trim(),

    summary: `
Oppgave:
- Oppsummer tråden i 3–5 korte bullets
- Ta med:
  1. hva saken gjelder
  2. hva som er gjort
  3. hva som mangler / neste steg
- Kun bullets, ingen innledning
    `.trim(),

    tasks: `
Oppgave:
- Foreslå konkrete oppgaver, avtaler og oppfølging
- Bruk korte bullets
- Skill tydelig mellom:
  - Oppgaver
  - Oppfølging
  - Eventuelle avtaler
- Kun output, ingen forklaring
    `.trim(),

    complaint: `
Oppgave:
- Lag en ny reklamasjonsmail til leverandør
- Tonen skal være tydelig, profesjonell og saklig
- Ta med relevant informasjon fra tråden
- Be om passende neste steg, vurdering eller løsning
- E-posten skal være klar til sending
    `.trim(),

    new_email: `
Oppgave:
- Lag en helt ny e-post basert på brukerens prompt og eventuell e-postkontekst
- E-posten skal være klar til sending
- Hold den konkret og profesjonell
    `.trim(),

    tools: `
Oppgave:
- Foreslå smarte handlinger for e-posten
- Returner korte bullets for:
  - Flagg: Ja/Nei + kort grunn
  - Pin: Ja/Nei + kort grunn
  - Foreslått mappe: navn + kort grunn
  - Neste steg: kort anbefaling
- Vær konkret
    `.trim(),
  };

  const draftSection = previousDraft?.trim()
    ? `\n\nEksisterende utkast:\n${previousDraft.trim()}`
    : "";

  const instructionSection = instruction?.trim()
    ? `\n\nEkstra instruksjon:\n${instruction.trim()}`
    : "";

  return `
${baseRules}

${intentRulesMap[intent]}

Kontekst:
${context.trim()}${draftSection}${instructionSection}
  `.trim();
}

export async function POST(req: NextRequest) {
  if (!isAllowedOrigin(req.headers.get("origin"))) {
    return NextResponse.json({ error: "Ikke tillatt." }, { status: 403 });
  }

  if (!rateLimiter.take(clientKey(req.headers))) {
    return NextResponse.json(
      { error: "For mange forespørsler. Prøv igjen om litt." },
      { status: 429, headers: { "Retry-After": "60" } }
    );
  }

  if (!process.env.OPENAI_API_KEY) {
    console.error("AI reply: OPENAI_API_KEY er ikke satt.");
    return NextResponse.json(
      { error: "Tjenesten er ikke konfigurert.", reply: "Feil ved generering av svar." },
      { status: 503 }
    );
  }

  const rawBody = await req.text();
  if (rawBody.length > MAX_BODY_CHARS) {
    return NextResponse.json(
      { error: "E-posten er for lang.", reply: "E-posten er for lang til å behandles." },
      { status: 413 }
    );
  }

  try {
    let payload: Record<string, unknown>;
    try {
      payload = JSON.parse(rawBody) as Record<string, unknown>;
    } catch {
      return NextResponse.json({ reply: "Ugyldig forespørsel." }, { status: 400 });
    }

    const legacyBody = String(payload?.body || "").trim();
    const context = String(payload?.context || "").trim();
    const intent = normalizeIntent(payload?.intent);
    const tone = normalizeTone(payload?.tone);
    const instruction = String(payload?.instruction || "").trim();
    const previousDraft = String(payload?.previousDraft || "").trim();

    const finalContext = context || buildLegacyContext(payload);

    if (!finalContext || (!context && !legacyBody)) {
      return NextResponse.json(
        { reply: "Ingen e-postinnhold funnet." },
        { status: 400 }
      );
    }

    const prompt = buildPrompt({
      intent,
      tone,
      context: finalContext,
      instruction,
      previousDraft,
    });

    const completion = await getOpenAI().chat.completions.create({
      model: "gpt-5-mini",
      messages: [
        {
          role: "system",
          content:
            "Du skriver korte, presise og handlingsorienterte tekster på norsk for Kontorcompaniet.",
        },
        {
          role: "user",
          content: prompt,
        },
      ],
    });

    const output =
      completion.choices?.[0]?.message?.content?.trim() ||
      "Kunne ikke generere svar.";

    return NextResponse.json({
      reply: output,
      output,
      intent,
      tone,
    });
  } catch (error) {
    // Logg detaljene på serveren, men send aldri interne feilmeldinger (som kan
    // inneholde detaljer om OpenAI-kontoen/oppsettet) tilbake til klienten.
    console.error("AI reply error:", error);

    return NextResponse.json(
      {
        error: "Ukjent serverfeil",
        reply: "Feil ved generering av svar.",
      },
      { status: 500 }
    );
  }
}
