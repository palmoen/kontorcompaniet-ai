import { NextResponse } from "next/server";

export const runtime = "nodejs";

/**
 * K-Pilot er faset ut — AI-hjelp til e-post er flyttet til Donna (inbox-pilot).
 *
 * Endepunktet var åpent uten innlogging, slik at hvem som helst kunne bruke
 * OpenAI-kontoen via det. Det svarer derfor nå alltid 410 Gone, uten å lese
 * forespørselen og uten å kontakte OpenAI. Se docs/SIKKERHET.md.
 */
function gone(): NextResponse {
  return NextResponse.json(
    {
      error: "K-Pilot er faset ut. Bruk Donna i stedet.",
      reply: "K-Pilot er faset ut. Bruk Donna i stedet.",
    },
    { status: 410 }
  );
}

export const GET = gone;
export const POST = gone;
