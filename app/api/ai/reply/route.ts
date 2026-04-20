import { NextRequest, NextResponse } from "next/server";
import { openai } from "@/lib/openai";

export const runtime = "nodejs";

export async function POST(req: NextRequest) {
  try {
    const { subject, body, from } = await req.json();

    if (!body) {
      return NextResponse.json({ reply: "Ingen e-postinnhold funnet" });
    }

    const prompt = `
Du er en profesjonell rådgiver i Kontorcompaniet.

Skriv et svar på denne e-posten:

Fra: ${from}
Emne: ${subject}
Innhold:
${body}

KRAV:
- Norsk språk
- Profesjonell men uformell tone
- Maks 6–8 linjer
- Vær konkret
- Foreslå tydelig neste steg
- Relevant for B2B kontormøbler
`;

    const completion = await openai.chat.completions.create({
      model: "gpt-5-mini",
      messages: [
        {
          role: "system",
          content: "Du skriver korte, presise e-postsvar.",
        },
        {
          role: "user",
          content: prompt,
        },
      ],
      temperature: 0.7,
    });

    const reply =
      completion.choices?.[0]?.message?.content ||
      "Kunne ikke generere svar";

    return NextResponse.json({ reply });
  } catch (err: any) {
    return NextResponse.json(
      { reply: "Feil ved generering av svar" },
      { status: 500 }
    );
  }
}