import { NextRequest, NextResponse } from "next/server";
import OpenAI from "openai";

export const runtime = "nodejs";

const openai = new OpenAI({
  apiKey: process.env.OPENAI_API_KEY,
});

export async function POST(req: NextRequest) {
  try {
    const { subject, body, from } = await req.json();

    if (!body) {
      return NextResponse.json({ reply: "Ingen e-postinnhold funnet" });
    }

    const prompt = `
Du er en profesjonell rådgiver i Kontorcompaniet.

Skriv et kort svar på e-posten under.

Fra: ${from}
Emne: ${subject}
Innhold:
${body}

KRAV:
- Norsk
- Profesjonell, men uformell tone
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
          content: "Du skriver korte, presise e-postsvar på norsk.",
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
  } catch (error) {
    return NextResponse.json(
      { reply: "Feil ved generering av svar" },
      { status: 500 }
    );
  }
}