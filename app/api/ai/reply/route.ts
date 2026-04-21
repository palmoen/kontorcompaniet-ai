import { NextRequest, NextResponse } from "next/server";
import OpenAI from "openai";

export const runtime = "nodejs";

const openai = new OpenAI({
  apiKey: process.env.OPENAI_API_KEY,
});

export async function POST(req: NextRequest) {
  try {
    const { subject, body, from } = await req.json();

    const prompt = `
Du er en rådgiver i Kontorcompaniet.

Skriv et kort svar på e-posten under.

Fra: ${from}
Emne: ${subject}
Innhold:
${body}

KRAV:
- Norsk
- Maks 6-8 linjer
- Profesjonell, men uformell
- Foreslå neste steg
`;

    const completion = await openai.chat.completions.create({
      model: "gpt-5-mini",
      messages: [
        { role: "system", content: "Du skriver korte e-postsvar." },
        { role: "user", content: prompt },
      ],
    });

    const reply =
      completion.choices?.[0]?.message?.content ||
      "Kunne ikke generere svar";

    return NextResponse.json({ reply });
  } catch (e) {
    return NextResponse.json(
      { reply: "Feil ved generering" },
      { status: 500 }
    );
  }
}