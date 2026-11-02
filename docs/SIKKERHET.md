# K-Pilot er faset ut

AI-hjelp til e-post er flyttet til Donna (`inbox-pilot`).

`/api/ai/reply` var åpent uten innlogging, slik at hvem som helst kunne bruke
OpenAI-kontoen via det. Endepunktet svarer derfor nå alltid **410 Gone**. Det leser ikke
forespørselen og kontakter ikke OpenAI. Taskpanelet i Outlook viser en melding om at
K-Pilot er faset ut.

## Gjenstående opprydding (manuelt)

1. **Vercel-prosjektet `kontorcompaniet-ai`:** fjern miljøvariabelen `OPENAI_API_KEY`.
   Dette lukker hullet umiddelbart, også før denne koden er deployet. Er nøkkelen den
   samme som brukes andre steder (f.eks. Donna), bør den heller **roteres** hos OpenAI.
2. **Deploy** denne versjonen, slik at brukerne ser utfasingsmeldingen.
3. **Microsoft 365 admin center → Integrerte apper:** fjern tillegget «AI E-postassistent»
   fra brukerne.
4. Når ingen bruker det lenger: slett Vercel-prosjektet og arkiver repoet på GitHub.
