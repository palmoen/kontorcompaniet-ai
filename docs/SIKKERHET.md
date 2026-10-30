# Sikkerhet — `/api/ai/reply`

## Status nå

Endepunktet har **ingen innlogging**. Det er beskyttet slik (`lib/requestGuard.ts`):

- **Origin-sjekk:** bare `KPILOT_ALLOWED_ORIGINS` (kommaseparert, standard
  `https://kontorcompaniet-ai.vercel.app`) godtas. Forespørsler uten `Origin` avvises.
  Taskpanelet kaller fra samme domene, og Outlooks nettleservisninger sender alltid
  `Origin` på POST.
- **Rate limiting:** 20 kall per minutt per IP. Telleren ligger i minnet, så grensen gjelder
  per serverinstans.
- **Størrelsesgrense:** 60 000 tegn per forespørsel.
- **Feilmeldinger:** interne feil (for eksempel fra OpenAI) logges på serveren og sendes
  aldri til klienten.
- **API-nøkkel:** `OPENAI_API_KEY` brukes bare på serveren og finnes ikke i `public/`.

**Dette stopper misbruk fra andre nettsteder og løpske klienter, men ikke en angriper som
kaller endepunktet direkte og forfalsker `Origin`-headeren.** Den egentlige løsningen er
innlogging, se under.

## Forslag: Office SSO (krever godkjenning før det bygges)

1. **Entra:** registrer (eller gjenbruk) en app med «Expose an API» og URI
   `api://kontorcompaniet-ai.vercel.app/<client-id>`, scope `access_as_user`, og
   forhåndsautoriser Office-klientene. Krever **admin-samtykke**.
2. **Manifest:** legg til `<WebApplicationInfo>` med client-id og ressurs-URI i
   `addin/manifest.xml`. Tillegget må deretter **rulles ut på nytt** til brukerne
   (sentralisert utrulling i Microsoft 365 admin center).
3. **Taskpanel:** hent token med `Office.auth.getAccessToken()` og send det som
   `Authorization: Bearer …` til `/api/ai/reply`. Legg inn fallback-melding for klienter
   uten SSO-støtte.
4. **Server:** verifiser JWT-en mot Entra sine JWKS (signatur, `aud`, `iss`/`tid` lik
   Kontorcompaniets tenant, `exp`), for eksempel med `jose`. Sett rate limiting per bruker
   (`oid`) i stedet for per IP.

**Konsekvens:** Brukere som har en eldre versjon av tillegget installert, slutter å virke
når serveren begynner å kreve token. Det bør derfor innføres bak et flagg
(`KPILOT_REQUIRE_SSO`) og slås på først når det nye manifestet er rullet ut.
