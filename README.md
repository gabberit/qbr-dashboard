# QBR — testopstelling (Supabase + GitHub)

Complete repo om de QBR-keten te testen: **Supabase (Postgres) -> GitHub Actions (collector) -> GitHub Pages (dashboard)**.

## Inhoud
- `index.html` — het dashboard (Supabase-loader ingebouwd; leeg = demo-modus)
- `schema.sql` — tabellen + RLS + synthetische testdata (2 klanten, 5 maand-snapshots)
- `collect_qbr_supabase.mjs` — maand-collector die snapshots naar Supabase schrijft
- `package.json` — npm-scripts (`npm run dryrun` = offline test)
- `.github/workflows/qbr.yml` — draait de collector maandelijks + handmatig

## Snelstart (na het aanmaken van je gratis accounts)
1. **Supabase**: nieuw project -> SQL Editor -> plak `schema.sql` -> Run. Noteer Project URL + anon key + service_role key.
2. **Dashboard**: vul in `index.html` bovenin `SUPA_URL` en `SUPA_ANON` in.
3. **GitHub**: maak een repo, upload deze bestanden. Settings -> Pages -> Deploy from branch `main`. Open de Pages-URL -> je ziet de 2 testklanten.
4. **Collector**: Settings -> Secrets and variables -> Actions -> voeg `SUPA_URL` en `SUPA_SERVICE_KEY` toe. Actions -> Run workflow (handmatig) om te testen.
5. **Bronnen**: koppel bron voor bron (zie het testplan `QBR_Testopstelling_SQL_GitHub.md`), telkens 1 secret + 1 klant-ID erbij.

> De testdata is synthetisch (verzonnen klanten/cijfers). Zet de repo op privé zodra je met echte klantdata werkt, en scherp de RLS aan voor de klantgerichte modus.

