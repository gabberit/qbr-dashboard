-- =====================================================================
--  QBR Admin — schrijf-rechten via Supabase Auth (Route A)
--  Draai dit in Supabase: SQL Editor -> New query -> plakken -> Run.
--  Effect: ingelogde medewerkers mogen snapshots bewerken; anonieme
--  bezoekers (het klant-dashboard) mogen alleen lezen.
-- =====================================================================

-- Anon (klant-dashboard): alleen lezen van gepubliceerde snapshots.
drop policy if exists "read snapshots" on qbr_snapshots;
create policy "read snapshots" on qbr_snapshots
  for select to anon using (status = 'published');

-- Ingelogde medewerkers: mogen lezen én bijwerken.
drop policy if exists "auth read snapshots" on qbr_snapshots;
create policy "auth read snapshots" on qbr_snapshots
  for select to authenticated using (true);

drop policy if exists "auth update snapshots" on qbr_snapshots;
create policy "auth update snapshots" on qbr_snapshots
  for update to authenticated using (true) with check (true);

-- Klantenlijst: lezen voor iedereen (dashboard + admin).
drop policy if exists "read clients" on clients;
create policy "read clients" on clients
  for select to anon using (true);
drop policy if exists "auth read clients" on clients;
create policy "auth read clients" on clients
  for select to authenticated using (true);

-- LET OP: de collector schrijft met de service_role key en negeert RLS -
-- die blijft dus gewoon werken. Deze policies regelen alleen anon vs. ingelogd.
