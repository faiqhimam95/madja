-- Sistem Nilai MDT: "jadwal_kiriman" inbox table
-- Bridges the offline Jadwal-only desktop build (E:\sistem-nilai-mdt-desktop-jadwal) and
-- this web app when the desktop build isn't run with a fully offline .env — its "Kirim ke
-- Web App" button inserts a row here, and the web app's Susun Jadwal → Cetak & Export →
-- "Kiriman dari Desktop" panel lists/reads/deletes rows. See src/App.jsx's
-- JADWAL_SYNC_URL/JADWAL_SYNC_KEY consts and sendJadwalKiriman/fetchJadwalKiriman/
-- deleteJadwalKiriman functions.
create table if not exists jadwal_kiriman (
  id uuid primary key default gen_random_uuid(),
  created_at timestamptz not null default now(),
  source text,
  jadwal jsonb not null
);

alter table jadwal_kiriman enable row level security;

drop policy if exists "anon_insert_jadwal_kiriman" on jadwal_kiriman;
create policy "anon_insert_jadwal_kiriman" on jadwal_kiriman for insert to anon with check (true);

drop policy if exists "anon_select_jadwal_kiriman" on jadwal_kiriman;
create policy "anon_select_jadwal_kiriman" on jadwal_kiriman for select to anon using (true);

drop policy if exists "anon_delete_jadwal_kiriman" on jadwal_kiriman;
create policy "anon_delete_jadwal_kiriman" on jadwal_kiriman for delete to anon using (true);
