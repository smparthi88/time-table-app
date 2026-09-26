-- Timetable Management System — Supabase schema
-- Already applied to the live project (jwodlfukmfdbkpbdsvoz). Kept here for
-- reference / reapplying on a fresh project.

create extension if not exists "pgcrypto";

create table if not exists app_settings (
  id boolean primary key default true check (id),
  academic_year text default '2025-2026',
  semester text default 'ODD',
  programme text default 'B.Tech - Information Technology'
);
insert into app_settings (id) values (true) on conflict do nothing;

create table if not exists faculty (
  id uuid primary key default gen_random_uuid(),
  name text not null unique,
  designation text,
  responsibility text,
  dob date,
  doj date,
  aicte_code text,
  au_code text,
  aadhar text,
  pan text,
  created_at timestamptz default now()
);

create table if not exists subject_master (
  id uuid primary key default gen_random_uuid(),
  class text not null,
  code text,
  title text not null,
  shortcut text,
  created_at timestamptz default now()
);
create index if not exists idx_subject_master_class on subject_master(class);

create table if not exists allotments (
  id uuid primary key default gen_random_uuid(),
  class text not null,
  academic_year text not null,
  subject_id uuid references subject_master(id) on delete cascade,
  faculty_name text,
  workload numeric,
  color text,
  created_at timestamptz default now(),
  unique (class, academic_year, subject_id)
);
create index if not exists idx_allotments_class_year on allotments(class, academic_year);

create table if not exists our_versions (
  id uuid primary key default gen_random_uuid(),
  class text not null,
  label text not null,
  meta jsonb not null default '{}'::jsonb,
  subjects jsonb not null default '[]'::jsonb,
  grid jsonb not null default '{}'::jsonb,
  is_active boolean not null default false,
  created_at timestamptz default now()
);
create index if not exists idx_our_versions_class on our_versions(class);
create unique index if not exists uniq_active_version_per_class
  on our_versions(class) where (is_active);

create table if not exists other_timetables (
  id uuid primary key default gen_random_uuid(),
  meta jsonb not null default '{}'::jsonb,
  subjects jsonb not null default '[]'::jsonb,
  grid jsonb not null default '{}'::jsonb,
  created_at timestamptz default now()
);

alter table app_settings enable row level security;
alter table faculty enable row level security;
alter table subject_master enable row level security;
alter table allotments enable row level security;
alter table our_versions enable row level security;
alter table other_timetables enable row level security;

create policy "auth read/write app_settings" on app_settings for all
  using (auth.role() = 'authenticated') with check (auth.role() = 'authenticated');
create policy "auth read/write faculty" on faculty for all
  using (auth.role() = 'authenticated') with check (auth.role() = 'authenticated');
create policy "auth read/write subject_master" on subject_master for all
  using (auth.role() = 'authenticated') with check (auth.role() = 'authenticated');
create policy "auth read/write allotments" on allotments for all
  using (auth.role() = 'authenticated') with check (auth.role() = 'authenticated');
create policy "auth read/write our_versions" on our_versions for all
  using (auth.role() = 'authenticated') with check (auth.role() = 'authenticated');
create policy "auth read/write other_timetables" on other_timetables for all
  using (auth.role() = 'authenticated') with check (auth.role() = 'authenticated');
