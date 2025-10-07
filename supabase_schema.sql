-- Schéma Supabase pour l'outil de gestion des gardes
-- À exécuter dans la console SQL de Supabase

-- Génération d'identifiants UUID
create extension if not exists "pgcrypto";

create table if not exists public.users (
  id uuid primary key default gen_random_uuid(),
  username text unique not null,
  password text not null,
  role text not null check (role in ('administrateur', 'medecin', 'remplacant')),
  created_at timestamptz not null default now()
);

insert into public.users (username, password, role)
values ('admin', 'Melatonine', 'administrateur')
on conflict (username) do nothing;

create table if not exists public.doctor_entries (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null references public.users(id) on delete cascade,
  date date not null,
  type text not null,
  notes text,
  created_at timestamptz not null default now()
);

create index if not exists doctor_entries_user_id_idx on public.doctor_entries (user_id, date);

create table if not exists public.replacement_entries (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null references public.users(id) on delete cascade,
  date date not null,
  slot text not null,
  notes text,
  created_at timestamptz not null default now()
);

create index if not exists replacement_entries_user_id_idx on public.replacement_entries (user_id, date);

-- Activez la Row Level Security si vous souhaitez utiliser une clé anon.
alter table public.users enable row level security;
alter table public.doctor_entries enable row level security;
alter table public.replacement_entries enable row level security;

-- Politiques d'exemple (à adapter selon vos besoins de sécurité)
create policy if not exists "Allow service role" on public.users
  for all
  using (true)
  with check (true);

create policy if not exists "Allow service role" on public.doctor_entries
  for all
  using (true)
  with check (true);

create policy if not exists "Allow service role" on public.replacement_entries
  for all
  using (true)
  with check (true);
