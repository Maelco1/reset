-- Schéma Supabase pour l'outil de gestion des gardes
-- À exécuter dans la console SQL de Supabase

-- Active l'extension nécessaire pour générer des UUID côté base
create extension if not exists "pgcrypto";

do $$
begin
  if not exists (
    select 1 from pg_type t
    join pg_namespace n on n.oid = t.typnamespace
    where t.typname = 'user_role' and n.nspname = 'public'
  ) then
    create type public.user_role as enum ('administrateur', 'medecin', 'remplacant');
  end if;
end $$;

-- Déclencheur générique pour mettre à jour updated_at
create or replace function public.set_updated_at()
returns trigger as $$
begin
  new.updated_at = now();
  return new;
end;
$$ language plpgsql;

create table if not exists public.users (
  id uuid primary key default gen_random_uuid(),
  username text unique not null,
  password text not null,
  role public.user_role not null,
  active boolean not null default true,
  display_name text,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

do $$
begin
  if not exists (
    select 1
    from pg_trigger tg
    join pg_class tbl on tbl.oid = tg.tgrelid
    join pg_namespace nsp on nsp.oid = tbl.relnamespace
    where tg.tgname = 'users_set_updated_at'
      and tbl.relname = 'users'
      and nsp.nspname = 'public'
  ) then
    create trigger users_set_updated_at
      before update on public.users
      for each row
      execute function public.set_updated_at();
  end if;
end $$;

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
do $$
begin
  if not exists (
    select 1 from information_schema.table_constraints
    where constraint_name = 'doctor_entries_unique_user_date'
      and table_name = 'doctor_entries'
      and table_schema = 'public'
  ) then
    alter table public.doctor_entries
      add constraint doctor_entries_unique_user_date unique (user_id, date, type);
  end if;
end $$;

create table if not exists public.replacement_entries (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null references public.users(id) on delete cascade,
  date date not null,
  slot text not null,
  notes text,
  created_at timestamptz not null default now()
);

create index if not exists replacement_entries_user_id_idx on public.replacement_entries (user_id, date);
do $$
begin
  if not exists (
    select 1 from information_schema.table_constraints
    where constraint_name = 'replacement_entries_unique_user_date'
      and table_name = 'replacement_entries'
      and table_schema = 'public'
  ) then
    alter table public.replacement_entries
      add constraint replacement_entries_unique_user_date unique (user_id, date, slot);
  end if;
end $$;

-- Activez la Row Level Security si vous souhaitez utiliser une clé anon.
alter table public.users enable row level security;
alter table public.doctor_entries enable row level security;
alter table public.replacement_entries enable row level security;

drop policy if exists "Allow service role" on public.users;
create policy "Allow service role" on public.users
  for all
  using (true)
  with check (true);

drop policy if exists "Admins can view users" on public.users;
create policy "Admins can view users" on public.users
  for select
  using (
    coalesce((auth.jwt() ->> 'role')::public.user_role, 'remplacant'::public.user_role) = 'administrateur'
  );

drop policy if exists "Allow service role" on public.doctor_entries;
create policy "Allow service role" on public.doctor_entries
  for all
  using (true)
  with check (true);

drop policy if exists "Admins can view doctor entries" on public.doctor_entries;
create policy "Admins can view doctor entries" on public.doctor_entries
  for select
  using (
    coalesce((auth.jwt() ->> 'role')::public.user_role, 'remplacant'::public.user_role) = 'administrateur'
  );

drop policy if exists "Allow service role" on public.replacement_entries;
create policy "Allow service role" on public.replacement_entries
  for all
  using (true)
  with check (true);

drop policy if exists "Admins can view replacement entries" on public.replacement_entries;
create policy "Admins can view replacement entries" on public.replacement_entries
  for select
  using (
    coalesce((auth.jwt() ->> 'role')::public.user_role, 'remplacant'::public.user_role) = 'administrateur'
  );
