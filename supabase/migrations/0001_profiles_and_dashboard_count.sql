-- Dash Architect auth: one row per user tracking how many dashboards
-- they've generated. See CLAUDE.md's "Supabase exception" section for why
-- this is the one place /addin is allowed to make network calls.

create table public.profiles (
  id uuid primary key references auth.users (id) on delete cascade,
  email text not null,
  dashboards_created integer not null default 0,
  created_at timestamptz not null default now()
);

alter table public.profiles enable row level security;

create policy "profiles: select own row"
  on public.profiles for select
  using (auth.uid() = id);

-- No insert/update policy for the client: rows are created only by the
-- trigger below (as the table owner, bypassing RLS) and updated only by
-- the increment_dashboard_count() RPC (security definer, see below) — the
-- client is never allowed to write dashboards_created directly, so it can't
-- reset or backdate its own counter.

-- Auto-create a profile row the moment someone signs up, mirroring their
-- email — otherwise a fresh signup would have a NULL profile until they
-- happen to generate a dashboard.
create function public.handle_new_user()
returns trigger
language plpgsql
security definer
set search_path = public
as $$
begin
  insert into public.profiles (id, email) values (new.id, new.email);
  return new;
end;
$$;

create trigger on_auth_user_created
  after insert on auth.users
  for each row execute function public.handle_new_user();

-- Atomic increment, callable by the signed-in user for their own row only
-- (security definer + the auth.uid() check inside, not an RLS update
-- policy — a plain "update own row" policy would let a compromised client
-- set dashboards_created to anything, not just +1).
create function public.increment_dashboard_count()
returns integer
language plpgsql
security definer
set search_path = public
as $$
declare
  new_count integer;
begin
  update public.profiles
    set dashboards_created = dashboards_created + 1
    where id = auth.uid()
    returning dashboards_created into new_count;
  return new_count;
end;
$$;
