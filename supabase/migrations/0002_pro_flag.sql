-- Free-tier gating: a plain boolean is enough for now (no partial/expired
-- states yet — that's for when the Paddle webhook lands and actually flips
-- this). Client can SELECT it (existing "profiles: select own row" policy
-- covers the new column too) but has no UPDATE policy on profiles at all,
-- so this can only ever be changed server-side (service_role) — a signed-in
-- user cannot grant themselves Pro by calling the REST API directly.
alter table public.profiles add column is_pro boolean not null default false;
