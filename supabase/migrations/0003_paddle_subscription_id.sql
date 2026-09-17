-- Needed so a signed-in user can cancel their own subscription from the
-- add-in's profile screen (supabase/functions/paddle-cancel-subscription)
-- without us storing anything more sensitive than the id itself. Same RLS
-- posture as dashboards_created/is_pro: client can read it, never write it
-- — only the webhook (service_role) sets it.
alter table public.profiles add column paddle_subscription_id text;
