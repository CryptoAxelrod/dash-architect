// Called from the add-in's profile screen's "Cancel subscription" button.
// Does NOT flip profiles.is_pro itself — that stays the webhook's job
// alone (supabase/functions/paddle-webhook), so there's exactly one place
// in the whole system that decides Pro status. This function only ever
// tells Paddle to cancel; the client polls profiles.is_pro afterward
// (same waitForProSync pattern as right after checkout) until the
// resulting subscription.canceled webhook has landed.
//
// Identity: verified via the caller's own Supabase access token (same
// token addin/auth.js already holds), not a service-role key — this
// function only ever acts on the token's own subscription, looked up
// through the existing "select own row" RLS policy on profiles rather
// than trusting a client-supplied subscription id (which would let
// anyone ask to cancel anyone else's).
//
// Unlike paddle-webhook (called server-to-server by Paddle, never subject
// to CORS), this one is called directly from the add-in's own page via
// fetch() with an Authorization header — that makes it a CORS-preflighted
// request, so every response path here needs the CORS headers below or
// the browser blocks it before our code ever sees the real request.

const SUPABASE_URL = Deno.env.get('SUPABASE_URL')!;
const SUPABASE_ANON_KEY = Deno.env.get('SUPABASE_ANON_KEY')!;
const PADDLE_API_KEY = Deno.env.get('PADDLE_API_KEY')!;
const PADDLE_API_BASE = Deno.env.get('PADDLE_API_BASE')!; // e.g. https://sandbox-api.paddle.com

const CORS_HEADERS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'authorization, apikey, content-type',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
};

function respond(body: string, status: number) {
  return new Response(body, { status, headers: CORS_HEADERS });
}

Deno.serve(async (req) => {
  if (req.method === 'OPTIONS') return new Response(null, { headers: CORS_HEADERS });
  if (req.method !== 'POST') return respond('Method not allowed', 405);

  const authHeader = req.headers.get('Authorization');
  if (!authHeader) return respond('Missing Authorization header', 401);

  const userRes = await fetch(`${SUPABASE_URL}/auth/v1/user`, {
    headers: { apikey: SUPABASE_ANON_KEY, Authorization: authHeader },
  });
  if (!userRes.ok) return respond('Invalid session', 401);
  const user = await userRes.json();

  const profileRes = await fetch(
    `${SUPABASE_URL}/rest/v1/profiles?select=paddle_subscription_id&id=eq.${encodeURIComponent(user.id)}`,
    { headers: { apikey: SUPABASE_ANON_KEY, Authorization: authHeader } },
  );
  if (!profileRes.ok) return respond('Could not read profile', 500);
  const rows = await profileRes.json();
  const subscriptionId = rows?.[0]?.paddle_subscription_id;
  if (!subscriptionId) return respond('No active subscription found', 400);

  const cancelRes = await fetch(`${PADDLE_API_BASE}/subscriptions/${encodeURIComponent(subscriptionId)}/cancel`, {
    method: 'POST',
    headers: { Authorization: `Bearer ${PADDLE_API_KEY}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ effective_from: 'immediately' }),
  });
  if (!cancelRes.ok) {
    console.error('Paddle cancel failed', cancelRes.status, await cancelRes.text());
    return respond('Could not cancel subscription', 502);
  }

  return respond('ok', 200);
});
