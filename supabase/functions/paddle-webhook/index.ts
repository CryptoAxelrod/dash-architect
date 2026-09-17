// Verifies and applies Paddle Billing webhook events, flipping
// profiles.is_pro for the user named in the checkout's custom_data.
// Implements the scenario discussion this came out of:
//   - transaction.completed / subscription.activated -> is_pro = true
//   - subscription.paused / subscription.canceled     -> is_pro = false
//   - subscription.past_due and a lone transaction.payment_failed are
//     deliberately NOT handled here — Paddle retries a failed renewal on
//     its own dunning schedule, and access should only be pulled once it
//     gives up (paused), not on the first missed attempt. Likewise
//     cancelling mid-period doesn't revoke immediately — Paddle keeps the
//     subscription "active" with a scheduled_change until the paid period
//     actually ends, and only then fires subscription.canceled.
// Every other event type (transaction.created, address.updated, etc.) is
// acknowledged and ignored — this function only cares about the ones that
// change what a user is entitled to.
//
// Identity: the add-in tags the checkout with
// custom_data.supabase_user_id when it initiates payment, so Paddle
// carries that through to every subsequent webhook for the same
// subscription — no separate Paddle-customer-id <-> Supabase-user table
// needed at this scale.

const SUPABASE_URL = Deno.env.get('SUPABASE_URL')!;
const SERVICE_ROLE_KEY = Deno.env.get('SUPABASE_SERVICE_ROLE_KEY')!;
const PADDLE_WEBHOOK_SECRET = Deno.env.get('PADDLE_WEBHOOK_SECRET')!;

const GRANT_EVENTS = new Set(['transaction.completed', 'subscription.activated']);
const REVOKE_EVENTS = new Set(['subscription.paused', 'subscription.canceled']);

// Paddle-Signature header looks like "ts=1671552777;h1=<hex hmac>". The
// signed message is "{ts}:{raw request body}" — must be the exact raw
// bytes Paddle sent, not a re-serialized JSON.parse/stringify round trip,
// or the digest won't match.
async function verifySignature(rawBody: string, header: string | null): Promise<boolean> {
  if (!header) return false;
  const parts = Object.fromEntries(header.split(';').map((p) => p.split('=') as [string, string]));
  const ts = parts['ts'];
  const h1 = parts['h1'];
  if (!ts || !h1) return false;

  const key = await crypto.subtle.importKey(
    'raw',
    new TextEncoder().encode(PADDLE_WEBHOOK_SECRET),
    { name: 'HMAC', hash: 'SHA-256' },
    false,
    ['sign'],
  );
  const sigBuf = await crypto.subtle.sign('HMAC', key, new TextEncoder().encode(`${ts}:${rawBody}`));
  const digest = Array.from(new Uint8Array(sigBuf)).map((b) => b.toString(16).padStart(2, '0')).join('');

  // Constant-time compare — a timing side-channel here would let an
  // attacker guess the correct signature byte by byte.
  if (digest.length !== h1.length) return false;
  let diff = 0;
  for (let i = 0; i < digest.length; i++) diff |= digest.charCodeAt(i) ^ h1.charCodeAt(i);
  return diff === 0;
}

Deno.serve(async (req) => {
  if (req.method !== 'POST') return new Response('Method not allowed', { status: 405 });

  const rawBody = await req.text();
  if (!(await verifySignature(rawBody, req.headers.get('paddle-signature')))) {
    return new Response('Invalid signature', { status: 401 });
  }

  let event: {
    event_type?: string;
    data?: { id?: string; subscription_id?: string; custom_data?: { supabase_user_id?: string } };
  };
  try {
    event = JSON.parse(rawBody);
  } catch {
    return new Response('Bad JSON', { status: 400 });
  }

  const eventType = event.event_type ?? '';
  if (!GRANT_EVENTS.has(eventType) && !REVOKE_EVENTS.has(eventType)) {
    return new Response('ignored', { status: 200 });
  }

  const userId = event.data?.custom_data?.supabase_user_id;
  if (!userId) {
    // Nothing a retry would fix — this checkout was never tagged with a
    // Supabase user id — so ack it rather than let Paddle keep resending.
    return new Response('no supabase_user_id in custom_data', { status: 200 });
  }

  const isPro = GRANT_EVENTS.has(eventType);
  // subscription.* events carry the subscription id as data.id;
  // transaction.completed carries it as data.subscription_id instead.
  // Stored so addin/paddle-cancel-subscription can find it later — a
  // canceled/paused event just leaves whatever was last stored, since
  // there's nothing better to overwrite it with and is_pro already gates
  // whether Cancel is even offered.
  const subscriptionId = eventType.startsWith('subscription.') ? event.data?.id : event.data?.subscription_id;
  const patch: Record<string, unknown> = { is_pro: isPro };
  if (isPro && subscriptionId) patch.paddle_subscription_id = subscriptionId;

  const res = await fetch(`${SUPABASE_URL}/rest/v1/profiles?id=eq.${encodeURIComponent(userId)}`, {
    method: 'PATCH',
    headers: {
      apikey: SERVICE_ROLE_KEY,
      Authorization: `Bearer ${SERVICE_ROLE_KEY}`,
      'Content-Type': 'application/json',
      Prefer: 'return=minimal',
    },
    body: JSON.stringify(patch),
  });

  if (!res.ok) {
    console.error('profiles update failed', res.status, await res.text());
    return new Response('database update failed', { status: 500 }); // non-2xx -> Paddle retries
  }

  return new Response('ok', { status: 200 });
});
