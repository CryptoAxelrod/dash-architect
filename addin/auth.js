/*
 * Dash Architect's one and only network client — see CLAUDE.md §3/§4 for
 * why this file exists at all: /addin is otherwise zero-network, and this
 * is the sole, explicitly-approved exception. Hand-rolled on `fetch`
 * against Supabase's plain HTTP API (Auth + PostgREST) rather than the
 * official supabase-js SDK, same reasoning as "no charting libraries" in
 * CLAUDE.md §5 — a handful of REST calls doesn't need a bundled client
 * library with its own realtime/storage/etc. surface.
 *
 * Owns exactly three things: sign up / sign in / sign out, persisting the
 * session (localStorage — same model supabase-js itself uses, so it
 * survives the task pane being closed and reopened), and the two calls
 * behind the per-user dashboard counter. Nothing here knows what a
 * dashboard is — addin/taskpane.js decides when to call
 * incrementDashboardCount(), same as it decides everything else about the
 * generate flow.
 */
(function (root) {
  'use strict';

  var SUPABASE_URL = 'https://qdubeyoavgbuyrkdixux.supabase.co';
  // Publishable/anon key — meant to ship in client code. Every table it can
  // reach is behind row-level security (see supabase/migrations); it grants
  // no access beyond what a signed-in user's own row allows.
  var ANON_KEY = 'sb_publishable_FoywjQsHTC0xiFvEbnoGvg_uzOnpAs5';
  var STORAGE_KEY = 'dash-architect-session';

  var session = null; // {access_token, refresh_token, expires_at (epoch seconds), user: {id, email}} | null

  function loadStoredSession() {
    try {
      var raw = localStorage.getItem(STORAGE_KEY);
      return raw ? JSON.parse(raw) : null;
    } catch (e) {
      return null; // private browsing / storage disabled — session just won't survive a reload
    }
  }

  function storeSession(s) {
    try {
      if (s) localStorage.setItem(STORAGE_KEY, JSON.stringify(s));
      else localStorage.removeItem(STORAGE_KEY);
    } catch (e) { /* see loadStoredSession */ }
  }

  function fromAuthResponse(body) {
    return {
      access_token: body.access_token,
      refresh_token: body.refresh_token,
      expires_at: Math.floor(Date.now() / 1000) + (body.expires_in || 3600),
      user: { id: body.user.id, email: body.user.email },
    };
  }

  async function request(path, opts) {
    var res = await fetch(SUPABASE_URL + path, opts);
    var body = null;
    try { body = await res.json(); } catch (e) { /* empty body, e.g. logout */ }
    if (!res.ok) {
      var message = (body && (body.error_description || body.msg || body.message)) || ('Request failed (' + res.status + ')');
      throw new Error(message);
    }
    return body;
  }

  function anonHeaders(extra) {
    return Object.assign({ apikey: ANON_KEY, 'Content-Type': 'application/json' }, extra || {});
  }

  // Refreshes in place if the access token is at or near expiry; otherwise
  // a no-op. Every authenticated call routes through this first so a task
  // pane left open past the token's ~1h lifetime doesn't just start failing.
  async function ensureFresh() {
    if (!session) return null;
    var now = Math.floor(Date.now() / 1000);
    if (session.expires_at - now > 60) return session;
    try {
      var body = await request('/auth/v1/token?grant_type=refresh_token', {
        method: 'POST', headers: anonHeaders(), body: JSON.stringify({ refresh_token: session.refresh_token }),
      });
      session = fromAuthResponse(body);
      storeSession(session);
      return session;
    } catch (e) {
      session = null;
      storeSession(null);
      return null;
    }
  }

  async function authedRequest(path, opts) {
    var s = await ensureFresh();
    if (!s) throw new Error('Not signed in.');
    var headers = Object.assign(anonHeaders({ Authorization: 'Bearer ' + s.access_token }), (opts && opts.headers) || {});
    return request(path, Object.assign({}, opts, { headers: headers }));
  }

  // Called once at startup to decide whether to show the sign-in gate or
  // go straight into the app — see addin/taskpane.js#init.
  async function restoreSession() {
    session = loadStoredSession();
    if (!session) return null;
    return ensureFresh();
  }

  async function signUp(email, password) {
    var body = await request('/auth/v1/signup', {
      method: 'POST', headers: anonHeaders(), body: JSON.stringify({ email: email, password: password }),
    });
    // With email confirmation off (CLAUDE.md's Supabase exception — see
    // this project's Auth settings), a fresh signup already comes back
    // with a live session. If it doesn't, the most common cause is the
    // email already being registered — Supabase avoids confirming that
    // directly (user enumeration), so this is a best-effort guess.
    if (!body || !body.access_token) throw new Error('Could not create an account with that email — it may already be registered.');
    session = fromAuthResponse(body);
    storeSession(session);
    return session;
  }

  async function signIn(email, password) {
    var body = await request('/auth/v1/token?grant_type=password', {
      method: 'POST', headers: anonHeaders(), body: JSON.stringify({ email: email, password: password }),
    });
    session = fromAuthResponse(body);
    storeSession(session);
    return session;
  }

  // Kicks off Google sign-in — the caller (addin/taskpane.js) navigates a
  // dialog window to this URL; Google eventually redirects the browser to
  // `redirectTo` (addin/auth-callback.html) with tokens in the URL
  // fragment, never through our own fetch/XHR code. Still entirely within
  // CLAUDE.md's Supabase exception: the only network call our code issues
  // is to this project's own /auth/v1/authorize — Google is a plain
  // browser navigation Supabase's redirect chain sends the dialog through,
  // the same way clicking a link would.
  function authorizeUrl(redirectTo) {
    return SUPABASE_URL + '/auth/v1/authorize?provider=google&redirect_to=' + encodeURIComponent(redirectTo);
  }

  // Turns the {access_token, refresh_token, expires_in, ...} the callback
  // page scraped out of the redirect's URL fragment into a stored session
  // — unlike signUp/signIn, that fragment never includes the user object
  // (too large for a URL), so this makes one extra call to fetch it.
  async function completeOAuthSession(params) {
    if (params.error) throw new Error(params.error_description || params.error);
    if (!params.access_token) throw new Error('Google sign-in did not return a session.');
    var expiresAt = params.expires_at ? parseInt(params.expires_at, 10) : Math.floor(Date.now() / 1000) + parseInt(params.expires_in || '3600', 10);
    var user = await request('/auth/v1/user', { headers: anonHeaders({ Authorization: 'Bearer ' + params.access_token }) });
    session = { access_token: params.access_token, refresh_token: params.refresh_token, expires_at: expiresAt, user: { id: user.id, email: user.email } };
    storeSession(session);
    return session;
  }

  async function signOut() {
    if (session) {
      try {
        await fetch(SUPABASE_URL + '/auth/v1/logout?scope=local', {
          method: 'POST', headers: anonHeaders({ Authorization: 'Bearer ' + session.access_token }),
        });
      } catch (e) { /* best-effort — clearing the local session below is what actually matters */ }
    }
    session = null;
    storeSession(null);
  }

  function getSession() { return session; }

  // Read-only, used right after sign-in to show a starting count/plan
  // without waiting for the first generate — addin/taskpane.js's free-tier
  // gate (Generate disabled + Upgrade CTA once the limit is hit) and the
  // account bar's usage indicator both read from this.
  async function fetchAccountStatus() {
    var s = await ensureFresh();
    if (!s) return null;
    var rows = await authedRequest('/rest/v1/profiles?select=dashboards_created,is_pro&id=eq.' + encodeURIComponent(s.user.id));
    if (!rows || !rows[0]) return null;
    return { dashboardsCreated: rows[0].dashboards_created, isPro: rows[0].is_pro };
  }

  // Atomic +1 via the increment_dashboard_count() RPC (supabase/migrations)
  // — never a read-modify-write from the client, so a stale local count or
  // a duplicate call can't stomp on a real value. Returns the new total.
  async function incrementDashboardCount() {
    return authedRequest('/rest/v1/rpc/increment_dashboard_count', { method: 'POST', body: '{}' });
  }

  // Tells Paddle to cancel this user's subscription (via
  // supabase/functions/paddle-cancel-subscription) — does not flip is_pro
  // itself, that only ever happens from the resulting webhook. Caller
  // (addin/taskpane.js) polls fetchAccountStatus afterward the same way it
  // does right after checkout.
  async function cancelSubscription() {
    var s = await ensureFresh();
    if (!s) throw new Error('Not signed in.');
    var res = await fetch(SUPABASE_URL + '/functions/v1/paddle-cancel-subscription', {
      method: 'POST',
      headers: anonHeaders({ Authorization: 'Bearer ' + s.access_token }),
    });
    if (!res.ok) {
      var text = await res.text().catch(function () { return ''; });
      throw new Error(text || ('Could not cancel subscription (' + res.status + ')'));
    }
  }

  root.DashAuth = {
    restoreSession: restoreSession,
    signUp: signUp,
    signIn: signIn,
    signOut: signOut,
    authorizeUrl: authorizeUrl,
    completeOAuthSession: completeOAuthSession,
    getSession: getSession,
    fetchAccountStatus: fetchAccountStatus,
    incrementDashboardCount: incrementDashboardCount,
    cancelSubscription: cancelSubscription,
  };
})(typeof window !== 'undefined' ? window : this);
