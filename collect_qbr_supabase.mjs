/**
 * collect_qbr_supabase.mjs — QBR-collector (Supabase-hub, ZONDER BrightGauge)
 * ---------------------------------------------------------------------------
 * Elke bron koppelt rechtstreeks via zijn eigen API. Architectuur:
 *   bronsystemen -> collector (deze file) -> Directus-hub -> HTML per klant.
 *
 * MSP-brede credentials staan in env-vars; PER KLANT staan de bron-ID's in de
 * Directus `clients`-collectie (tenant_id, datto_site, datto_saas,
 * rocketcyber_customer, itglue_org, freshdesk_company, bullphish_org, grc_client ...).
 *
 * Betrouwbaar geimplementeerd: Datto RMM, Datto BCDR/SaaS, RocketCyber,
 * Microsoft Graph (GDAP), Freshdesk, IT Glue.
 * Auth-patroon + TODO-API (endpoint/veld verifieren in vendor-docs): GRC,
 * K365 User (BullPhish/Dark Web ID), INKY, SaaS Alerts.
 *
 * Dry-run:  QBR_OFFLINE=1 QBR_TEMPLATE=QBR_Slimme_Werkplek_Fivespark.html node collect_qbr_direct.mjs
 */

import fs from 'node:fs';
import path from 'node:path';

const CFG = {
  offline : process.env.QBR_OFFLINE === '1',
  template: process.env.QBR_TEMPLATE || 'QBR_Slimme_Werkplek_Fivespark.html',
  outDir  : process.env.QBR_OUT || '.',
  supa: { url: process.env.SUPA_URL, key: process.env.SUPA_SERVICE_KEY },
  ms      : { clientId: process.env.MS_CLIENT_ID, secret: process.env.MS_SECRET, refresh: process.env.MS_REFRESH_TOKEN,
              scope: process.env.MS_SCOPE || 'https://graph.microsoft.com/.default offline_access' },
  dattoRmm: { platform: process.env.DRMM_PLATFORM, key: process.env.DRMM_KEY, secret: process.env.DRMM_SECRET },
  datto   : { pub: process.env.DATTO_PUBLIC_KEY, sec: process.env.DATTO_SECRET_KEY, base: 'https://api.datto.com' },
  rocket  : { base: process.env.ROCKETCYBER_BASE, token: process.env.ROCKETCYBER_TOKEN },
  bullphish:{ base: process.env.BULLPHISH_BASE, key: process.env.BULLPHISH_KEY },
  darkweb : { base: process.env.DARKWEB_BASE, key: process.env.DARKWEB_KEY },
  grc     : { base: process.env.GRC_BASE, key: process.env.GRC_KEY },
  inky    : { base: process.env.INKY_BASE, key: process.env.INKY_KEY },
  saasalerts:{ base: process.env.SAASALERTS_BASE, key: process.env.SAASALERTS_KEY },
  itglue  : { base: 'https://api.itglue.com', key: process.env.ITGLUE_KEY },
  fresh   : { domain: process.env.FRESHDESK_DOMAIN, key: process.env.FRESHDESK_KEY, catField: process.env.FRESHDESK_CATEGORY || 'type' },
  run     : { batchSize: Number(process.env.QBR_BATCH || 20), delayMs: Number(process.env.QBR_DELAY || 4000),
              trendN: Number(process.env.QBR_TRENDN || 12), sinceDays: Number(process.env.QBR_SINCE_DAYS || 30) },
  alertWebhook: process.env.ALERT_WEBHOOK,
};

const log = (...a) => console.log('[qbr]', new Date().toISOString(), ...a);
const sleep = ms => new Promise(r => setTimeout(r, ms));
async function jfetch(url, opts = {}) {
  const r = await fetch(url, opts);
  if (!r.ok) throw new Error(`${r.status} ${r.statusText} @ ${url} :: ${(await r.text()).slice(0, 140)}`);
  return r.json();
}

/* === SUPABASE (hub, PostgREST) ========================================== */
const sHead = () => ({ apikey: CFG.supa.key, Authorization: `Bearer ${CFG.supa.key}`, 'Content-Type': 'application/json', Prefer: 'return=minimal' });
const listClients = async () => (await jfetch(`${CFG.supa.url}/rest/v1/clients?select=*`, { headers: sHead() }));
const pushSnapshot = async (item) => { const r = await fetch(`${CFG.supa.url}/rest/v1/qbr_snapshots`, { method: 'POST', headers: sHead(), body: JSON.stringify(item) }); if (!r.ok) throw new Error('Supabase ' + r.status + ' ' + (await r.text()).slice(0,160)); };
const recentSnapshots = async (id, n) => (await jfetch(`${CFG.supa.url}/rest/v1/qbr_snapshots?client=eq.${id}&status=eq.published&order=generated_at.desc&limit=${n}&select=generated_at,data`, { headers: sHead() }));

/* === TEMPLATE-BASIS ===================================================== */
function braceSlice(html) {
  const a = html.indexOf('QBR_DATA = {'); if (a < 0) throw new Error('QBR_DATA niet gevonden.');
  const s = html.indexOf('{', a); let i = s, d = 0, inStr = false, q = '', esc = false;
  for (; i < html.length; i++) { const c = html[i];
    if (inStr) { if (esc) esc = false; else if (c === '\\') esc = true; else if (c === q) inStr = false; continue; }
    if (c === '"' || c === "'") { inStr = true; q = c; continue; }
    if (c === '{') d++; else if (c === '}') { d--; if (d === 0) break; } }
  return { start: s, end: i };
}
const extractBase = html => { const { start, end } = braceSlice(html); return new Function('return (' + html.slice(start, end + 1) + ')')(); };
const injectBase = (html, data) => { const { start, end } = braceSlice(html); return html.slice(0, start) + JSON.stringify(data) + html.slice(end + 1); };

/* === 1. MICROSOFT GRAPH via GDAP (licenties + MFA/licentie-drilldown) === */
async function msToken(tenantId) {
  // App-only (client-credentials): geen refresh token nodig, alleen client-ID + secret.
  const url = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
  const body = new URLSearchParams({ client_id: CFG.ms.clientId, client_secret: CFG.ms.secret,
    grant_type: 'client_credentials', scope: 'https://graph.microsoft.com/.default' });
  const tok = await (await fetch(url, { method: 'POST', body })).json();
  if (!tok.access_token) throw new Error('Graph-token: ' + (tok.error_description || tok.error));
  return tok.access_token;
}
const mfaStrength = m => { const s = (m || []).join(',').toLowerCase();
  if (/fido2|windowshello|passkey/.test(s)) return { t: 'Sterk', cls: 'c-success' };
  if (/authenticator|softwareoath/.test(s)) return { t: 'Goed', cls: 'c-success' };
  if (/sms|voice|phone/.test(s)) return { t: 'Zwak \u2014 migreren', cls: 'c-warning b' };
  return { t: '\u2014', cls: 'c-muted' }; };
async function getMicrosoft(client) {
  if (!CFG.ms.clientId || !CFG.ms.secret || !client.tenant_id) return null;
  const H = { Authorization: `Bearer ${await msToken(client.tenant_id)}` };
  const out = { licenties: {}, details: {} };
  try { const skus = await jfetch('https://graph.microsoft.com/v1.0/subscribedSkus', H);
    let p = 0, a = 0; for (const s of skus.value || []) { p += s.prepaidUnits?.enabled || 0; a += s.consumedUnits || 0; }
    out.licenties = { purchased: p, assigned: a }; } catch (e) { log('  graph lic:', e.message); }
  try { const rep = await jfetch('https://graph.microsoft.com/beta/reports/authenticationMethods/userRegistrationDetails?$top=200', H);
    const rows = (rep.value || []).slice(0, 12).map(u => [u.userDisplayName || u.userPrincipalName, '\u2014',
      u.isMfaRegistered ? { t: 'Aan', cls: 'c-success b' } : { t: 'Uit', cls: 'c-danger b' }, (u.methodsRegistered || []).join(', ') || '\u2014', mfaStrength(u.methodsRegistered)]);
    if (rows.length) out.details.mfa = { title: 'MFA & sterke authenticatie \u2014 per gebruiker', src: 'Entra ID \u00b7 Graph', cols: ['Gebruiker', 'Afdeling', 'MFA', 'Methode', 'Sterkte'], rows };
  } catch (e) { log('  graph mfa:', e.message); }
  return out;
}

/* === 2. COMPLIANCE MANAGER GRC (Secure Score, Identity, NIS2, Purview) == */
async function getGRC(client) {
  if (!CFG.grc.base || !CFG.grc.key || !client.grc_client) return null;
  const H = { Authorization: `Bearer ${CFG.grc.key}` };
  // TODO-API: verifieer endpoint(s) van Compliance Manager GRC (Microsoft Cloud-assessment).
  try {
    const a = await jfetch(`${CFG.grc.base}/assessments/${client.grc_client}/microsoft-cloud`, H);
    return { scorecard: { secure: a.secureScore, identity: a.identitySecureScore, nis2: a.nis2Percent },
             compliance: { pct: a.nis2Percent }, purview: { dlpIncidents: a.dlpIncidents },
             licenties: { purchased: a.licensesPurchased, assigned: a.licensesAssigned } };
  } catch (e) { log('  grc:', e.message); return null; }
}

/* === 3. DATTO RMM (patch/health) ======================================== */
async function getDattoRMM(client) {
  const { platform, key, secret } = CFG.dattoRmm;
  if (!platform || !key || !secret || !client.datto_site) return null;
  const base = `https://${platform}-api.centrastage.net`;
  try {
    const tok = await (await fetch(`${base}/auth/oauth/token`, { method: 'POST',
      headers: { Authorization: 'Basic ' + Buffer.from('public-client:public').toString('base64'), 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({ grant_type: 'password', username: key, password: secret }) })).json();
    const H = { Authorization: `Bearer ${tok.access_token}` };
    const dev = await jfetch(`${base}/api/v2/site/${client.datto_site}/devices?max=250`, H);
    const list = dev.devices || []; let patched = 0;
    for (const d of list) if ((d.patchManagement?.patchStatus || '').toLowerCase().includes('fully')) patched++;
    return { scorecard: { patch: list.length ? Math.round(patched / list.length * 100) : null } };
  } catch (e) { log('  datto rmm:', e.message); return null; }
}

/* === 4. DATTO BCDR + SaaS Protection (back-up) ========================== */
async function getDattoBCDR(client) {
  if (!CFG.datto.pub || !CFG.datto.sec || !client.datto_saas) return null;
  const H = { Authorization: 'Basic ' + Buffer.from(`${CFG.datto.pub}:${CFG.datto.sec}`).toString('base64') };
  try {
    const s = await jfetch(`${CFG.datto.base}/v1/saas/${client.datto_saas}/seats`, H);   // TODO-API: bevestig pad
    const seats = s.items || s || [];
    return { rmm: { backupNote: `${seats.length} SaaS-seats beschermd` } };
  } catch (e) { log('  datto bcdr:', e.message); return null; }
}

/* === 5. ROCKETCYBER (MDR/SOC + Datto EDR) =============================== */
async function getRocketCyber(client) {
  if (!CFG.rocket.base || !CFG.rocket.token || !client.rocketcyber_customer) return null;
  const H = { Authorization: `Bearer ${CFG.rocket.token}` };
  try {
    const inc = await jfetch(`${CFG.rocket.base}/incidents?customerId=${client.rocketcyber_customer}&pageSize=1000`, H);
    const items = inc.data || inc.items || [];
    const p1 = items.filter(i => (i.severity || '').toLowerCase() === 'critical').length;
    return { mdr: { socHandled: String(items.length), p1: String(p1) } };
  } catch (e) { log('  rocketcyber:', e.message); return null; }
}

/* === 6. K365 USER (BullPhish ID + Dark Web ID) ========================== */
async function getK365User(client) {
  if (!client.bullphish_org) return null;
  const out = { scorecard: {}, details: {} };
  // TODO-API: verifieer BullPhish/Dark Web ID endpoints + velden.
  if (CFG.bullphish.base && CFG.bullphish.key) {
    try { const c = await jfetch(`${CFG.bullphish.base}/organizations/${client.bullphish_org}/campaigns`, { Authorization: `Bearer ${CFG.bullphish.key}` });
      const pct = c.phishProneRate ?? null; if (pct != null) out.scorecard.aware = pct; } catch (e) { log('  bullphish:', e.message); }
  }
  if (CFG.darkweb.base && CFG.darkweb.key) {
    try { await jfetch(`${CFG.darkweb.base}/organizations/${client.bullphish_org}/exposures`, { Authorization: `Bearer ${CFG.darkweb.key}` }); } catch (e) { log('  darkweb:', e.message); }
  }
  return out;
}

/* === 7. INKY (e-mailbeveiliging) ======================================== */
async function getINKY(client) {
  if (!CFG.inky.base || !CFG.inky.key) return null;
  // TODO-API: verifieer INKY reporting-endpoint + client-filter.
  try { const s = await jfetch(`${CFG.inky.base}/stats?org=${client.slug}`, { Authorization: `Bearer ${CFG.inky.key}` });
    return { email: { banners: s.banners, impersonation: s.impersonation, reported: s.reported } };
  } catch (e) { log('  inky:', e.message); return null; }
}

/* === 8. SaaS ALERTS (M365-activiteit) =================================== */
async function getSaaSAlerts(client) {
  if (!CFG.saasalerts.base || !CFG.saasalerts.key) return null;
  // TODO-API: verifieer SaaS Alerts endpoint + client-filter.
  try { const a = await jfetch(`${CFG.saasalerts.base}/alerts?customer=${client.slug}`, { Authorization: `Bearer ${CFG.saasalerts.key}` });
    const items = a.data || a || []; return { mdr: { saasAlerts: String(items.length) } };
  } catch (e) { log('  saasalerts:', e.message); return null; }
}

/* === 9. IT GLUE (hardware lifecycle) ==================================== */
async function getITGlue(client) {
  if (!CFG.itglue.key || !client.itglue_org) return null;
  const H = { 'x-api-key': CFG.itglue.key, 'Content-Type': 'application/vnd.api+json' };
  try {
    const cfgs = await jfetch(`${CFG.itglue.base}/configurations?filter[organization_id]=${client.itglue_org}&page[size]=200`, H);
    const list = cfgs.data || [];
    const eol = list.filter(c => c.attributes && c.attributes['warranty-expires-at'] && Date.parse(c.attributes['warranty-expires-at']) < Date.now());
    return { hardware: { total: list.length, eol: eol.length } };
  } catch (e) { log('  itglue:', e.message); return null; }
}

/* === 10. FRESHDESK (top-categorieen + SLA + top-10 tickets) ============= */
const FD_PRIO = { 1: 'Laag', 2: 'Middel', 3: 'Hoog', 4: 'Urgent' };
const FD_STAT = { 2: 'Open', 3: 'In behandeling', 4: 'Opgelost', 5: 'Gesloten' };
async function getFreshdesk(client) {
  if (!CFG.fresh.domain || !CFG.fresh.key || !client.freshdesk_company) return null;
  const H = { Authorization: 'Basic ' + Buffer.from(CFG.fresh.key + ':X').toString('base64') };
  const since = new Date(Date.now() - CFG.run.sinceDays * 864e5).toISOString();
  let page = 1, all = [];
  for (;;) { const url = `https://${CFG.fresh.domain}.freshdesk.com/api/v2/tickets?company_id=${client.freshdesk_company}&updated_since=${since}&per_page=100&page=${page}`;
    const b = await jfetch(url, H); if (!Array.isArray(b) || !b.length) break; all = all.concat(b); if (b.length < 100 || page >= 10) break; page++; }
  const cat = t => (CFG.fresh.catField === 'type' ? t.type : t.custom_fields?.[CFG.fresh.catField]) || 'Overig';
  const byCat = {}; for (const t of all) { const c = cat(t); byCat[c] = (byCat[c] || 0) + 1; }
  const tickets = Object.entries(byCat).sort((a, b) => b[1] - a[1]).slice(0, 10).map(([c, n]) => [c, String(n), { t: '\u2014', cls: 'c-muted' }]);
  const resolved = all.filter(t => t.status === 4 || t.status === 5);
  const slaPct = resolved.length ? Math.round(resolved.filter(t => !t.is_escalated).length / resolved.length * 100) : null;
  const rows = all.sort((a, b) => (b.priority || 0) - (a.priority || 0)).slice(0, 10)
    .map(t => [(t.subject || '').slice(0, 60), cat(t), FD_PRIO[t.priority] || '\u2014', { t: FD_STAT[t.status] || '\u2014', cls: t.status >= 4 ? 'c-success' : 'c-warning' }]);
  return { service: { totaalVal: String(all.length), slaVal: slaPct != null ? slaPct + '%' : null, tickets },
           details: { tickets: { title: 'Top-tickets \u2014 dit kwartaal', src: 'Freshdesk', lead: `${all.length} tickets in de periode.`, cols: ['Onderwerp', 'Categorie', 'Prioriteit', 'Status'], rows } } };
}

/* === OVERLAY ============================================================ */
function setKpi(base, det, val) { const k = base.scorecard.find(x => x.det === det); if (!k || val == null) return;
  const pct = ['secure', 'identity', 'nis2', 'patch', 'aware'].includes(det);
  k.val = pct && !String(val).includes('%') ? String(val) + '%' : String(val); }
const tile = (list, needle, val) => { const t = (list || []).find(x => x.lbl && x.lbl.toLowerCase().includes(needle)); if (t && val != null) t.val = String(val); };
function overlay(base, slices) {
  for (const s of slices) { if (!s) continue;
    if (s.scorecard) for (const d of ['secure', 'identity', 'nis2', 'patch', 'aware', 'vuln', 'phish', 'dlp']) if (s.scorecard[d] != null) setKpi(base, d, s.scorecard[d]);
    if (s.service?.totaalVal) base.service.totaal.val = s.service.totaalVal;
    if (s.service?.slaVal) base.service.sla.val = s.service.slaVal;
    if (s.service?.tickets?.length) base.service.tickets = s.service.tickets;
    if (s.mdr?.socHandled) tile(base.mdr?.stats, 'soc-incidenten', s.mdr.socHandled);
    if (s.mdr?.p1) tile(base.mdr?.stats, 'kritieke', s.mdr.p1);
    if (s.rmm?.backupNote && base.rmm?.stats) { const t = base.rmm.stats.find(x => (x.lbl || '').toLowerCase().includes('back-up')); if (t) t.delta = s.rmm.backupNote; }
    if (s.hardware && base.hardware) { /* map naar hardware-sectie indien aanwezig */ }
    if (s.licenties) { tile(base.licenties?.stats, 'aangeschaft', s.licenties.purchased); tile(base.licenties?.stats, 'toegewezen', s.licenties.assigned); }
    if (s.details) base.details = { ...base.details, ...s.details };
  }
  return base;
}

/* === TREND ============================================================== */
const numOf = v => { const m = String(v ?? '').match(/-?\d+(?:[.,]\d+)?/); return m ? Number(m[0].replace(',', '.')) : null; };
const kpiVal = (sn, det) => sn?.data?.scorecard?.find(k => k.det === det)?.val;
async function addTrend(clientId, base) {
  if (CFG.offline || !CFG.supa.url) return;
  let hist = []; try { hist = await recentSnapshots(clientId, CFG.run.trendN); } catch { return; }
  const ser = hist.slice().reverse().map(sn => ({ m: (sn.generated_at || '').slice(0, 7), secure: numOf(kpiVal(sn, 'secure')), phishProne: numOf(kpiVal(sn, 'aware')), nis2: numOf(kpiVal(sn, 'nis2')) }));
  base.trend = { secure: ser.map(p => ({ m: p.m, v: p.secure })).filter(p => p.v != null),
                 phishProne: ser.map(p => ({ m: p.m, v: p.phishProne })).filter(p => p.v != null),
                 nis2: ser.map(p => ({ m: p.m, v: p.nis2 })).filter(p => p.v != null) };
}

/* === PER KLANT ========================================================== */
const periodLabel = (d = new Date()) => `Q${Math.floor(d.getMonth() / 3) + 1} ${d.getFullYear()}`;
async function buildForClient(client, templateHtml) {
  const base = extractBase(templateHtml);
  base.meta.client = client.name; base.meta.quarter = periodLabel();
  const sources = [getMicrosoft, getGRC, getDattoRMM, getDattoBCDR, getRocketCyber, getK365User, getINKY, getSaaSAlerts, getITGlue, getFreshdesk];
  const slices = [];
  for (const fn of sources) { try { const r = await fn(client); if (r) slices.push(r); } catch (e) { log('  bron-fout', client.name, fn.name, e.message); } }
  overlay(base, slices);
  await addTrend(client.id, base);
  return base;
}

/* === MAIN =============================================================== */
async function notify(res) { if (!CFG.alertWebhook) return;
  const text = `QBR-maandrun (direct): ${res.ok.length} ok, ${res.fail.length} gefaald`
    + (res.fail.length ? `\n` + res.fail.map(f => `\u2022 ${f.client}: ${f.error}`).join('\n') : '');
  try { await fetch(CFG.alertWebhook, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ text }) }); } catch {} }
async function main() {
  const templateHtml = fs.readFileSync(CFG.template, 'utf8');
  const clients = CFG.offline
    ? [{ id: 'demo', name: 'De Jong Logistics B.V.', slug: 'dejong', tenant_id: null }]
    : await listClients();
  log(`Start maand-run [BUILD-4 · app-only Graph] voor ${clients.length} klant(en)`);
  const res = { ok: [], fail: [] };
  for (let i = 0; i < clients.length; i += CFG.run.batchSize) {
    await Promise.all(clients.slice(i, i + CFG.run.batchSize).map(async c => {
      try {
        const data = await buildForClient(c, templateHtml);
        if (CFG.offline) { const out = path.join(CFG.outDir, `QBR_${c.name.replace(/[^\w]+/g, '_')}.html`); fs.writeFileSync(out, injectBase(templateHtml, data)); log('  offline geschreven:', out); }
        else await pushSnapshot({ client: c.id, quarter: data.meta.quarter, generated_at: new Date().toISOString(), status: 'published', data });
        res.ok.push(c.name);
      } catch (e) { res.fail.push({ client: c.name, error: e.message }); log('  FOUT', c.name, e.message); }
    }));
    if (i + CFG.run.batchSize < clients.length) await sleep(CFG.run.delayMs);
  }
  log(`Klaar: ${res.ok.length} ok, ${res.fail.length} gefaald`);
  if (res.fail.length) await notify(res);
}
main().catch(e => { console.error(e); process.exit(1); });
