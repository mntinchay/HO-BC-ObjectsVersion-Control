/*
 * BC Extension Snapshot — unattended runner for GitHub Actions.
 *
 * Runs the same logic as "Run Sync Now" inside the app (viewExtMgmt /
 * runExtensionSnapshot in index.html), but from GitHub's own servers on a
 * cron schedule, so it fires whether or not anyone has the app open.
 *
 * Reads its config from the SAME SharePoint database file the app uses
 * (settings.extMgmt.environments / retentionDays), so the environments
 * you configure in the UI are exactly what gets snapshotted here — no
 * separate config to keep in sync.
 *
 * Required repo secrets (Settings > Secrets and variables > Actions):
 *   BC_TENANT_ID      Azure AD tenant (directory) ID
 *   BC_CLIENT_ID      Azure AD app registration (client) ID
 *   BC_CLIENT_SECRET  Client secret value for that app registration
 *
 * That single app registration needs BOTH of these, admin-consented:
 *   - Microsoft Graph  > Application > Sites.ReadWrite.All   (read/write the SharePoint file)
 *   - Business Central API access — either:
 *       (a) Custom API mode (default): normal BC user permissions in each
 *           environment for the deployed AL extension's API (publisher/
 *           group/version and each environment's Company ID come from
 *           settings.extMgmt in the database file itself), or
 *       (b) Admin Center API mode: BC Admin Center / Global Admin role,
 *           if settings.extMgmt.apiMode is set to "admin"
 *
 * Optional repo variables (Settings > Secrets and variables > Actions > Variables)
 * override where the database file lives; defaults match the values baked
 * into index.html:
 *   SP_HOSTNAME     default: healthyoptionsph.sharepoint.com
 *   SP_SITE_PATH    default: /sites/IT-SP
 *   SP_FOLDER_PATH  default: IT Team Folder/27 BC_LS/DevEnvironment/BC Object and Version Control Monitoring
 *   SP_FILE_NAME    default: bcovc-database.json
 */

const TENANT_ID = process.env.BC_TENANT_ID;
const CLIENT_ID = process.env.BC_CLIENT_ID;
const CLIENT_SECRET = process.env.BC_CLIENT_SECRET;

const SP_HOSTNAME = process.env.SP_HOSTNAME || 'healthyoptionsph.sharepoint.com';
const SP_SITE_PATH = process.env.SP_SITE_PATH || '/sites/IT-SP';
const SP_FOLDER_PATH = process.env.SP_FOLDER_PATH || 'IT Team Folder/27 BC_LS/DevEnvironment/BC Object and Version Control Monitoring';
const SP_FILE_NAME = process.env.SP_FILE_NAME || 'bcovc-database.json';

function fail(msg) {
  console.error(`::error::${msg}`);
  process.exit(1);
}

function encodeGraphPath(path) {
  return path.split('/').map(encodeURIComponent).join('/');
}

async function getToken(scope) {
  const res = await fetch(`https://login.microsoftonline.com/${encodeURIComponent(TENANT_ID)}/oauth2/v2.0/token`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      grant_type: 'client_credentials',
      client_id: CLIENT_ID,
      client_secret: CLIENT_SECRET,
      scope,
    }),
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(data.error_description || `Token request failed (HTTP ${res.status}) for scope ${scope}`);
  return data.access_token;
}

async function graphFetch(graphToken, path, options) {
  options = options || {};
  const res = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    ...options,
    headers: { Authorization: `Bearer ${graphToken}`, ...(options.headers || {}) },
  });
  return res;
}

async function getSiteId(graphToken) {
  const res = await graphFetch(graphToken, `/sites/${SP_HOSTNAME}:${SP_SITE_PATH}`);
  if (!res.ok) throw new Error(`Could not reach the SharePoint site (HTTP ${res.status}).`);
  const data = await res.json();
  return data.id;
}

async function loadDb(graphToken, siteId) {
  const res = await graphFetch(graphToken, `/sites/${siteId}/drive/root:/${encodeGraphPath(SP_FOLDER_PATH)}/${SP_FILE_NAME}:/content`);
  if (res.status === 404) throw new Error('Database file not found in SharePoint yet — open the app once first so it can create it.');
  if (!res.ok) throw new Error(`Could not read the database file (HTTP ${res.status}).`);
  return res.json();
}

async function saveDb(graphToken, siteId, db) {
  const res = await graphFetch(graphToken, `/sites/${siteId}/drive/root:/${encodeGraphPath(SP_FOLDER_PATH)}/${SP_FILE_NAME}:/content`, {
    method: 'PUT',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(db),
  });
  if (!res.ok) throw new Error(`Could not write the database file (HTTP ${res.status}).`);
}

async function fetchBcExtensions(bcToken, env, em) {
  let url;
  if (em.apiMode === 'admin') {
    url = `https://api.businesscentral.dynamics.com/admin/v2.18/applications/businesscentral/environments/${encodeURIComponent(env.name)}/extensions`;
  } else {
    if (!em.apiPublisher || !em.apiGroup) throw new Error('Missing API Publisher/Group in settings.extMgmt for custom API mode.');
    if (!env.companyId) throw new Error('Missing Company ID for this environment (required for custom API mode).');
    url = `https://api.businesscentral.dynamics.com/v2.0/${encodeURIComponent(em.tenantId)}/${encodeURIComponent(env.name)}/api/${encodeURIComponent(em.apiPublisher)}/${encodeURIComponent(em.apiGroup)}/${encodeURIComponent(em.apiVersion || 'v2.0')}/companies(${env.companyId})/installedExtensions`;
  }
  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${bcToken}` },
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error((data.error && data.error.message) || `HTTP ${res.status}`);
  return (data.value || []).map((e) => ({
    id: e.id || e.appId || e.name,
    name: e.displayName || e.name || 'Unnamed extension',
    publisher: e.publisher || '',
    version: e.version || [e.versionMajor, e.versionMinor, e.versionBuild, e.versionRevision].filter((v) => v != null).join('.'),
    state: e.state || (e.installed === false ? 'not installed' : 'installed'),
  }));
}

function uid() {
  return Math.random().toString(36).slice(2, 10) + Date.now().toString(36).slice(-4);
}

function pruneSnapshots(db) {
  const days = (db.settings.extMgmt && db.settings.extMgmt.retentionDays) || 7;
  const cutoff = Date.now() - days * 24 * 60 * 60 * 1000;
  db.extSnapshots = (db.extSnapshots || []).filter((s) => new Date(s.timestamp).getTime() >= cutoff);
}

async function main() {
  if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) {
    fail('Missing BC_TENANT_ID / BC_CLIENT_ID / BC_CLIENT_SECRET repo secrets.');
  }

  console.log('Requesting Microsoft Graph token...');
  const graphToken = await getToken('https://graph.microsoft.com/.default');
  const siteId = await getSiteId(graphToken);

  console.log('Loading database from SharePoint...');
  const db = await loadDb(graphToken, siteId);
  db.settings = db.settings || {};
  db.settings.extMgmt = db.settings.extMgmt || { environments: [], retentionDays: 7 };
  db.extSnapshots = db.extSnapshots || [];

  const environments = (db.settings.extMgmt.environments || []).filter((e) => e.name);
  if (!environments.length) {
    fail('No environments configured. Add at least one environment in the app\'s Extension Management Comparison tab first.');
  }

  console.log('Requesting Business Central token...');
  const bcToken = await getToken('https://api.businesscentral.dynamics.com/.default');

  const results = [];
  for (const env of environments) {
    console.log(`Fetching extensions for ${env.label || env.name}...`);
    try {
      const extensions = await fetchBcExtensions(bcToken, env, db.settings.extMgmt);
      results.push({ envId: env.id, envName: env.name, envLabel: env.label || env.name, envType: env.type, ok: true, extensions });
    } catch (e) {
      console.warn(`  failed: ${e.message}`);
      results.push({ envId: env.id, envName: env.name, envLabel: env.label || env.name, envType: env.type, ok: false, error: e.message, extensions: [] });
    }
  }

  const snapshot = { id: uid(), timestamp: new Date().toISOString(), results };
  db.extSnapshots.push(snapshot);
  pruneSnapshots(db);

  const anyFailed = results.some((r) => !r.ok);
  db.settings.extMgmt.lastRunAt = snapshot.timestamp;
  db.settings.extMgmt.lastRunStatus = anyFailed ? 'error' : 'success';
  db.settings.extMgmt.lastRunError = anyFailed
    ? results.filter((r) => !r.ok).map((r) => `${r.envLabel}: ${r.error}`).join(' | ')
    : '';

  db.logs = db.logs || [];
  db.logs.unshift({
    id: uid(),
    timestamp: new Date().toISOString(),
    user: 'GitHub Actions',
    action: 'Synced',
    entity: 'extmgmt',
    entityName: 'Extension Management',
    details: `Captured ${results.filter((r) => r.ok).length}/${results.length} environment(s) (scheduled run)`,
  });

  console.log('Saving database back to SharePoint...');
  await saveDb(graphToken, siteId, db);

  console.log(anyFailed ? 'Done, with some environment errors — see settings.extMgmt.lastRunError.' : 'Done.');
  if (anyFailed) process.exitCode = 1; // mark the Action run as failed so it's visible, without blocking the snapshot that did save
}

main().catch((e) => fail(e.message || String(e)));
