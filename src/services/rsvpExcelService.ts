import * as XLSX from 'xlsx';

export interface RsvpRecord {
  id: string;
  submitted_at: string;
  guest_name: string;
  phone?: string | null;
  attending: 'yes' | 'no';
  guest_count: number;
  events: string[];
  dietary?: string | null;
  message?: string | null;
}

export interface GitHubSyncConfig {
  enabled: boolean;
  owner: string;
  repo: string;
  branch: string;
  filePath: string;
  token: string;
  autoSyncOnSubmit: boolean;
  lastSyncedAt?: string;
  lastCommitUrl?: string;
}

const STORAGE_KEY_RSVPS = 'wedding_rsvps';
const STORAGE_KEY_GH_CONFIG = 'wedding_github_sync_config';

const DEFAULT_GH_CONFIG: GitHubSyncConfig = {
  enabled: false,
  owner: 'auraweddingsandevents21-lang',
  repo: 'Basti-Ambiya',
  branch: 'main',
  filePath: 'wedding-rsvps.xlsx',
  token: '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',
  autoSyncOnSubmit: true,
};

/**
 * Retrieves all stored RSVPs from localStorage
 */
export function getStoredRsvps(): RsvpRecord[] {
  try {
    const raw = localStorage.getItem(STORAGE_KEY_RSVPS);
    if (!raw) return [];
    const parsed = JSON.parse(raw);
    if (Array.isArray(parsed)) {
      return parsed.map((item, index) => ({
        id: item.id || `rsvp-${index + 1}-${Date.now()}`,
        submitted_at: item.submitted_at || new Date().toISOString(),
        guest_name: item.guest_name || 'Anonymous Guest',
        phone: item.phone || '',
        attending: item.attending === 'no' ? 'no' : 'yes',
        guest_count: Number(item.guest_count) || (item.attending === 'no' ? 0 : 1),
        events: Array.isArray(item.events) ? item.events : [],
        dietary: item.dietary || '',
        message: item.message || '',
      }));
    }
  } catch (err) {
    console.error('Error reading wedding_rsvps from localStorage:', err);
  }
  return [];
}

/**
 * Saves all RSVPs to localStorage
 */
export function saveAllRsvps(rsvps: RsvpRecord[]): void {
  try {
    localStorage.setItem(STORAGE_KEY_RSVPS, JSON.stringify(rsvps));
    window.dispatchEvent(new CustomEvent('wedding_rsvp_updated', { detail: rsvps }));
  } catch (err) {
    console.error('Error saving wedding_rsvps to localStorage:', err);
  }
}

/**
 * Retrieves stored GitHub sync settings
 */
export function getGitHubConfig(): GitHubSyncConfig {
  let config: GitHubSyncConfig = {
    ...DEFAULT_GH_CONFIG,
    owner: (import.meta as any).env?.VITE_GITHUB_OWNER || 'ashutoshs019',
    repo: (import.meta as any).env?.VITE_GITHUB_REPO || 'wedding-invitation',
    branch: (import.meta as any).env?.VITE_GITHUB_BRANCH || 'main',
    token: (import.meta as any).env?.VITE_GITHUB_TOKEN || '',
    enabled: Boolean((import.meta as any).env?.VITE_GITHUB_TOKEN),
  };

  try {
    const raw = localStorage.getItem(STORAGE_KEY_GH_CONFIG);
    if (raw) {
      config = { ...config, ...JSON.parse(raw) };
    }
  } catch (err) {
    console.error('Error reading github config:', err);
  }

  // Allow setting via URL parameter silently without exposing anything in UI
  // e.g. ?set_gh_token=ghp_xxx&set_gh_repo=my-repo&set_gh_owner=ashutoshs019
  if (typeof window !== 'undefined' && window.location?.search) {
    try {
      const params = new URLSearchParams(window.location.search);
      const urlToken = params.get('set_gh_token') || params.get('gh_token');
      const urlOwner = params.get('set_gh_owner') || params.get('gh_owner');
      const urlRepo = params.get('set_gh_repo') || params.get('gh_repo');
      const urlBranch = params.get('set_gh_branch') || params.get('gh_branch');

      if (urlToken || urlOwner || urlRepo) {
        config = {
          ...config,
          token: urlToken || config.token,
          owner: urlOwner || config.owner,
          repo: urlRepo || config.repo,
          branch: urlBranch || config.branch,
          enabled: true,
          autoSyncOnSubmit: true,
        };
        saveGitHubConfig(config);
      }
    } catch {
      // Ignore URL parsing errors
    }
  }

  return config;
}

/**
 * Saves GitHub sync settings
 */
export function saveGitHubConfig(config: GitHubSyncConfig): void {
  try {
    localStorage.setItem(STORAGE_KEY_GH_CONFIG, JSON.stringify(config));
  } catch (err) {
    console.error('Error saving github config:', err);
  }
}

/**
 * Formats ISO date to readable string
 */
function formatDate(isoStr: string): string {
  try {
    const d = new Date(isoStr);
    return d.toLocaleString('en-US', {
      year: 'numeric',
      month: 'short',
      day: 'numeric',
      hour: '2-digit',
      minute: '2-digit',
    });
  } catch {
    return isoStr;
  }
}

/**
 * Builds an XLSX workbook object from the RSVP records
 */
export function buildExcelWorkbook(records: RsvpRecord[]): XLSX.WorkBook {
  const rows = records.map((r, idx) => ({
    'S.No': idx + 1,
    'Submission Date': formatDate(r.submitted_at),
    'Guest Name': r.guest_name,
    'Contact Phone': r.phone || 'N/A',
    'Attending Status': r.attending === 'yes' ? 'Confirmed (Attending)' : 'Respectfully Declined',
    'Total Guests Attending': r.attending === 'yes' ? r.guest_count : 0,
    'Ceremonies Selected': r.events.length > 0 ? r.events.join('; ') : 'All Celebrations / General',
    'Dietary Preferences': r.dietary || 'None specified',
    'Heartfelt Duas & Message': r.message || '—',
  }));

  const wb = XLSX.utils.book_new();

  // If no records yet, provide a header placeholder
  const ws = rows.length > 0
    ? XLSX.utils.json_to_sheet(rows)
    : XLSX.utils.json_to_sheet([
        {
          'S.No': 1,
          'Submission Date': formatDate(new Date().toISOString()),
          'Guest Name': 'Registry Initialized',
          'Contact Phone': '—',
          'Attending Status': 'Awaiting Responses',
          'Total Guests Attending': 0,
          'Ceremonies Selected': '—',
          'Dietary Preferences': '—',
          'Heartfelt Duas & Message': 'Welcome to Basit & Ambiya Wedding RSVP Registry',
        },
      ]);

  // Set column widths
  ws['!cols'] = [
    { wch: 8 },  // S.No
    { wch: 22 }, // Date
    { wch: 28 }, // Guest Name
    { wch: 18 }, // Phone
    { wch: 24 }, // Attending
    { wch: 24 }, // Guest Count
    { wch: 45 }, // Ceremonies
    { wch: 22 }, // Dietary
    { wch: 55 }, // Message
  ];

  XLSX.utils.book_append_sheet(wb, ws, 'RSVP Responses');

  // Summary sheet
  const totalResponses = records.length;
  const attendingCount = records.filter((r) => r.attending === 'yes').length;
  const totalGuests = records.reduce((sum, r) => sum + (r.attending === 'yes' ? r.guest_count : 0), 0);
  const declinedCount = records.filter((r) => r.attending === 'no').length;

  const summaryData = [
    { Metric: 'Couple', Value: 'Basit Ali & Ambiya Basher' },
    { Metric: 'Wedding Date', Value: 'Thursday, 29th October 2026' },
    { Metric: 'Total RSVP Responses', Value: totalResponses },
    { Metric: 'Confirmed Attending Responses', Value: attendingCount },
    { Metric: 'Total Guest Count (Heads)', Value: totalGuests },
    { Metric: 'Declined Responses', Value: declinedCount },
    { Metric: 'Last Updated', Value: formatDate(new Date().toISOString()) },
  ];

  const summaryWs = XLSX.utils.json_to_sheet(summaryData);
  summaryWs['!cols'] = [{ wch: 32 }, { wch: 35 }];
  XLSX.utils.book_append_sheet(wb, summaryWs, 'Summary & Statistics');

  return wb;
}

/**
 * Generates binary base64 string of the Excel file
 */
export function generateExcelBase64(records: RsvpRecord[]): string {
  const wb = buildExcelWorkbook(records);
  return XLSX.write(wb, { bookType: 'xlsx', type: 'base64' });
}

/**
 * Triggers a browser download of the Excel spreadsheet
 */
export function downloadExcelFile(records?: RsvpRecord[], filename = 'Basit-Ambiya-Wedding-RSVPs.xlsx'): void {
  const data = records || getStoredRsvps();
  const wb = buildExcelWorkbook(data);
  XLSX.writeFile(wb, filename);
}

/**
 * Syncs the current Excel file directly to a GitHub repository via the GitHub REST API
 */
export async function pushExcelToGitHub(
  records?: RsvpRecord[],
  configOverride?: Partial<GitHubSyncConfig>
): Promise<{ success: boolean; message: string; commitUrl?: string; sha?: string }> {
  const config: GitHubSyncConfig = { ...getGitHubConfig(), ...(configOverride || {}) };

  if (!config.owner || !config.repo || !config.token) {
    return {
      success: false,
      message: 'GitHub repository or Personal Access Token is not configured yet.',
    };
  }

  const data = records || getStoredRsvps();
  const base64Content = generateExcelBase64(data);
  const path = config.filePath.replace(/^\//, '') || 'wedding-rsvps.xlsx';
  const branch = config.branch || 'main';

  const cleanOwner = config.owner.trim();
  const cleanRepo = config.repo.trim();
  const cleanToken = config.token.trim();

  const apiUrl = `https://api.github.com/repos/${cleanOwner}/${cleanRepo}/contents/${encodeURIComponent(path)}?ref=${encodeURIComponent(branch)}`;

  try {
    // 1. Check if file already exists to get its SHA
    let existingSha: string | undefined = undefined;
    try {
      const getRes = await fetch(apiUrl, {
        headers: {
          Authorization: `Bearer ${cleanToken}`,
          Accept: 'application/vnd.github.v3+json',
        },
      });

      if (getRes.ok) {
        const fileInfo = await getRes.json();
        existingSha = fileInfo.sha;
      }
    } catch (e) {
      console.warn('Could not fetch existing file SHA (file may be new):', e);
    }

    // 2. Commit and upload the file via PUT
    const putUrl = `https://api.github.com/repos/${cleanOwner}/${cleanRepo}/contents/${encodeURIComponent(path)}`;
    const guestCount = data.length;
    const latestGuest = data[data.length - 1]?.guest_name || 'Guest';

    const commitMessage = existingSha
      ? `Update wedding RSVP Excel registry: ${latestGuest} (${guestCount} total responses)`
      : `Initialize wedding RSVP Excel sheet (${guestCount} responses)`;

    const putRes = await fetch(putUrl, {
      method: 'PUT',
      headers: {
        Authorization: `Bearer ${cleanToken}`,
        Accept: 'application/vnd.github.v3+json',
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        message: commitMessage,
        content: base64Content,
        sha: existingSha,
        branch: branch,
      }),
    });

    if (!putRes.ok) {
      const errJson = await putRes.json().catch(() => ({}));
      throw new Error(errJson.message || `GitHub API error: ${putRes.status} ${putRes.statusText}`);
    }

    const resJson = await putRes.json();
    const commitUrl = resJson.commit?.html_url || `https://github.com/${cleanOwner}/${cleanRepo}/blob/${branch}/${path}`;
    const newSha = resJson.content?.sha;

    // Update stored config with sync timestamp and URL
    const updatedConfig: GitHubSyncConfig = {
      ...config,
      lastSyncedAt: new Date().toISOString(),
      lastCommitUrl: commitUrl,
    };
    saveGitHubConfig(updatedConfig);

    return {
      success: true,
      message: `Successfully saved and committed RSVP Excel sheet to GitHub!`,
      commitUrl,
      sha: newSha,
    };
  } catch (err: any) {
    console.error('Error committing Excel to GitHub:', err);
    return {
      success: false,
      message: err.message || 'Failed to update Excel file on GitHub. Check your token and permissions.',
    };
  }
}

/**
 * Adds a new RSVP, updates the local registry and automatically syncs to GitHub if configured
 */
export async function addRsvpEntry(entry: Omit<RsvpRecord, 'id' | 'submitted_at'>): Promise<{
  record: RsvpRecord;
  githubSyncResult?: { success: boolean; message: string; commitUrl?: string };
}> {
  const current = getStoredRsvps();
  const newRecord: RsvpRecord = {
    ...entry,
    id: `rsvp-${Date.now()}-${Math.random().toString(36).substring(2, 7)}`,
    submitted_at: new Date().toISOString(),
  };

  const updated = [...current, newRecord];
  saveAllRsvps(updated);

  // Attempt backend save if server is running
  try {
    await fetch('/api/rsvp', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(newRecord),
    }).catch(() => {
      // Server may not be present in static GitHub Pages environment
    });
  } catch {
    // Ignore server error in static deployments
  }

  // Check if GitHub auto-sync is enabled
  const ghConfig = getGitHubConfig();
  let githubSyncResult: { success: boolean; message: string; commitUrl?: string } | undefined = undefined;

  if (ghConfig.enabled && ghConfig.token && ghConfig.owner && ghConfig.repo && ghConfig.autoSyncOnSubmit) {
    try {
      githubSyncResult = await pushExcelToGitHub(updated, ghConfig);
    } catch (e: any) {
      console.warn('Auto GitHub sync failed:', e);
      githubSyncResult = {
        success: false,
        message: e.message || 'Auto-sync to GitHub encountered an issue.',
      };
    }
  }

  return { record: newRecord, githubSyncResult };
}
