import {
  AccountInfo,
  AuthenticationResult,
  BrowserCacheLocation,
  PublicClientApplication,
} from "@azure/msal-browser";
import {
  Apps24Regular,
  ArrowLeft24Regular,
  ArrowRight24Regular,
  BookOpen24Regular,
  ChevronDown20Regular,
  ChevronRight20Regular,
  CheckmarkCircle24Regular,
  Code24Regular,
  Database24Regular,
  Dismiss24Regular,
  DocumentText24Regular,
  FullScreenMaximize24Regular,
  History24Regular,
  Person24Regular,
  QuestionCircle24Regular,
  Search24Regular,
  Send24Regular,
  Share24Regular,
  SignOut24Regular,
  Star24Filled,
  Star24Regular,
} from "@fluentui/react-icons";
import type { ReactNode } from "react";
import { useEffect, useMemo, useState } from "react";
import { partnerCenterAuthConfig } from "./authConfig";
import { learnBaseUrl, Scenario, scenarios } from "./data/scenarios";

type HttpMethod = "GET" | "POST" | "PATCH" | "PUT" | "DELETE";

type HistoryEntry = {
  id: string;
  createdAt: string;
  method: HttpMethod;
  url: string;
  status: number | "network";
  durationMs: number;
  requestBody?: string;
  responsePreview: string;
};

type SharedQuery = {
  method: HttpMethod;
  endpoint: string;
  urlParameterValues: Record<string, string>;
  body: string;
  requestHeaders: string;
  scenarioId?: string;
};

const historyKey = "partner-center-api-explorer-history";
const favoritesKey = "partner-center-api-explorer-favorites";
const shareQueryParam = "pcq";
const wizardDismissedKey = "partner-center-api-explorer-wizard-dismissed";
const defaultRequestHeaders = `Accept: application/json
Content-Type: application/json`;

const wizardSteps = [
  {
    title: "Sign in with your work account",
    description:
      "Start by selecting Sign in in the top-right corner. The explorer uses your Microsoft 365 work account to request Partner Center API access.",
    icon: Person24Regular,
  },
  {
    title: "Choose a scenario or search",
    description:
      "Use the Scenarios tab on the left to browse Partner Center API areas, or search for APIs such as billing, orders, customers, or growth margins.",
    icon: Search24Regular,
  },
  {
    title: "Edit URL parameters and send",
    description:
      "When an API URL contains placeholders, text boxes appear so you can enter values. You can also edit the URL directly, then select Send.",
    icon: Send24Regular,
  },
  {
    title: "Review body, headers, and expanded response",
    description:
      "Use Request Body and Request Headers tabs before sending. After a response arrives, use Expand to view response body and response headers in a larger dialog.",
    icon: FullScreenMaximize24Regular,
  },
  {
    title: "Share queries and open documentation",
    description:
      "Use the inline Share icon to copy a deep link for the prepared query. Use the Docs icon to open the matching Microsoft Learn article.",
    icon: Share24Regular,
  },
] as const;

function isHttpMethod(value: string): value is HttpMethod {
  return ["GET", "POST", "PATCH", "PUT", "DELETE"].includes(value);
}

function encodeSharePayload(payload: SharedQuery) {
  const json = JSON.stringify(payload);
  return btoa(encodeURIComponent(json))
    .replace(/\+/g, "-")
    .replace(/\//g, "_")
    .replace(/=+$/, "");
}

function decodeSharePayload(value: string): SharedQuery | null {
  try {
    const padded = value.replace(/-/g, "+").replace(/_/g, "/").padEnd(Math.ceil(value.length / 4) * 4, "=");
    const parsed = JSON.parse(decodeURIComponent(atob(padded))) as Partial<SharedQuery>;

    if (!parsed.endpoint || !parsed.method || !isHttpMethod(parsed.method)) {
      return null;
    }

    return {
      method: parsed.method,
      endpoint: parsed.endpoint,
      urlParameterValues: parsed.urlParameterValues ?? {},
      body: parsed.body ?? "",
      requestHeaders: parsed.requestHeaders ?? defaultRequestHeaders,
      scenarioId: parsed.scenarioId,
    };
  } catch {
    return null;
  }
}

function readSharedQuery() {
  return decodeSharePayload(new URLSearchParams(window.location.search).get(shareQueryParam) ?? "");
}

function loadHistory(): HistoryEntry[] {
  const saved = localStorage.getItem(historyKey);
  if (!saved) {
    return [];
  }

  try {
    const parsed = JSON.parse(saved) as HistoryEntry[];
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function saveHistory(entries: HistoryEntry[]) {
  localStorage.setItem(historyKey, JSON.stringify(entries.slice(0, 25)));
}

function loadFavoriteIds(): string[] {
  const saved = localStorage.getItem(favoritesKey);
  if (!saved) {
    return [];
  }

  try {
    const parsed = JSON.parse(saved) as string[];
    return Array.isArray(parsed) ? parsed.filter((id) => scenarios.some((scenario) => scenario.id === id)) : [];
  } catch {
    return [];
  }
}

function saveFavoriteIds(ids: string[]) {
  localStorage.setItem(favoritesKey, JSON.stringify(ids));
}

function normalizeEndpoint(endpoint: string) {
  if (/^https?:\/\//i.test(endpoint)) {
    return endpoint;
  }

  return `${partnerCenterAuthConfig.baseUrl.replace(/\/$/, "")}/${endpoint.replace(/^\//, "")}`;
}

function getUrlParameters(endpoint: string) {
  return Array.from(endpoint.matchAll(/\{([^}]+)\}/g))
    .map((match) => match[1])
    .filter((name, index, names) => names.indexOf(name) === index);
}

function applyUrlParameters(endpoint: string, values: Record<string, string>) {
  return endpoint.replace(/\{([^}]+)\}/g, (_, name: string) => {
    const value = values[name]?.trim();
    return value ? encodeURIComponent(value) : `{${name}}`;
  });
}

function formatJson(value: string) {
  if (!value.trim()) {
    return "";
  }

  return JSON.stringify(JSON.parse(value), null, 2);
}

function formatResponseBody(value: string) {
  if (!value.trim()) {
    return "";
  }

  try {
    return JSON.stringify(JSON.parse(value), null, 2);
  } catch {
    return value;
  }
}

function getInitials(account: AccountInfo | null) {
  const displayName = account?.name || account?.username || "";
  const initials = displayName
    .split(/[.\s@_-]+/)
    .filter(Boolean)
    .slice(0, 2)
    .map((part) => part[0]?.toUpperCase())
    .join("");

  return initials || "PC";
}

function getLearnUrl(learnPath: string) {
  if (/^https?:\/\//i.test(learnPath)) {
    return learnPath;
  }

  if (learnPath.startsWith("/")) {
    return `https://learn.microsoft.com${learnPath}`;
  }

  return `${learnBaseUrl}${learnPath}`;
}

function getRedirectUri() {
  return window.location.origin + window.location.pathname;
}

function getAuthErrorMessage(error: unknown) {
  const message = error instanceof Error ? error.message : "Sign-in failed.";
  if (message.includes("AADSTS9002326") || message.includes("9002326")) {
    return `Sign-in failed because this Entra app registration is not configured as a Single-page application. In Entra ID, add this exact redirect URI under Authentication > Single-page application: ${getRedirectUri()}. If it exists under Web, remove it from Web and add it under Single-page application.`;
  }

  return message;
}

function renderJsonTokens(value: string) {
  const tokenPattern =
    /("(?:\\u[\da-fA-F]{4}|\\[^u]|[^\\"])*"(?=\s*:)|"(?:\\u[\da-fA-F]{4}|\\[^u]|[^\\"])*"|true|false|null|-?\d+(?:\.\d+)?(?:[eE][+-]?\d+)?)/g;
  const parts: ReactNode[] = [];
  let lastIndex = 0;
  let index = 0;

  for (const match of value.matchAll(tokenPattern)) {
    const token = match[0];
    const offset = match.index ?? 0;
    if (offset > lastIndex) {
      parts.push(value.slice(lastIndex, offset));
    }

    const className = token.startsWith("\"")
      ? value.slice(offset + token.length).trimStart().startsWith(":")
        ? "json-key"
        : "json-string"
      : token === "true" || token === "false"
        ? "json-boolean"
        : token === "null"
          ? "json-null"
          : "json-number";

    parts.push(
      <span className={className} key={`${token}-${index}`}>
        {token}
      </span>,
    );
    lastIndex = offset + token.length;
    index += 1;
  }

  if (lastIndex < value.length) {
    parts.push(value.slice(lastIndex));
  }

  return parts;
}

function renderResponseBody(value: string) {
  if (!value) {
    return "Run a request to see the Partner Center API response.";
  }

  try {
    return renderJsonTokens(JSON.stringify(JSON.parse(value), null, 2));
  } catch {
    return value;
  }
}

function getResponseLines(value: string) {
  if (!value) {
    return ["Run a request to see the Partner Center API response."];
  }

  try {
    return JSON.stringify(JSON.parse(value), null, 2).split("\n");
  } catch {
    return value.split(/\r?\n/);
  }
}

function JsonFormatter({ value }: { value: string }) {
  return (
    <div className="json-formatter" role="region" aria-label="Formatted JSON response">
      {getResponseLines(value).map((line, index) => (
        <div className="json-line" key={`${index}-${line}`}>
          <span className="line-number">{index + 1}</span>
          <code>{renderJsonTokens(line)}</code>
        </div>
      ))}
    </div>
  );
}

function parseRequestHeaders(value: string) {
  const headers: Record<string, string> = {};

  for (const line of value.split(/\r?\n/)) {
    if (!line.trim()) {
      continue;
    }

    const separatorIndex = line.indexOf(":");
    if (separatorIndex <= 0) {
      throw new Error(`Invalid request header: "${line}". Use "Header-Name: value".`);
    }

    const name = line.slice(0, separatorIndex).trim();
    const headerValue = line.slice(separatorIndex + 1).trim();
    if (!name || !headerValue) {
      throw new Error(`Invalid request header: "${line}". Use "Header-Name: value".`);
    }

    headers[name] = headerValue;
  }

  return headers;
}

export default function App() {
  const sharedQuery = useMemo(() => readSharedQuery(), []);
  const initialScenario = scenarios.find((scenario) => scenario.id === sharedQuery?.scenarioId) ?? scenarios[3];
  const [account, setAccount] = useState<AccountInfo | null>(null);
  const [method, setMethod] = useState<HttpMethod>(sharedQuery?.method ?? "GET");
  const [endpoint, setEndpoint] = useState(sharedQuery?.endpoint ?? "/customers");
  const [urlParameterValues, setUrlParameterValues] = useState<Record<string, string>>(
    sharedQuery?.urlParameterValues ?? {},
  );
  const [body, setBody] = useState(sharedQuery?.body ?? "");
  const [requestHeaders, setRequestHeaders] = useState(sharedQuery?.requestHeaders ?? defaultRequestHeaders);
  const [query, setQuery] = useState("");
  const [selectedScenario, setSelectedScenario] = useState<Scenario>(initialScenario);
  const [history, setHistory] = useState<HistoryEntry[]>(loadHistory);
  const [result, setResult] = useState<{
    status: string;
    body: string;
    headers: string;
    durationMs?: number;
  } | null>(null);
  const [busy, setBusy] = useState(false);
  const [authMessage, setAuthMessage] = useState("");
  const [profileOpen, setProfileOpen] = useState(false);
  const [sideTab, setSideTab] = useState<"scenarios" | "collection" | "history">("scenarios");
  const [requestTab, setRequestTab] = useState<"body" | "headers">("body");
  const [responseTab, setResponseTab] = useState<"body" | "headers">("body");
  const [expandedResponseTab, setExpandedResponseTab] = useState<"body" | "headers">("body");
  const [responseExpanded, setResponseExpanded] = useState(false);
  const [shareMessage, setShareMessage] = useState("");
  const [favoriteIds, setFavoriteIds] = useState<string[]>(loadFavoriteIds);
  const [wizardOpen, setWizardOpen] = useState(() => localStorage.getItem(wizardDismissedKey) !== "true");
  const [wizardStep, setWizardStep] = useState(0);
  const [expandedCategories, setExpandedCategories] = useState<Record<string, boolean>>({});

  const msal = useMemo(
    () =>
      new PublicClientApplication({
        auth: {
          clientId: partnerCenterAuthConfig.clientId,
          authority: `https://login.microsoftonline.com/${partnerCenterAuthConfig.tenantId}`,
          redirectUri: getRedirectUri(),
        },
        cache: {
          cacheLocation: BrowserCacheLocation.LocalStorage,
        },
      }),
    [],
  );

  useEffect(() => {
    let mounted = true;

    async function restoreSession() {
      await msal.initialize();
      const activeAccount = msal.getActiveAccount() ?? msal.getAllAccounts()[0] ?? null;
      if (mounted) {
        setAccount(activeAccount);
      }
    }

    restoreSession().catch((error: unknown) => {
      if (mounted) {
        setAuthMessage(error instanceof Error ? error.message : "Could not restore the sign-in session.");
      }
    });

    return () => {
      mounted = false;
    };
  }, [msal]);

  const filteredScenarios = scenarios.filter((scenario) => {
    const haystack =
      `${scenario.category} ${scenario.subcategory ?? ""} ${scenario.title} ${scenario.description} ${scenario.tags.join(" ")}`.toLowerCase();
    return haystack.includes(query.toLowerCase());
  });

  const scenarioGroups = filteredScenarios.reduce<Record<string, Scenario[]>>((groups, scenario) => {
    groups[scenario.category] = [...(groups[scenario.category] ?? []), scenario];
    return groups;
  }, {});
  const favoriteScenarios = favoriteIds
    .map((id) => scenarios.find((scenario) => scenario.id === id))
    .filter((scenario): scenario is Scenario => Boolean(scenario));

  const urlParameters = getUrlParameters(endpoint);

  useEffect(() => {
    setUrlParameterValues((current) =>
      Object.fromEntries(urlParameters.map((name) => [name, current[name] ?? ""])),
    );
  }, [endpoint]);

  async function signIn() {
    if (!partnerCenterAuthConfig.clientId.trim()) {
      setAuthMessage("This site needs an internal Entra app registration client ID before sign-in can be used.");
      return;
    }

    setBusy(true);
    setAuthMessage("");
    try {
      await msal.initialize();
      const response = await msal.loginPopup({
        scopes: partnerCenterAuthConfig.scopes,
        prompt: "select_account",
      });
      msal.setActiveAccount(response.account);
      setAccount(response.account);
      setProfileOpen(false);
    } catch (error) {
      setAuthMessage(getAuthErrorMessage(error));
    } finally {
      setBusy(false);
    }
  }

  async function signOut() {
    if (!account) {
      return;
    }

    setBusy(true);
    setAuthMessage("");
    try {
      await msal.logoutPopup({
        account,
        mainWindowRedirectUri: getRedirectUri(),
      });
      msal.setActiveAccount(null);
      setAccount(null);
      setProfileOpen(false);
    } catch (error) {
      setAuthMessage(error instanceof Error ? error.message : "Sign-out failed.");
    } finally {
      setBusy(false);
    }
  }

  async function acquireToken(): Promise<AuthenticationResult> {
    if (!account) {
      throw new Error("Sign in before sending a request.");
    }

    await msal.initialize();
    try {
      return await msal.acquireTokenSilent({ scopes: partnerCenterAuthConfig.scopes, account });
    } catch {
      return msal.acquireTokenPopup({ scopes: partnerCenterAuthConfig.scopes, account });
    }
  }

  async function sendRequest() {
    setBusy(true);
    setResult(null);

    const missingParameters = urlParameters.filter((name) => !urlParameterValues[name]?.trim());
    if (missingParameters.length > 0) {
      setResult({
        status: "Missing URL parameters",
        body: `Enter values for: ${missingParameters.join(", ")}`,
        headers: "",
      });
      setBusy(false);
      return;
    }

    const url = normalizeEndpoint(applyUrlParameters(endpoint, urlParameterValues));
    const startedAt = performance.now();
    try {
      const parsedHeaders = parseRequestHeaders(requestHeaders);
      const token = await acquireToken();
      const requestBody = ["POST", "PATCH", "PUT"].includes(method) ? formatJson(body) || body : undefined;
      const response = await fetch(url, {
        method,
        headers: {
          ...parsedHeaders,
          Authorization: `Bearer ${token.accessToken}`,
          "MS-RequestId": crypto.randomUUID(),
          "MS-CorrelationId": crypto.randomUUID(),
        },
        body: requestBody,
      });
      const durationMs = Math.round(performance.now() - startedAt);
      const responseText = await response.text();
      const formattedBody = formatResponseBody(responseText);
      const responseHeaders = Array.from(response.headers.entries())
        .map(([key, value]) => `${key}: ${value}`)
        .join("\n");
      const nextHistory = [
        {
          id: crypto.randomUUID(),
          createdAt: new Date().toISOString(),
          method,
          url,
          status: response.status,
          durationMs,
          requestBody,
          responsePreview: formattedBody.slice(0, 500),
        },
        ...history,
      ];
      setHistory(nextHistory);
      saveHistory(nextHistory);
      setResult({
        status: `${response.status} ${response.statusText}`,
        body: formattedBody,
        headers: responseHeaders,
        durationMs,
      });
    } catch (error) {
      const durationMs = Math.round(performance.now() - startedAt);
      const message = error instanceof Error ? error.message : "The request failed.";
      const nextHistory = [
        {
          id: crypto.randomUUID(),
          createdAt: new Date().toISOString(),
          method,
          url,
          status: "network" as const,
          durationMs,
          requestBody: body,
          responsePreview: message,
        },
        ...history,
      ];
      setHistory(nextHistory);
      saveHistory(nextHistory);
      setResult({
        status: "Request failed",
        body: message,
        headers: "If this is a browser CORS failure, use a lightweight authenticated proxy for production deployments.",
        durationMs,
      });
    } finally {
      setBusy(false);
    }
  }

  function applyScenario(scenario: Scenario) {
    setSelectedScenario(scenario);
    setMethod(scenario.method);
    setEndpoint(scenario.endpoint);
    setUrlParameterValues({});
    setBody(scenario.body ?? "");
    setRequestTab(scenario.body ? "body" : "headers");
    setExpandedCategories((current) => ({ ...current, [scenario.category]: true }));
  }

  function toggleFavorite(scenario: Scenario) {
    setFavoriteIds((current) => {
      const next = current.includes(scenario.id)
        ? current.filter((id) => id !== scenario.id)
        : [...current, scenario.id];
      saveFavoriteIds(next);
      return next;
    });
  }

  function toggleCategory(category: string) {
    setExpandedCategories((current) => ({ ...current, [category]: !current[category] }));
  }

  function groupBySubcategory(group: Scenario[]) {
    return group.reduce<Record<string, Scenario[]>>((subgroups, scenario) => {
      const subcategory = scenario.subcategory ?? "APIs";
      subgroups[subcategory] = [...(subgroups[subcategory] ?? []), scenario];
      return subgroups;
    }, {});
  }

  async function shareQuery() {
    const payload: SharedQuery = {
      method,
      endpoint,
      urlParameterValues,
      body,
      requestHeaders,
      scenarioId: selectedScenario.id,
    };
    const url = new URL(window.location.href);
    url.searchParams.set(shareQueryParam, encodeSharePayload(payload));

    try {
      await navigator.clipboard.writeText(url.toString());
      setShareMessage("Share link copied to clipboard.");
      window.setTimeout(() => setShareMessage(""), 2600);
    } catch {
      setShareMessage(url.toString());
    }
  }

  function clearHistory() {
    setHistory([]);
    saveHistory([]);
  }

  function closeWizard() {
    localStorage.setItem(wizardDismissedKey, "true");
    setWizardOpen(false);
  }

  function launchWizard() {
    setWizardStep(0);
    setWizardOpen(true);
  }

  return (
    <main className="app-shell">
      <header className="suite-bar">
        <div className="suite-left">
          <button className="app-launcher" aria-label="App launcher">
            <Apps24Regular />
          </button>
          <div className="suite-brand">
            <strong>Partner Center API Explorer</strong>
          </div>
        </div>
        <div className="suite-actions">
          <button className="header-icon-button" onClick={launchWizard} aria-label="Launch getting started wizard" title="Getting started">
            <QuestionCircle24Regular />
          </button>
          {account ? (
            <div className="profile-area">
              <button className="profile-button" onClick={() => setProfileOpen(!profileOpen)}>
                <span className="avatar">{getInitials(account)}</span>
                <span className="profile-text">
                  <strong>{account.name ?? "Signed in user"}</strong>
                  <small>{account.username}</small>
                </span>
              </button>
              {profileOpen && (
                <div className="profile-menu">
                  <div className="profile-menu-head">
                    <span className="avatar large">{getInitials(account)}</span>
                    <div>
                      <strong>{account.name ?? "Signed in user"}</strong>
                      <small>{account.username}</small>
                    </div>
                  </div>
                  <button className="ghost full-width" onClick={signOut} disabled={busy}>
                    <span className="button-content">
                      <SignOut24Regular />
                      Sign out
                    </span>
                  </button>
                </div>
              )}
            </div>
          ) : (
            <button className="primary" onClick={signIn} disabled={busy}>
              <span className="button-content">
                <Person24Regular />
                {busy ? "Opening sign in..." : "Sign in"}
              </span>
            </button>
          )}
        </div>
      </header>

      {authMessage && <p className="message danger">{authMessage}</p>}
      {shareMessage && <div className="toast success-toast" role="status">{shareMessage}</div>}

      <section className="layout">
        <aside className="panel side-panel">
          <div className="tab-list" role="tablist" aria-label="Explorer side panel">
            <button
              className={`tab-button ${sideTab === "scenarios" ? "active" : ""}`}
              onClick={() => setSideTab("scenarios")}
              role="tab"
              aria-selected={sideTab === "scenarios"}
            >
              <span className="button-content">
                <Database24Regular />
                Scenarios
              </span>
            </button>
            <button
              className={`tab-button ${sideTab === "collection" ? "active" : ""}`}
              onClick={() => setSideTab("collection")}
              role="tab"
              aria-selected={sideTab === "collection"}
            >
              <span className="button-content">
                <Star24Regular />
                Collection
              </span>
            </button>
            <button
              className={`tab-button ${sideTab === "history" ? "active" : ""}`}
              onClick={() => setSideTab("history")}
              role="tab"
              aria-selected={sideTab === "history"}
            >
              <span className="button-content">
                <History24Regular />
                History
              </span>
            </button>
          </div>

          {sideTab === "scenarios" ? (
            <section>
              <div className="panel-heading compact">
                <div>
                  <p className="eyebrow">Examples</p>
                  <h2>Scenarios</h2>
                </div>
                <a href={getLearnUrl("scenarios")} target="_blank" rel="noreferrer">
                  <span className="button-content">
                    <BookOpen24Regular />
                    Learn
                  </span>
                </a>
              </div>
              <div className="search-input-wrap">
                <Search24Regular />
                <input
                  className="input scenario-filter"
                  aria-label="Filter scenarios"
                  placeholder="Filter by customers, billing, orders..."
                  value={query}
                  onChange={(event) => setQuery(event.target.value)}
                />
              </div>
              <div className="scenario-list">
                {Object.entries(scenarioGroups).map(([category, group]) => (
                  <section className="scenario-group" key={category}>
                    <button
                      className="scenario-group-header"
                      onClick={() => toggleCategory(category)}
                      aria-expanded={Boolean(expandedCategories[category])}
                    >
                      <span className="chevron">
                        {expandedCategories[category] ? <ChevronDown20Regular /> : <ChevronRight20Regular />}
                      </span>
                      <strong>{category}</strong>
                      <small>{group.length}</small>
                    </button>
                    {expandedCategories[category] && (
                      <div className="scenario-group-content">
                        {Object.entries(groupBySubcategory(group)).map(([subcategory, subgroup]) => (
                          <div className="scenario-subgroup" key={`${category}-${subcategory}`}>
                            <h3>{subcategory}</h3>
                            {subgroup.map((scenario) => (
                              <button
                                className={`scenario-card ${selectedScenario.id === scenario.id ? "active" : ""}`}
                                key={scenario.id}
                                onClick={() => applyScenario(scenario)}
                              >
                                <span className="method-pill">{scenario.method}</span>
                                <span
                                  className={`favorite-action ${favoriteIds.includes(scenario.id) ? "favorited" : ""}`}
                                  role="button"
                                  tabIndex={0}
                                  title={favoriteIds.includes(scenario.id) ? "Remove from collection" : "Add to collection"}
                                  aria-label={favoriteIds.includes(scenario.id) ? "Remove from collection" : "Add to collection"}
                                  onClick={(event) => {
                                    event.stopPropagation();
                                    toggleFavorite(scenario);
                                  }}
                                  onKeyDown={(event) => {
                                    if (event.key === "Enter" || event.key === " ") {
                                      event.preventDefault();
                                      event.stopPropagation();
                                      toggleFavorite(scenario);
                                    }
                                  }}
                                >
                                  {favoriteIds.includes(scenario.id) ? <Star24Filled /> : <Star24Regular />}
                                </span>
                                <strong>{scenario.title}</strong>
                                <small>{scenario.description}</small>
                              </button>
                            ))}
                          </div>
                        ))}
                      </div>
                    )}
                  </section>
                ))}
              </div>
            </section>
          ) : sideTab === "collection" ? (
            <section>
              <div className="panel-heading compact">
                <div>
                  <p className="eyebrow">Saved</p>
                  <h2>My API Collection</h2>
                </div>
                <span className="pill">{favoriteScenarios.length}</span>
              </div>
              <div className="history-list">
                {favoriteScenarios.length === 0 && (
                  <p className="empty">Select the star on any scenario to add it to your collection.</p>
                )}
                {favoriteScenarios.map((scenario) => (
                  <button
                    className={`history-item ${selectedScenario.id === scenario.id ? "active" : ""}`}
                    key={scenario.id}
                    onClick={() => applyScenario(scenario)}
                  >
                    <span className="history-meta">
                      <strong>{scenario.method}</strong>
                      <span
                        className="favorite-action favorited inline"
                        role="button"
                        tabIndex={0}
                        title="Remove from collection"
                        aria-label="Remove from collection"
                        onClick={(event) => {
                          event.stopPropagation();
                          toggleFavorite(scenario);
                        }}
                        onKeyDown={(event) => {
                          if (event.key === "Enter" || event.key === " ") {
                            event.preventDefault();
                            event.stopPropagation();
                            toggleFavorite(scenario);
                          }
                        }}
                      >
                        <Star24Filled />
                      </span>
                    </span>
                    <span className="history-url">{scenario.title}</span>
                    <span className="history-status">{scenario.category}{scenario.subcategory ? ` · ${scenario.subcategory}` : ""}</span>
                  </button>
                ))}
              </div>
            </section>
          ) : (
            <section>
              <div className="panel-heading compact">
                <div>
                  <p className="eyebrow">Local</p>
                  <h2>History</h2>
                </div>
                <button className="ghost" onClick={clearHistory} disabled={!history.length}>
                  <span className="button-content">
                    <History24Regular />
                    Clear
                  </span>
                </button>
              </div>
              <div className="history-list">
                {history.length === 0 && <p className="empty">Your last 25 API actions will appear here.</p>}
                {history.map((entry) => (
                  <button
                    className="history-item"
                    key={entry.id}
                    onClick={() => {
                      setMethod(entry.method);
                      setEndpoint(entry.url.replace(partnerCenterAuthConfig.baseUrl, ""));
                      setBody(entry.requestBody ?? "");
                    }}
                  >
                    <span className="history-meta">
                      <strong>{entry.method}</strong>
                      <small>{new Date(entry.createdAt).toLocaleString()}</small>
                    </span>
                    <span className="history-url">{entry.url}</span>
                    <span className="history-status">
                      Status: {entry.status} · {entry.durationMs} ms
                    </span>
                  </button>
                ))}
              </div>
            </section>
          )}
        </aside>

        <section className="main-stack">
          <section className="panel">
            <div className="request-line">
              <select
                className="input method"
                value={method}
                onChange={(event) => setMethod(event.target.value as HttpMethod)}
              >
                <option>GET</option>
                <option>POST</option>
                <option>PATCH</option>
                <option>PUT</option>
                <option>DELETE</option>
              </select>
              <div className="endpoint-group">
                <input
                  className="input endpoint-input"
                  aria-label="API endpoint"
                  value={endpoint}
                  onChange={(event) => setEndpoint(event.target.value)}
                />
                <button className="inline-input-button" onClick={shareQuery} aria-label="Share query" title="Share query">
                  <Share24Regular />
                </button>
                <a
                  className="inline-input-button"
                  href={getLearnUrl(selectedScenario.learnPath)}
                  target="_blank"
                  rel="noreferrer"
                  aria-label="Open scenario documentation"
                  title="Scenario docs"
                >
                  <BookOpen24Regular />
                </a>
              </div>
              <button className="primary icon-button" onClick={sendRequest} disabled={busy || !account} aria-label="Send request" title={busy ? "Working..." : "Send request"}>
                <Send24Regular />
              </button>
            </div>
            {urlParameters.length > 0 && (
              <div className="parameter-grid">
                {urlParameters.map((name) => (
                  <label key={name}>
                    {name}
                    <input
                      className="input"
                      value={urlParameterValues[name] ?? ""}
                      onChange={(event) =>
                        setUrlParameterValues({
                          ...urlParameterValues,
                          [name]: event.target.value,
                        })
                      }
                      placeholder={`Enter ${name}`}
                    />
                  </label>
                ))}
              </div>
            )}
            <div className="content-tabs" role="tablist" aria-label="Request editor">
              <button
                className={`content-tab ${requestTab === "body" ? "active" : ""}`}
                onClick={() => setRequestTab("body")}
                role="tab"
                aria-selected={requestTab === "body"}
              >
                <span className="button-content">
                  <DocumentText24Regular />
                  Request Body
                </span>
              </button>
              <button
                className={`content-tab ${requestTab === "headers" ? "active" : ""}`}
                onClick={() => setRequestTab("headers")}
                role="tab"
                aria-selected={requestTab === "headers"}
              >
                <span className="button-content">
                  <Code24Regular />
                  Request Headers
                </span>
              </button>
            </div>
            {requestTab === "body" ? (
              <div>
                <textarea
                  className="input code-box"
                  aria-label="JSON request body"
                  value={body}
                  onChange={(event) => setBody(event.target.value)}
                  placeholder="Required for POST, PATCH, and PUT requests"
                />
              </div>
            ) : (
              <label>
                Headers
                <textarea
                  className="input code-box header-box"
                  value={requestHeaders}
                  onChange={(event) => setRequestHeaders(event.target.value)}
                  placeholder="Header-Name: value"
                />
              </label>
            )}
          </section>

          <section className="panel result-panel">
            <div className="panel-heading">
              <div>
                <span className="response-status">{result?.status ?? "No request sent"}</span>
              </div>
              <div className="request-actions">
                {result?.durationMs && <span className="pill">{result.durationMs} ms</span>}
                <button
                  className="ghost icon-button"
                  onClick={() => {
                    setExpandedResponseTab(responseTab);
                    setResponseExpanded(true);
                  }}
                  aria-label="Expand response"
                  title="Expand response"
                >
                  <FullScreenMaximize24Regular />
                </button>
              </div>
            </div>
            <div className="content-tabs" role="tablist" aria-label="Response viewer">
              <button
                className={`content-tab ${responseTab === "body" ? "active" : ""}`}
                onClick={() => setResponseTab("body")}
                role="tab"
                aria-selected={responseTab === "body"}
              >
                <span className="button-content">
                  <DocumentText24Regular />
                  Response Body
                </span>
              </button>
              <button
                className={`content-tab ${responseTab === "headers" ? "active" : ""}`}
                onClick={() => setResponseTab("headers")}
                role="tab"
                aria-selected={responseTab === "headers"}
              >
                <span className="button-content">
                  <Code24Regular />
                  Response Headers
                </span>
              </button>
            </div>
            {responseTab === "body" ? (
              <JsonFormatter value={result?.body ?? ""} />
            ) : (
              <pre className="response-box headers-view">
                {result?.headers || "Run a request to see the response headers."}
              </pre>
            )}
          </section>
        </section>
      </section>
      {responseExpanded && (
        <div className="modal-backdrop" role="presentation">
          <section className="response-modal" role="dialog" aria-modal="true" aria-label="Expanded response viewer">
            <div className="modal-header">
              <div>
                <p className="eyebrow">Response</p>
                <span className="response-status">{result?.status ?? "No request sent"}</span>
              </div>
              <button className="ghost icon-button" onClick={() => setResponseExpanded(false)} aria-label="Close expanded response" title="Close">
                <Dismiss24Regular />
              </button>
            </div>
            <div className="content-tabs" role="tablist" aria-label="Expanded response viewer">
              <button
                className={`content-tab ${expandedResponseTab === "body" ? "active" : ""}`}
                onClick={() => setExpandedResponseTab("body")}
                role="tab"
                aria-selected={expandedResponseTab === "body"}
              >
                <span className="button-content">
                  <DocumentText24Regular />
                  Response Body
                </span>
              </button>
              <button
                className={`content-tab ${expandedResponseTab === "headers" ? "active" : ""}`}
                onClick={() => setExpandedResponseTab("headers")}
                role="tab"
                aria-selected={expandedResponseTab === "headers"}
              >
                <span className="button-content">
                  <Code24Regular />
                  Response Headers
                </span>
              </button>
            </div>
            <div className="modal-content">
              {expandedResponseTab === "body" ? (
                <JsonFormatter value={result?.body ?? ""} />
              ) : (
                <pre className="response-box headers-view">
                  {result?.headers || "Run a request to see the response headers."}
                </pre>
              )}
            </div>
          </section>
        </div>
      )}
      {wizardOpen && (
        <div className="modal-backdrop wizard-backdrop" role="presentation">
          <section className="wizard-modal" role="dialog" aria-modal="true" aria-label="Getting started wizard">
            <div className="wizard-header">
              <div>
                <p className="eyebrow">Getting started</p>
                <h2>Learn the Partner Center API Explorer</h2>
              </div>
              <button className="ghost icon-button" onClick={closeWizard} aria-label="Close wizard" title="Close">
                <Dismiss24Regular />
              </button>
            </div>

            <div className="wizard-progress" aria-label="Wizard progress">
              {wizardSteps.map((step, index) => (
                <button
                  className={`wizard-dot ${index === wizardStep ? "active" : ""} ${index < wizardStep ? "complete" : ""}`}
                  key={step.title}
                  onClick={() => setWizardStep(index)}
                  aria-label={`Go to step ${index + 1}: ${step.title}`}
                >
                  {index < wizardStep ? <CheckmarkCircle24Regular /> : index + 1}
                </button>
              ))}
            </div>

            <div className="wizard-content">
              {(() => {
                const step = wizardSteps[wizardStep];
                const StepIcon = step.icon;
                return (
                  <>
                    <div className="wizard-illustration">
                      <StepIcon />
                    </div>
                    <p className="wizard-step-count">
                      Step {wizardStep + 1} of {wizardSteps.length}
                    </p>
                    <h3>{step.title}</h3>
                    <p>{step.description}</p>
                  </>
                );
              })()}
            </div>

            <div className="wizard-actions">
              <button
                className="ghost"
                onClick={() => setWizardStep(Math.max(0, wizardStep - 1))}
                disabled={wizardStep === 0}
              >
                <span className="button-content">
                  <ArrowLeft24Regular />
                  Back
                </span>
              </button>
              {wizardStep === wizardSteps.length - 1 ? (
                <button className="primary" onClick={closeWizard}>
                  <span className="button-content">
                    <CheckmarkCircle24Regular />
                    Finish
                  </span>
                </button>
              ) : (
                <button
                  className="primary"
                  onClick={() => setWizardStep(Math.min(wizardSteps.length - 1, wizardStep + 1))}
                >
                  <span className="button-content">
                    Next
                    <ArrowRight24Regular />
                  </span>
                </button>
              )}
            </div>
          </section>
        </div>
      )}
      <footer className="footer">
        <span>
          Designed and developed by{" "}
          <a href="https://www.linkedin.com/in/abpuro/" target="_blank" rel="noreferrer">
            Abhishek Purohit (AB)
          </a>
        </span>
      </footer>
    </main>
  );
}
