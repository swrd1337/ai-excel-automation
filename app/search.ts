/*
 * Web search providers — server-side only.
 *
 * This runs on the server, never in the taskpane: the search key must not enter
 * the WebView, and the taskpane's job is to touch Excel, nothing else.
 *
 * Auth shapes were verified by probing the live endpoints rather than trusting
 * the docs (both providers' docs sites are JavaScript-rendered and return no
 * usable schema to a fetch):
 *
 *   Brave  — GET /res/v1/web/search?q=…
 *            no header    -> 422 "Field required" (x-subscription-token)
 *            bad token    -> 422 SUBSCRIPTION_TOKEN_INVALID
 *            So: Accept + X-Subscription-Token. Confirmed.
 *
 *   Tavily — POST /search
 *            Returns the SAME "missing or invalid API key" for absent, Bearer and
 *            legacy body auth, so the correct mechanism cannot be distinguished
 *            from outside. We therefore send both forms; the extra field is
 *            ignored by whichever version is current.
 */

const DEFAULT_MAX_RESULTS = 6;
const MAX_RESULTS_LIMIT = 10;
const SNIPPET_CHARS = 700;
const ANSWER_CHARS = 1200;

export interface SearchResult {
  title: string;
  url: string;
  snippet: string;
}

export interface SearchResponse {
  query: string;
  provider: string;
  answer?: string;
  results: SearchResult[];
}

export interface SearchClient {
  provider: string;
  search(query: string, limit?: number): Promise<SearchResponse>;
}

export interface SearchOptions {
  provider?: string;
  apiKey: string;
  fetchImpl?: typeof fetch;
  maxResults?: number;
}

function clip(text: unknown, limit: number): string {
  const s = String(text ?? '')
    .replace(/\s+/g, ' ')
    .trim();
  return s.length > limit ? `${s.slice(0, limit)}…` : s;
}

export function createSearchClient({
  provider = 'tavily',
  apiKey,
  fetchImpl = globalThis.fetch,
  maxResults = DEFAULT_MAX_RESULTS,
}: SearchOptions): SearchClient {
  if (!apiKey) throw new Error('a search API key is required');
  if (!['tavily', 'brave'].includes(provider)) {
    throw new Error(`unknown search provider "${provider}" (expected tavily or brave)`);
  }

  async function tavily(query: string, limit: number): Promise<Omit<SearchResponse, 'query' | 'provider'>> {
    const res = await fetchImpl('https://api.tavily.com/search', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: `Bearer ${apiKey}`, // current API form
      },
      // api_key is the legacy body form; harmless if ignored, and saves a round
      // trip if this key belongs to an older account.
      body: JSON.stringify({
        api_key: apiKey,
        query,
        max_results: limit,
        search_depth: 'basic',
        include_answer: true,
      }),
    });
    const text = await res.text();
    if (!res.ok) throw new Error(`Tavily HTTP ${res.status}: ${clip(text, 300)}`);

    let data: { answer?: string; results?: { title?: string; url?: string; content?: string }[] };
    try {
      data = JSON.parse(text) as typeof data;
    } catch {
      throw new Error(`Tavily returned non-JSON: ${clip(text, 200)}`);
    }

    return {
      ...(data.answer ? { answer: clip(data.answer, ANSWER_CHARS) } : {}),
      results: (data.results || []).map((r) => ({
        title: clip(r.title, 200),
        url: r.url || '',
        snippet: clip(r.content, SNIPPET_CHARS),
      })),
    };
  }

  async function brave(query: string, limit: number): Promise<Omit<SearchResponse, 'query' | 'provider'>> {
    const url = new URL('https://api.search.brave.com/res/v1/web/search');
    url.searchParams.set('q', query);
    url.searchParams.set('count', String(limit));

    const res = await fetchImpl(url, {
      headers: { Accept: 'application/json', 'X-Subscription-Token': apiKey },
    });
    const text = await res.text();
    if (!res.ok) throw new Error(`Brave HTTP ${res.status}: ${clip(text, 300)}`);

    let data: { web?: { results?: { title?: string; url?: string; description?: string }[] } };
    try {
      data = JSON.parse(text) as typeof data;
    } catch {
      throw new Error(`Brave returned non-JSON: ${clip(text, 200)}`);
    }

    const items = data.web?.results || [];
    return {
      results: items.map((r) => ({
        title: clip(r.title, 200),
        url: r.url || '',
        snippet: clip(r.description, SNIPPET_CHARS),
      })),
    };
  }

  return {
    provider,
    async search(query: string, limit: number = maxResults): Promise<SearchResponse> {
      const q = String(query || '').trim();
      if (!q) throw new Error('query is required');
      const n = Math.max(1, Math.min(Number(limit) || maxResults, MAX_RESULTS_LIMIT));
      const out = provider === 'brave' ? await brave(q, n) : await tavily(q, n);
      return { query: q, provider, ...out };
    },
  };
}
