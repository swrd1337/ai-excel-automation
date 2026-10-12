/*
 * Does DeepSeek's Responses API actually do server-side web search?
 *
 * The official compatibility table says the `web_search` tool type is "Ignored",
 * but two issues from 2026-08-06 (hermes-agent #79820, oh-my-pi #7794) describe
 * it working on `deepseek-v4-flash`, emitting `response.web_search_call.*`
 * events. Those cannot both be current, and the answer decides whether this app
 * can drop its Tavily/Brave dependency.
 *
 * This makes TWO requests at most, both read-only, and prints what came back.
 *
 *   node scripts/probe-search.js
 *
 * It uses the key from ~/.ai-excel/config.json. Cost is a fraction of a cent:
 * roughly 100 input tokens and a few hundred output tokens.
 */

import { readConfig, mask } from '../app/config.js';

const KEY = process.env.DEEPSEEK_API_KEY || readConfig().deepseekApiKey;
if (!KEY) {
  console.error('No DeepSeek key configured. Run: node app/setkey.js');
  process.exit(1);
}

/** The Responses API documents only deepseek-flash, not the chat alias. */
const MODEL = process.env.DEEPSEEK_MODEL || 'deepseek-flash';

const PROMPT =
  'Search the web and tell me one current fact that changes often, such as a major ' +
  'stock index level or a recent release version. Say if you cannot search.';

/** The hermes issue flagged that /v1/responses and /responses may differ. */
const ENDPOINTS = ['https://api.deepseek.com/responses', 'https://api.deepseek.com/v1/responses'];

interface Probe {
  endpoint: string;
  status: number;
  error?: string;
  itemTypes: string[];
  searched: boolean;
  text: string;
  usage?: unknown;
}

async function probe(endpoint: string): Promise<Probe> {
  const result: Probe = { endpoint, status: 0, itemTypes: [], searched: false, text: '' };

  let res: Response;
  try {
    res = await fetch(endpoint, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${KEY}` },
      body: JSON.stringify({
        model: MODEL,
        input: PROMPT,
        tools: [{ type: 'web_search' }],
        tool_choice: 'auto',
        stream: false,
      }),
    });
  } catch (err) {
    result.error = `request failed: ${(err as Error).message}`;
    return result;
  }

  result.status = res.status;
  const text = await res.text();

  if (!res.ok) {
    result.error = text.slice(0, 400);
    return result;
  }

  let body: Record<string, unknown>;
  try {
    body = JSON.parse(text) as Record<string, unknown>;
  } catch {
    result.error = `non-JSON response: ${text.slice(0, 200)}`;
    return result;
  }

  const output = Array.isArray(body.output) ? (body.output as Record<string, unknown>[]) : [];
  result.itemTypes = output.map((item) => String(item.type ?? '?'));

  // A web_search_call output item is the signal that the server actually searched.
  result.searched = output.some((item) => String(item.type ?? '').includes('web_search'));

  result.text = output
    .flatMap((item) => (Array.isArray(item.content) ? (item.content as Record<string, unknown>[]) : []))
    .map((part) => String(part.text ?? ''))
    .join(' ')
    .trim();

  result.usage = body.usage;
  return result;
}

console.log(`Probing DeepSeek native web_search`);
console.log(`  model: ${MODEL}`);
console.log(`  key:   ${mask(KEY)}`);
console.log('');

let anySearched = false;

for (const endpoint of ENDPOINTS) {
  const r = await probe(endpoint);
  console.log(`${endpoint}`);
  console.log(`  status:     ${r.status}`);
  if (r.error) console.log(`  error:      ${r.error}`);
  console.log(`  item types: ${r.itemTypes.length ? r.itemTypes.join(', ') : '(none)'}`);
  console.log(`  searched:   ${r.searched ? 'YES — web_search_call came back' : 'no'}`);
  if (r.text) console.log(`  reply:      ${r.text.slice(0, 200)}`);
  if (r.usage) console.log(`  usage:      ${JSON.stringify(r.usage)}`);
  console.log('');
  if (r.searched) anySearched = true;
}

console.log('─'.repeat(60));
if (anySearched) {
  console.log('VERDICT: native search WORKS. The docs table is out of date.');
  console.log('Using it means a second transport in ai.ts (Responses API shape),');
  console.log('but it would remove the need for a Tavily/Brave key.');
} else {
  console.log('VERDICT: no web_search_call came back — the tool was ignored, as the');
  console.log('current docs say. The August reports describe behaviour that no longer');
  console.log('applies to this model. Keep Tavily/Brave.');
}
