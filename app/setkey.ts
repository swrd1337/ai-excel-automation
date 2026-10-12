/*
 * Store API keys once, in ~/.ai-excel/config.json (mode 0600) — deliberately
 * outside this project so a key can never be committed by accident.
 *
 * Usage:
 *   pbpaste | node app/setkey.js                    # DeepSeek key (avoids shell history)
 *   pbpaste | node app/setkey.js --search           # search key, provider defaults to tavily
 *   node app/setkey.js sk-xxxxxxxx                  # works, but lands in shell history
 *   node app/setkey.js --search tvly-xxxx --provider tavily
 *   node app/setkey.js --unset --search             # remove the search key again
 *   node app/setkey.js                              # show what is configured
 *
 * Get keys from https://platform.deepseek.com and https://tavily.com (or
 * https://brave.com/search/api).
 */

import {
  CONFIG_FILE,
  SEARCH_PROVIDERS,
  checkCredential,
  mask,
  readConfig,
  resolveConfig,
  writeConfig,
  type SearchProvider,
} from './config.js';

const argv = process.argv.slice(2);
let wantSearch = false;
let wantUnset = false;
let provider: string | null = null;
const positional: string[] = [];

for (let i = 0; i < argv.length; i++) {
  const arg = argv[i] as string;
  if (arg === '--search') wantSearch = true;
  else if (arg === '--unset') wantUnset = true;
  else if (arg === '--provider') provider = argv[++i] ?? null;
  else if (arg.startsWith('--')) {
    console.error(`Unknown option: ${arg}`);
    process.exit(1);
  } else positional.push(arg);
}

function status(): void {
  const cfg = resolveConfig();
  console.log(`Config: ${CONFIG_FILE}`);
  console.log(
    `  DeepSeek : ${cfg.deepseek.apiKey ? `${mask(cfg.deepseek.apiKey)} (${cfg.deepseek.apiKey.length} chars)` : '(not set)'}`
  );
  console.log(
    `  Search   : ${cfg.search.apiKey ? `${mask(cfg.search.apiKey)} — provider ${cfg.search.provider}` : '(not set)'}`
  );

  const missing: [string, string][] = [];
  if (!cfg.deepseek.apiKey) missing.push(['  pbpaste | node app/setkey.js', 'DeepSeek key — platform.deepseek.com']);
  if (!cfg.search.apiKey) missing.push(['  pbpaste | node app/setkey.js --search', 'search key — tavily.com or brave.com']);
  if (missing.length) {
    console.log('\nStill to set:');
    for (const [cmd, why] of missing) console.log(`${cmd}\n      ${why}`);
  }
}

async function readStdin(): Promise<string> {
  const chunks: Buffer[] = [];
  for await (const chunk of process.stdin) chunks.push(chunk as Buffer);
  return Buffer.concat(chunks).toString('utf8').trim();
}

if (wantUnset) {
  if (wantSearch) {
    writeConfig({ search: { provider: readConfig().search?.provider || 'tavily', apiKey: null } });
    console.log('Cleared the search key — the assistant will have no web_search tool.');
  } else {
    writeConfig({ deepseekApiKey: null });
    console.log('Cleared the DeepSeek key.');
  }
  console.log('');
  status();
  process.exit(0);
}

let key: string | undefined = positional[0];
const fromArg = Boolean(key);

if (!key) {
  if (process.stdin.isTTY) {
    status();
    process.exit(0);
  }
  key = await readStdin();
}

if (!key) {
  status();
  process.exit(0);
}

/*
 * The same checks the window's settings form runs, so a key the CLI accepts
 * cannot be one the form refuses, or the other way round.
 */
const chosen = (provider || readConfig().search?.provider || 'tavily').toLowerCase();
if (wantSearch && !SEARCH_PROVIDERS.includes(chosen as SearchProvider)) {
  console.error(`Unknown provider "${chosen}". Expected one of: ${SEARCH_PROVIDERS.join(', ')}.`);
  process.exit(1);
}

const check = checkCredential(key, wantSearch ? 'search' : 'deepseek', chosen);
if (check.error) {
  console.error(check.error);
  process.exit(1);
}
if (check.warning) console.warn(`Note: ${check.warning}`);

if (wantSearch) {
  writeConfig({ search: { provider: chosen, apiKey: key } });
  console.log(`Stored ${chosen} search key ${mask(key)} in ${CONFIG_FILE}`);
} else {
  writeConfig({ deepseekApiKey: key });
  console.log(`Stored DeepSeek key ${mask(key)} in ${CONFIG_FILE}`);
}

console.log('');
if (fromArg) {
  console.log('Heads up: passing a key as an argument puts it in your shell history.');
  console.log('Next time:  pbpaste | node app/setkey.js' + (wantSearch ? ' --search' : ''));
  console.log('');
}
status();
