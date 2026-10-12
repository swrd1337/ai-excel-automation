/*
 * Config lives at ~/.ai-excel/config.json, mode 0600, deliberately OUTSIDE this
 * project so a key can never be committed by accident.
 *
 * Shape:
 * {
 *   "deepseekApiKey": "sk-…",
 *   "search": { "provider": "tavily" | "brave", "apiKey": "…" }
 * }
 *
 * Environment variables win over the file, so a throwaway key never has to
 * disturb your saved one.
 */

import { chmodSync, existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { homedir } from 'node:os';
import { join } from 'node:path';

export const CONFIG_DIR = process.env.AI_EXCEL_HOME || join(homedir(), '.ai-excel');
export const CONFIG_FILE = join(CONFIG_DIR, 'config.json');

/** Written by the first version of setkey.ts; migrated on read so nobody loses a key. */
const LEGACY_KEY_FILE = join(CONFIG_DIR, 'key');

export const SEARCH_PROVIDERS = ['tavily', 'brave'] as const;
export type SearchProvider = (typeof SEARCH_PROVIDERS)[number];

export interface StoredConfig {
  deepseekApiKey?: string | null;
  model?: string | null;
  search?: {
    provider?: string | null;
    apiKey?: string | null;
  } | null;
}

export interface ResolvedConfig {
  configFile: string;
  deepseek: { apiKey: string | null; model: string | null };
  search: { provider: string; apiKey: string | null };
}

export function readConfig(): StoredConfig {
  let cfg: StoredConfig = {};
  if (existsSync(CONFIG_FILE)) {
    try {
      cfg = JSON.parse(readFileSync(CONFIG_FILE, 'utf8')) as StoredConfig;
    } catch (err) {
      throw new Error(`${CONFIG_FILE} is not valid JSON: ${(err as Error).message}`);
    }
  }
  if (!cfg.deepseekApiKey && existsSync(LEGACY_KEY_FILE)) {
    try {
      const legacy = readFileSync(LEGACY_KEY_FILE, 'utf8').trim();
      if (legacy) cfg.deepseekApiKey = legacy;
    } catch {
      // Unreadable legacy file: nothing to migrate.
    }
  }
  return cfg;
}

export function writeConfig(patch: StoredConfig): StoredConfig {
  const merged: StoredConfig = { ...readConfig(), ...patch };
  mkdirSync(CONFIG_DIR, { recursive: true, mode: 0o700 });
  writeFileSync(CONFIG_FILE, `${JSON.stringify(merged, null, 2)}\n`, { mode: 0o600 });
  chmodSync(CONFIG_FILE, 0o600); // writeFileSync only applies mode when creating
  return merged;
}

/** Effective values after environment overrides. */
export function resolveConfig(): ResolvedConfig {
  const cfg = readConfig();
  return {
    configFile: CONFIG_FILE,
    deepseek: {
      apiKey: (process.env.DEEPSEEK_API_KEY || '').trim() || cfg.deepseekApiKey || null,
      model: (process.env.DEEPSEEK_MODEL || '').trim() || cfg.model || null,
    },
    search: {
      provider: (process.env.SEARCH_PROVIDER || '').trim() || cfg.search?.provider || 'tavily',
      apiKey: (process.env.SEARCH_API_KEY || '').trim() || cfg.search?.apiKey || null,
    },
  };
}

export function mask(key: string | null | undefined): string {
  if (!key) return '(none)';
  if (key.length <= 12) return `${key.slice(0, 3)}…`;
  return `${key.slice(0, 6)}…${key.slice(-4)}`;
}

/* ------------------------------------------------------------ credentials */

/**
 * What a caller may change. `null` clears a value, absent leaves it alone —
 * which is what lets the window save one key without having to echo the other
 * back, and therefore without the other key ever reaching the window.
 */
export interface CredentialPatch {
  deepseekApiKey?: string | null;
  searchProvider?: string;
  searchApiKey?: string | null;
}

export interface KeyCheck {
  error?: string;
  warning?: string;
}

/**
 * Format checks only — nothing here proves a key works, and this project has
 * been burnt by "seems to work" before. It catches the mistakes that are
 * actually made: a pasted command instead of a key, and a key for the wrong
 * provider. The provider is the only thing that can say whether a key is real,
 * and that happens on the first call.
 */
export function checkCredential(key: string, kind: 'deepseek' | 'search', provider = 'tavily'): KeyCheck {
  const trimmed = key.trim();
  if (!trimmed) return { error: 'a key is required' };
  if (trimmed !== key) return { error: 'the key has leading or trailing whitespace' };

  /*
   * A real API key is one token with no whitespace. This is what catches "the
   * command got pasted instead of the key", which is easy to do given how these
   * are documented — `pbpaste | node app/setkey.js` in a copy-paste is a very
   * plausible accident.
   */
  if (/\s/.test(trimmed)) {
    return {
      error: `that is not a key — it contains whitespace: ${JSON.stringify(trimmed.slice(0, 80))}. It looks like a command was pasted rather than the key itself.`,
    };
  }
  if (trimmed.length > 200) return { error: `that is ${trimmed.length} characters long — far too long for an API key` };

  if (kind === 'deepseek' && !/^sk-[A-Za-z0-9_-]{16,}$/.test(trimmed)) {
    return { error: 'a DeepSeek key starts with "sk-" and is at least 20 characters (platform.deepseek.com)' };
  }

  // Warnings, not errors: providers change key formats without notice.
  if (kind === 'search' && provider === 'tavily' && !trimmed.startsWith('tvly-')) {
    return { warning: `Tavily keys normally start with "tvly-" — got ${mask(trimmed)}. Stored anyway.` };
  }
  if (kind === 'search' && provider === 'brave' && trimmed.length < 20) {
    return { warning: `that looks short for a Brave key — got ${mask(trimmed)}. Stored anyway.` };
  }
  return {};
}

export interface KeyState {
  set: boolean;
  /** Safe to show, and the only form of a key that ever leaves the server. */
  masked: string | null;
  /**
   * Which value is actually in charge. An environment variable wins over the
   * file, so a key saved in the window can look like it did nothing — the window
   * has to be able to say why.
   */
  source: 'env' | 'file' | null;
  envVar: string | null;
}

export interface ConfigState {
  file: string;
  deepseek: KeyState & { model: string | null };
  search: KeyState & { provider: string };
}

function keyState(fileValue: string | null | undefined, envName: string): KeyState {
  const fromEnv = (process.env[envName] || '').trim();
  if (fromEnv) return { set: true, masked: mask(fromEnv), source: 'env', envVar: envName };
  if (fileValue) return { set: true, masked: mask(fileValue), source: 'file', envVar: null };
  return { set: false, masked: null, source: null, envVar: null };
}

/** The console view of the config. Never includes a key, by construction. */
export function describeConfig(): ConfigState {
  const cfg = readConfig();
  const effective = resolveConfig();
  return {
    file: CONFIG_FILE,
    deepseek: {
      ...keyState(cfg.deepseekApiKey, 'DEEPSEEK_API_KEY'),
      model: effective.deepseek.model,
    },
    search: {
      ...keyState(cfg.search?.apiKey, 'SEARCH_API_KEY'),
      provider: effective.search.provider,
    },
  };
}

/**
 * Validate a patch before anything is written. Returns the problem or the
 * courtesy warning, never both. Lives here so the window's form, the CLI and the
 * tests all apply one set of rules.
 */
export function checkCredentialPatch(patch: CredentialPatch): KeyCheck {
  if (patch.deepseekApiKey === undefined && patch.searchApiKey === undefined && patch.searchProvider === undefined) {
    return { error: 'nothing to change' };
  }

  if (patch.searchProvider !== undefined && !SEARCH_PROVIDERS.includes(patch.searchProvider as SearchProvider)) {
    return {
      error: `unknown search provider "${patch.searchProvider}" — expected one of ${SEARCH_PROVIDERS.join(', ')}`,
    };
  }

  if (typeof patch.deepseekApiKey === 'string') {
    const check = checkCredential(patch.deepseekApiKey, 'deepseek');
    if (check.error) return { error: `DeepSeek key: ${check.error}` };
  }

  if (typeof patch.searchApiKey === 'string') {
    const provider = patch.searchProvider ?? readConfig().search?.provider ?? 'tavily';
    const check = checkCredential(patch.searchApiKey, 'search', provider);
    if (check.warning) return { warning: check.warning };
    if (check.error) return { error: `search key: ${check.error}` };
  }

  return {};
}

/**
 * Write a validated patch. A flat patch because the caller has one field per
 * control, and `undefined` means "leave it alone" while `null` means "clear it" —
 * which is what lets the window save one key without the other key ever being
 * sent to it.
 */
export function applyCredentialPatch(patch: CredentialPatch): void {
  const current = readConfig();
  const next: StoredConfig = {};

  if (patch.deepseekApiKey !== undefined) next.deepseekApiKey = patch.deepseekApiKey;

  if (patch.searchApiKey !== undefined || patch.searchProvider !== undefined) {
    // writeConfig merges shallowly, so the whole search object must be rebuilt.
    next.search = {
      provider: patch.searchProvider ?? current.search?.provider ?? 'tavily',
      apiKey: patch.searchApiKey === undefined ? current.search?.apiKey ?? null : patch.searchApiKey,
    };
  }

  writeConfig(next);
}
