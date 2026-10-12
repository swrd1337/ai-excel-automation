/*
 * DeepSeek pricing.
 *
 * Source: https://api-docs.deepseek.com/quick_start/pricing
 * Prices are USD per 1M tokens.
 *
 * Two things make any figure here an ESTIMATE, and the UI says so rather than
 * presenting a made-up number as fact:
 *
 *  1. The model string this app sends is `deepseek-chat`, which is no longer in
 *     the published line-up — the current models are `deepseek-flash` and
 *     `deepseek-v4-pro`. It still resolves (the app works), but we cannot know
 *     from here which model actually served the request, so it is priced on the
 *     cheapest tier. `resolveModel` marks the match inexact so the UI can say so.
 *
 *  2. Peak/off-peak depends on Chinese public holidays, which we cannot know
 *     offline. Weekends and public holidays are off-peak; on a public holiday
 *     this will over-estimate.
 *
 * If prices change, this file is the only place to edit.
 */

export const PRICING_AS_OF = '2026-09-23';
export const PRICING_SOURCE = 'https://api-docs.deepseek.com/quick_start/pricing';

export interface Rates {
  cacheHit: number;
  cacheMiss: number;
  output: number;
}

/** Off-peak rates. Peak is exactly double. */
const OFF_PEAK_RATES: Record<string, Rates> = {
  'deepseek-flash': { cacheHit: 0.003, cacheMiss: 0.15, output: 0.6 },
  'deepseek-v4-pro': { cacheHit: 0.022, cacheMiss: 0.66, output: 1.98 },
};

/*
 * `deepseek-chat` is the app's default and is no longer published. It is priced
 * as Flash — a deliberate assumption, surfaced to the user as an estimate.
 * `deepseek-v4-flash` is documented as a retired legacy name served by V4.1-Flash.
 */
const ALIASES: Record<string, string> = {
  'deepseek-chat': 'deepseek-flash',
  'deepseek-v4-flash': 'deepseek-flash',
  'deepseek-v4-flash-vision-exp': 'deepseek-flash',
};

const FALLBACK = 'deepseek-flash';

/** Peak: 01:00-04:00 and 06:00-10:00 UTC, Monday-Friday. */
export function isPeak(at: Date = new Date()): boolean {
  const day = at.getUTCDay(); // 0 = Sunday
  if (day === 0 || day === 6) return false;
  const hour = at.getUTCHours();
  return (hour >= 1 && hour < 4) || (hour >= 6 && hour < 10);
}

export interface ResolvedModel {
  key: string;
  exact: boolean;
}

export function resolveModel(model?: string | null): ResolvedModel {
  const name = String(model ?? '').trim().toLowerCase();
  if (OFF_PEAK_RATES[name]) return { key: name, exact: true };
  const alias = ALIASES[name];
  if (alias) return { key: alias, exact: false };
  return { key: FALLBACK, exact: false };
}

export interface TokenUsage {
  promptTokens?: number;
  completionTokens?: number;
  cacheHitTokens?: number;
  cacheMissTokens?: number | null;
}

export interface CostEstimate {
  usd: number;
  model: string;
  exact: boolean;
  peak: boolean;
  rates: Rates;
}

export function estimateCost(usage: TokenUsage = {}, model?: string | null, at: Date = new Date()): CostEstimate {
  const { key, exact } = resolveModel(model);
  const rates = OFF_PEAK_RATES[key] as Rates; // FALLBACK is always present
  const peak = isPeak(at);
  const multiplier = peak ? 2 : 1;

  const promptTokens = usage.promptTokens || 0;
  const cacheHit = usage.cacheHitTokens || 0;
  // Without cache reporting, assume every input token was a cache miss. That
  // over-estimates rather than flattering the number.
  const cacheMiss = usage.cacheMissTokens ?? Math.max(0, promptTokens - cacheHit);
  const output = usage.completionTokens || 0;

  const usd =
    (cacheHit / 1e6) * rates.cacheHit * multiplier +
    (cacheMiss / 1e6) * rates.cacheMiss * multiplier +
    (output / 1e6) * rates.output * multiplier;

  return {
    usd,
    model: key,
    exact,
    peak,
    rates: {
      cacheHit: rates.cacheHit * multiplier,
      cacheMiss: rates.cacheMiss * multiplier,
      output: rates.output * multiplier,
    },
  };
}
