/** Compact number/currency formatting for the usage display. */

export function formatTokens(n: number | undefined): string {
  const value = Number(n) || 0;
  if (value < 1000) return String(value);
  if (value < 1_000_000) return `${(value / 1000).toFixed(value < 10_000 ? 2 : 1)}k`;
  return `${(value / 1_000_000).toFixed(2)}M`;
}

export function formatUsd(n: number | undefined): string {
  const value = Number(n) || 0;
  if (value === 0) return '$0.00';
  if (value < 0.01) return `$${value.toFixed(5)}`;
  if (value < 1) return `$${value.toFixed(4)}`;
  return `$${value.toFixed(2)}`;
}

export function formatCount(n: number | undefined): string {
  return (Number(n) || 0).toLocaleString();
}

export function formatRate(n: number | undefined): string {
  const value = Number(n) || 0;
  return value >= 10 ? value.toFixed(0) : value.toFixed(1);
}

export function formatPercent(fraction: number | undefined): string {
  return `${Math.round((Number(fraction) || 0) * 100)}%`;
}
