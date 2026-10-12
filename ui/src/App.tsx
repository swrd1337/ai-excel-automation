import { useEffect, useRef, useState } from 'react';
import { Activity, Info, Settings, Sparkles, Table2, Undo2 } from 'lucide-react';
import { Badge, Switch } from '@/components/ui/switch';
import { Button } from '@/components/ui/button';
import { Card, CardContent, Separator } from '@/components/ui/primitives';
import { Hint } from '@/components/ui/tooltip';
import { ApprovalCard } from '@/components/ApprovalCard';
import { Composer } from '@/components/Composer';
import { Entry } from '@/components/Entry';
import { SettingsDialog } from '@/components/SettingsDialog';
import { SetupPanel } from '@/components/SetupPanel';
import { useServer, type ServerStatus, type UsageSummary } from '@/lib/api';
import { formatCount, formatPercent, formatRate, formatTokens, formatUsd } from '@/lib/format';
import { cn } from '@/lib/utils';

const SESSION_ID = 'default'; // wire contract shared with app/web/taskpane.ts

const SUGGESTIONS = [
  "What's in this workbook?",
  'Add a total row under the table in Sheet1 and bold it',
  'Format column C as currency with 2 decimals',
  'Add a header row in A1:C1 saying Region, Units, Price and make it bold',
];

function DetailRow({ label, value, hint }: { label: string; value: string; hint?: boolean }) {
  return (
    <div className="flex items-baseline justify-between gap-3 py-1">
      <span className="text-muted-foreground">{label}</span>
      <span className={cn('truncate font-mono text-[11px]', hint && 'text-warning')}>{value}</span>
    </div>
  );
}

/** Token and cost accounting for this session. */
function UsageSection({ usage }: { usage: UsageSummary | null }) {
  if (!usage || !usage.calls) {
    return (
      <p className="py-1 text-[11px] leading-relaxed text-muted-foreground">
        No requests yet. Token counts, speed and estimated cost appear here once you send something.
      </p>
    );
  }

  return (
    <div>
      <DetailRow label="Requests" value={formatCount(usage.calls)} />
      <DetailRow
        label="Input tokens"
        value={`${formatTokens(usage.promptTokens)}${usage.cacheHitTokens ? ` · ${formatPercent(usage.cacheHitRate)} cached` : ''}`}
      />
      <DetailRow label="Output tokens" value={formatTokens(usage.completionTokens)} />
      <DetailRow label="Total tokens" value={formatTokens(usage.totalTokens)} />
      {usage.searches > 0 && (
        <DetailRow label="Web searches" value={`${formatCount(usage.searches)} (provider quota)`} />
      )}
      <DetailRow label="Speed" value={`≈ ${formatRate(usage.tokensPerSecond)} tokens/sec`} />
      <DetailRow label="Estimated cost" value={`${formatUsd(usage.usd)}${usage.estimated ? ' *' : ''}`} />
      <p className="mt-1.5 text-[10.5px] leading-relaxed text-muted-foreground">
        {usage.peak ? 'Peak' : 'Off-peak'} rates, priced as <span className="font-mono">{usage.pricedAs}</span>, as of{' '}
        {usage.pricingAsOf}.
        {usage.estimated && (
          <>
            {' '}
            <span className="text-warning">*</span> Estimated: this app sends{' '}
            <span className="font-mono">deepseek-chat</span>, which DeepSeek no longer publishes, so it is priced on the
            cheapest tier.
          </>
        )}{' '}
        Speed is wall-clock including network latency, so it is a lower bound.
      </p>
    </div>
  );
}

function Header({
  status,
  statusLoaded,
  reachable,
  busy,
  review,
  setReview,
  webSearch,
  setWebSearch,
  onUndo,
  loadConfig,
  saveConfig,
}: {
  status: ServerStatus;
  statusLoaded: boolean;
  reachable: boolean;
  busy: boolean;
  review: boolean;
  setReview: (value: boolean) => void;
  webSearch: boolean;
  setWebSearch: (value: boolean) => void;
  onUndo: () => void;
  loadConfig: ReturnType<typeof useServer>['loadConfig'];
  saveConfig: ReturnType<typeof useServer>['saveConfig'];
}) {
  const [showInfo, setShowInfo] = useState(false);
  const [showSettings, setShowSettings] = useState(false);

  const connected = reachable && status.excelConnected;
  const state = !reachable ? 'offline' : connected ? 'connected' : 'noExcel';

  const badge = {
    offline: { variant: 'destructive' as const, label: 'Server unreachable' },
    noExcel: { variant: 'warning' as const, label: 'Excel not connected' },
    connected: { variant: 'success' as const, label: 'Excel connected' },
  }[state];

  const badgeHint = {
    offline: 'The local server is not responding. Restart the app.',
    noExcel:
      'Open the add-in taskpane in Excel (Insert → Add-ins → My Add-ins) and keep it open — it is the bridge to your workbook.',
    connected: `Connected through the add-in taskpane. Model: ${status.model || 'unknown'}.`,
  }[state];

  return (
    <header className="flex flex-wrap items-center justify-between gap-x-4 gap-y-2 border-b bg-card/50 px-5 py-3">
      <div className="flex min-w-0 items-center gap-2.5">
        <div className="grid size-7 shrink-0 place-items-center rounded-md bg-primary/12 text-primary">
          <Table2 className="size-4" />
        </div>
        <div className="min-w-0">
          <h1 className="truncate text-[13.5px] font-semibold leading-tight">AI Spreadsheet Automation</h1>
          <p className="truncate text-[11px] text-muted-foreground">
            {status.model
              ? `${status.model}${status.search ? ` · search ${webSearch ? 'on' : 'off'}` : ''}`
              : 'connecting…'}
          </p>
        </div>
      </div>

      <div className="flex items-center gap-2">
        <Hint text={badgeHint}>
          <Badge variant={badge.variant}>
            <span
              className={cn(
                'size-1.5 rounded-full',
                state === 'connected' ? 'bg-success' : state === 'noExcel' ? 'bg-warning' : 'bg-destructive'
              )}
            />
            {badge.label}
          </Badge>
        </Hint>

        <Hint text="Show what each change will do and wait for your approval before anything is written.">
          <label className="flex cursor-pointer select-none items-center gap-2 rounded-md px-1.5 py-1 text-[12px] text-muted-foreground hover:bg-accent">
            <Switch checked={review} onCheckedChange={setReview} aria-label="Review changes" />
            Review
          </label>
        </Hint>

        <Hint
          side="bottom"
          text={
            !statusLoaded
              ? 'Checking whether a search provider is configured…'
              : status.search
                ? 'Let the assistant look things up on the web. Off by default — turning it on sends your questions to the search provider.'
                : 'No search key configured. Add one with:  npm run setkey -- --search'
          }
        >
          <label
            className={cn(
              'flex select-none items-center gap-2 rounded-md px-1.5 py-1 text-[12px]',
              status.search ? 'cursor-pointer text-muted-foreground hover:bg-accent' : 'cursor-not-allowed opacity-50'
            )}
          >
            <Switch
              checked={webSearch && Boolean(status.search)}
              onCheckedChange={setWebSearch}
              disabled={!status.search}
              aria-label="Web search"
            />
            Search
          </label>
        </Hint>

        <Hint
          text={
            status.undoAvailable > 0
              ? `Put back everything the last task changed (${status.undoAvailable} task${status.undoAvailable > 1 ? 's' : ''} available).`
              : 'Nothing to undo yet — this becomes available once the assistant changes something.'
          }
        >
          <Button variant="outline" size="sm" disabled={busy || status.undoAvailable === 0} onClick={onUndo}>
            <Undo2 />
            Undo
          </Button>
        </Hint>

        {status.usage && status.usage.calls > 0 && (
          <Hint text="Tokens used and estimated cost this session. Click for the full breakdown.">
            <Button variant="ghost" size="sm" onClick={() => setShowInfo((v) => !v)} className="text-muted-foreground">
              <Activity />
              {formatTokens(status.usage.totalTokens)} tok · {formatUsd(status.usage.usd)}
              {status.usage.estimated ? '*' : ''}
            </Button>
          </Hint>
        )}

        <Hint text="Add or change the DeepSeek and web-search API keys">
          <Button
            variant="ghost"
            size="icon"
            onClick={() => setShowSettings(true)}
            aria-label="API keys"
            data-settings-open
          >
            <Settings />
          </Button>
        </Hint>

        <div className="relative">
          <Hint text="Connection, configuration and AI usage">
            <Button variant="ghost" size="icon" onClick={() => setShowInfo((v) => !v)} aria-label="Details">
              <Info />
            </Button>
          </Hint>

          {showInfo && (
            <>
              <div className="fixed inset-0 z-40" onClick={() => setShowInfo(false)} />
              <Card className="absolute right-0 top-10 z-50 w-80 shadow-lg">
                <CardContent className="p-3 text-[12px]">
                  <DetailRow label="Model" value={status.model || '—'} />
                  <DetailRow
                    label="Web search"
                    value={status.search ? `${status.search} · ${webSearch ? 'on' : 'off'}` : 'not configured'}
                  />
                  <DetailRow label="Excel" value={connected ? 'connected' : 'not connected'} />
                  <DetailRow label="Undo available" value={String(status.undoAvailable)} />
                  <DetailRow
                    label="Certificate"
                    value={status.tlsSource ?? 'none (Excel needs HTTPS)'}
                    hint={!status.tlsSource}
                  />

                  <Separator className="my-2.5" />
                  <p className="mb-1.5 text-[10.5px] font-semibold uppercase tracking-wide text-muted-foreground">
                    AI usage
                  </p>
                  <UsageSection usage={status.usage} />

                  <Separator className="my-2.5" />
                  <Button
                    variant="outline"
                    size="sm"
                    className="w-full"
                    onClick={() => {
                      setShowInfo(false);
                      setShowSettings(true);
                    }}
                  >
                    <Settings />
                    API keys
                  </Button>

                  <p className="mt-2 text-[11px] leading-relaxed text-muted-foreground">
                    The pane in Excel is the bridge — it carries every change to your workbook, so leave it open
                    while the assistant works.
                  </p>
                </CardContent>
              </Card>
            </>
          )}
        </div>
      </div>

      {showSettings && (
        <SettingsDialog onClose={() => setShowSettings(false)} loadConfig={loadConfig} saveConfig={saveConfig} />
      )}
    </header>
  );
}

function EmptyState({ onPick }: { onPick: (text: string) => void }) {
  return (
    <div className="w-full max-w-xl text-center">
      <div className="mx-auto mb-4 grid size-11 place-items-center rounded-xl bg-primary/12 text-primary">
        <Sparkles className="size-5" />
      </div>
      <h2 className="text-[15px] font-semibold">What should I do in your workbook?</h2>
      <p className="mx-auto mt-1.5 max-w-sm text-[13px] leading-relaxed text-muted-foreground">
        Describe it in plain words. Every change is snapshotted first, so <strong>Undo</strong> will put things back.
      </p>
      <div className="mt-6 grid gap-2 text-left sm:grid-cols-2">
        {SUGGESTIONS.map((s) => (
          <Button
            key={s}
            variant="outline"
            size="sm"
            onClick={() => onPick(s)}
            className="h-auto justify-start whitespace-normal py-2 text-left font-normal leading-snug"
          >
            {s}
          </Button>
        ))}
      </div>
    </div>
  );
}

export default function App() {
  const {
    reachable,
    statusLoaded,
    status,
    entries,
    approvals,
    busy,
    review,
    setReview,
    webSearch,
    setWebSearch,
    send,
    decide,
    undo,
    loadConfig,
    saveConfig,
    loadSetup,
  } = useServer(SESSION_ID);
  const [draft, setDraft] = useState('');
  const scrollRef = useRef<HTMLElement>(null);

  useEffect(() => {
    const el = scrollRef.current;
    if (el) el.scrollTop = el.scrollHeight;
  }, [entries, approvals]);

  const handleSend = (text: string): void => {
    if (!text.trim()) return;
    setDraft('');
    void send(text);
  };

  return (
    <div className="flex h-full flex-col bg-background">
      <Header
        status={status}
        statusLoaded={statusLoaded}
        reachable={reachable}
        busy={busy}
        review={review}
        setReview={setReview}
        webSearch={webSearch}
        setWebSearch={setWebSearch}
        onUndo={() => void undo()}
        loadConfig={loadConfig}
        saveConfig={saveConfig}
      />

      <SetupPanel connected={reachable && status.excelConnected} loadSetup={loadSetup} />

      <main ref={scrollRef} className="min-h-0 flex-1 overflow-y-auto px-5 py-5">
        {entries.length === 0 ? (
          <div className="grid min-h-full place-items-center">
            <EmptyState onPick={setDraft} />
          </div>
        ) : (
          <div className="mx-auto flex max-w-3xl flex-col gap-2.5">
            {entries.map((e) => (
              <Entry key={e.id} entry={e} />
            ))}
          </div>
        )}
      </main>

      {approvals.length > 0 && (
        <div className="border-t bg-background px-5 py-3">
          <div className="mx-auto flex max-w-3xl flex-col gap-2">
            {approvals.map((a) => (
              <ApprovalCard key={a.id} approval={a} onDecide={(id, d) => void decide(id, d)} />
            ))}
          </div>
        </div>
      )}

      <Composer value={draft} onChange={setDraft} onSend={handleSend} busy={busy} />
    </div>
  );
}
