import { useEffect, useState } from 'react';
import { AlertTriangle, MonitorCheck, X } from 'lucide-react';

import { Button } from '@/components/ui/button';
import { Card, CardContent } from '@/components/ui/primitives';
import type { SetupState } from '@/lib/api';

/**
 * The steps the app could not do for you.
 *
 * It exists for one situation: the window says "Excel not connected" and the
 * usual cause is that the add-in has never been inserted, which Office gives no
 * way to automate. So the app does everything it can — certificate, manifest —
 * and this says the one or two things a person has to click.
 *
 * It renders **nothing** once Excel is connected. A panel that nags after success
 * is worse than no panel, and the common case after the first run is success.
 */
export function SetupPanel({
  connected,
  loadSetup,
}: {
  connected: boolean;
  loadSetup: () => Promise<{ available: boolean; state: SetupState | null }>;
}) {
  const [state, setState] = useState<SetupState | null>(null);
  const [dismissed, setDismissed] = useState(false);

  useEffect(() => {
    let live = true;
    void loadSetup()
      .then((res) => live && setState(res.available ? res.state : null))
      .catch(() => live && setState(null));
    return () => {
      live = false;
    };
  }, [loadSetup]);

  if (dismissed || connected || !state) return null;

  /*
   * Something the app could not do for itself. These are the only messages worth
   * showing before the checklist: they mean the add-in cannot work no matter what
   * the user clicks, so pointing at Excel first would waste their time.
   */
  const blocked: string[] = [];
  if (state.certificate.ready && !state.certificate.trusted) {
    blocked.push(
      'Excel will not accept the add-in until macOS trusts the local certificate. ' +
        'Run this once — macOS will ask for your password:' +
        `\n     npm run trust-cert`
    );
  } else if (!state.certificate.ready) {
    blocked.push(`The local HTTPS certificate is not ready: ${state.certificate.detail}`);
  }
  if (state.manifest.blocked) {
    blocked.push(state.manifest.detail);
  } else if (!state.manifest.installed) {
    blocked.push(`The Excel add-in is not installed: ${state.manifest.detail}`);
  }

  const steps = blocked.length
    ? []
    : [
        ...(state.manifest.current && state.excelRunning
          ? ['Quit Excel completely (Cmd-Q) and reopen it, so it notices the add-in.']
          : []),
        'In Excel: Insert → Add-ins → My Add-ins → AI Spreadsheet Automation — once per workbook.',
        'Keep that pane open; it is the bridge to your workbook.',
      ];

  return (
    <Card data-setup className="mx-5 mt-4 border-warning/45 bg-warning/[0.07]">
      <CardContent className="flex items-start gap-3 p-3.5">
        {blocked.length ? (
          <AlertTriangle className="mt-[2px] size-4 shrink-0 text-warning" />
        ) : (
          <MonitorCheck className="mt-[2px] size-4 shrink-0 text-warning" />
        )}
        <div className="min-w-0 flex-1">
          <p className="text-[13px] font-medium">
            {blocked.length ? 'Excel cannot connect yet' : 'Finish connecting Excel'}
          </p>
          <ol className="mt-1 list-inside list-decimal space-y-0.5 text-[12px] text-muted-foreground">
            {(blocked.length ? blocked : steps).map((line, i) => (
              <li key={i}>{line}</li>
            ))}
          </ol>
          {state.manifest.path && (
            <p className="mt-1.5 truncate font-mono text-[10.5px] text-muted-foreground">
              add-in {state.manifest.version ?? '?'} · certificate {state.tlsSource ?? 'none'}
            </p>
          )}
        </div>
        <Button variant="ghost" size="icon" aria-label="Dismiss setup" onClick={() => setDismissed(true)}>
          <X />
        </Button>
      </CardContent>
    </Card>
  );
}
