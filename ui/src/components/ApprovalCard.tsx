import { AlertTriangle } from 'lucide-react';
import { Button } from '@/components/ui/button';
import { Card, CardContent } from '@/components/ui/primitives';
import type { Approval } from '@/lib/api';

/** Long values are unreadable in a card; keep both sides to one short line. */
function describe(value: unknown, limit = 96): string {
  if (value === undefined || value === null) return '—';
  const text = typeof value === 'string' ? value : JSON.stringify(value);
  return text.length > limit ? `${text.slice(0, limit)}…` : text;
}

function shapeOf(value: unknown): string | null {
  if (!Array.isArray(value) || !value.length) return null;
  const rows = value.length;
  const cols = Math.max(...value.map((r) => (Array.isArray(r) ? r.length : 1)));
  return `${rows} row${rows === 1 ? '' : 's'} × ${cols} column${cols === 1 ? '' : 's'}`;
}

/**
 * The review gate. Nothing has been written when this is on screen — the server
 * is holding the turn until a decision arrives.
 */
export function ApprovalCard({
  approval,
  onDecide,
}: {
  approval: Approval;
  onDecide: (id: string, decision: 'approve' | 'decline') => void;
}) {
  const { id, tool, target, args, before } = approval;
  const after = args ? (args.values ?? args.formulas ?? args) : undefined;
  const shape = shapeOf(after);

  return (
    <Card data-approval={id} className="border-warning/45 bg-warning/[0.07] shadow-sm">
      <CardContent className="p-3.5">
        <div className="flex items-start gap-2.5">
          <AlertTriangle className="mt-[3px] size-4 shrink-0 text-warning" />
          <div className="min-w-0 flex-1">
            <p className="text-[13px] font-medium">
              Apply <span className="font-mono text-[12.5px]">{tool}</span> to{' '}
              <span className="font-mono text-[12.5px]">{target}</span>?
              {shape && <span className="ml-2 text-[11.5px] font-normal text-muted-foreground">{shape}</span>}
            </p>

            <div className="mt-2 space-y-0.5 font-mono text-[11px] leading-relaxed">
              <div className="flex gap-2">
                <span className="w-8 shrink-0 text-muted-foreground">now</span>
                <span className="min-w-0 flex-1 truncate text-muted-foreground">{describe(before)}</span>
              </div>
              <div className="flex gap-2">
                <span className="w-8 shrink-0 text-muted-foreground">then</span>
                <span className="min-w-0 flex-1 truncate">{describe(after)}</span>
              </div>
            </div>

            <div className="mt-3 flex items-center gap-2">
              <Button size="sm" variant="success" onClick={() => onDecide(id, 'approve')}>
                Apply
              </Button>
              <Button size="sm" variant="outline" onClick={() => onDecide(id, 'decline')}>
                Skip
              </Button>
              <span className="text-[11px] text-muted-foreground">Nothing is written until you choose.</span>
            </div>
          </div>
        </div>
      </CardContent>
    </Card>
  );
}
