import { useEffect, useRef } from 'react';
import { Loader2, Send } from 'lucide-react';
import { Button } from '@/components/ui/button';
import { Textarea } from '@/components/ui/primitives';
import { Hint } from '@/components/ui/tooltip';

export interface ComposerProps {
  value: string;
  onChange: (value: string) => void;
  onSend: (text: string) => void;
  busy: boolean;
}

export function Composer({ value, onChange, onSend, busy }: ComposerProps) {
  const ref = useRef<HTMLTextAreaElement>(null);

  // Grow with the content, up to a ceiling, so long instructions stay readable.
  useEffect(() => {
    const el = ref.current;
    if (!el) return;
    el.style.height = '0px';
    el.style.height = `${Math.min(el.scrollHeight, 180)}px`;
  }, [value]);

  const canSend = value.trim().length > 0 && !busy;

  return (
    <div className="border-t bg-card/40 px-5 py-3.5">
      <div className="mx-auto flex max-w-3xl items-end gap-2.5">
        <Textarea
          ref={ref}
          rows={1}
          value={value}
          disabled={busy}
          placeholder="Describe what to do in the open workbook…"
          onChange={(e) => onChange(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === 'Enter' && !e.shiftKey) {
              e.preventDefault();
              if (canSend) onSend(value);
            }
          }}
          className="max-h-[180px] min-h-[42px]"
        />

        <Hint text={busy ? 'Working on the current task…' : 'Send (Enter)'} side="top">
          <Button size="lg" onClick={() => onSend(value)} disabled={!canSend} className="h-[42px]">
            {busy ? <Loader2 className="animate-spin" /> : <Send />}
            <span>Send</span>
          </Button>
        </Hint>
      </div>

      <p className="mx-auto mt-2 max-w-3xl text-[11px] text-muted-foreground">
        <kbd className="rounded border px-1 py-px font-sans text-[10px]">Enter</kbd> to send ·{' '}
        <kbd className="rounded border px-1 py-px font-sans text-[10px]">Shift</kbd>+
        <kbd className="rounded border px-1 py-px font-sans text-[10px]">Enter</kbd> for a new line
      </p>
    </div>
  );
}
