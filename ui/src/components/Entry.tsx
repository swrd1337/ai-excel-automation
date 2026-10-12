import { AlertTriangle, Check, Search, Undo2, X } from 'lucide-react';
import type { ReactNode } from 'react';
import type { ChatEntry } from '@/lib/api';
import { MarkdownText } from '@/components/Markdown';
import { cn } from '@/lib/utils';

function ActivityRow({
  icon,
  name,
  detail,
  tone = 'muted',
}: {
  icon: ReactNode;
  name?: string;
  detail?: string;
  tone?: 'muted' | 'error';
}) {
  return (
    <div
      className={cn(
        'flex items-baseline gap-2 px-1 py-[3px] text-[11.5px]',
        tone === 'error' ? 'text-destructive' : 'text-muted-foreground'
      )}
    >
      <span className="relative top-[2px] shrink-0">{icon}</span>
      {name && <span className="shrink-0 font-mono">{name}</span>}
      {detail && <span className="min-w-0 flex-1 truncate opacity-75">{detail}</span>}
    </div>
  );
}

/** One row of the conversation: a message, or a line of what happened. */
export function Entry({ entry }: { entry: ChatEntry }) {
  if (entry.kind === 'user') {
    return (
      <div className="flex justify-end">
        <div className="max-w-[85%] rounded-2xl rounded-br-md bg-primary px-3.5 py-2 text-[13.5px] leading-relaxed text-primary-foreground shadow-sm whitespace-pre-wrap break-words">
          {entry.text}
        </div>
      </div>
    );
  }

  if (entry.kind === 'assistant') {
    return (
      <div className="flex justify-start">
        <div
          className={cn(
            // No whitespace-pre-wrap: markdown supplies its own block structure.
            'max-w-[85%] rounded-2xl rounded-bl-md px-3.5 py-2.5 text-[13.5px] leading-relaxed shadow-sm break-words',
            entry.failed ? 'bg-destructive text-destructive-foreground' : 'border bg-card text-card-foreground'
          )}
        >
          {/*
            Render markdown, except for failures: an error string is not markdown
            and formatting it could mangle paths or quotes.
            No streaming cursor either — the growing text is the signal.
          */}
          {entry.failed ? entry.text : <MarkdownText>{entry.text ?? ''}</MarkdownText>}
        </div>
      </div>
    );
  }

  if (entry.kind === 'tool') {
    return (
      <ActivityRow
        icon={entry.ok ? <Check className="size-3.5 text-success" /> : <X className="size-3.5 text-destructive" />}
        name={entry.name}
        detail={entry.detail}
        tone={entry.ok ? 'muted' : 'error'}
      />
    );
  }

  if (entry.kind === 'search') {
    return <ActivityRow icon={<Search className="size-3.5" />} name="web_search" detail={entry.query} />;
  }

  if (entry.kind === 'approval') {
    return (
      <ActivityRow
        icon={
          entry.ok ? <Check className="size-3.5 text-success" /> : <AlertTriangle className="size-3.5 text-warning" />
        }
        name={entry.name}
        detail={entry.detail}
        tone={entry.ok ? 'muted' : 'error'}
      />
    );
  }

  if (entry.kind === 'undo') {
    return (
      <ActivityRow
        icon={
          entry.ok ? <Undo2 className="size-3.5 text-success" /> : <AlertTriangle className="size-3.5 text-destructive" />
        }
        name="undo"
        detail={entry.detail}
        tone={entry.ok ? 'muted' : 'error'}
      />
    );
  }

  return null;
}
