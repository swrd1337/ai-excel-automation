import Markdown from 'react-markdown';
import remarkGfm from 'remark-gfm';
import remarkBreaks from 'remark-breaks';
import type { Components } from 'react-markdown';
import { cn } from '@/lib/utils';

/*
 * Model replies arrive as markdown — bold labels, lists, and tables when the
 * answer is tabular, which for a spreadsheet assistant is a real case.
 *
 * react-markdown is used rather than a hand-rolled formatter: it is safe by
 * default (raw HTML is not rendered unless you opt in), and it handles the cases
 * a regex would get wrong — nested lists, GFM tables, emphasis inside code.
 *
 * Styling is mapped element by element instead of adding @tailwindcss/typography,
 * so the output matches the bubble it sits in and there is no second theme.
 *
 * remark-breaks treats a single newline as a line break. Models write that way
 * expecting it to be honoured, and without it adjacent lines silently join into
 * one paragraph — the same class of bug as leaving the markers visible.
 */
const COMPONENTS: Components = {
  // Everything is compact: these render inside a chat bubble, not a document.
  p: ({ ...props }) => <p className="mb-1.5 leading-relaxed last:mb-0" {...props} />,
  strong: ({ ...props }) => <strong className="font-semibold" {...props} />,
  em: ({ ...props }) => <em className="italic" {...props} />,

  ul: ({ ...props }) => <ul className="mb-1.5 list-disc space-y-0.5 pl-4 last:mb-0" {...props} />,
  ol: ({ ...props }) => <ol className="mb-1.5 list-decimal space-y-0.5 pl-4 last:mb-0" {...props} />,
  li: ({ ...props }) => <li className="leading-relaxed [&>p]:mb-0" {...props} />,

  h1: ({ ...props }) => <h3 className="mb-1 mt-2 text-[14px] font-semibold first:mt-0" {...props} />,
  h2: ({ ...props }) => <h3 className="mb-1 mt-2 text-[14px] font-semibold first:mt-0" {...props} />,
  h3: ({ ...props }) => <h4 className="mb-1 mt-2 text-[13.5px] font-semibold first:mt-0" {...props} />,
  h4: ({ ...props }) => <h4 className="mb-1 mt-2 text-[13.5px] font-semibold first:mt-0" {...props} />,

  code: ({ className, ...props }) => (
    <code
      className={cn(
        // Tight horizontal padding: 'px-1' made the following punctuation look
        // detached from the code span.
        'rounded bg-foreground/8 px-[0.3em] py-px font-mono text-[12px]',
        // Fenced blocks carry a language class and are styled by `pre` instead.
        className
      )}
      {...props}
    />
  ),
  pre: ({ ...props }) => (
    <pre
      className="mb-1.5 overflow-x-auto rounded-md border bg-foreground/5 p-2 font-mono text-[12px] leading-relaxed last:mb-0 [&>code]:bg-transparent [&>code]:p-0"
      {...props}
    />
  ),

  a: ({ ...props }) => (
    // Opens in the real browser: Electron denies in-window navigation and hands
    // the URL to the OS.
    <a className="underline underline-offset-2 hover:opacity-80" target="_blank" rel="noreferrer" {...props} />
  ),

  table: ({ ...props }) => (
    <div className="mb-1.5 overflow-x-auto last:mb-0">
      <table className="w-full border-collapse text-[12.5px]" {...props} />
    </div>
  ),
  th: ({ ...props }) => (
    <th className="border-b border-foreground/20 px-2 py-1 text-left font-semibold" {...props} />
  ),
  td: ({ ...props }) => <td className="border-b border-foreground/10 px-2 py-1 align-top" {...props} />,

  blockquote: ({ ...props }) => (
    <blockquote className="mb-1.5 border-l-2 border-foreground/20 pl-3 text-muted-foreground last:mb-0" {...props} />
  ),
  hr: ({ ...props }) => <hr className="my-2 border-foreground/15" {...props} />,
};

export function MarkdownText({ children }: { children: string }) {
  return (
    <Markdown remarkPlugins={[remarkGfm, remarkBreaks]} components={COMPONENTS}>
      {children}
    </Markdown>
  );
}
