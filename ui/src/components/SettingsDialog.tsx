import { useEffect, useState, type ReactNode } from 'react';
import { Check, Loader2, X } from 'lucide-react';

import { Button } from '@/components/ui/button';
import { Card, CardContent, Separator } from '@/components/ui/primitives';
import type { ConfigState, CredentialPatch, KeyState } from '@/lib/api';

/**
 * Where the keys are set. Before this existed the only way in was
 * `node app/setkey.ts`, which is fine for whoever wrote the app and wrong for
 * anyone else — including the person who wrote it, six months later.
 *
 * Two rules this component exists to respect:
 *
 *   1. **A saved key is never shown again.** The server sends a masked form and
 *      nothing else, so the field is always empty and an empty field means
 *      "leave it alone". That is also why clearing is its own button rather than
 *      an empty submit.
 *   2. **A key that is not the one in charge has to say so.** An environment
 *      variable beats the file, so saving here can look like it did nothing; the
 *      form names the variable that is winning.
 */
export function SettingsDialog({
  onClose,
  loadConfig,
  saveConfig,
}: {
  onClose: () => void;
  loadConfig: () => Promise<{ configurable: boolean; config: ConfigState | null }>;
  saveConfig: (patch: CredentialPatch) => Promise<{ config: ConfigState; warning?: string | null }>;
}) {
  const [config, setConfig] = useState<ConfigState | null>(null);
  const [configurable, setConfigurable] = useState(true);
  const [loadError, setLoadError] = useState<string | null>(null);

  const [deepseek, setDeepseek] = useState('');
  const [provider, setProvider] = useState('tavily');
  const [searchKey, setSearchKey] = useState('');

  const [busy, setBusy] = useState<string | null>(null);
  const [notice, setNotice] = useState<{ kind: 'ok' | 'error' | 'warn'; text: string } | null>(null);

  useEffect(() => {
    let live = true;
    void loadConfig()
      .then((res) => {
        if (!live) return;
        setConfig(res.config);
        setConfigurable(res.configurable);
        if (res.config) setProvider(res.config.search.provider);
      })
      .catch((err: Error) => live && setLoadError(err.message));
    return () => {
      live = false;
    };
  }, [loadConfig]);

  useEffect(() => {
    const onKey = (e: KeyboardEvent): void => {
      if (e.key === 'Escape') onClose();
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [onClose]);

  async function submit(what: 'deepseek' | 'search', patch: CredentialPatch): Promise<void> {
    setBusy(what);
    setNotice(null);
    try {
      const body = await saveConfig(patch);
      setConfig(body.config);
      setProvider(body.config.search.provider);
      setDeepseek('');
      setSearchKey('');
      setNotice(
        body.warning
          ? { kind: 'warn', text: `Saved. ${body.warning}` }
          : { kind: 'ok', text: 'Saved. It is in effect now — no restart needed.' }
      );
    } catch (err) {
      setNotice({ kind: 'error', text: (err as Error).message });
    } finally {
      setBusy(null);
    }
  }

  return (
    <div className="fixed inset-0 z-50 flex items-start justify-center overflow-y-auto bg-black/40 p-6">
      <div className="fixed inset-0" onClick={onClose} />
      <Card role="dialog" aria-label="API keys" data-settings className="relative z-10 mt-10 w-full max-w-md shadow-xl">
        <div className="flex items-center justify-between p-4 pb-2">
          <h2 className="text-[13.5px] font-semibold">API keys</h2>
          <Button variant="ghost" size="icon" onClick={onClose} aria-label="Close settings">
            <X />
          </Button>
        </div>

        <CardContent className="space-y-3 text-[12.5px]">
          {loadError && <Notice kind="error" text={loadError} />}

          {!configurable && !loadError && (
            <Notice
              kind="warn"
              text="This server was given its model already, so it does not manage keys. Start it with `npm run desktop` to set keys here, or use: node app/setkey.js"
            />
          )}

          {config && (
            <>
              <Field
                label="DeepSeek"
                state={config.deepseek}
                hint="Used for every reply. platform.deepseek.com"
                placeholder="sk-…"
                value={deepseek}
                onChange={setDeepseek}
                disabled={!configurable || busy !== null}
                onSave={() => void submit('deepseek', { deepseekApiKey: deepseek.trim() })}
                onClear={() => void submit('deepseek', { deepseekApiKey: null })}
                busy={busy === 'deepseek'}
                saveDisabled={!deepseek.trim()}
              />

              <Separator />

              <Field
                label="Web search"
                state={config.search}
                hint="Optional. Only used when the Search toggle is on."
                placeholder={provider === 'brave' ? 'Brave Search API key' : 'tvly-…'}
                value={searchKey}
                onChange={setSearchKey}
                disabled={!configurable || busy !== null}
                onSave={() => void submit('search', { searchProvider: provider, searchApiKey: searchKey.trim() })}
                onClear={() => void submit('search', { searchApiKey: null })}
                busy={busy === 'search'}
                saveDisabled={!searchKey.trim()}
                extra={
                  <label className="flex items-center gap-2">
                    <span className="text-muted-foreground">Provider</span>
                    <select
                      value={provider}
                      onChange={(e) => setProvider(e.target.value)}
                      disabled={!configurable || busy !== null}
                      className="h-7 rounded-md border bg-background px-1.5 text-[12px]"
                      aria-label="Search provider"
                    >
                      <option value="tavily">Tavily</option>
                      <option value="brave">Brave</option>
                    </select>
                  </label>
                }
              />
            </>
          )}

          {notice && <Notice kind={notice.kind === 'warn' ? 'warn' : notice.kind} text={notice.text} />}

          <p className="pt-1 text-[11px] leading-relaxed text-muted-foreground">
            Stored in <span className="font-mono">{config?.file ?? '~/.ai-excel/config.json'}</span>, readable only by
            you. Keys are sent to the server over this computer's loopback connection and are never shown again once
            saved.
          </p>
        </CardContent>
      </Card>
    </div>
  );
}

/** One credential: what is set, where it came from, and how to change it. */
function Field({
  label,
  state,
  hint,
  placeholder,
  value,
  onChange,
  onSave,
  onClear,
  busy,
  disabled,
  saveDisabled,
  extra,
}: {
  label: string;
  state: KeyState;
  hint: string;
  placeholder: string;
  value: string;
  onChange: (value: string) => void;
  onSave: () => void;
  onClear: () => void;
  busy: boolean;
  disabled: boolean;
  saveDisabled: boolean;
  extra?: ReactNode;
}) {
  return (
    <div className="space-y-1.5">
      <div className="flex items-baseline gap-2">
        <span className="font-semibold">{label}</span>
        {state.set ? (
          <span className="font-mono text-[11.5px] text-muted-foreground">{state.masked}</span>
        ) : (
          <span className="text-[11.5px] text-warning">not set</span>
        )}
      </div>

      <p className="text-[11px] leading-relaxed text-muted-foreground">{hint}</p>

      {state.source === 'env' && (
        <p className="text-[11px] leading-relaxed text-warning">
          In effect from the <span className="font-mono">{state.envVar}</span> environment variable. A key saved here
          is stored but will not be used until that variable is unset.
        </p>
      )}

      <div className="flex flex-wrap items-center gap-2">
        <input
          type="password"
          value={value}
          onChange={(e) => onChange(e.target.value)}
          placeholder={placeholder}
          autoComplete="off"
          spellCheck={false}
          disabled={disabled}
          aria-label={`${label} key`}
          className="h-8 min-w-[9rem] flex-1 rounded-md border bg-background px-2 font-mono text-[12px] disabled:opacity-50"
        />
        <Button size="sm" variant="success" onClick={onSave} disabled={disabled || saveDisabled || busy}>
          {busy ? <Loader2 className="animate-spin" /> : <Check />}
          Save
        </Button>
        {state.set && state.source === 'file' && (
          <Button size="sm" variant="outline" onClick={onClear} disabled={disabled}>
            Remove
          </Button>
        )}
      </div>

      {extra}
    </div>
  );
}

function Notice({ kind, text }: { kind: 'ok' | 'error' | 'warn'; text: string }) {
  const tone =
    kind === 'error' ? 'border-destructive/40 text-destructive' : kind === 'warn' ? 'border-warning/45 text-warning' : 'border-success/40 text-success';
  return <p className={`rounded-md border px-2.5 py-1.5 text-[11.5px] leading-relaxed ${tone}`}>{text}</p>;
}
