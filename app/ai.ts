/*
 * DeepSeek client — OpenAI-compatible /chat/completions with tool calling.
 *
 * Always streamed. Text deltas go to `onDelta` as they arrive so the window can
 * print the reply as it is written, and tool-call fragments are accumulated into
 * a complete call before anything is executed.
 *
 * Plain fetch, no SDK: the payload is small, and a dependency-free server means
 * no node_modules to ship inside the app later.
 *
 * `fetchImpl` is injectable purely so the self-test can drive the whole loop
 * against a scripted model.
 */

const DEFAULT_BASE_URL = 'https://api.deepseek.com';
export const DEFAULT_MODEL = 'deepseek-chat';

/* ------------------------------------------------------------- wire format */

export interface WireToolCall {
  index?: number;
  id?: string;
  type?: string;
  function?: { name?: string; arguments?: string };
}

export interface WireMessage {
  role: 'system' | 'user' | 'assistant' | 'tool';
  content?: string | null;
  tool_calls?: WireToolCall[];
  tool_call_id?: string;
  name?: string;
}

export interface WireUsage {
  prompt_tokens?: number;
  completion_tokens?: number;
  total_tokens?: number;
  prompt_cache_hit_tokens?: number;
  prompt_cache_miss_tokens?: number;
}

export interface StreamDelta {
  content?: string | null;
  tool_calls?: WireToolCall[];
  finish_reason?: string | null;
}

interface StreamChunk {
  choices?: { delta?: StreamDelta; finish_reason?: string | null }[];
  usage?: WireUsage | null;
}

export interface ApiTool {
  type: 'function';
  function: { name: string; description: string; parameters: unknown };
}

export interface ParsedToolCall {
  id: string;
  name: string;
  args: Record<string, unknown>;
  argError: string | null;
  rawArgs: string;
}

export interface Completion {
  text: string;
  toolCalls: ParsedToolCall[];
  finishReason: string;
  usage: WireUsage | null;
  assistantMessage: WireMessage;
}

export interface CompleteRequest {
  messages: WireMessage[];
  tools?: ApiTool[];
  toolChoice?: string;
  onDelta?: (text: string) => void;
}

export interface AiClient {
  model: string;
  complete(req: CompleteRequest): Promise<Completion>;
}

/* ----------------------------------------------------------------- helpers */

/** Tool-call arguments arrive as a JSON *string* and are not always valid JSON. */
export function parseToolArgs(raw: unknown): { args: Record<string, unknown>; error: string | null } {
  if (raw == null || raw === '') return { args: {}, error: null };
  if (typeof raw === 'object') return { args: raw as Record<string, unknown>, error: null };
  try {
    const parsed: unknown = JSON.parse(String(raw));
    if (parsed === null || typeof parsed !== 'object' || Array.isArray(parsed)) {
      return {
        args: {},
        error: `arguments must be a JSON object, got ${Array.isArray(parsed) ? 'array' : typeof parsed}`,
      };
    }
    return { args: parsed as Record<string, unknown>, error: null };
  } catch (err) {
    return { args: {}, error: `arguments are not valid JSON (${(err as Error).message}): ${String(raw).slice(0, 200)}` };
  }
}

export interface FoldedStream {
  text: string;
  toolCalls: { id: string; name: string; rawArgs: string }[];
  finishReason: string;
}

/**
 * Fold streamed deltas into a complete reply.
 *
 * A tool call arrives in fragments spread across chunks, keyed by `index`: the
 * name may be split, and the arguments JSON is streamed a few characters at a
 * time. Anything that treats one chunk as the whole call will produce truncated
 * JSON, so this is separated out and unit-tested.
 */
export function foldDeltas(deltas: StreamDelta[]): FoldedStream {
  let text = '';
  let finishReason = '';
  const byIndex = new Map<number, { id: string; name: string; args: string }>();

  for (const delta of deltas) {
    if (typeof delta.content === 'string') text += delta.content;
    if (delta.finish_reason) finishReason = delta.finish_reason;

    for (const call of delta.tool_calls || []) {
      const index = call.index ?? 0;
      const entry = byIndex.get(index) || { id: '', name: '', args: '' };
      if (call.id) entry.id = call.id;
      if (call.function?.name) entry.name += call.function.name;
      if (call.function?.arguments) entry.args += call.function.arguments;
      byIndex.set(index, entry);
    }
  }

  const toolCalls = [...byIndex.entries()]
    .sort((a, b) => a[0] - b[0])
    .map(([index, entry]) => ({
      id: entry.id || `call_${index}`,
      name: entry.name,
      rawArgs: entry.args,
    }));

  return { text, toolCalls, finishReason };
}

/* ------------------------------------------------------------------ client */

export interface DeepSeekOptions {
  apiKey: string;
  baseUrl?: string;
  model?: string;
  fetchImpl?: typeof fetch;
  temperature?: number;
}

export function createDeepSeekClient({
  apiKey,
  baseUrl = DEFAULT_BASE_URL,
  model = DEFAULT_MODEL,
  fetchImpl = globalThis.fetch,
  temperature = 0,
}: DeepSeekOptions): AiClient {
  if (!apiKey) throw new Error('DEEPSEEK_API_KEY is required (or use --mock for the self-test)');

  return {
    model,

    async complete({ messages, tools, toolChoice = 'auto', onDelta }: CompleteRequest): Promise<Completion> {
      const body: Record<string, unknown> = {
        model,
        messages,
        temperature,
        stream: true,
        // Without this, a streamed response carries no token counts and the usage
        // panel would silently report zero.
        stream_options: { include_usage: true },
      };
      if (tools && tools.length) {
        body.tools = tools;
        body.tool_choice = toolChoice;
      }

      const res = await fetchImpl(`${baseUrl}/chat/completions`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          Authorization: `Bearer ${apiKey}`,
        },
        body: JSON.stringify(body),
      });

      if (!res.ok) {
        const text = await res.text().catch(() => '');
        throw new Error(`DeepSeek HTTP ${res.status}: ${text.slice(0, 500)}`);
      }
      if (!res.body) throw new Error('the response had no body to stream');

      const reader = res.body.getReader();
      const decoder = new TextDecoder();
      const deltas: StreamDelta[] = [];
      let usage: WireUsage | null = null;
      let buffer = '';

      for (;;) {
        const { done, value } = await reader.read();
        if (done) break;

        buffer += decoder.decode(value, { stream: true });
        const lines = buffer.split('\n');
        buffer = lines.pop() ?? ''; // keep the partial last line for the next chunk

        for (const line of lines) {
          const trimmed = line.trim();
          if (!trimmed.startsWith('data:')) continue; // comments and keep-alives
          const payload = trimmed.slice(5).trim();
          if (!payload || payload === '[DONE]') continue;

          let chunk: StreamChunk;
          try {
            chunk = JSON.parse(payload) as StreamChunk;
          } catch {
            continue; // a malformed frame should not kill the whole reply
          }

          if (chunk.usage) usage = chunk.usage;
          const choice = chunk.choices?.[0];
          if (choice?.delta) {
            deltas.push(choice.delta);
            const piece = choice.delta.content;
            if (typeof piece === 'string' && piece && onDelta) onDelta(piece);
          }
          if (choice?.finish_reason) deltas.push({ finish_reason: choice.finish_reason });
        }
      }

      const folded = foldDeltas(deltas);
      const toolCalls: ParsedToolCall[] = folded.toolCalls.map((call) => {
        const { args, error } = parseToolArgs(call.rawArgs);
        return { id: call.id, name: call.name, args, argError: error, rawArgs: call.rawArgs };
      });

      const assistantMessage: WireMessage = {
        role: 'assistant',
        content: folded.text || null,
      };
      if (folded.toolCalls.length) {
        assistantMessage.tool_calls = folded.toolCalls.map((call) => ({
          id: call.id,
          type: 'function',
          function: { name: call.name, arguments: call.rawArgs },
        }));
      }

      return {
        text: folded.text,
        toolCalls,
        finishReason: folded.finishReason,
        usage,
        assistantMessage,
      };
    },
  };
}
