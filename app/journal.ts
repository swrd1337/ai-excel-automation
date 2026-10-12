/*
 * The undo journal, on disk.
 *
 * In memory it was capped at 20 tasks and died with the process, which made the
 * one safety net this app has evaporate every time it restarted — the most
 * likely way a person could actually lose work here.
 *
 * It lives beside the config at ~/.ai-excel/journal.json, mode 0600, for the
 * same reason the key does: outside any checkout, so it can never be committed.
 *
 * Two properties matter more than the file format:
 *
 *   - A missing or corrupt file must never stop the app starting. Losing the
 *     ability to undo is bad; refusing to open is worse, and a half-written file
 *     is exactly what a crash leaves behind.
 *   - The store never reads inside a record. The taskpane wrote it and the
 *     taskpane replays it, so `{action, label, expect}` is an opaque blob here,
 *     for the same reason the server does not interpret it.
 *
 * The one question this module cannot answer is whether a restored record is
 * still true of the workbook — that is decided in the taskpane, by comparing the
 * record's `expect` fingerprint against the cells as they are now. Persisting a
 * snapshot without that check would mean writing week-old state over newer work,
 * which is the data loss the journal exists to prevent.
 */

import { chmodSync, existsSync, mkdirSync, readFileSync, renameSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';

import { CONFIG_DIR } from './config.js';

export const JOURNAL_FILE = join(CONFIG_DIR, 'journal.json');
export const JOURNAL_VERSION = 1;
export const MAX_JOURNAL_TASKS = 20;

/**
 * One user turn's worth of change, replayed newest-first.
 *
 * Generic over the record because the store never looks inside one: the caller
 * knows what it put in, and the taskpane re-validates whatever comes back out.
 */
export interface JournalTask<T = Record<string, unknown>> {
  at: number;
  entries: T[];
}

export interface JournalStore<T = Record<string, unknown>> {
  readonly file: string;
  /** The tasks recorded for a session, oldest first. Empty when none are usable. */
  load(sessionId: string): JournalTask<T>[];
  /** Replace a session's tasks. Best effort: a failed write must not fail a turn. */
  save(sessionId: string, tasks: JournalTask<T>[]): void;
}

interface JournalFile<T> {
  version: number;
  sessions: Record<string, JournalTask<T>[]>;
}

function isTask<T>(value: unknown): value is JournalTask<T> {
  const task = value as JournalTask<T> | null;
  return Boolean(task) && typeof task?.at === 'number' && Array.isArray(task.entries);
}

export function createJournalStore<T = Record<string, unknown>>(file: string = JOURNAL_FILE): JournalStore<T> {
  // Read once, then keep in memory: the file is small, written only when the
  // journal changes, and re-reading it per turn would buy nothing.
  let data: JournalFile<T> | null = null;

  function read(): JournalFile<T> {
    if (data) return data;
    data = { version: JOURNAL_VERSION, sessions: {} };
    if (!existsSync(file)) return data;

    try {
      const raw = JSON.parse(readFileSync(file, 'utf8')) as Partial<JournalFile<T>>;
      if (raw && typeof raw === 'object' && raw.sessions && typeof raw.sessions === 'object') {
        data.sessions = raw.sessions as Record<string, JournalTask<T>[]>;
      } else {
        console.warn(`[journal] ${file} is not in the expected shape; starting with nothing to undo`);
      }
    } catch (err) {
      console.warn(`[journal] ignoring ${file}: ${(err as Error).message}`);
    }
    return data;
  }

  return {
    file,

    load(sessionId) {
      const tasks = read().sessions[sessionId];
      if (!Array.isArray(tasks)) return [];
      // A hand-edited or truncated file must not become a crash one turn later.
      return tasks.filter(isTask).slice(-MAX_JOURNAL_TASKS);
    },

    save(sessionId, tasks) {
      const fileData = read();
      const kept = tasks.filter(isTask).slice(-MAX_JOURNAL_TASKS);
      if (kept.length) fileData.sessions[sessionId] = kept;
      else delete fileData.sessions[sessionId];

      try {
        mkdirSync(dirname(file), { recursive: true, mode: 0o700 });
        // Write-then-rename, because a crash mid-write would otherwise leave a
        // half-file that reads back as "nothing to undo" — quietly turning a
        // recoverable situation into a lost one.
        const tmp = `${file}.tmp`;
        writeFileSync(tmp, `${JSON.stringify(fileData, null, 2)}\n`, { mode: 0o600 });
        chmodSync(tmp, 0o600); // writeFileSync only applies mode when creating
        renameSync(tmp, file);
      } catch (err) {
        console.warn(`[journal] could not save ${file}: ${(err as Error).message}`);
      }
    },
  };
}
