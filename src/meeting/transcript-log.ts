import {
  appendFileSync,
  closeSync,
  existsSync,
  mkdirSync,
  openSync,
  readFileSync,
  readSync,
  renameSync,
  statSync,
  writeFileSync,
} from 'node:fs';
import { dirname, join } from 'node:path';
import { log } from '../core/logger';

/**
 * The full verbatim record of one meeting, on disk: every committed
 * transcript sentence and every in-meeting chat line, with speaker and time.
 *
 * Why it exists: the in-memory buffer used to be the only copy and was capped
 * at `transcript.keep` lines, so the minutes of a long meeting silently lost
 * its first half, and a restart lost everything before it.
 *
 * Format: `<dataDir>/transcripts/<meetingId>.jsonl`, append-only. The first
 * line is a `meta` record; a sentence that is re-sent with more text is
 * appended again and the last copy wins on load. When the meeting ends the
 * file is compacted to one line per sentence and a readable `.txt` is written
 * next to it.
 */

export interface TranscriptMeta {
  k: 'meta';
  meetingId: string;
  meetingNo: string;
  topic?: string;
  /** When this bridge first joined (ISO). */
  joinedAt: string;
}

export interface TranscriptLine {
  k: 'speech' | 'chat';
  /** sentence_id (speech only). */
  sid?: string;
  /** Speaker / sender display name (`?` when unknown). */
  who: string;
  whoId?: string;
  text: string;
  /** Epoch ms: the event's own time when it has one, else the arrival time. */
  at: number;
  /** Dedup key of the activity item it came from, so a replay is not stored twice. */
  eid?: string;
}

/**
 * Event ids that produced no line of their own (a sentence revision that was
 * superseded, an older revision skipped) — kept so a reload still knows them.
 */
export interface TranscriptSeen {
  k: 'seen';
  seen: string[];
}

export type TranscriptRecord = TranscriptLine | TranscriptSeen;

/** `<dataDir>/transcripts`. */
export function transcriptDir(dataDir: string): string {
  return join(dataDir, 'transcripts');
}

/** File names are built from the meeting id; anything unusual in it is neutralised. */
function safeName(meetingId: string): string {
  return meetingId.replace(/[^A-Za-z0-9_-]/g, '_') || 'meeting';
}

export function transcriptFile(dir: string, meetingId: string): string {
  return join(dir, `${safeName(meetingId)}.jsonl`);
}

export function transcriptTextFile(dir: string, meetingId: string): string {
  return join(dir, `${safeName(meetingId)}.txt`);
}

/** `HH:mm:ss`, local time. */
export function clockLabel(ms: number): string {
  const d = new Date(ms);
  if (Number.isNaN(d.getTime())) return '--:--:--';
  const p = (n: number) => String(n).padStart(2, '0');
  return `${p(d.getHours())}:${p(d.getMinutes())}:${p(d.getSeconds())}`;
}

/**
 * Read a transcript file. Tolerates a torn last line (crash mid-append) and
 * unknown records; a missing file is empty.
 */
export function loadTranscript(file: string): { meta?: TranscriptMeta; records: TranscriptRecord[] } {
  let raw: string;
  try {
    raw = readFileSync(file, 'utf8');
  } catch (err) {
    if ((err as NodeJS.ErrnoException)?.code !== 'ENOENT') {
      log.warn('meeting', 'transcript-read-failed', { file: file.slice(-60), err: String(err) });
    }
    return { records: [] };
  }
  let meta: TranscriptMeta | undefined;
  const records: TranscriptRecord[] = [];
  for (const line of raw.split('\n')) {
    if (!line.trim()) continue;
    let rec: unknown;
    try {
      rec = JSON.parse(line);
    } catch {
      continue;
    }
    if (!rec || typeof rec !== 'object') continue;
    const r = rec as Record<string, unknown>;
    if (r.k === 'meta' && typeof r.meetingId === 'string') {
      meta ??= r as unknown as TranscriptMeta;
      continue;
    }
    if (r.k === 'seen' && Array.isArray(r.seen)) {
      records.push({ k: 'seen', seen: r.seen.filter((x): x is string => typeof x === 'string') });
      continue;
    }
    if ((r.k === 'speech' || r.k === 'chat') && typeof r.text === 'string' && typeof r.at === 'number') {
      records.push({
        k: r.k,
        ...(typeof r.sid === 'string' ? { sid: r.sid } : {}),
        who: typeof r.who === 'string' ? r.who : '?',
        ...(typeof r.whoId === 'string' ? { whoId: r.whoId } : {}),
        text: r.text,
        at: r.at,
        ...(typeof r.eid === 'string' ? { eid: r.eid } : {}),
      });
    }
  }
  return { ...(meta ? { meta } : {}), records };
}

/**
 * Append-only writer. Never throws: a disk problem is logged once and the
 * meeting carries on from memory (losing the file must not lose the meeting).
 */
export class TranscriptLog {
  private ready = false;
  private broken = false;

  constructor(
    readonly file: string,
    private readonly meta: TranscriptMeta,
  ) {}

  append(rec: TranscriptRecord): void {
    if (this.broken) return;
    try {
      if (!this.ready) {
        mkdirSync(dirname(this.file), { recursive: true });
        if (!existsSync(this.file)) appendFileSync(this.file, `${JSON.stringify(this.meta)}\n`);
        // A crash mid-append leaves a torn last line; start on a fresh one so
        // the next record isn't glued to it and lost with it.
        else if (!endsWithNewline(this.file)) appendFileSync(this.file, '\n');
        this.ready = true;
      }
      appendFileSync(this.file, `${JSON.stringify(rec)}\n`);
    } catch (err) {
      this.broken = true;
      log.warn('meeting', 'transcript-write-failed', { file: this.file.slice(-60), err: String(err) });
    }
  }

  /**
   * Replace the file with one record per line (the meeting's final state) and
   * write the readable `.txt` next to it. Atomic per file. Returns the text
   * file path, or undefined when nothing could be written.
   */
  finalize(records: TranscriptRecord[], textFile: string, text: string): string | undefined {
    try {
      mkdirSync(dirname(this.file), { recursive: true });
      const body = [this.meta, ...records].map((r) => JSON.stringify(r)).join('\n') + '\n';
      atomicWrite(this.file, body);
      this.ready = true;
      atomicWrite(textFile, text);
      return textFile;
    } catch (err) {
      log.warn('meeting', 'transcript-finalize-failed', { file: this.file.slice(-60), err: String(err) });
      return undefined;
    }
  }
}

function endsWithNewline(file: string): boolean {
  const { size } = statSync(file);
  if (size === 0) return true;
  const fd = openSync(file, 'r');
  try {
    const buf = Buffer.alloc(1);
    readSync(fd, buf, 0, 1, size - 1);
    return buf[0] === 0x0a;
  } finally {
    closeSync(fd);
  }
}

function atomicWrite(file: string, body: string): void {
  const tmp = `${file}.tmp-${process.pid}`;
  writeFileSync(tmp, body);
  renameSync(tmp, file);
}

/** The readable transcript: a header, then `[HH:mm:ss] 说话人：内容` per line. */
export function renderTranscriptText(
  meta: Pick<TranscriptMeta, 'meetingNo' | 'topic'>,
  records: TranscriptLine[],
): string {
  const first = records[0]?.at;
  const last = records[records.length - 1]?.at;
  const day = first ? new Date(first).toLocaleDateString('zh-CN') : '';
  const head = [
    `会议逐字稿 · ${meta.topic ?? meta.meetingNo}`,
    `会议号 ${meta.meetingNo}${day ? ` · ${day} ${clockLabel(first!)}–${clockLabel(last!)}` : ''} · ${records.length} 条`,
    '（标「弹幕」的是会中聊天；时间是本机时间）',
    '',
  ];
  const body = records.map(
    (r) => `[${clockLabel(r.at)}] ${r.who}${r.k === 'chat' ? '（弹幕）' : ''}：${r.text}`,
  );
  return [...head, ...body].join('\n') + '\n';
}
