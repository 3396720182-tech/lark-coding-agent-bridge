import { log } from '../core/logger';
import type { MeetingConfig } from '../config/profile-schema';
import {
  fetchMeetingEvents,
  leaveMeeting,
  sendMeetingText,
  VcApiError,
  type RawActivityItem,
  type VcRequestClient,
} from './api';
import { unpackActivity, type MeetingEvent, type TranscriptEvent } from './types';
import {
  clockLabel,
  loadTranscript,
  renderTranscriptText,
  TranscriptLog,
  transcriptFile,
  transcriptTextFile,
  type TranscriptLine,
  type TranscriptRecord,
} from './transcript-log';

export type MeetingEventSourceKind = 'push' | 'poll';

/** Consecutive poll failures before the owner is told. */
export const POLL_FAILURE_ALERT = 5;

/**
 * Once push delivers, the poller keeps running this slowly as a gap-fill lane:
 * pushes can be dropped, and `event_id` dedup makes the overlap harmless.
 */
export const GAP_FILL_INTERVAL_MS = 60_000;

/** A full re-read of the meeting's events from the start, at least this often while in it. */
export const FULL_BACKFILL_INTERVAL_MS = 15 * 60_000;

/** After someone leaves (the meeting may be ending), re-read everything this soon. */
export const LEAVE_BACKFILL_DELAY_MS = 5_000;

/** Two full re-reads no closer than this, except the explicit ones (leave, resume). */
export const BACKFILL_MIN_GAP_MS = 60_000;

/** Memory guard only: far beyond any real meeting (a 3-hour one is a few thousand lines). */
const MAX_LINES = 50_000;

/** Dedup ids kept per meeting. Large on purpose: a full re-read replays every event. */
const MAX_SEEN_EVENTS = 100_000;

/** One line of the meeting record: a transcript sentence or a chat message. */
interface MeetingLine {
  kind: 'speech' | 'chat';
  sentenceId?: string;
  speaker: string;
  speakerId?: string;
  text: string;
  /** Epoch ms used for ordering and the time label (event time, else arrival). */
  at: number;
  /** Arrival order — the tie-breaker when two lines share a timestamp. */
  seq: number;
  /** Chat only: what kind of dedup key it arrived under (see {@link hasChat}). */
  keyKind?: KeyKind;
  /** Dedup key of the item it (last) came from. */
  eid?: string;
}

/** How an item reached {@link MeetingSession.ingest}. */
export interface IngestOptions {
  /**
   * Store it, but tell no subscriber: a re-read of old events must not answer
   * old questions or fire old alerts. A sentence already held is only
   * replaced by a longer version, so a replayed early prefix cannot overwrite
   * the final text.
   */
  quiet?: boolean;
  /**
   * Came through the poller while push is the main lane: anything here was
   * either already pushed or dropped by push. Subscribers still hear what is
   * new (a dropped question should be answered), but a sentence already held
   * is only replaced by a longer version — this lane runs behind push and
   * would otherwise roll a sentence back to an older revision.
   */
  gapFill?: boolean;
}

/** `content`: dedup key built from the item's content (it had no `event_id`). */
type KeyKind = 'id' | 'content';
const keyKind = (eid: string | undefined): KeyKind | undefined =>
  eid === undefined ? undefined : /^[a-z_]+#/.test(eid) ? 'content' : 'id';

export interface MeetingSessionStatus {
  meetingId: string;
  meetingNo: string;
  topic?: string;
  startedAt: string;
  /** Where in-meeting content is currently coming from. */
  source: MeetingEventSourceKind;
  transcriptLines: number;
  participants: number;
  /** Pushes/polls ingested — a quick health signal for the console. */
  ingested: number;
  /**
   * Count of raw activity items per `activity_event_type`, including types we
   * don't handle (prefixed `?`). This is what distinguishes "Feishu never sent
   * transcripts" from "it sent them and we failed to parse" — without it the
   * only symptom is a silent zero.
   */
  eventCounts: Record<string, number>;
  ended: boolean;
  /** In-meeting triggers refused by the authorization gate. */
  deniedCount: number;
  /** Transcript lines that fired a keyword / speaker alert. */
  alertHits: number;
  /** Chat lines held (reactions excluded). */
  chatLines: number;
  /** Full re-reads done, and lines they recovered that the live lanes had missed. */
  backfills: number;
  backfilled: number;
  /** Activity items fetched by the poller (not catch-up or re-reads). */
  polled: number;
  /** Where the full transcript is being written; absent when saving is off. */
  transcriptFile?: string;
  /** The readable copy, once written (meeting ended or left). */
  transcriptText?: string;
}

export interface MeetingSessionDeps {
  client: VcRequestClient;
  meetingId: string;
  meetingNo: string;
  topic?: string;
  config: MeetingConfig;
  /** Bot's own open_id, used to drop self-echo transcript lines. */
  botOpenId?: string;
  /** Chat that started this session — where IM answers go back to. */
  originChatId?: string;
  /**
   * Directory for the full transcript (`transcript-log.ts`). Omit, or set
   * `transcript.save: false`, to keep it in memory only.
   */
  transcriptDir?: string;
  now?: () => number;
}

/**
 * One live meeting. Owns the transcript buffer, de-duplication and the
 * end-of-life cleanup for a single meeting; the manager owns routing across
 * meetings. Deliberately has no knowledge of the agent — the orchestrator
 * subscribes via {@link on}.
 */
export class MeetingSession {
  readonly meetingId: string;
  readonly meetingNo: string;
  readonly topic?: string;
  readonly startedAt: string;
  readonly originChatId?: string;
  /** Set when the session was re-adopted after a bridge restart (epoch ms). */
  resumedAt?: number;
  /**
   * Called when the session hits a problem the user should hear about (the
   * orchestrator turns it into an owner DM). Currently: polling failed
   * {@link POLL_FAILURE_ALERT} times in a row.
   */
  onTrouble?: (kind: 'poll-failed', err: unknown) => void;
  /**
   * Called once when the poller sees the meeting has ended (`122001`). The
   * `meeting_ended` push is the other way to learn it, but pushes may never
   * arrive (no console subscription), and without this the session would poll
   * a dead meeting forever, never summarize, and alarm the owner about it.
   */
  onGone?: (err: VcApiError) => void;

  private readonly client: VcRequestClient;
  private readonly config: MeetingConfig;
  private readonly botOpenId?: string;
  private readonly now: () => number;

  /**
   * The whole meeting so far — transcript sentences and chat — not a rolling
   * window: the minutes need all of it. `transcript.keep` only bounds what an
   * in-meeting answer is given ({@link recentContext}). Chat is meeting content
   * too: a meeting without transcription (captions off) only has chat.
   */
  private lines: MeetingLine[] = [];
  /** sentence_id → its line, so a re-sent sentence replaces in place. */
  private bySentence = new Map<string, MeetingLine>();
  /**
   * Arrival order shared by transcript and chat lines. A counter, not a clock:
   * a poll page or a catch-up delivers many events in the same millisecond.
   */
  private seq = 0;
  private readonly transcriptLog?: TranscriptLog;
  private readonly textFile?: string;
  private backfillRun?: Promise<number>;
  private backfillTimer?: ReturnType<typeof setTimeout>;
  private periodic?: ReturnType<typeof setTimeout>;
  private lastBackfillAt = 0;
  private backfills = 0;
  private backfilled = 0;
  private polled = 0;
  private savedText?: string;
  /** Lines the current quiet ingest added (counts recoveries; longer versions don't count). */
  private quietGain = 0;
  /** sentence_id → latest text, so a re-sent sentence overwrites rather than appends. */
  private sentences = new Map<string, string>();
  /** event_id of already-ingested activity items (push replay / poll overlap). */
  private seenEvents = new Set<string>();
  /** Pending debounce timers when `stabilizeMs > 0`. */
  private pending = new Map<
    string,
    { timer: ReturnType<typeof setTimeout>; event: TranscriptEvent; eid: string | undefined }
  >();
  private handlers = new Map<string, Set<(e: MeetingEvent) => void>>();
  private participantIds = new Set<string>();

  private source: MeetingEventSourceKind = 'poll';
  private ingested = 0;
  private eventCounts: Record<string, number> = {};
  private cursor?: string;
  private poller?: ReturnType<typeof setTimeout>;
  private idleRounds = 0;
  private pollFailures = 0;
  private stopped = false;
  private leaving = false;
  ended = false;
  /** Bumped by the in-meeting authorization gate (see authz.ts). */
  deniedCount = 0;
  /** Bumped by the alert matcher (see alerts.ts). */
  alertHits = 0;
  /** Alert state lives with the meeting: a new meeting alerts afresh. */
  readonly alertFired = new Map<string, Set<string>>();
  readonly alertCooldown = new Map<string, number>();

  constructor(deps: MeetingSessionDeps) {
    this.client = deps.client;
    this.meetingId = deps.meetingId;
    this.meetingNo = deps.meetingNo;
    if (deps.topic) this.topic = deps.topic;
    this.config = deps.config;
    if (deps.botOpenId) this.botOpenId = deps.botOpenId;
    if (deps.originChatId) this.originChatId = deps.originChatId;
    this.now = deps.now ?? (() => Date.now());
    this.startedAt = new Date(this.now()).toISOString();
    if (deps.transcriptDir && this.config.transcript.save) {
      const file = transcriptFile(deps.transcriptDir, this.meetingId);
      this.textFile = transcriptTextFile(deps.transcriptDir, this.meetingId);
      // A restarted bridge picks up where the file left off, so the minutes
      // still cover the part before the restart.
      const prior = loadTranscript(file);
      for (const rec of prior.records) this.restore(rec);
      this.transcriptLog = new TranscriptLog(file, {
        k: 'meta',
        meetingId: this.meetingId,
        meetingNo: this.meetingNo,
        ...(this.topic ? { topic: this.topic } : {}),
        joinedAt: prior.meta?.joinedAt ?? this.startedAt,
      });
      if (prior.records.length) {
        log.info('meeting', 'transcript-restored', { meetingId: this.meetingId, lines: this.lines.length });
      }
    }
  }

  /** Rebuild one line from the transcript file (no event, no write). */
  private restore(rec: TranscriptRecord): void {
    if (rec.k === 'seen') {
      for (const id of rec.seen) this.seenEvents.add(id);
      return;
    }
    if (rec.eid) this.seenEvents.add(rec.eid);
    if (rec.k === 'speech' && rec.sid) {
      this.sentences.set(rec.sid, rec.text);
      const known = this.bySentence.get(rec.sid);
      if (known) {
        known.text = rec.text;
        return;
      }
    }
    const line: MeetingLine = {
      kind: rec.k,
      ...(rec.sid ? { sentenceId: rec.sid } : {}),
      speaker: rec.who,
      ...(rec.whoId ? { speakerId: rec.whoId } : {}),
      text: rec.text,
      at: rec.at,
      seq: ++this.seq,
      ...(rec.k === 'chat' && rec.eid ? { keyKind: keyKind(rec.eid)! } : {}),
      ...(rec.eid ? { eid: rec.eid } : {}),
    };
    this.lines.push(line);
    if (line.sentenceId) this.bySentence.set(line.sentenceId, line);
  }

  status(): MeetingSessionStatus {
    return {
      meetingId: this.meetingId,
      meetingNo: this.meetingNo,
      ...(this.topic ? { topic: this.topic } : {}),
      startedAt: this.startedAt,
      source: this.source,
      transcriptLines: this.transcriptLineCount(),
      participants: this.participantIds.size,
      ingested: this.ingested,
      eventCounts: { ...this.eventCounts },
      ended: this.ended,
      deniedCount: this.deniedCount,
      alertHits: this.alertHits,
      chatLines: this.lines.length - this.transcriptLineCount(),
      backfills: this.backfills,
      backfilled: this.backfilled,
      polled: this.polled,
      ...(this.transcriptLog ? { transcriptFile: this.transcriptLog.file } : {}),
      ...(this.savedText ? { transcriptText: this.savedText } : {}),
    };
  }

  /** Subscribe to a normalized event kind. Multiple handlers per kind. */
  on(kind: MeetingEvent['kind'], handler: (e: MeetingEvent) => void): () => void {
    let set = this.handlers.get(kind);
    if (!set) this.handlers.set(kind, (set = new Set()));
    set.add(handler);
    return () => set!.delete(handler);
  }

  /** Lines in meeting order: by time, then arrival. */
  private ordered(): MeetingLine[] {
    return [...this.lines].sort((a, b) => a.at - b.at || a.seq - b.seq);
  }

  /** Latest transcript as `说话人: 内容` lines, oldest first; all of it without `limit`. */
  recentTranscript(limit?: number): string[] {
    const speech = this.ordered().filter((l) => l.kind === 'speech');
    const slice = !limit || limit >= speech.length ? speech : speech.slice(-limit);
    return slice.map((l) => `${l.speaker}: ${l.text}`);
  }

  /**
   * Transcript and in-meeting chat merged in meeting order, oldest first — the
   * agent's context for in-meeting answers (pass `transcript.keep`). Chat lines
   * are tagged so the agent can tell typed messages from speech.
   */
  recentContext(limit?: number): string[] {
    const all = this.ordered();
    const slice = !limit || limit >= all.length ? all : all.slice(-limit);
    return slice.map(formatLine);
  }

  /**
   * The whole meeting with a `[HH:mm:ss]` label per line — what the minutes
   * are written from.
   */
  fullContext(): string[] {
    return this.ordered().map((l) => `[${clockLabel(l.at)}] ${formatLine(l)}`);
  }

  /**
   * The latest lines with a `[HH:mm:ss]` label each — what in-meeting answers
   * are given, so they can cite who said what when (〔14:32 张三〕).
   */
  recentTimedContext(limit?: number): string[] {
    const all = this.ordered();
    const slice = !limit || limit >= all.length ? all : all.slice(-limit);
    return slice.map((l) => `[${clockLabel(l.at)}] ${formatLine(l)}`);
  }

  /** Display name → open_id for people heard in this meeting; names shared by two ids are left out. */
  speakerDirectory(): Map<string, string> {
    const out = new Map<string, string>();
    const clash = new Set<string>();
    for (const l of this.lines) {
      if (!l.speakerId || !l.speaker || l.speaker === '?') continue;
      const known = out.get(l.speaker);
      if (known && known !== l.speakerId) clash.add(l.speaker);
      else out.set(l.speaker, l.speakerId);
    }
    for (const n of clash) out.delete(n);
    return out;
  }

  /** Arrival number of the newest line (0 when empty); see {@link linesAfter}. */
  lastSeq(): number {
    return this.seq;
  }

  /** Arrival number of a transcript sentence, if held. */
  seqOf(sentenceId: string): number | undefined {
    return this.bySentence.get(sentenceId)?.seq;
  }

  /** Lines that arrived after `seq`, oldest first. */
  linesAfter(seq: number): { kind: 'speech' | 'chat'; speaker: string; speakerId?: string; text: string }[] {
    return this.lines
      .filter((l) => l.seq > seq)
      .sort((a, b) => a.seq - b.seq)
      .map((l) => ({ kind: l.kind, speaker: l.speaker, ...(l.speakerId ? { speakerId: l.speakerId } : {}), text: l.text }));
  }

  /** Number of transcript lines held (chat excluded). */
  transcriptLineCount(): number {
    return this.bySentence.size;
  }

  /** Number of lines held, transcript and chat. */
  lineCount(): number {
    return this.lines.length;
  }

  /**
   * Events are arriving via push. Polling was the primary lane until now; keep
   * both and every event is delivered twice, so hand over and stop the poller.
   * The push lane is what the protocol intends for TAT; `bots/events` stays
   * available for an explicit gap-fill.
   */
  markPushActive(): void {
    if (this.source === 'push') return;
    this.source = 'push';
    // The poller is not stopped, only slowed (see GAP_FILL_INTERVAL_MS): a
    // dropped push would otherwise be lost for good. Dedup makes overlap free.
    if (this.poller) log.info('meeting', 'poll-slowed-push-active', { meetingId: this.meetingId });
  }

  /**
   * Ingest one raw activity item. Idempotent per `event_id`, so replayed pushes
   * and poll/push overlap during gap-fill collapse to one delivery.
   */
  ingest(item: RawActivityItem, opts: IngestOptions = {}): void {
    if (this.ended) return;
    // Push and poll deliver the *same* content by design (poll is the gap-fill
    // lane), so every item can arrive twice. `event_id` is the intended dedup
    // key, but it isn't always present — without a fallback a duplicate chat
    // event fires a second agent run, which then gets rejected as
    // "another run is already active" and looks like a phantom stuck task.
    const eventId = item.event_id ?? contentDedupKey(item);
    if (eventId) {
      if (this.seenEvents.has(eventId)) return;
      this.seenEvents.add(eventId);
      // Bound the dedup set. Generous: a full re-read replays the whole
      // meeting, and an evicted id lets an old chat line in a second time.
      if (this.seenEvents.size > MAX_SEEN_EVENTS) {
        for (const id of this.seenEvents) {
          this.seenEvents.delete(id);
          if (this.seenEvents.size <= MAX_SEEN_EVENTS * 0.8) break;
        }
      }
    }
    this.ingested += 1;
    const rawType = item.activity_event_type ?? (item as { event_type?: string }).event_type ?? 'unknown';
    const events = unpackActivity(item, this.botOpenId);
    // An item that yields nothing is either an unhandled type or a shape we
    // failed to read; mark it with `?` so the console shows it instead of
    // silently dropping it.
    const key = events.length === 0 ? `?${rawType}` : rawType;
    this.eventCounts[key] = (this.eventCounts[key] ?? 0) + 1;
    if (events.length === 0) {
      log.warn('meeting', 'activity-unparsed', {
        meetingId: this.meetingId,
        type: rawType,
        keys: Object.keys(item).join(','),
      });
    }
    for (const event of events) {
      this.handle(event, eventId, opts);
    }
  }

  private handle(event: MeetingEvent, eid: string | undefined, opts: IngestOptions): void {
    const quiet = opts.quiet === true;
    if (event.kind === 'transcript') {
      this.handleTranscript(event, eid, quiet ? 'quiet' : opts.gapFill ? 'gap-fill' : 'live');
      return;
    }
    // message_type 3 is a reaction, not text.
    if (event.kind === 'chat' && event.messageType !== 3 && event.content) {
      // Push and poll may carry the same message under different dedup keys
      // (one with event_id, one without). Stored twice it would also be
      // answered twice.
      const kind = keyKind(eid);
      if (this.hasChat(event.from.id ?? event.from.name, event.content, quiet ? event.atMs : (event.atMs ?? this.now()), kind)) {
        return;
      }
      const line = this.addLine({
        kind: 'chat',
        speaker: event.from.name ?? event.from.id ?? '?',
        ...(event.from.id ? { speakerId: event.from.id } : {}),
        text: event.content,
        at: event.atMs ?? this.now(),
        seq: ++this.seq,
        ...(kind ? { keyKind: kind } : {}),
        ...(eid ? { eid } : {}),
      });
      this.persistLine(line, eid);
      if (quiet) this.quietGain += 1;
    }
    if (event.kind === 'participant') {
      const id = event.user.id;
      if (id) {
        if (event.action === 'joined') this.participantIds.add(id);
        else this.participantIds.delete(id);
      }
      // People leaving is the only warning before a meeting ends: re-read
      // everything while the events API still answers.
      if (event.action === 'left' && !quiet && id !== this.botOpenId) {
        this.scheduleBackfill('participant-left', LEAVE_BACKFILL_DELAY_MS);
      }
    }
    if (!quiet) this.emit(event);
  }

  /**
   * The same message already held under the *other* kind of dedup key: same
   * sender, same text, within two minutes of `at` (any time when unknown).
   * Two messages that both carry their own `event_id` are two messages, even
   * when someone types the same thing twice.
   */
  private hasChat(who: string | undefined, text: string, at: number | undefined, kind: KeyKind | undefined): boolean {
    return this.lines.some(
      (l) =>
        l.kind === 'chat' &&
        l.text === text &&
        (l.speakerId ?? l.speaker) === (who ?? '?') &&
        (kind === undefined || l.keyKind === undefined || l.keyKind !== kind) &&
        (at === undefined || Math.abs(l.at - at) < 120_000),
    );
  }

  private addLine(line: MeetingLine): MeetingLine {
    this.lines.push(line);
    if (line.sentenceId) this.bySentence.set(line.sentenceId, line);
    if (this.lines.length > MAX_LINES) {
      const drop = this.lines.splice(0, this.lines.length - MAX_LINES);
      for (const l of drop) if (l.sentenceId) this.bySentence.delete(l.sentenceId);
    }
    return line;
  }

  private persistLine(line: MeetingLine, eid: string | undefined): void {
    this.transcriptLog?.append({
      k: line.kind,
      ...(line.sentenceId ? { sid: line.sentenceId } : {}),
      who: line.speaker,
      ...(line.speakerId ? { whoId: line.speakerId } : {}),
      text: line.text,
      at: line.at,
      ...(eid ? { eid } : {}),
    });
  }

  private handleTranscript(
    event: TranscriptEvent,
    eid: string | undefined,
    lane: 'live' | 'gap-fill' | 'quiet',
  ): void {
    // The bot's own speech comes back through the meeting's transcription; not
    // filtering it here would let "answer whatever you hear" self-trigger.
    if (event.selfEcho) return;
    const latest = this.sentences.get(event.sentenceId);
    if (latest === event.text) return; // exact repeat
    if (lane !== 'live') {
      // These lanes replay older revisions too (a re-read from the first, the
      // gap-fill behind push); only a version longer than the newest one seen
      // — held or still settling — can be news. Known limit: a final shorter
      // correction that was itself missed can lose to a longer draft.
      const held = this.bySentence.get(event.sentenceId)?.text.length ?? -1;
      if (latest !== undefined && event.text.length <= Math.max(held, latest.length)) {
        this.noteSeen(eid);
        return;
      }
    }
    this.sentences.set(event.sentenceId, event.text);
    const isNew = !this.bySentence.has(event.sentenceId);

    const stabilizeMs = this.config.transcript.stabilizeMs;
    if (stabilizeMs <= 0 || lane === 'quiet') {
      // A quiet line is final as it is: drop the (older) revision still settling.
      const prior = this.pending.get(event.sentenceId);
      if (prior) {
        clearTimeout(prior.timer);
        this.pending.delete(event.sentenceId);
        this.noteSeen(prior.eid);
      }
      this.commitTranscript(event, eid, lane !== 'quiet');
      if (lane === 'quiet' && isNew) this.quietGain += 1;
      return;
    }
    // Debounce: a sentence is "final" once it stops growing for stabilizeMs.
    const prior = this.pending.get(event.sentenceId);
    if (prior) {
      clearTimeout(prior.timer);
      // The superseded revision is never written as a line; remember its id
      // so a restarted bridge doesn't take it for news on catch-up.
      this.noteSeen(prior.eid);
    }
    this.pending.set(event.sentenceId, {
      event,
      eid,
      timer: setTimeout(() => {
        this.pending.delete(event.sentenceId);
        this.commitTranscript(event, eid);
      }, stabilizeMs),
    });
  }

  /** Record an event id that produced no line, so a reload still knows it. */
  private noteSeen(eid: string | undefined): void {
    if (eid) this.transcriptLog?.append({ k: 'seen', seen: [eid] });
  }

  /** Commit what is still settling (no fan-out): the meeting is over or the process stops. */
  private flushPending(): void {
    for (const { timer, event, eid } of this.pending.values()) {
      clearTimeout(timer);
      this.commitTranscript(event, eid, false);
    }
    this.pending.clear();
  }

  /** Append (or replace) the sentence in the record, write it, then fan out. */
  private commitTranscript(event: TranscriptEvent, eid: string | undefined, emit = true): void {
    const speaker = event.speaker.name ?? event.speaker.id ?? '?';
    // Same sentence growing -> replace in place, so the record holds one entry
    // per sentence instead of every intermediate prefix. Keep the first time
    // and position so the sentence doesn't jump past chat sent while it was spoken.
    let line = this.bySentence.get(event.sentenceId);
    if (line) {
      line.text = event.text;
      line.speaker = speaker;
      if (eid) line.eid = eid;
    } else {
      line = this.addLine({
        kind: 'speech',
        sentenceId: event.sentenceId,
        speaker,
        ...(event.speaker.id ? { speakerId: event.speaker.id } : {}),
        text: event.text,
        at: event.atMs ?? this.now(),
        seq: ++this.seq,
        ...(eid ? { eid } : {}),
      });
    }
    this.persistLine(line, eid);
    if (emit) this.emit(event);
  }

  private emit(event: MeetingEvent): void {
    for (const handler of this.handlers.get(event.kind) ?? []) {
      try {
        handler(event);
      } catch (err) {
        log.warn('meeting', 'handler-failed', { kind: event.kind, err: String(err) });
      }
    }
  }

  /** Send a message into the meeting chat. */
  async sendMessage(text: string): Promise<void> {
    if (this.ended) return;
    await sendMeetingText(this.client, this.meetingId, text);
  }

  /**
   * Read the events the meeting already has, page by page, and leave the cursor
   * at the end. Used when re-adopting a meeting after a restart: it runs before
   * the agent subscribes, so old `@bot` questions fill the context without
   * being answered a second time. Throws on API errors — the caller decides
   * whether the meeting is gone.
   */
  async catchUp(maxPages = 1000): Promise<void> {
    for (let i = 0; i < maxPages; i++) {
      const page = await fetchMeetingEvents(this.client, {
        meetingId: this.meetingId,
        ...(this.cursor ? { pageToken: this.cursor } : {}),
      });
      if (page.pageToken) this.cursor = page.pageToken;
      for (const item of page.items) this.ingest(item, { quiet: true });
      if (!page.hasMore) return;
    }
    log.warn('meeting', 'catchup-truncated', { meetingId: this.meetingId, pages: maxPages });
  }

  /**
   * Re-read the meeting's events from the very first page and keep whatever
   * the live lanes missed (quietly: no answers to old questions, no alerts).
   * The poller's cursor is left alone. Concurrent calls share one run.
   * Resolves to the number of lines recovered; a `122001` ends the session
   * like the poller would; other errors are logged and resolve to 0.
   */
  backfill(reason: string, maxPages = 1000): Promise<number> {
    if (this.backfillRun) return this.backfillRun;
    if (this.ended) return Promise.resolve(0);
    this.lastBackfillAt = this.now();
    const run = (async (): Promise<number> => {
      const before = this.backfilled;
      let token: string | undefined;
      let pages = 0;
      try {
        for (; pages < maxPages; ) {
          const page = await fetchMeetingEvents(this.client, {
            meetingId: this.meetingId,
            ...(token ? { pageToken: token } : {}),
          });
          pages += 1;
          this.quietGain = 0;
          for (const item of page.items) this.ingest(item, { quiet: true });
          this.backfilled += this.quietGain;
          // Stop on the last page, and on a cursor that does not move (no loop).
          if (!page.hasMore || !page.pageToken || page.pageToken === token) break;
          token = page.pageToken;
        }
      } catch (err) {
        if (err instanceof VcApiError && err.meetingGone) {
          this.gone(err);
        } else {
          log.warn('meeting', 'backfill-failed', { meetingId: this.meetingId, reason, err: String(err) });
        }
      }
      this.backfills += 1;
      const added = this.backfilled - before;
      log.info('meeting', 'backfill', { meetingId: this.meetingId, reason, pages, added });
      return added;
    })();
    this.backfillRun = run.finally(() => {
      this.backfillRun = undefined;
    });
    return this.backfillRun;
  }

  /**
   * Re-read soon, unless one ran recently or one is already scheduled. Timers
   * die with the session ({@link stopTimers}).
   */
  private scheduleBackfill(reason: string, delayMs: number): void {
    if (this.stopped || this.ended || this.backfillTimer) return;
    const wait = Math.max(delayMs, this.lastBackfillAt + BACKFILL_MIN_GAP_MS - this.now());
    this.backfillTimer = setTimeout(() => {
      this.backfillTimer = undefined;
      void this.backfill(reason);
    }, wait);
    this.backfillTimer.unref?.();
  }

  /**
   * Start the periodic full re-read (every {@link FULL_BACKFILL_INTERVAL_MS}).
   * The manager calls this next to {@link startPolling}.
   */
  startBackfill(): void {
    if (this.stopped || this.ended) return;
    const tick = () => {
      this.periodic = undefined;
      if (this.stopped || this.ended) return;
      void this.backfill('periodic').finally(() => {
        if (this.stopped || this.ended) return;
        this.periodic = setTimeout(tick, FULL_BACKFILL_INTERVAL_MS);
        this.periodic.unref?.();
      });
    };
    this.periodic = setTimeout(tick, FULL_BACKFILL_INTERVAL_MS);
    this.periodic.unref?.();
  }

  /** The events API answered "meeting over": stop, and let the manager wrap up. */
  private gone(err: VcApiError): void {
    if (this.ended) return;
    log.info('meeting', 'ended-observed', { meetingId: this.meetingId, code: err.code });
    this.markEnded();
    this.onGone?.(err);
  }

  /**
   * Compact the transcript file to its final state and write the readable
   * `.txt` next to it. Returns the `.txt` path, or undefined when saving is off
   * or failed. Safe to call more than once.
   */
  finalizeTranscript(): string | undefined {
    if (!this.transcriptLog || !this.textFile) return undefined;
    const ordered = this.ordered();
    const lines: TranscriptLine[] = ordered.map((l) => ({
      k: l.kind,
      ...(l.sentenceId ? { sid: l.sentenceId } : {}),
      who: l.speaker,
      ...(l.speakerId ? { whoId: l.speakerId } : {}),
      text: l.text,
      at: l.at,
      ...(l.eid ? { eid: l.eid } : {}),
    }));
    // Every id seen stays in the file: the bot may join this meeting again,
    // and a new session reading from the first page must not take old events
    // for new ones (re-answering old questions, rolling sentences back).
    const records: TranscriptRecord[] = [...lines, { k: 'seen', seen: [...this.seenEvents] }];
    const text = renderTranscriptText({ meetingNo: this.meetingNo, ...(this.topic ? { topic: this.topic } : {}) }, lines);
    const out = this.transcriptLog.finalize(records, this.textFile, text);
    if (out) {
      this.savedText = out;
      log.info('meeting', 'transcript-saved', { meetingId: this.meetingId, lines: lines.length });
    }
    return out;
  }

  /**
   * Start the `bots/events` poller. Used as the primary source until push is
   * proven, and as gap-fill afterwards. Idle rounds back off to 10s.
   */
  startPolling(): void {
    if (this.poller || this.stopped) return;
    const tick = async (): Promise<void> => {
      if (this.stopped || this.ended) return;
      try {
        const page = await fetchMeetingEvents(this.client, {
          meetingId: this.meetingId,
          ...(this.cursor ? { pageToken: this.cursor } : {}),
        });
        this.pollFailures = 0;
        this.polled += page.items.length;
        if (page.pageToken) this.cursor = page.pageToken;
        const gapFill = this.source === 'push';
        for (const item of page.items) this.ingest(item, { gapFill });
        this.idleRounds = page.items.length > 0 ? 0 : Math.min(this.idleRounds + 1, 3);
        // More pages queued → come back immediately instead of sleeping.
        if (page.hasMore) this.idleRounds = -1;
      } catch (err) {
        if (err instanceof VcApiError && err.meetingGone) {
          this.gone(err);
          return;
        }
        this.idleRounds = Math.min(this.idleRounds + 1, 3);
        log.warn('meeting', 'poll-failed', { meetingId: this.meetingId, err: String(err) });
        // Report once per streak: a poller that keeps failing means the bot is
        // deaf to the meeting, which is otherwise only visible in the log.
        this.pollFailures += 1;
        if (this.pollFailures === POLL_FAILURE_ALERT) this.onTrouble?.('poll-failed', err);
      }
      if (this.stopped || this.ended) return;
      const base = this.config.pollIntervalMs;
      const delay =
        this.idleRounds < 0
          ? 0
          : this.source === 'push'
            ? GAP_FILL_INTERVAL_MS
            : Math.min(base * 2 ** this.idleRounds, 10_000);
      this.poller = setTimeout(() => void tick(), delay);
    };
    this.poller = setTimeout(() => void tick(), 0);
  }

  private stopPolling(): void {
    if (this.poller) clearTimeout(this.poller);
    this.poller = undefined;
  }

  /** Called when the meeting ends server-side: stop work, keep the transcript. */
  markEnded(): void {
    this.ended = true;
    this.stopTimers();
  }

  private stopTimers(): void {
    this.stopped = true;
    this.stopPolling();
    if (this.backfillTimer) clearTimeout(this.backfillTimer);
    this.backfillTimer = undefined;
    if (this.periodic) clearTimeout(this.periodic);
    this.periodic = undefined;
    // Sentences still settling are part of the meeting: keep them.
    this.flushPending();
  }

  /** Leave the meeting. Idempotent — safe to call after the meeting ended. */
  async leave(): Promise<void> {
    if (this.leaving) return;
    this.leaving = true;
    this.stopTimers();
    if (this.ended) return; // already over; nothing to leave
    this.ended = true;
    await leaveMeeting(this.client, this.meetingId).catch((err) =>
      log.warn('meeting', 'leave-failed', { meetingId: this.meetingId, err: String(err) }),
    );
  }

  /** Tear down local timers without calling the API (process shutdown). */
  dispose(): void {
    this.stopTimers();
    this.handlers.clear();
  }
}

function formatLine(l: MeetingLine): string {
  return l.kind === 'chat' ? `${l.speaker}（弹幕）: ${l.text}` : `${l.speaker}: ${l.text}`;
}

/**
 * Stable key for an activity item that carries no `event_id`. Built from the
 * type plus the identity-bearing fields of its items, so the same event
 * arriving over both lanes collapses to one.
 */
function contentDedupKey(item: RawActivityItem): string | undefined {
  const type = item.activity_event_type ?? (item as { event_type?: string }).event_type;
  if (!type) return undefined;
  const payload = item as Record<string, unknown>;
  const nested = (payload.payload && typeof payload.payload === 'object' ? payload.payload : {}) as Record<string, unknown>;
  const arr = Object.entries({ ...payload, ...nested }).find(
    ([k, v]) => k.endsWith('_items') && Array.isArray(v),
  )?.[1] as unknown[] | undefined;
  if (!arr || arr.length === 0) return undefined;
  const parts = arr.map((entry) => {
    const it = (entry && typeof entry === 'object' ? entry : {}) as Record<string, unknown>;
    // sentence_id / content identify a transcript line or chat message; the
    // timestamp disambiguates repeated identical text.
    return [
      it.sentence_id ?? it.message_id ?? '',
      it.content ?? it.text ?? '',
      it.start_time_ms ?? it.create_time ?? it.time ?? '',
      it.open_id ?? (it.sender as { open_id?: string } | undefined)?.open_id ?? '',
    ].join('|');
  });
  return `${type}#${parts.join('~')}`;
}
