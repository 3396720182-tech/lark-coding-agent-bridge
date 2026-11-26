import type { LarkChannel } from '@larksuite/channel';
import type { AgentEvent } from '../agent/types';
import { claudeCapability, codexCapability } from '../agent/capability';
import type { Controls } from '../commands';
import { log } from '../core/logger';
import type { RunExecutor } from '../runtime/run-executor';
import type { ActiveRuns } from '../bot/active-runs';
import type { SessionCatalog } from '../session/catalog';
import type { SessionStore } from '../session/store';
import type { WorkspaceStore } from '../workspace/store';
import { startRunFlow } from '../bot/run-flow';
import { describeMeetingError } from './manager';
import { attachMeetingAlerts } from './alerts';
import { attachCalendarDocs } from './calendar';
import { attachProactive } from './proactive';
import { attachVoiceWake } from './voice';
import { meetingAuthzGate, meetingEffectiveAccess, meetingReadOnlyAccess } from './authz';
import type { MeetingSession } from './session';
import type { AccessMode, MeetingSummaryTarget, ProfileConfig } from '../config/profile-schema';
import type { ChatEvent, MeetingEvent } from './types';
import {
  adoptPendingDocsOnJoin,
  attachShareDocs,
  docCitations,
  docLabel,
  meetingDocContext,
  meetingDocRefs,
  recordMeetingQa,
  releaseMeetingDocs,
  sendAnswerToIm,
  type DocRef,
  type UsedChunk,
} from './docs';
import {
  MEETING_MINUTES_QUESTION,
  meetingMinutesQuestion,
  meetingArchiveFile,
  meetingArchiveId,
  meetingMinutesCard,
  localDate,
  splitMinutesTodos,
  TO_CONFIRM,
  type MinutesRecord,
} from './minutes';
import { meetingDataDir, writeMeetingJson } from './store';

/**
 * Turns in-meeting content into agent runs.
 *
 * Deliberate policy: the transcript is only ever **context**, never a trigger.
 * Running the coding agent because somebody said something out loud would be
 * both expensive and unsafe (an offhand remark could start editing code), so a
 * run needs an explicit ask — an in-meeting chat message prefixed with the
 * configured trigger, or a `/meeting ask` from IM.
 */

export interface MeetingAgentDeps {
  session: MeetingSession;
  channel: LarkChannel;
  controls: Controls;
  executor: RunExecutor;
  /** Needed to interrupt this meeting's run from inside the meeting. */
  activeRuns: ActiveRuns;
  sessions: SessionStore;
  sessionCatalog?: SessionCatalog;
  workspaces: WorkspaceStore;
}

/** Scope id for a meeting's agent session — one conversation per meeting. */
export function meetingScopeId(meetingId: string): string {
  return `meeting:${meetingId}`;
}

/**
 * Scope of one asker's in-meeting questions. Per asker, not per meeting: two
 * people asking at once each get their own run; the same person asking again
 * while their first answer is still running is told to wait (or stop it).
 */
export function meetingAskScopeId(meetingId: string, askerId: string): string {
  return `${meetingScopeId(meetingId)}:ask:${askerId}`;
}

/**
 * In-meeting questions running (or waiting for a process slot) at once, per
 * meeting. Without a cap, a crowd asking together — in team mode anyone in the
 * meeting may — would take every agent process the bridge has.
 */
export const MEETING_MAX_PARALLEL = 3;

/** meetingId → scope → state of each in-meeting question run, from submit until it finishes. */
const askRuns = new Map<string, Map<string, { cancelled: boolean }>>();

/** Questions of this meeting in flight (for tests and status). */
export function meetingAsksInFlight(meetingId: string): number {
  return askRuns.get(meetingId)?.size ?? 0;
}

/**
 * Stop a meeting's runs: one asker's (`askerId`), or every question run of
 * the meeting (`/meeting stop`). The minutes run is never touched. A question
 * still waiting for a process slot is cancelled too (it stops as soon as it
 * starts). Returns how many were stopped.
 */
export function interruptMeetingRuns(
  activeRuns: Pick<ActiveRuns, 'interrupt'> & Partial<Pick<ActiveRuns, 'scopes'>>,
  meetingId: string,
  askerId?: string,
): number {
  const tracked = askRuns.get(meetingId) ?? new Map<string, { cancelled: boolean }>();
  const stopped = new Set<string>();
  const stop = (scope: string) => {
    if (activeRuns.interrupt(scope)) stopped.add(scope);
    const pending = tracked.get(scope);
    if (pending && !pending.cancelled) {
      pending.cancelled = true;
      stopped.add(scope);
    }
  };
  if (askerId) {
    stop(meetingAskScopeId(meetingId, askerId));
    return stopped.size;
  }
  const base = meetingScopeId(meetingId);
  const minutes = meetingMinutesScopeId(meetingId);
  const scopes = new Set([...(activeRuns.scopes?.() ?? [base]), ...tracked.keys()]);
  for (const scope of scopes) {
    if (scope === minutes || (scope !== base && !scope.startsWith(`${base}:`))) continue;
    stop(scope);
  }
  return stopped.size;
}

/**
 * Words that mean "stop the current run" when sent after the trigger. Runs are
 * serialized per meeting, so a hung run (e.g. an agent whose SessionEnd hook
 * never exits) would otherwise block every later question with no way out —
 * there is no `/stop` inside a meeting, only the chat lane.
 */
const STOP_WORDS = new Set(['stop', '停', '停止', '中断', '取消', 'cancel', 'abort']);

/**
 * Match an in-meeting chat line against the accepted trigger prefixes and
 * return what follows.
 *
 * Returns `undefined` when nothing matched, or the remaining text (possibly
 * empty) when it did — the caller needs to tell "not for me" apart from "for me
 * but nothing was asked".
 *
 * In-meeting chat is plain text with no mention autocomplete, so people type
 * whatever the bot is called in the participant list. Accepting `@<botName>`
 * alongside the configured prefix means the natural thing works with no config.
 * Punctuation right after the prefix is dropped, so a prefix followed by a
 * comma behaves the same as one followed by a space.
 */
export function matchTrigger(text: string, prefixes: string[]): string | undefined {
  const trimmed = text.trim();
  const lower = trimmed.toLowerCase();
  const head = matchPrefix(trimmed, prefixes);
  if (head !== undefined) return head;
  // People also mention the bot mid- or end-of-sentence ("说句话听听@助手").
  // Accept the mention anywhere, unless it is glued to a latin word or an
  // email/domain on either side ("a@bot.com", "@bottle" don't count).
  for (const prefix of prefixes) {
    const p = prefix.trim();
    if (!p) continue;
    const pl = p.toLowerCase();
    const endsWord = LATIN_WORD.test(p[p.length - 1] ?? '');
    for (let idx = lower.indexOf(pl); idx >= 0; idx = lower.indexOf(pl, idx + 1)) {
      const before = trimmed[idx - 1] ?? '';
      const after = trimmed[idx + p.length] ?? '';
      const afterNext = trimmed[idx + p.length + 1] ?? '';
      if (EMAIL_LOCAL.test(before)) continue;
      if (endsWord && (LATIN_WORD.test(after) || (after === '.' && LATIN_WORD.test(afterNext))))
        continue;
      return `${trimmed.slice(0, idx)} ${trimmed.slice(idx + p.length)}`
        .replace(/^[\s,，、:：。.!！?？~～-]+/, '')
        .replace(/\s+/g, ' ')
        .trim();
    }
  }
  return undefined;
}

const escapeRe = (s: string): string => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

/**
 * The elevation word counts only directly after a trigger mention
 * (`@bot !write 改 README`), never elsewhere in the sentence ("@bot 别用
 * !write"), and only as a whole word (`!writer` is not `!write`). Matched on
 * the raw text: trigger matching strips leading punctuation, which would eat
 * the `!` of `!write`. Returns the text without trigger and word, or undefined.
 */
export function matchElevation(content: string, prefixes: string[], word: string): string | undefined {
  const w = word.trim();
  if (!w) return undefined;
  for (const prefix of prefixes) {
    const p = prefix.trim();
    if (!p) continue;
    const re = new RegExp(`${escapeRe(p)}[\\s,，、:：]*${escapeRe(w)}(?=$|[\\s,，、:：。.!！?？])`, 'gi');
    let m: RegExpExecArray | null;
    // Same left boundary as the trigger: `foo@bot !write` is an email-ish token, not a mention.
    while ((m = re.exec(content)) && EMAIL_LOCAL.test(content[m.index - 1] ?? '')) {
      /* keep looking */
    }
    if (!m) continue;
    return `${content.slice(0, m.index)} ${content.slice(m.index + m[0].length)}`
      .replace(/^[\s,，、:：。.!！?？~～-]+/, '')
      .replace(/\s+/g, ' ')
      .trim();
  }
  return undefined;
}

/**
 * Prefix-only match: the text must start with one of `prefixes`. Returns the
 * rest (punctuation after the prefix dropped), or `undefined`.
 */
export function matchPrefix(text: string, prefixes: string[]): string | undefined {
  const trimmed = text.trim();
  const lower = trimmed.toLowerCase();
  for (const prefix of prefixes) {
    const p = prefix.trim();
    if (!p) continue;
    if (!lower.startsWith(p.toLowerCase())) continue;
    return trimmed
      .slice(p.length)
      .replace(/^[\s,，、:：。.!！?？~～-]+/, '')
      .trim();
  }
  return undefined;
}

const LATIN_WORD = /[A-Za-z0-9_]/;
const EMAIL_LOCAL = /[A-Za-z0-9_.+-]/;

/** Configured prefix plus `@<botName>`, de-duplicated. */
export function triggerPrefixes(configured: string, botName?: string): string[] {
  const list = [configured, botName ? `@${botName}` : ''];
  return [...new Set(list.map((p) => p.trim()).filter(Boolean))];
}

/** Subscribe the agent to a freshly-created session. */
export function attachMeetingAgent(deps: MeetingAgentDeps): void {
  const { session, controls } = deps;
  session.onTrouble = (kind, err) =>
    notifyMeetingOwner(
      deps,
      kind,
      `会中事件拉取连续失败（${describeMeetingError(err)}），bot 可能收不到会中消息。`,
    );
  session.on('chat', (event) => {
    const chat = event as ChatEvent;
    // message_type 3 is an in-meeting reaction, not text — never a question.
    if (chat.messageType === 3) return;
    // Our own answers come back as chat too. Now that a mention anywhere
    // counts, an answer quoting "@<botName>" must not trigger another run.
    const selfId = deps.channel.botIdentity?.openId;
    if (selfId && chat.from.id === selfId) return;
    // Read the bot name live: it is only known after connect and can change.
    const prefixes = triggerPrefixes(
      controls.profileConfig.meeting.trigger,
      deps.channel.botIdentity?.name,
    );
    const matched = matchTrigger(chat.content, prefixes);
    if (matched === undefined) return; // not addressed to us
    if (!matched) return; // addressed but empty
    const usedPrefix = prefixes.find((p) =>
      chat.content.trim().toLowerCase().includes(p.toLowerCase()),
    );
    // Authorization comes before the stop words: otherwise anyone in the
    // meeting could interrupt a run somebody else started.
    const gate = meetingAuthzGate(deps, chat, matched, (word) => matchElevation(chat.content, prefixes, word));
    if (!gate.ok) return;
    const question = gate.question;

    if (STOP_WORDS.has(question.toLowerCase())) {
      // Your own question only: others' answers keep going.
      const stopped = interruptMeetingRuns(deps.activeRuns, session.meetingId, gate.senderId) > 0;
      void session
        .sendMessage(stopped ? '已中断你的问题。' : '你现在没有正在回答的问题。')
        .catch((err) => log.warn('meeting', 'stop-reply-failed', { err: String(err) }));
      return;
    }

    if (gate.elevate) {
      // Elevation is announced to everyone in the meeting, so an attempt to
      // go beyond read-only is never silent. The session switches with the
      // policy: read-only history is not carried into a writable session.
      void session
        .sendMessage(
          `⚠️ 本次以 ${meetingEffectiveAccess(controls.profileConfig, true)} 权限执行（由 ${chat.from.name ?? '授权用户'} 发起）。提权会切换会话，不带之前的会中对话历史。`,
        )
        .catch((err) => log.warn('meeting', 'elevate-notice-failed', { err: String(err) }));
    }

    // Asked from inside the meeting → answer where the conversation is.
    void answerInMeeting(deps, question, {
      deliver: 'broadcast',
      elevate: gate.elevate,
      ...(usedPrefix ? { usedPrefix } : {}),
      ...(chat.from.name ? { askedBy: chat.from.name } : {}),
      askerId: gate.senderId,
    }).catch((err) =>
      log.warn('meeting', 'answer-failed', { meetingId: session.meetingId, err: String(err) }),
    );
  });
  // Say who we are on arrival — not again when a restart re-adopts the meeting.
  if (!session.resumedAt && controls.profileConfig.meeting.announceOnJoin) {
    const text = meetingJoinNotice(controls.profileConfig, deps.channel.botIdentity?.name);
    void session
      .sendMessage(text)
      .then(() => log.info('meeting', 'join-notice-sent', { meetingId: session.meetingId }))
      .catch((err) =>
        log.warn('meeting', 'join-notice-failed', { meetingId: session.meetingId, err: describeMeetingError(err) }),
      );
  }
  // Additional subscriber: session.on keeps a Set per kind.
  attachMeetingAlerts(deps);
  // Pre-meeting docs: take the ones waiting for this meeting, bind docs
  // shared on screen in the meeting, and the ones linked in its calendar event.
  adoptPendingDocsOnJoin(deps);
  attachShareDocs(deps);
  void attachCalendarDocs(deps);
  // Both opt-in; they check the live config on every event.
  attachVoiceWake(
    { session, controls, botName: () => deps.channel.botIdentity?.name },
    (ask) => askByVoice(deps, ask),
  );
  attachProactive({
    session,
    controls,
    selfId: () => deps.channel.botIdentity?.openId,
    addressedToBot: (text) =>
      matchTrigger(text, triggerPrefixes(controls.profileConfig.meeting.trigger, deps.channel.botIdentity?.name)) !==
      undefined,
    run: (prompt) =>
      runReadOnlyPrompt(deps, prompt, `${meetingScopeId(session.meetingId)}:proactive`).then((r) => r.text),
  });
}

/** A spoken question from an allowed speaker (see voice.ts): stop words, an acknowledgement, then the answer. */
function askByVoice(deps: MeetingAgentDeps, ask: { question: string; speakerId: string; speakerName?: string }): void {
  const { session } = deps;
  const who = ask.speakerName ?? '你';
  const q = ask.question.replace(/[。.!！?？\s]+$/, '');
  if (STOP_WORDS.has(q.toLowerCase())) {
    const stopped = interruptMeetingRuns(deps.activeRuns, session.meetingId, ask.speakerId) > 0;
    void session
      .sendMessage(stopped ? `已中断 ${who} 的问题。` : `${who} 现在没有正在回答的问题。`)
      .catch((err) => log.warn('meeting', 'stop-reply-failed', { err: String(err) }));
    return;
  }
  // Speech can be misheard: say what was heard, so a wrong wake is easy to see (and stop).
  void session
    .sendMessage(`🎙️ 听到 ${who} 叫我：「${ask.question.slice(0, 60)}」，正在查…`)
    .catch((err) => log.warn('meeting', 'voice-ack-failed', { err: String(err) }));
  void answerInMeeting(deps, ask.question, {
    deliver: 'broadcast',
    via: 'voice',
    askerId: ask.speakerId,
    ...(ask.speakerName ? { askedBy: ask.speakerName } : {}),
  }).catch((err) => log.warn('meeting', 'answer-failed', { meetingId: session.meetingId, err: String(err) }));
}

/**
 * The notice posted in the meeting chat on joining: who the bot is, that it is
 * recording, how to ask it, and who may. Disclosure first — participants see a
 * bot in the list but not what it does with what it hears.
 */
export function meetingJoinNotice(profile: ProfileConfig, botName?: string): string {
  const m = profile.meeting;
  const name = botName?.trim() || 'AI 会议助手';
  const ask = triggerPrefixes(m.trigger, botName)[botName ? 1 : 0] ?? m.trigger;
  // Team mode lets anyone in the meeting drive it (canUseDm), outsiders included.
  const who =
    m.access === 'admins-only'
      ? '目前只有机器人的主人和管理员能让我执行，其他人发了不会执行'
      : profile.mode === 'team'
        ? '会里的每个人都可以问我'
        : '目前只有机器人的主人和授权名单里的人能让我执行，其他人发了不会执行';
  const where =
    m.respondIn === 'im' ? '回答会发到发起入会的聊天里' : m.respondIn === 'both' ? '回答发在这里，也会发到发起入会的聊天里' : '我会回在这里';
  const lines = [
    botName?.trim() ? `👋 我是「${name}」，AI 会议助手，已加入本场会议。` : `👋 我是 AI 会议助手，已加入本场会议。`,
    `· 我在记录本场的字幕和会中聊天（带发言人和时间），用来回答问题${m.summaryOnEnd ? '，散会后整理成纪要' : ''}。`,
    `· 有问题在聊天里发「${ask} 你的问题」，${where}${m.respondIn !== 'im' && m.longAnswerChars > 0 ? '；长回答这里只发要点，全文私聊发给提问的人' : ''}。`,
    ...(m.voiceWake.enabled
      ? [`· 授权名单里的人也可以开口叫我：句子开头说「${(m.voiceWake.names.length ? m.voiceWake.names : [name]).join('」或「')}」再说问题。`]
      : []),
    ...(m.proactive.enabled
      ? [`· 会上有人提问、过一会儿没人答，而会前资料里正好有答案时，我会插一句（每场最多 ${m.proactive.maxPerMeeting} 次）。`]
      : []),
    `· ${who}。`,
    '· 不希望被记录，可以请主持人把我移出会议。',
  ];
  return lines.join('\n');
}

export interface AnswerOptions {
  /** Name of the participant who asked (in-meeting triggers only). */
  askedBy?: string;
  /**
   * open_id of whoever asked: their questions run in their own scope (people
   * asking at once don't queue behind each other), and the full text of a
   * long answer goes to their DM.
   */
  askerId?: string;
  /** How the question was put: typed in the meeting chat (default) or said out loud. */
  via?: 'chat' | 'voice';
  /**
   * Who sees the answer.
   *  - `broadcast`: deliver per `respondIn` — used when the ask came from
   *    inside the meeting, where everyone is already in the conversation.
   *  - `caller`: return the text and deliver nothing — used by `/meeting
   *    notes|ask`, so a privately-typed command does NOT push its answer into
   *    the meeting where everybody would see it.
   */
  deliver: 'broadcast' | 'caller';
  /** Prefix the asker actually typed, so hints quote their wording back. */
  usedPrefix?: string;
  /**
   * Run with the profile's default access instead of the meeting tier. Only
   * the in-meeting elevation word sets this; `/meeting ask|notes` and the
   * end-of-meeting summary stay read-only (summarizing needs no writes).
   */
  elevate?: boolean;
  /**
   * `minutes`: the end-of-meeting summary. Not bound by the in-meeting length
   * limit, given the doc list instead of doc chunks, not logged as Q&A, and
   * run in its own agent scope (an in-meeting run still going must not block
   * it, and a rejection is an error rather than an answer).
   */
  purpose?: 'minutes';
  /**
   * The meeting record to give instead of the last `transcript.keep` lines:
   * the whole meeting (`full`) or per-segment summaries of it (`segments`).
   */
  context?: { kind: 'full' | 'segments'; lines: string[] };
}

/**
 * Run the agent with the meeting transcript as context. Returns the answer;
 * delivery depends on {@link AnswerOptions.deliver}.
 */
export async function answerInMeeting(
  deps: MeetingAgentDeps,
  question: string,
  opts: AnswerOptions,
): Promise<string> {
  const { session, controls } = deps;
  const config = controls.profileConfig.meeting;
  // Time-labelled, so the answer can cite who said it when.
  const transcript = opts.context?.lines ?? session.recentTimedContext(config.transcript.keep);
  const docs = meetingDocContext(deps, question, transcript, opts.purpose);
  // Only an answer posted into the meeting chat is split; IM copies and
  // `/meeting ask|notes` get everything.
  const intoMeeting =
    opts.deliver === 'broadcast' && opts.purpose !== 'minutes' && config.respondIn !== 'im';
  const prompt = buildMeetingPrompt({
    question,
    transcript,
    ...(opts.context ? { contextKind: opts.context.kind } : {}),
    topic: session.topic,
    ...(opts.askedBy ? { askedBy: opts.askedBy } : {}),
    ...(opts.purpose ? { purpose: opts.purpose } : {}),
    ...(opts.via === 'voice' ? { via: 'voice' as const } : {}),
    lane: opts.deliver === 'broadcast' ? 'meeting' : 'private',
    ...(intoMeeting && opts.askerId ? { longAnswerChars: config.longAnswerChars } : {}),
    ...docs.prompt,
  });

  // In-meeting chat runs at the meeting tier (or elevated); everything else
  // (`/meeting ask|notes`, the summary) is read-only: its input is meeting content.
  // Spoken questions stay read-only too: their text is whatever speech recognition heard.
  const tier =
    opts.deliver === 'broadcast' && opts.via !== 'voice'
      ? meetingEffectiveAccess(controls.profileConfig, opts.elevate === true)
      : meetingReadOnlyAccess(controls.profileConfig);
  const raw = await runMeetingAgent(deps, prompt, {
    ...(opts.usedPrefix ? { usedPrefix: opts.usedPrefix } : {}),
    elevate: opts.elevate === true,
    tier,
    ...(opts.purpose === 'minutes'
      ? { minutes: true }
      : opts.deliver === 'broadcast'
        ? { scopeId: meetingAskScopeId(session.meetingId, opts.askerId ?? '?'), track: true }
        : // `/meeting ask` from a chat: its own scope, not the asker's in-meeting one.
          { scopeId: meetingAskScopeId(session.meetingId, `im:${opts.askerId ?? '?'}`) }),
  });
  if (!raw) return '';
  if (opts.purpose === 'minutes') return raw;
  // Citations of what was said: one that matches no line given is marked, not trusted.
  const said = transcriptCitations(raw, transcript);
  const parts = splitAnswer(said.text, intoMeeting && opts.askerId ? config.longAnswerChars : 0);
  const answer = parts.full;
  // Which of the provided chunks the answer actually cites — those are its
  // grounds; citations matching nothing provided are logged as unverified.
  const { cited, unverified } = docCitations(answer, docs.used);
  recordMeetingQa(deps, {
    question,
    answer,
    ...(opts.askedBy ? { askedBy: opts.askedBy } : {}),
    provided: docs.used,
    cited,
    unverified,
    said: { cited: said.cited, unverified: said.unverified },
  });
  // The caller relays it themselves; never echo into the meeting.
  if (opts.deliver === 'caller') return answer;

  const target = config.respondIn;
  if (target === 'meeting' || target === 'both') {
    let inMeeting = answer;
    if (parts.points !== undefined && opts.askerId) {
      // Long answer: the full text to the asker's DM first; the meeting only
      // gets the points once that worked, otherwise the whole answer.
      const dm = await sendAnswerToIm(deps, opts.askerId, question, answer, opts.askedBy, cited)
        .then(() => true)
        .catch((err) => {
          log.warn('meeting', 'answer-dm-failed', { meetingId: session.meetingId, err: String(err) });
          return false;
        });
      if (dm) inMeeting = `${parts.points}\n（完整回答已私聊发给${opts.askedBy ? ` ${opts.askedBy}` : '你'}）`;
      log.info('meeting', 'answer-split', { meetingId: session.meetingId, chars: answer.length, dm });
    }
    await session.sendMessage(inMeeting).catch(async (err) => {
      const reason = describeMeetingError(err);
      log.warn('meeting', 'send-meeting-message-failed', { meetingId: session.meetingId, err: reason });
      // An answer that only exists in the log is lost. With respondIn=both and
      // an origin chat, the IM copy below already covers it. Otherwise DM the
      // owner only: the origin chat may be a group whose members aren't in the
      // meeting.
      if (target === 'both' && session.originChatId) return;
      const to = controls.botOwnerId;
      if (!to) return;
      await deps.channel
        .send(to, { markdown: `⚠️ 回答没能发进会议（${reason}），改私聊发给你：\n\n${answer}` })
        .catch((e) => log.warn('meeting', 'send-fallback-failed', { err: String(e) }));
    });
  }
  if ((target === 'im' || target === 'both') && session.originChatId) {
    await sendAnswerToIm(deps, session.originChatId, question, answer, opts.askedBy, cited).catch((err) =>
      log.warn('meeting', 'send-im-failed', { err: String(err) }),
    );
  }
  return answer;
}

/** Line in the model's answer that separates the key points from the full text. */
export const FULL_TEXT_MARK = '===全文===';

/**
 * Split an answer for the meeting chat. `points` is set when it should be
 * split: the model wrote {@link FULL_TEXT_MARK} and the whole is longer than
 * `max`, or it is simply longer than `max` (then the points are its opening,
 * cut at a sentence end). `full` is always the complete answer without the
 * marker. `max` 0 never splits.
 */
export function splitAnswer(answer: string, max: number): { full: string; points?: string } {
  const text = String(answer ?? '').trim();
  const at = text.indexOf(FULL_TEXT_MARK);
  const before = at >= 0 ? text.slice(0, at).trim() : '';
  const after = at >= 0 ? text.slice(at + FULL_TEXT_MARK.length).trim() : '';
  // With a marker the full text is what follows it (it restates the points);
  // a marker with nothing after it is just noise.
  const full = at < 0 ? text : after || before;
  if (max <= 0 || full.length <= max) return { full };
  if (before && before.length < full.length) return { full, points: before };
  return { full, points: openingOf(full, Math.min(max, 200)) };
}

/** The first `max` characters, cut back to the last sentence end when there is one. */
function openingOf(text: string, max: number): string {
  const head = text.slice(0, max);
  const cut = Math.max(...['。', '！', '？', '；', '\n', '. ', '! ', '? '].map((p) => head.lastIndexOf(p)));
  return `${(cut >= max / 3 ? head.slice(0, cut + 1) : head).trim()}…`;
}

/**
 * Citations of what was said in the meeting, 〔14:32 张三〕. One is `cited`
 * when a line given to the model has that speaker within a minute of that
 * time; otherwise it is `unverified` and marked in the text, so nobody takes
 * a made-up quote for a real one. Lines are `[HH:mm:ss] who: text`.
 */
export function transcriptCitations(
  answer: string,
  lines: string[],
): { text: string; cited: string[]; unverified: string[] } {
  const given = lines
    .map((l) => /^\[(\d{2}):(\d{2}):\d{2}\]\s*(.+?)(?:（弹幕）)?:\s/.exec(l))
    .filter((m): m is RegExpExecArray => Boolean(m))
    .map((m) => ({ min: Number(m[1]) * 60 + Number(m[2]), who: m[3]!.trim() }));
  const cited: string[] = [];
  const unverified: string[] = [];
  const text = String(answer ?? '').replace(
    /〔\s*(\d{1,2})[:：](\d{2})(?:[:：]\d{2})?\s*([^〔〕《》§\d][^〔〕《》§]{0,19}?)\s*〕/g,
    (all, h: string, m: string, who: string) => {
      const label = `${h.padStart(2, '0')}:${m} ${who.trim()}`;
      const min = Number(h) * 60 + Number(m);
      const name = who.trim();
      const ok = given.some((g) => Math.abs(g.min - min) <= 1 && (g.who === name || g.who.includes(name) || name.includes(g.who)));
      if (ok) {
        if (!cited.includes(label)) cited.push(label);
        return all;
      }
      if (!unverified.includes(label)) unverified.push(label);
      return `〔${label}（未核对到原话）〕`;
    },
  );
  return { text, cited, unverified };
}

export interface MeetingPromptInput {
  question: string;
  transcript: string[];
  /**
   * What `transcript` holds: the latest lines (default), the whole meeting
   * with time labels, or per-segment summaries of a long meeting.
   */
  contextKind?: 'recent' | 'full' | 'segments';
  topic?: string;
  askedBy?: string;
  purpose?: 'minutes';
  /** Doc chunks picked for this question, each shown with its citation. */
  materials?: UsedChunk[];
  /** Docs bound to the meeting (listed when no chunk matched, and for minutes). */
  materialList?: DocRef[];
  /** Where the answer goes: the meeting chat (default) or only the asker. */
  lane?: 'meeting' | 'private';
  /**
   * Set (above 0) when a long answer may be split: points in the meeting, the
   * full text by DM. The model is told how to mark the split.
   */
  longAnswerChars?: number;
  /** The question was said out loud (the transcript may garble it). */
  via?: 'voice';
}

/** Compose the agent prompt: meeting context first, then the actual ask. */
export function buildMeetingPrompt(input: MeetingPromptInput): string {
  const parts: string[] = [];
  const kind = input.contextKind ?? 'recent';
  if (kind === 'segments') {
    parts.push('你参加了一场飞书视频会议。会议很长，记录已按时间分段整理成摘要，以下是各段摘要。');
  } else {
    parts.push('你正在参加一场飞书视频会议，以下是会议的实时字幕和会中弹幕（标「弹幕」的是文字消息）。');
  }
  if (input.topic) parts.push(`会议主题：${input.topic}`);
  parts.push('');
  parts.push(
    kind === 'segments'
      ? '=== 分段摘要（按时间顺序） ==='
      : kind === 'full'
        ? '=== 会议记录（整场，按时间顺序，方括号里是时间） ==='
        : '=== 会议记录（最近的部分，按时间顺序，方括号里是时间） ===',
  );
  parts.push(input.transcript.length ? input.transcript.join('\n') : '（暂无字幕和弹幕）');
  parts.push(kind === 'segments' ? '=== 摘要结束 ===' : '=== 记录结束 ===');
  if (input.materials?.length) {
    parts.push('');
    parts.push('=== 会前资料（按和问题的相关度摘录，方括号里是出处） ===');
    for (const m of input.materials) parts.push(`〔${docLabel(m.doc)} §${m.i}〕\n${m.text}`);
    parts.push('=== 资料结束 ===');
  } else if (input.materialList?.length) {
    parts.push('');
    parts.push(
      `本场会议绑定的会前资料：${input.materialList.map(docLabel).join('、')}（这次没有摘到和问题相关的段落）`,
    );
  }
  parts.push('');
  parts.push(
    input.askedBy
      ? input.via === 'voice'
        ? `参会人「${input.askedBy}」在会上口头叫你、向你提问（语音转写，可能有错字）：`
        : `参会人「${input.askedBy}」在会中向你提问：`
      : '收到的请求：',
  );
  parts.push(input.question);
  parts.push('');
  if (input.purpose === 'minutes') {
    parts.push('这是会后纪要，会发到飞书聊天里，不受 200 字限制。小标题用加粗，不要用 # 标题。');
    return parts.join('\n');
  }
  if (input.lane === 'private') {
    parts.push('请基于上面的会议上下文作答。回答只发给提问的人，不发进会议，可以完整作答，不要用 markdown 标题。');
  } else if (input.longAnswerChars && input.longAnswerChars > 0) {
    parts.push(
      '请基于上面的会议上下文作答。回答会发到会议聊天里，大家都看得到。简单问题直接答，200 字内，不要用 markdown 标题。',
      `需要展开时（清单、步骤、长解释），先写不超过 150 字的要点，然后单独一行写「${FULL_TEXT_MARK}」，后面写完整回答：会议里只显示要点，全文私聊发给提问人。`,
    );
  } else {
    parts.push(
      '请基于上面的会议上下文作答。回答会被发到会议里，请简洁（尽量 200 字内），不要用 markdown 标题。',
    );
  }
  parts.push('引用会上说过的话时，在那句话后面标出处〔时:分 发言人〕，时间取会议记录方括号里的时和分，如〔14:32 张三〕。');
  if (input.materials?.length) {
    parts.push('用到会前资料时，在那句话后面照抄方括号里的出处，如〔《标题》r12 §3〕。');
  }
  parts.push(
    `问的是会上说过的事或资料里的内容时，只依据会议记录${input.materials?.length || input.materialList?.length ? '和会前资料' : ''}作答；里面没有的，直接说「会上${input.materials?.length || input.materialList?.length ? '和资料里' : ''}都没有提到」，不要编，也不要拿常识补。`,
  );
  return parts.join('\n');
}

const notifiedKinds = new WeakMap<MeetingSession, Set<string>>();

/**
 * Tell the bot owner about a meeting problem over DM. Once per meeting per
 * `kind`: these are "something is wrong" signals, not a stream.
 */
export function notifyMeetingOwner(
  deps: Pick<MeetingAgentDeps, 'session' | 'controls' | 'channel'>,
  kind: string,
  text: string,
): void {
  const { session, controls } = deps;
  const notified = notifiedKinds.get(session) ?? new Set<string>();
  notifiedKinds.set(session, notified);
  if (notified.has(kind)) return;
  notified.add(kind);
  const to = controls.botOwnerId;
  log.info('meeting', 'owner-notified', { meetingId: session.meetingId, kind, sent: Boolean(to) });
  if (!to) return;
  const title = session.topic ?? session.meetingNo;
  void deps.channel
    .send(to, { markdown: `⚠️ **会议智能体** · ${title}（${session.meetingNo}）\n${text}` })
    .catch((err) => log.warn('meeting', 'owner-notify-failed', { kind, err: String(err) }));
}

/**
 * Summarize a finished meeting and deliver it over IM.
 *
 * Runs after the meeting ended, so the in-meeting lane is gone — the summary
 * can only go to a chat. Target order: the chat the bot was told to join from,
 * else the bot owner's DM (`send` accepts an `ou_*` id as a direct message).
 * With neither, there is nowhere to put it, so we skip loudly instead of
 * silently dropping the work.
 *
 * Input is transcript + in-meeting chat: with captions off a meeting has no
 * transcript at all, and "no summary, no message" looked exactly like a broken
 * bot. When there is nothing, or the run / delivery fails, the owner is told.
 */
export async function summarizeEndedMeeting(deps: MeetingAgentDeps): Promise<void> {
  try {
    await summarize(deps);
  } finally {
    // The end of the meeting releases its doc bindings whether or not a
    // summary went out (summary off, nothing heard, run failed). Snapshots stay.
    try {
      releaseMeetingDocs(meetingDataDir(deps.controls), deps.session.meetingId);
    } catch (err) {
      log.warn('meeting', 'docs-release-failed', { err: String(err) });
    }
  }
}

async function summarize(deps: MeetingAgentDeps): Promise<void> {
  const { session, controls } = deps;
  if (!controls.profileConfig.meeting.summaryOnEnd) return;

  const transcriptLines = session.transcriptLineCount();
  const context = session.fullContext();
  if (context.length === 0) {
    log.info('meeting', 'summary-skipped', { meetingId: session.meetingId, reason: 'empty-transcript' });
    notifyMeetingOwner(
      deps,
      'summary-skipped',
      '会议已结束，但本场没收到任何字幕或弹幕，没有生成纪要。\n如果会上有人说话，多半是会议没开字幕/转写，bot 收不到语音内容。',
    );
    return;
  }
  const target = resolveSummaryTarget(
    controls.profileConfig.meeting.summaryTarget,
    session.originChatId,
    controls.botOwnerId,
  );
  if (!target) {
    log.warn('meeting', 'summary-no-target', { meetingId: session.meetingId });
    return;
  }

  // Never the in-meeting lane: the meeting is over.
  const answer = await summarizeMeetingRecord(deps, 'minutes').catch((err) => {
    log.warn('meeting', 'summary-run-failed', { meetingId: session.meetingId, err: String(err) });
    notifyMeetingOwner(deps, 'summary-run-failed', `会议已结束，但生成纪要失败：${String(err)}`);
    return '';
  });
  if (!answer) {
    notifyMeetingOwner(deps, 'summary-run-failed', '会议已结束，但生成纪要失败：agent 没有返回内容。');
    return;
  }

  const notes: string[] = [];
  if (transcriptLines === 0) notes.push('本场没收到字幕（可能没开转写），以上纪要只基于会中弹幕。');
  if (session.resumedAt) {
    const hhmm = new Date(session.resumedAt).toLocaleTimeString('zh-CN', {
      hour: '2-digit',
      minute: '2-digit',
      hour12: false,
    });
    // With the transcript on disk, a restart loses at most what was said
    // while the bridge was down — and the catch-up re-read usually has that.
    notes.push(
      session.status().transcriptFile
        ? `bridge 在 ${hhmm} 重启后接回本场会议，重启期间的内容已从会议事件补回，个别语句可能缺失。`
        : `bridge 在 ${hhmm} 重启后接回本场会议，此前的部分内容可能缺失。`,
    );
  }
  const backfilled = session.status().backfilled;
  if (backfilled > 0) notes.push(`会中实时漏收的 ${backfilled} 条记录，已靠重新拉取整场事件补回。`);
  const sent = await deliverMeetingMinutes(deps, target, answer, notes);
  if (!sent) return;
  log.info('meeting', 'summary-sent', {
    meetingId: session.meetingId,
    lines: context.length,
    transcriptLines,
    target: target.kind,
    fellBack: target.fellBack,
  });
}

/**
 * Archive the minutes and send them as a card: prose, structured todos with a
 * "派成工单" button each, notes and bound docs in the footer. When the card is
 * rejected, the old markdown goes out instead. The archive holds minutes and
 * todos only, never transcript or chat text; failing to write it never blocks
 * the minutes (the card then has no buttons). Returns whether anything was sent.
 */
async function deliverMeetingMinutes(
  deps: MeetingAgentDeps,
  target: ResolvedSummaryTarget,
  answer: string,
  notes: string[],
): Promise<boolean> {
  const { session, controls } = deps;
  const title = session.topic ?? session.meetingNo;
  const split = splitMinutesTodos(answer);
  let dir: string | undefined;
  try {
    dir = meetingDataDir(controls);
  } catch (err) {
    log.warn('meeting', 'data-dir-failed', { err: String(err) });
  }
  let docs: DocRef[] = [];
  try {
    if (dir) docs = meetingDocRefs(dir, session.meetingId);
  } catch (err) {
    log.warn('meeting', 'doc-context-failed', { meetingId: session.meetingId, err: String(err) });
  }
  const people = session.speakerDirectory();
  const rec: MinutesRecord = {
    id: meetingArchiveId(session.meetingNo),
    meetingId: session.meetingId,
    meetingNo: session.meetingNo,
    ...(session.topic ? { topic: session.topic } : {}),
    endedAt: new Date().toISOString(),
    ...(session.originChatId ? { originChatId: session.originChatId } : {}),
    minutes: split.minutes,
    todosParsed: split.parsed,
    // An owner named in the meeting becomes assignable when the name is one participant's.
    todos: split.todos.map((t) => {
      const id = t.owner ? people.get(t.owner) : undefined;
      return id ? { ...t, ownerId: id } : t;
    }),
    docs,
    ...(people.size ? { participants: [...people.keys()] } : {}),
    ...(session.status().transcriptText ? { transcript: session.status().transcriptText } : {}),
  };
  let archived = false;
  if (dir) {
    try {
      writeMeetingJson(meetingArchiveFile(dir, rec.id), rec);
      archived = true;
    } catch (err) {
      log.warn('meeting', 'archive-failed', { err: String(err) });
    }
  }
  const body =
    `**会议纪要 · ${title}**\n\n${answer}` + (notes.length ? `\n\n> ${notes.join('\n> ')}` : '');
  let via: 'card' | 'markdown' = 'card';
  const sent = await deps.channel
    .send(target.to, { card: meetingMinutesCard(rec, notes, archived) })
    .then(() => true)
    .catch(async (cardErr) => {
      log.warn('meeting', 'minutes-card-failed', { err: String(cardErr) });
      via = 'markdown';
      return deps.channel
        .send(target.to, { markdown: body })
        .then(() => true)
        .catch((err) => {
          log.warn('meeting', 'summary-send-failed', { err: String(err) });
          if (target.to !== controls.botOwnerId) {
            notifyMeetingOwner(deps, 'summary-send-failed', `纪要发到来源会话失败（${String(err)}），改发这里：\n\n${body}`);
          }
          return false;
        });
    });
  log.info('meeting', 'minutes-archived', {
    meetingId: session.meetingId,
    archive: archived ? rec.id : null,
    todos: rec.todos.length,
    parsed: split.parsed,
    via,
    sent,
  });
  return sent;
}

export interface ResolvedSummaryTarget {
  to: string;
  kind: MeetingSummaryTarget;
  /** True when the preferred target was unavailable and the other one was used. */
  fellBack: boolean;
}

/**
 * Pick where the summary goes: the configured preference first, the other one as
 * a fallback. Keeping the fallback means a summary is never thrown away just
 * because the preferred lane is missing — a console-initiated join has no origin
 * chat, and an owner that never resolved has no DM.
 */
export function resolveSummaryTarget(
  preference: MeetingSummaryTarget,
  originChatId: string | undefined,
  botOwnerId: string | undefined,
): ResolvedSummaryTarget | undefined {
  const order: MeetingSummaryTarget[] =
    preference === 'owner' ? ['owner', 'origin'] : ['origin', 'owner'];
  for (const [index, kind] of order.entries()) {
    const to = kind === 'origin' ? originChatId : botOwnerId;
    if (to) return { to, kind, fellBack: index > 0 };
  }
  return undefined;
}

/** Longest meeting record (chars) the minutes are written from in one pass. */
export const MINUTES_SINGLE_PASS_CHARS = 60_000;
/** Size of one segment when the record is longer: summarized first, then merged. */
export const MINUTES_SEGMENT_CHARS = 30_000;

/**
 * Split the meeting record into segments of at most `max` characters, on line
 * boundaries. A single line longer than that is cut. One segment when it all fits.
 */
export function splitMeetingRecord(lines: string[], max = MINUTES_SEGMENT_CHARS): string[][] {
  const out: string[][] = [];
  let cur: string[] = [];
  let size = 0;
  for (const raw of lines) {
    const line = raw.length > max ? `${raw.slice(0, max - 1)}…` : raw;
    if (cur.length && size + line.length + 1 > max) {
      out.push(cur);
      cur = [];
      size = 0;
    }
    cur.push(line);
    size += line.length + 1;
  }
  if (cur.length) out.push(cur);
  return out;
}

const recordChars = (lines: string[]): number => lines.reduce((n, l) => n + l.length + 1, 0);

/** `[HH:mm:ss]` of a record line, for naming a segment. */
const lineClock = (line: string | undefined): string => /^\[([0-9:-]+)\]/.exec(line ?? '')?.[1] ?? '?';

/** The ask for one segment of a long meeting. */
export function buildSegmentPrompt(input: {
  topic?: string;
  index: number;
  total: number;
  lines: string[];
}): string {
  const range = `${lineClock(input.lines[0])}–${lineClock(input.lines[input.lines.length - 1])}`;
  return [
    `你在帮忙整理一场飞书视频会议的纪要。会议很长，分成 ${input.total} 段处理，这是第 ${input.index} 段（${range}）。`,
    ...(input.topic ? [`会议主题：${input.topic}`] : []),
    '',
    `=== 第 ${input.index} 段会议记录（方括号里是时间，标「弹幕」的是文字消息） ===`,
    input.lines.join('\n'),
    '=== 记录结束 ===',
    '',
    '请把这一段整理成要点摘要，最后会和其他段合并成整场纪要：',
    '- 讨论了什么、谁的观点（保留关键数字、名字、时间点）',
    '- 做了哪些决定',
    `- 提到的待办：做什么、负责人、截止时间；负责人或截止时间没明确说出来的写「${TO_CONFIRM}」，不要猜`,
    '只依据这一段的记录，不要编。直接输出摘要，不要寒暄，不要写 json。',
  ].join('\n');
}

/**
 * Summarize the whole meeting record: in one pass when it fits
 * {@link MINUTES_SINGLE_PASS_CHARS}, otherwise segment by segment and then
 * merged. `minutes` is the end-of-meeting minutes (prose + todo json);
 * `notes` is `/meeting notes` (a short summary for whoever asked). Read-only,
 * own agent scope, never delivered anywhere — the caller does that. Throws when
 * a step fails or comes back empty: half a meeting must not pass as minutes.
 */
export async function summarizeMeetingRecord(
  deps: MeetingAgentDeps,
  kind: 'minutes' | 'notes',
): Promise<string> {
  const { session } = deps;
  const lines = session.fullContext();
  let context: { kind: 'full' | 'segments'; lines: string[] } = { kind: 'full', lines };
  if (recordChars(lines) > MINUTES_SINGLE_PASS_CHARS) {
    const segments = splitMeetingRecord(lines);
    log.info('meeting', 'minutes-segmented', { meetingId: session.meetingId, segments: segments.length, lines: lines.length });
    const parts: string[] = [];
    for (const [i, seg] of segments.entries()) {
      const prompt = buildSegmentPrompt({
        ...(session.topic ? { topic: session.topic } : {}),
        index: i + 1,
        total: segments.length,
        lines: seg,
      });
      const part = await runMeetingAgent(deps, prompt, { minutes: true });
      if (!part) throw new Error(`第 ${i + 1}/${segments.length} 段摘要没有返回内容`);
      parts.push(`【第 ${i + 1} 段 · ${lineClock(seg[0])}–${lineClock(seg[seg.length - 1])}】\n${part}`);
    }
    context = { kind: 'segments', lines: parts };
  }
  const question =
    kind === 'minutes'
      ? meetingMinutesQuestion(context.kind === 'segments', localDate(Date.parse(session.startedAt) || Date.now()))
      : `请基于以上${context.kind === 'segments' ? '分段摘要' : '会议记录'}做一份简洁纪要：讨论了什么、结论、待办（如有）。`;
  return answerInMeeting(deps, question, { deliver: 'caller', purpose: 'minutes', context });
}

interface RunOptions {
  /** Agent scope; in-meeting questions run per asker (see {@link meetingAskScopeId}). */
  scopeId?: string;
  /** Chat the run belongs to when there is no meeting session. */
  chatId?: string;
  /** A rejected run (scope busy, reconnecting) throws {@link RunRefused} instead of answering with the reason. */
  silentReject?: boolean;
  /** An in-meeting question: counted against {@link MEETING_MAX_PARALLEL} and stoppable while waiting. */
  track?: boolean;
  usedPrefix?: string;
  elevate?: boolean;
  tier?: AccessMode;
  /**
   * Minutes work: runs in the meeting's minutes scope (not blocked by an
   * in-meeting run that is still going) and a rejected run throws instead of
   * coming back as an answer.
   */
  minutes?: boolean;
}

/** Scope for the minutes runs of a meeting — separate from in-meeting answers. */
export function meetingMinutesScopeId(meetingId: string): string {
  return `meeting:${meetingId}:minutes`;
}

/** What a run needs; `session` is absent for runs outside a meeting (history questions). */
export type RunDeps = Pick<MeetingAgentDeps, 'controls' | 'executor' | 'sessions' | 'sessionCatalog' | 'workspaces' | 'channel'> & {
  session?: Pick<MeetingSession, 'meetingId' | 'meetingNo' | 'originChatId'>;
};

/** A run the bridge would not start (scope busy, runs paused, workspace missing…). */
export class RunRefused extends Error {
  constructor(
    readonly code: string,
    readonly userVisible: string,
  ) {
    super(userVisible);
    this.name = 'RunRefused';
  }
}

/**
 * One read-only run in its own scope, drained into the answer text; when the
 * bridge would not start it, `refused` says why. For work the meeting did not
 * ask for: the unasked interjection, history questions over the archive.
 */
export async function runReadOnlyPrompt(
  deps: RunDeps,
  prompt: string,
  scopeId: string,
  chatId?: string,
): Promise<{ text: string; refused?: { code: string; userVisible: string } }> {
  try {
    const text = await runMeetingAgent(deps, prompt, {
      scopeId,
      tier: meetingReadOnlyAccess(deps.controls.profileConfig),
      silentReject: true,
      ...(chatId ? { chatId } : {}),
    });
    return { text };
  } catch (err) {
    if (err instanceof RunRefused) return { text: '', refused: { code: err.code, userVisible: err.userVisible } };
    throw err;
  }
}

/** Submit one run and drain it into a single answer string. */
async function runMeetingAgent(deps: RunDeps, prompt: string, opts: RunOptions = {}): Promise<string> {
  const meetingId = deps.session?.meetingId;
  const scopeId = opts.scopeId;
  if (!opts.track || !meetingId || !scopeId) return runOnce(deps, prompt, opts);
  const runs = askRuns.get(meetingId) ?? new Map<string, { cancelled: boolean }>();
  // The same asker again: let the scope reservation refuse it ("你上一个问题还在回答").
  if (runs.has(scopeId)) return runOnce(deps, prompt, opts);
  if (runs.size >= MEETING_MAX_PARALLEL) {
    log.info('meeting', 'ask-over-cap', { meetingId, running: runs.size });
    return `现在已经有 ${runs.size} 个问题在回答，等其中一个答完再问我。`;
  }
  const state = { cancelled: false };
  runs.set(scopeId, state);
  askRuns.set(meetingId, runs);
  try {
    return await runOnce(deps, prompt, opts, state);
  } finally {
    runs.delete(scopeId);
    if (runs.size === 0 && askRuns.get(meetingId) === runs) askRuns.delete(meetingId);
  }
}

async function runOnce(
  deps: RunDeps,
  prompt: string,
  opts: RunOptions,
  state?: { cancelled: boolean },
): Promise<string> {
  const { controls } = deps;
  // Without a meeting (history questions) the caller always names the scope.
  const session = deps.session ?? { meetingId: '-', meetingNo: '-', originChatId: opts.chatId };
  const usedPrefix = opts.usedPrefix;
  const elevate = opts.elevate === true;
  const tier = opts.tier ?? meetingReadOnlyAccess(controls.profileConfig);
  const scopeId =
    opts.scopeId ?? (opts.minutes ? meetingMinutesScopeId(session.meetingId) : meetingScopeId(session.meetingId));
  const base =
    controls.profileConfig.agentKind === 'codex'
      ? codexCapability(controls.profileConfig)
      : claudeCapability(controls.profileConfig);
  // Lowering capability.permissions.maxAccess is enough: clampAccess takes
  // min(defaultAccess, profile max, capability max).
  const capability = { ...base, permissions: { ...base.permissions, maxAccess: tier } };
  if (elevate) {
    log.warn('meeting', 'elevated', { meetingId: session.meetingId, meetingNo: session.meetingNo, tier, stage: 'start' });
  }
  const result = await startRunFlow({
    scopeId,
    scope: {
      source: 'meeting',
      actorId: controls.botOwnerId ?? 'meeting',
      ...(session.originChatId ? { chatId: session.originChatId } : {}),
    },
    // Say so in read-only runs; otherwise the agent may fail silently or pretend it did the write.
    prompt:
      tier === 'read-only'
        ? `${prompt}\n\n[执行环境] 当前以只读权限运行，无法写文件或执行有副作用的命令。若请求需要写操作，直接说明「本次为只读模式，未执行」并给出建议，不要假装已完成。`
        : prompt,
    attachments: [],
    // Being in the meeting means "visible to every participant", not "anyone
    // may drive it": the asker was already authorized by meetingAuthzGate.
    access: { ok: true, reason: 'allowed-user' },
    capability,
    profileConfig: controls.profileConfig,
    sessions: deps.sessions,
    ...(deps.sessionCatalog ? { sessionCatalog: deps.sessionCatalog } : {}),
    workspaces: deps.workspaces,
    executor: deps.executor,
    now: Date.now(),
    observability: {
      profile: controls.profile,
      agent: capability.agentId,
      source: 'meeting',
      stage: 'submit',
    },
  });

  if (!result.ok) {
    log.info('meeting', 'run-rejected', {
      meetingId: session.meetingId,
      reason: result.rejectReason.code,
    });
    if (opts.minutes) throw new Error(result.rejectReason.userVisible);
    if (opts.silentReject) throw new RunRefused(result.rejectReason.code, result.rejectReason.userVisible);
    // The generic "already running" text is a dead end inside a meeting, where
    // /stop isn't reachable — point at the in-meeting escape hatch instead.
    if (result.rejectReason.code === 'run-already-active') {
      // Quote back whatever the asker typed: answering someone who used the
      // bot-name prefix with the configured one reads like a different bot.
      const prefix =
        usedPrefix ??
        triggerPrefixes(controls.profileConfig.meeting.trigger, deps.channel.botIdentity?.name)[0] ??
        controls.profileConfig.meeting.trigger;
      // Scopes are per asker: only their own earlier question can be in the way.
      return `你上一个问题还在回答，答完再问；发「${prefix} stop」可以中断它。`;
    }
    return result.rejectReason.userVisible;
  }

  // Stopped while it waited for a process slot: end it before it says anything.
  if (state?.cancelled) {
    await result.execution.stop().catch(() => {});
    log.info('meeting', 'ask-cancelled-before-start', { meetingId: session.meetingId });
    return '';
  }

  let answer = '';
  let failed = false;
  // Minutes must not be half an answer: an errored run throws (see below).
  try {
    for await (const event of result.execution.subscribe()) {
      const chunk = textOf(event);
      if (chunk) answer += chunk;
      if (event.type === 'error') {
        failed = true;
        log.warn('meeting', 'run-error', { meetingId: session.meetingId });
      }
    }
  } finally {
    // Every elevated run can be found by meetingId alone.
    if (elevate) {
      log.warn('meeting', 'elevated', {
        meetingId: session.meetingId,
        meetingNo: session.meetingNo,
        tier,
        stage: 'end',
        failed,
      });
    }
  }
  if (failed && opts.minutes) throw new Error('agent 运行出错，没有完整的结果');
  return answer.trim();
}

/** Collect assistant text from the agent event stream. */
function textOf(event: AgentEvent): string {
  return event.type === 'text' ? event.delta : '';
}

export type { MeetingEvent };
