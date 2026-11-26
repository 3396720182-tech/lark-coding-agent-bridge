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
  splitMinutesTodos,
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
      const stopped = deps.activeRuns.interrupt(meetingScopeId(session.meetingId));
      void session
        .sendMessage(stopped ? '已中断当前任务。' : '当前没有正在执行的任务。')
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
  // Pre-meeting docs: take the ones waiting for this meeting, and bind docs
  // shared on screen in the meeting.
  adoptPendingDocsOnJoin(deps);
  attachShareDocs(deps);
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
    `· 有问题在聊天里发「${ask} 你的问题」，${where}。`,
    `· ${who}。`,
    '· 不希望被记录，可以请主持人把我移出会议。',
  ];
  return lines.join('\n');
}

export interface AnswerOptions {
  /** Name of the participant who asked (in-meeting triggers only). */
  askedBy?: string;
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
  const transcript = opts.context?.lines ?? session.recentContext(config.transcript.keep);
  const docs = meetingDocContext(deps, question, transcript, opts.purpose);
  const prompt = buildMeetingPrompt({
    question,
    transcript,
    ...(opts.context ? { contextKind: opts.context.kind } : {}),
    topic: session.topic,
    ...(opts.askedBy ? { askedBy: opts.askedBy } : {}),
    ...(opts.purpose ? { purpose: opts.purpose } : {}),
    ...docs.prompt,
  });

  // In-meeting chat runs at the meeting tier (or elevated); everything else
  // (`/meeting ask|notes`, the summary) is read-only: its input is meeting content.
  const tier =
    opts.deliver === 'broadcast'
      ? meetingEffectiveAccess(controls.profileConfig, opts.elevate === true)
      : meetingReadOnlyAccess(controls.profileConfig);
  const answer = await runMeetingAgent(deps, prompt, {
    ...(opts.usedPrefix ? { usedPrefix: opts.usedPrefix } : {}),
    elevate: opts.elevate === true,
    tier,
    ...(opts.purpose === 'minutes' ? { minutes: true } : {}),
  });
  if (!answer) return '';
  // Which of the provided chunks the answer actually cites — those are its
  // grounds; citations matching nothing provided are logged as unverified.
  const { cited, unverified } = docCitations(answer, docs.used);
  if (opts.purpose !== 'minutes') {
    recordMeetingQa(deps, {
      question,
      answer,
      ...(opts.askedBy ? { askedBy: opts.askedBy } : {}),
      provided: docs.used,
      cited,
      unverified,
    });
  }
  // The caller relays it themselves; never echo into the meeting.
  if (opts.deliver === 'caller') return answer;

  const target = config.respondIn;
  if (target === 'meeting' || target === 'both') {
    await session.sendMessage(answer).catch(async (err) => {
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
        : '=== 会议记录（按时间顺序，可能不完整） ===',
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
      ? `参会人「${input.askedBy}」在会中向你提问：`
      : '收到的请求：',
  );
  parts.push(input.question);
  parts.push('');
  if (input.purpose === 'minutes') {
    parts.push('这是会后纪要，会发到飞书聊天里，不受 200 字限制。小标题用加粗，不要用 # 标题。');
  } else {
    parts.push(
      '请基于上面的会议上下文作答。回答会被发到会议里，请简洁（尽量 200 字内），不要用 markdown 标题。',
    );
  }
  if (input.materials?.length) {
    parts.push(
      '用到会前资料时，在那句话后面照抄方括号里的出处，如〔《标题》r12 §3〕。资料和会议记录里都没有的，直接说不知道，不要编。',
    );
  }
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
  const rec: MinutesRecord = {
    id: meetingArchiveId(session.meetingNo),
    meetingId: session.meetingId,
    meetingNo: session.meetingNo,
    ...(session.topic ? { topic: session.topic } : {}),
    endedAt: new Date().toISOString(),
    ...(session.originChatId ? { originChatId: session.originChatId } : {}),
    minutes: split.minutes,
    todosParsed: split.parsed,
    todos: split.todos,
    docs,
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
    '- 提到的待办：做什么、负责人、截止时间；没提到的写「未明确」',
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
      ? meetingMinutesQuestion(context.kind === 'segments')
      : `请基于以上${context.kind === 'segments' ? '分段摘要' : '会议记录'}做一份简洁纪要：讨论了什么、结论、待办（如有）。`;
  return answerInMeeting(deps, question, { deliver: 'caller', purpose: 'minutes', context });
}

interface RunOptions {
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

/** Submit one run and drain it into a single answer string. */
async function runMeetingAgent(deps: MeetingAgentDeps, prompt: string, opts: RunOptions = {}): Promise<string> {
  const { session, controls } = deps;
  const usedPrefix = opts.usedPrefix;
  const elevate = opts.elevate === true;
  const tier = opts.tier ?? meetingReadOnlyAccess(controls.profileConfig);
  const scopeId = opts.minutes ? meetingMinutesScopeId(session.meetingId) : meetingScopeId(session.meetingId);
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
    // The generic "already running" text is a dead end inside a meeting, where
    // /stop isn't reachable — point at the in-meeting escape hatch instead.
    if (result.rejectReason.code === 'run-already-active') {
      // Quote back whatever the asker typed: answering someone who used the
      // bot-name prefix with the configured one reads like a different bot.
      const prefix =
        usedPrefix ??
        triggerPrefixes(controls.profileConfig.meeting.trigger, deps.channel.botIdentity?.name)[0] ??
        controls.profileConfig.meeting.trigger;
      return `上一个任务还在执行。发「${prefix} stop」可以中断它，然后再问我。`;
    }
    return result.rejectReason.userVisible;
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
