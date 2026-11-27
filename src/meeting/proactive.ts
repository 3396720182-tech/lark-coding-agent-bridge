import { log } from '../core/logger';
import type { ProfileConfig } from '../config/profile-schema';
import { canRunAdminCommand, canUseDm, type RuntimeControls } from '../policy/access';
import { docCitations, docLabel, meetingDocContext, recordMeetingQa, type UsedChunk } from './docs';
import type { MeetingSession } from './session';
import type { ChatEvent, TranscriptEvent } from './types';

/**
 * Speaking up unasked, Facilitator-style and deliberately rare. Off by
 * default. The bot says something only when all of these hold:
 *   - someone asked a question (speech or chat, not addressed to the bot),
 *   - for `waitMs` nobody else said anything (the asker going on doesn't count),
 *   - the meeting has docs bound and some of them match the question,
 *   - a read-only run, told to answer only from the docs, answered instead of
 *     passing (`SKIP`) and cited a chunk it was actually given,
 *   - still nobody else spoke while it ran,
 *   - fewer than `maxPerMeeting` interjections so far.
 */

export const PROACTIVE_SKIP = 'SKIP';
const MAX_ANSWER_CHARS = 300;
const MIN_ANSWER_LEN = 2;

/**
 * Questions to the room rather than about facts — "大家还有别的问题吗", "对吧",
 * "行不行", "能听到吗". Docs never answer them; checking costs a run.
 */
const SMALL_TALK = [
  /还有(?:没有)?(?:别的|其他|其它|什么)?(?:问题|补充|意见|要说的|想法|要讲的)/,
  /^大家(?:觉得|看|怎么看|有没有|还有|同意|没意见|都)/,
  /(?:怎么看|觉得呢|觉得怎么样|觉得如何|意见呢|有意见吗|同意吗|没问题吧)/,
  /(?:对吧|对不对|是吧|好吧|好不好|好吗|行不行|行吗|可以吗|可以不|成吗|ok吗|明白吗|清楚吗|懂了吗)[?？。!！\s]*$/i,
  // "能听到我说话吗" / "大家看得到屏幕吗"
  /^(?:大家|你们)?(?:能|可以)?(?:听|看)得?(?:到|见|清)了?(?:我|我们)?的?(?:说话|讲话|声音|屏幕|画面|共享)?(?:吗|么|没|不)[?？。!！\s]*$/,
];

/** A sentence that is plainly a question: a question mark, a question particle at the end, or a "请问" opening. Small talk is not. */
export function looksLikeQuestion(text: string): boolean {
  const t = String(text ?? '').trim();
  if (t.length < 5 || t.length > 200) return false;
  if (SMALL_TALK.some((re) => re.test(t))) return false;
  if (/[?？]\s*$/.test(t)) return true;
  if (/(吗|么|呢)[。.!！\s]*$/.test(t)) return true;
  return /^(请问|谁知道|有没有人知道|有人知道|我想问一下|问一下)/.test(t);
}

/** The ask for the interjection run. */
export function buildProactivePrompt(input: {
  question: string;
  askedBy?: string;
  topic?: string;
  recent: string[];
  materials: UsedChunk[];
  /** The bot's display name: its own messages in the record are under it. */
  botName?: string;
}): string {
  const self = input.botName?.trim();
  return [
    '你在旁听一场飞书视频会议。会上有人提了一个问题，过了一会儿没人回答。',
    ...(input.topic ? [`会议主题：${input.topic}`] : []),
    ...(self ? [`你在这场会里的名字是「${self}」：记录里「${self}」发的消息就是你之前说的。`] : []),
    '',
    '=== 最近的会议记录（方括号里是时间） ===',
    input.recent.length ? input.recent.join('\n') : '（暂无）',
    '=== 记录结束 ===',
    '',
    '=== 会前资料（方括号里是出处） ===',
    ...input.materials.map((m) => `〔${docLabel(m.doc)} §${m.i}〕\n${m.text}`),
    '=== 资料结束 ===',
    '',
    `${input.askedBy ? `「${input.askedBy}」` : '有人'}问：${input.question}`,
    '',
    `只有当上面的会前资料能直接、明确地回答这个问题时，才用一两句话（100 字内）回答，并在句后照抄出处，如〔《标题》r12 §3〕。`,
    `资料回答不了、只能部分回答、问题不是在问事实（比如在征求意见、问在场的人）、或者会上其实已经有人回答了，都只输出 ${PROACTIVE_SKIP}，不要输出别的。`,
    '不要寒暄，不要解释你为什么插话。',
  ].join('\n');
}

/**
 * Only questions from people who may ask the bot anyway (the in-meeting gate:
 * DM policy, or admins under `admins-only`) are considered. Anyone else's
 * words would otherwise start a run whose prompt they wrote. Silent.
 */
export function proactiveAskerAllowed(
  controls: RuntimeControls & { profileConfig: ProfileConfig },
  id: string | undefined,
): boolean {
  if (!id) return false;
  const p = controls.profileConfig;
  return (p.meeting.access === 'admins-only' ? canRunAdminCommand(p, controls, id) : canUseDm(p, controls, id)).ok;
}

/**
 * The model's own lead-in ("资料里写的是：", "根据会前资料，") would repeat the
 * line's "会前资料里有：".
 */
export function tidyInterjection(answer: string): string {
  return String(answer ?? '')
    .trim()
    .replace(/^(?:根据|按照|依据)?(?:会前)?资料(?:里|中|上)?(?:写的是|写着|写了|显示|提到|说|记录)?[：:，,\s]*/, '')
    .trim();
}

export interface ProactiveDeps {
  session: MeetingSession;
  controls: RuntimeControls & { configPath: string; profile: string; profileConfig: ProfileConfig };
  /** Is this chat message addressed to the bot (`@bot …`)? Those are answered anyway. */
  addressedToBot: (text: string) => boolean;
  /** Is this spoken sentence someone calling the bot by name? Answered by the voice path, never an interjection. */
  calledByName?: (text: string) => boolean;
  /** The bot's own open_id: its own messages are never questions. */
  selfId: () => string | undefined;
  /** The bot's display name, told to the model. */
  botName?: () => string | undefined;
  /** One read-only run; '' when it could not run. */
  run: (prompt: string) => Promise<string>;
}

interface Candidate {
  question: string;
  askedBy?: string;
  askerKey: string;
  seq: number;
  at?: number;
  timer: ReturnType<typeof setTimeout>;
}

/** Subscribe; returns a function that cancels pending checks. */
export function attachProactive(deps: ProactiveDeps): () => void {
  const { session, controls } = deps;
  const cfg = () => controls.profileConfig.meeting.proactive;
  const pending = new Map<string, Candidate>();
  let spoken = 0;
  let busy = false;

  /** Did anyone other than the asker say something after `seq`? */
  const answeredSince = (c: Pick<Candidate, 'seq' | 'at' | 'askerKey'>): boolean =>
    session
      .linesAfter(c.seq, c.at)
      .some((l) => (l.speakerId ?? l.speaker) !== c.askerKey && l.text.replace(/[\s,，、。.!！?？~～]/g, '').length >= 2);

  const check = async (key: string): Promise<void> => {
    const c = pending.get(key);
    pending.delete(key);
    if (!c || session.ended || busy || spoken >= cfg().maxPerMeeting || !cfg().enabled) return;
    if (answeredSince(c)) return;
    const docs = meetingDocContext(deps, c.question, []);
    if (docs.used.length === 0) return;
    busy = true;
    try {
      const recent = session.recentTimedContext(30);
      const botName = deps.botName?.();
      const raw = tidyInterjection(await deps.run(buildProactivePrompt({
        question: c.question,
        ...(c.askedBy ? { askedBy: c.askedBy } : {}),
        ...(session.topic ? { topic: session.topic } : {}),
        ...(botName ? { botName } : {}),
        recent,
        materials: docs.used,
      })));
      if (!raw || raw.length < MIN_ANSWER_LEN || /^[\s"'`*「『【(（]*SKIP\b/i.test(raw) || raw.toUpperCase() === PROACTIVE_SKIP) {
        log.info('meeting', 'proactive-skip', { meetingId: session.meetingId, reason: raw ? 'model' : 'no-run' });
        return;
      }
      const { cited, unverified } = docCitations(raw, docs.used);
      if (cited.length === 0) {
        log.info('meeting', 'proactive-skip', { meetingId: session.meetingId, reason: 'no-citation' });
        return;
      }
      // Somebody answered while the run was going: the moment has passed.
      if (session.ended || answeredSince(c) || spoken >= cfg().maxPerMeeting) return;
      const answer = raw.length > MAX_ANSWER_CHARS ? `${raw.slice(0, MAX_ANSWER_CHARS)}…` : raw;
      spoken += 1;
      await session.sendMessage(`💡 刚才${c.askedBy ? ` ${c.askedBy} ` : ''}问「${c.question.slice(0, 40)}」，会前资料里有：${answer}`);
      log.info('meeting', 'proactive-spoke', { meetingId: session.meetingId, count: spoken });
      recordMeetingQa(deps, {
        question: c.question,
        answer,
        ...(c.askedBy ? { askedBy: c.askedBy } : {}),
        provided: docs.used,
        cited,
        unverified,
      });
    } catch (err) {
      log.warn('meeting', 'proactive-failed', { meetingId: session.meetingId, err: String(err) });
    } finally {
      busy = false;
    }
  };

  /** (Re)start the wait for a question; a revision of the same sentence restarts it. */
  const consider = (key: string, question: string, askerKey: string, seq: number | undefined, askedBy?: string) => {
    const prior = pending.get(key);
    if (prior) clearTimeout(prior.timer);
    pending.delete(key);
    if (!cfg().enabled || spoken >= cfg().maxPerMeeting || seq === undefined || !looksLikeQuestion(question)) return;
    if (!proactiveAskerAllowed(controls, askerKey)) return;
    const at = session.atOf(seq);
    const timer = setTimeout(() => void check(key), cfg().waitMs);
    timer.unref?.();
    pending.set(key, {
      question: question.trim(),
      askerKey,
      seq,
      timer,
      ...(at !== undefined ? { at } : {}),
      ...(askedBy ? { askedBy } : {}),
    });
  };

  const offSpeech = session.on('transcript', (e) => {
    const t = e as TranscriptEvent;
    if (t.selfEcho) return;
    if (!t.speaker.id) return; // no id, no way to tell who asked
    // "小会，回滚负责人是谁？" is a question to the bot: the voice path answers it (or
    // refuses it); it must not also use up the meeting's interjection.
    if (deps.calledByName?.(t.text)) {
      const prior = pending.get(`s:${t.sentenceId}`);
      if (prior) clearTimeout(prior.timer);
      pending.delete(`s:${t.sentenceId}`);
      return;
    }
    consider(`s:${t.sentenceId}`, t.text, t.speaker.id, session.seqOf(t.sentenceId), t.speaker.name);
  });
  const offChat = session.on('chat', (e) => {
    const c = e as ChatEvent;
    if (c.messageType === 3 || deps.addressedToBot(c.content)) return;
    const self = deps.selfId();
    if (self && c.from.id === self) return;
    if (!c.from.id) return;
    consider(`c:${session.lastSeq()}`, c.content, c.from.id, session.lastSeq(), c.from.name);
  });

  return () => {
    offSpeech();
    offChat();
    for (const c of pending.values()) clearTimeout(c.timer);
    pending.clear();
  };
}
