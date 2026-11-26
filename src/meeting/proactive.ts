import { log } from '../core/logger';
import type { ProfileConfig } from '../config/profile-schema';
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

/** A sentence that is plainly a question: a question mark, a question particle at the end, or a "请问" opening. */
export function looksLikeQuestion(text: string): boolean {
  const t = String(text ?? '').trim();
  if (t.length < 5 || t.length > 200) return false;
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
}): string {
  return [
    '你在旁听一场飞书视频会议。会上有人提了一个问题，过了一会儿没人回答。',
    ...(input.topic ? [`会议主题：${input.topic}`] : []),
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

export interface ProactiveDeps {
  session: MeetingSession;
  controls: { configPath: string; profile: string; profileConfig: ProfileConfig };
  /** Is this chat message addressed to the bot (`@bot …`)? Those are answered anyway. */
  addressedToBot: (text: string) => boolean;
  /** The bot's own open_id: its own messages are never questions. */
  selfId: () => string | undefined;
  /** One read-only run; '' when it could not run. */
  run: (prompt: string) => Promise<string>;
}

interface Candidate {
  question: string;
  askedBy?: string;
  askerKey: string;
  seq: number;
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
  const answeredSince = (seq: number, askerKey: string): boolean =>
    session
      .linesAfter(seq)
      .some((l) => (l.speakerId ?? l.speaker) !== askerKey && l.text.replace(/[\s,，、。.!！?？~～]/g, '').length >= 2);

  const check = async (key: string): Promise<void> => {
    const c = pending.get(key);
    pending.delete(key);
    if (!c || session.ended || busy || spoken >= cfg().maxPerMeeting || !cfg().enabled) return;
    if (answeredSince(c.seq, c.askerKey)) return;
    const docs = meetingDocContext(deps, c.question, []);
    if (docs.used.length === 0) return;
    busy = true;
    try {
      const recent = session.recentTimedContext(30);
      const raw = (await deps.run(buildProactivePrompt({
        question: c.question,
        ...(c.askedBy ? { askedBy: c.askedBy } : {}),
        ...(session.topic ? { topic: session.topic } : {}),
        recent,
        materials: docs.used,
      }))).trim();
      if (!raw || raw.length < MIN_ANSWER_LEN || raw.toUpperCase().includes(PROACTIVE_SKIP)) {
        log.info('meeting', 'proactive-skip', { meetingId: session.meetingId, reason: raw ? 'model' : 'no-run' });
        return;
      }
      const { cited, unverified } = docCitations(raw, docs.used);
      if (cited.length === 0) {
        log.info('meeting', 'proactive-skip', { meetingId: session.meetingId, reason: 'no-citation' });
        return;
      }
      // Somebody answered while the run was going: the moment has passed.
      if (session.ended || answeredSince(c.seq, c.askerKey) || spoken >= cfg().maxPerMeeting) return;
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
    const timer = setTimeout(() => void check(key), cfg().waitMs);
    timer.unref?.();
    pending.set(key, { question: question.trim(), askerKey, seq, timer, ...(askedBy ? { askedBy } : {}) });
  };

  const offSpeech = session.on('transcript', (e) => {
    const t = e as TranscriptEvent;
    if (t.selfEcho) return;
    consider(`s:${t.sentenceId}`, t.text, t.speaker.id ?? t.speaker.name ?? '?', session.seqOf(t.sentenceId), t.speaker.name);
  });
  const offChat = session.on('chat', (e) => {
    const c = e as ChatEvent;
    if (c.messageType === 3 || deps.addressedToBot(c.content)) return;
    const self = deps.selfId();
    if (self && c.from.id === self) return;
    consider(`c:${session.lastSeq()}`, c.content, c.from.id ?? c.from.name ?? '?', session.lastSeq(), c.from.name);
  });

  return () => {
    offSpeech();
    offChat();
    for (const c of pending.values()) clearTimeout(c.timer);
    pending.clear();
  };
}
