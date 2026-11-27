import { log } from '../core/logger';
import type { ProfileConfig } from '../config/profile-schema';
import type { RuntimeControls } from '../policy/access';
import { meetingVoiceAllowed } from './authz';
import type { MeetingSession } from './session';
import { sameInitial, sameSound } from './pinyin';
import type { TranscriptEvent } from './types';

/**
 * Waking the bot by voice: a transcript sentence that calls one of the wake
 * names ("小助手，帮我查一下…", "这个问题小助手你查一下") is a question, like an
 * `@bot` chat message. Speech recognition often writes the name with other
 * characters of the same sound; those count too, unless that spelling is the
 * name of someone in the meeting. Off by default, and only for people on the allowlist (see
 * {@link meetingVoiceAllowed}); everyone else is ignored without a word.
 *
 * ASR revises a sentence while it is being spoken, so a sentence is acted on
 * once it has stopped changing for {@link VOICE_SETTLE_MS}. "小助手" on its own
 * waits for the same speaker's next sentence ({@link VOICE_FOLLOW_MS}).
 */

export const VOICE_SETTLE_MS = 1500;
export const VOICE_FOLLOW_MS = 10_000;

/** Interjections and politeness that may come before the name. */
const LEAD = /^(?:(?:嗨|嘿|哈喽|喂|hi|hey|hello|ok|okay|那个|请问|麻烦|劳驾)[\s,，、。.!！~～]*)*/i;
const SEP = /^[\s,，、:：。.!！?？~～-]+/;
const SEP_CHAR = /[\s,，、:：。.!！?？~～-]/;
/** After a name: "小会你查一下" / "小会，您看下" — someone talking *to* it. */
const ADDRESSED_AFTER = /^(?:你|您)(?:帮|查|看|说|讲|总结|确认|找|列|给|算|搜|翻|整理)/;
/** Before a name in mid-sentence: "那个小会" / "开个小会" / "下午的小会" use it as a word, not a call. */
const AS_A_WORD_BEFORE = /(?:个|的|次|场|这|那|开|些)$/;

const squeeze = (s: string) => s.replace(/\s+/g, '').toLowerCase();

/**
 * Where `name` (squeezed) is heard starting at `from`: the index just past it,
 * skipping spaces, or undefined. A character matches the name's when it is the
 * same or sounds the same (speech recognition writes 「小会」 as 「小惠」).
 */
function nameEnd(t: string, from: number, name: string): { end: number; exact: boolean } | undefined {
  let i = from;
  let k = 0;
  let exact = true;
  while (i < t.length && k < name.length) {
    const c = t[i]!;
    if (/\s/.test(c)) {
      if (k === 0) return undefined;
      i += 1;
      continue;
    }
    const want = name[k]!;
    if (c.toLowerCase() !== want) {
      if (!sameSound(c, want)) return undefined;
      exact = false;
    }
    i += 1;
    k += 1;
  }
  if (k < name.length) return undefined;
  // "小助手们" / "botty": the name has to end there.
  const next = t[i] ?? '';
  if (/[A-Za-z0-9_]/.test(next) && /[a-z0-9_]$/.test(name)) return undefined;
  if (next === '们') return undefined;
  return { end: i, exact };
}

export interface WakeMatch {
  /** The question: the sentence without the name ('' when the name is all there is). */
  rest: string;
  /** The name as written in the transcript. */
  heard: string;
  /** Written exactly as configured (not a same-sound variant). */
  exact: boolean;
}

/**
 * Find a wake name in a sentence. At the start (after "嘿 / 请问 …"), or in
 * mid-sentence when it is plainly spoken to: followed by "你查 / 您看 …"
 * ("这个问题小会你查一下", "回滚那块，小会，你看下…"). The everyday word in
 * passing ("开个小会讨论一下", "那个小会你帮我约下会议室") is not a call. Names
 * match case-, space- and tone-insensitively; a same-sound spelling must
 * stand alone (a pause or "你…" after it).
 */
export function findWakeName(text: string, names: string[]): WakeMatch | undefined {
  const t = String(text ?? '').trim().replace(LEAD, '');
  const list = names.map(squeeze).filter((n) => n.length >= 2);
  for (const name of list) {
    const hit = nameEnd(t, 0, name);
    if (!hit) continue;
    const after = t.slice(hit.end).replace(/^\s+/, '');
    // A same-sound spelling straight into the next word is a word ("销毁数据",
    // "小汇报一下"): it has to stand alone — a pause, the end, or "你查…".
    if (!hit.exact && after && !SEP_CHAR.test(after[0]!) && !ADDRESSED_AFTER.test(after)) continue;
    return { rest: after.replace(SEP, '').trim(), heard: t.slice(0, hit.end), exact: hit.exact };
  }
  // In mid-sentence only when spoken to: "（，）小会（，）你查一下…".
  for (const name of list) {
    for (let i = 1; i < t.length; i++) {
      const hit = nameEnd(t, i, name);
      if (!hit) continue;
      if (AS_A_WORD_BEFORE.test(t.slice(0, i))) continue;
      const after = t.slice(hit.end).replace(SEP, '');
      if (!ADDRESSED_AFTER.test(after)) continue;
      return { rest: `${t.slice(0, i)}${after}`.trim(), heard: t.slice(i, hit.end), exact: hit.exact };
    }
  }
  return undefined;
}

/**
 * The text after a wake name ('' when the name is all there is), or undefined
 * when the sentence doesn't call the bot. See {@link findWakeName}.
 */
export function matchWakeName(text: string, names: string[]): string | undefined {
  return findWakeName(text, names)?.rest;
}

/**
 * A sentence that opens like a call to the bot but did not match a name: a
 * word of the same length before the first pause, at least half of it the
 * same sound and the rest starting with the same consonant ("小黑，回滚条件
 * 是什么？" for 「小会」; not "小李，" or "销量，"). Not a wake, but not a question for the room
 * either: speech recognition may have mangled the name.
 */
export function looksLikeWakeAttempt(text: string, names: string[]): boolean {
  const t = String(text ?? '').trim().replace(LEAD, '');
  const head = /^([^\s,，、:：。.!！?？~～-]{2,6})[\s]*[,，、:：]/.exec(t)?.[1];
  if (!head) return false;
  return names.map(squeeze).some((name) => {
    if (name.length < 2 || head.length !== name.length) return false;
    let same = 0;
    for (let k = 0; k < name.length; k++) {
      if (sameSound(head[k]!, name[k]!)) same += 1;
      else if (!sameInitial(head[k]!, name[k]!)) return false;
    }
    return same * 2 >= name.length;
  });
}

/** The names that wake the bot: the configured ones, else its display name. */
export function voiceWakeNames(profile: ProfileConfig, botName: string | undefined): string[] {
  const cfg = profile.meeting.voiceWake;
  const own = botName?.trim();
  return cfg.names.length ? cfg.names : own ? [own] : [];
}

/**
 * Is this sentence someone calling the bot by name, or something that sounds
 * like a mangled call (voice wake on)? Either way it is not a question for the
 * room: the voice path answers it, or nobody does.
 */
export function calledByName(profile: ProfileConfig, botName: string | undefined, text: string): boolean {
  if (!profile.meeting.voiceWake.enabled) return false;
  const names = voiceWakeNames(profile, botName);
  return matchWakeName(text, names) !== undefined || looksLikeWakeAttempt(text, names);
}

export interface VoiceWakeDeps {
  session: MeetingSession;
  controls: RuntimeControls & { profileConfig: ProfileConfig };
  /** Bot's current display name: the wake name when none are configured. */
  botName: () => string | undefined;
  now?: () => number;
}

export interface VoiceAsk {
  question: string;
  speakerId: string;
  speakerName?: string;
}

/**
 * Subscribe to the transcript and call `onAsk` for each spoken question from
 * an allowed speaker. `onAsk` does the stop words, the acknowledgement and the
 * answer. Returns a function that cancels pending timers.
 */
export function attachVoiceWake(deps: VoiceWakeDeps, onAsk: (ask: VoiceAsk) => void): () => void {
  const { session, controls } = deps;
  const now = deps.now ?? (() => Date.now());
  const settling = new Map<string, { timer: ReturnType<typeof setTimeout>; event: TranscriptEvent }>();
  /** Sentences already acted on (a late revision must not ask again). */
  const done = new Set<string>();
  /** Speakers who said only the name, waiting for their next sentence. */
  const waiting = new Map<string, { until: number; after: string }>();

  const names = (): string[] => voiceWakeNames(controls.profileConfig, deps.botName());
  /** The question after a wake name; a same-sound spelling that is a participant's name ("小慧") is that person, not the bot. */
  const wakeRest = (text: string): string | undefined => {
    const hit = findWakeName(text, names());
    if (!hit) return undefined;
    if (!hit.exact && [...session.speakerDirectory().keys()].some((n) => n.includes(hit.heard))) return undefined;
    return hit.rest;
  };

  const nameOnly = (text: string): boolean => {
    const rest = wakeRest(text);
    return rest !== undefined && rest.replace(/[\s,，、。.!！?？~～]/g, '').length < 2;
  };
  /** This speaker said only the name a moment ago and that sentence is still settling. */
  const nameStillSettling = (speakerId: string | undefined, sentenceId: string): boolean =>
    Boolean(speakerId) &&
    [...settling.values()].some(
      (p) => p.event.speaker.id === speakerId && p.event.sentenceId !== sentenceId && nameOnly(p.event.text),
    );
  const arm = (event: TranscriptEvent): void => {
    const prior = settling.get(event.sentenceId);
    if (prior) clearTimeout(prior.timer);
    const timer = setTimeout(() => fire(event), VOICE_SETTLE_MS);
    timer.unref?.();
    settling.set(event.sentenceId, { timer, event });
  };

  const fire = (event: TranscriptEvent): void => {
    settling.delete(event.sentenceId);
    if (session.ended || done.has(event.sentenceId)) return;
    const speakerId = event.speaker.id;
    if (!speakerId) return;
    const follow = waiting.get(speakerId);
    let question: string | undefined;
    if (follow && follow.after !== event.sentenceId && now() <= follow.until) {
      waiting.delete(speakerId);
      question = event.text.trim();
    } else {
      const rest = wakeRest(event.text);
      if (rest === undefined) {
        // The follow-up settled before the name-only sentence did: wait for it.
        if (nameStillSettling(speakerId, event.sentenceId)) arm(event);
        return;
      }
      done.add(event.sentenceId);
      if (!meetingVoiceAllowed({ session, controls }, speakerId)) return;
      if (rest.replace(/[\s,，、。.!！?？~～]/g, '').length < 2) {
        waiting.set(speakerId, { until: now() + VOICE_FOLLOW_MS, after: event.sentenceId });
        return;
      }
      question = rest;
    }
    done.add(event.sentenceId);
    if (!question) return;
    log.info('meeting', 'voice-wake', { meetingId: session.meetingId, sender: speakerId.slice(-6), chars: question.length });
    onAsk({ question, speakerId, ...(event.speaker.name ? { speakerName: event.speaker.name } : {}) });
  };

  const off = session.on('transcript', (e) => {
    if (!controls.profileConfig.meeting.voiceWake.enabled) return;
    const event = e as TranscriptEvent;
    if (event.selfEcho || done.has(event.sentenceId)) return;
    const speakerId = event.speaker.id;
    const follow = speakerId ? waiting.get(speakerId) : undefined;
    const candidate =
      wakeRest(event.text) !== undefined ||
      Boolean(follow && follow.after !== event.sentenceId && now() <= follow.until) ||
      // Said right after the name, before the name-only sentence settled.
      nameStillSettling(speakerId, event.sentenceId);
    if (!candidate) {
      // A revision that no longer starts with the name ("小助手" → "小猪手").
      const prior = settling.get(event.sentenceId);
      if (prior) clearTimeout(prior.timer);
      settling.delete(event.sentenceId);
      return;
    }
    arm(event);
  });

  return () => {
    off();
    for (const { timer } of settling.values()) clearTimeout(timer);
    settling.clear();
  };
}
