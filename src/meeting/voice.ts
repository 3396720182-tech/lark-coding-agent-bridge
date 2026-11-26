import { log } from '../core/logger';
import type { ProfileConfig } from '../config/profile-schema';
import type { RuntimeControls } from '../policy/access';
import { meetingVoiceAllowed } from './authz';
import type { MeetingSession } from './session';
import type { TranscriptEvent } from './types';

/**
 * Waking the bot by voice: a transcript sentence that *starts* with one of the
 * wake names ("小助手，帮我查一下…") is a question, like an `@bot` chat
 * message. Off by default, and only for people on the allowlist (see
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

/**
 * The text after a wake name at the start of `text` ('' when the name is all
 * there is), or undefined when the sentence doesn't start with one. Names are
 * matched case- and space-insensitively.
 */
export function matchWakeName(text: string, names: string[]): string | undefined {
  const t = String(text ?? '').trim().replace(LEAD, '');
  const squeeze = (s: string) => s.replace(/\s+/g, '').toLowerCase();
  for (const raw of names) {
    const name = squeeze(raw);
    if (name.length < 2) continue;
    // Walk the text, skipping spaces, until the name is consumed.
    let i = 0;
    let k = 0;
    while (i < t.length && k < name.length) {
      const c = t[i]!;
      if (/\s/.test(c)) {
        i += 1;
        continue;
      }
      if (c.toLowerCase() !== name[k]) break;
      i += 1;
      k += 1;
    }
    if (k < name.length) continue;
    // "小助手们" / "botty": the name has to end there.
    const next = t[i] ?? '';
    if (/[A-Za-z0-9_]/.test(next) && /[a-z0-9_]$/.test(name)) continue;
    if (next === '们') continue;
    return t.slice(i).replace(SEP, '').trim();
  }
  return undefined;
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

  const names = (): string[] => {
    const cfg = controls.profileConfig.meeting.voiceWake;
    const own = deps.botName()?.trim();
    return cfg.names.length ? cfg.names : own ? [own] : [];
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
      const rest = matchWakeName(event.text, names());
      if (rest === undefined) return;
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
      matchWakeName(event.text, names()) !== undefined ||
      Boolean(follow && follow.after !== event.sentenceId && now() <= follow.until);
    const prior = settling.get(event.sentenceId);
    if (prior) clearTimeout(prior.timer);
    if (!candidate) {
      // A revision that no longer starts with the name ("小助手" → "小猪手").
      settling.delete(event.sentenceId);
      return;
    }
    const timer = setTimeout(() => fire(event), VOICE_SETTLE_MS);
    timer.unref?.();
    settling.set(event.sentenceId, { timer, event });
  });

  return () => {
    off();
    for (const { timer } of settling.values()) clearTimeout(timer);
    settling.clear();
  };
}
