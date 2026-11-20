import { log } from '../core/logger';
import { ACCESS_ORDER, type AccessMode } from '../config/permissions';
import type { ProfileConfig } from '../config/profile-schema';
import { canRunAdminCommand, canUseDm, type RuntimeControls } from '../policy/access';
import type { MeetingSession } from './session';
import type { ChatEvent } from './types';

/**
 * Who may drive the agent from in-meeting chat.
 *
 * Joining a meeting makes the bot visible to every participant; it does not
 * make every participant allowed to run it. In-meeting chat is the one
 * untrusted entry point, so the decision is made here, at the entry, rather
 * than in shared downstream code.
 */

export const MEETING_DENY_REPLY = '该请求未执行：你尚未获得驱动本智能体的授权。如需开通请联系机器人管理员。';

export interface MeetingAuthzDeps {
  session: Pick<MeetingSession, 'meetingId' | 'sendMessage'> & { deniedCount?: number };
  controls: RuntimeControls & { profileConfig: ProfileConfig };
}

export type MeetingGate =
  | { ok: false }
  | { ok: true; question: string; elevate: boolean; senderId: string; reason: string };

/** Senders already told "not authorized", per meeting session. */
const deniedOnce = new WeakMap<object, Set<string>>();

/** Refusal: logged every time, the in-meeting reply goes once per sender per meeting. */
function deny(
  deps: MeetingAuthzDeps,
  reason: string,
  senderId: string | undefined,
  extra: Record<string, unknown> = {},
): void {
  const { session } = deps;
  session.deniedCount = (session.deniedCount ?? 0) + 1;
  log.warn('meeting', 'trigger-denied', {
    meetingId: session.meetingId,
    // Same redaction as command/admin-deny.
    sender: senderId ? String(senderId).slice(-6) : '?',
    reason,
    total: session.deniedCount,
    ...extra,
  });
  const told = deniedOnce.get(session) ?? new Set<string>();
  deniedOnce.set(session, told);
  const key = senderId ?? '?';
  if (told.has(key)) return;
  told.add(key);
  void session
    .sendMessage(MEETING_DENY_REPLY)
    .catch((err) => log.warn('meeting', 'deny-reply-failed', { err: String(err) }));
}

/**
 * Access tier for an in-meeting run: `executionAccess`, or the profile's
 * default when elevated — never above the profile's `maxAccess`
 * (`clampAccess` clamps once more downstream).
 */
export function meetingEffectiveAccess(profile: ProfileConfig, elevate: boolean): AccessMode {
  const wanted = elevate ? profile.permissions.defaultAccess : profile.meeting.executionAccess;
  const max = profile.permissions.maxAccess;
  return ACCESS_ORDER[wanted] <= ACCESS_ORDER[max] ? wanted : max;
}

/** True when the configured tier was lowered by the profile's maxAccess. */
export function meetingAccessConverged(profile: ProfileConfig): boolean {
  return ACCESS_ORDER[profile.meeting.executionAccess] > ACCESS_ORDER[profile.permissions.maxAccess];
}

/** One line for `/meeting status`. */
export function meetingAuthzSummary(profile: ProfileConfig): string {
  const m = profile.meeting;
  return (
    '验证：' +
    (m.access === 'admins-only' ? '仅管理员' : '同私聊白名单') +
    ' · 执行档位 ' +
    meetingEffectiveAccess(profile, false) +
    (meetingAccessConverged(profile) ? '（已被 profile maxAccess 收敛）' : '') +
    ' · 提权词 ' +
    (m.elevateWord || '已关闭')
  );
}

/**
 * Gate one in-meeting trigger. `question` is the text after the trigger.
 * On `{ ok: false }` the refusal has already been logged and answered.
 *
 * "May DM the bot" and "may drive it in a meeting" are the same question, so
 * this reuses canUseDm (owner, team mode, admins, allowedUsers).
 */
export function meetingAuthzGate(
  deps: MeetingAuthzDeps,
  chat: Pick<ChatEvent, 'from'>,
  question: string,
  /** Text with the elevation word removed when it directly follows the trigger, else undefined. */
  matchElevation: (word: string) => string | undefined,
): MeetingGate {
  const { controls } = deps;
  const profile = controls.profileConfig;
  const meeting = profile.meeting;
  const senderId = chat.from?.id;
  // Fail closed: no sender, no run. Log the keys we did get so "the event
  // shape changed" is distinguishable from "this person has no access".
  if (!senderId) {
    deny(deps, 'no-sender-id', undefined, { keys: Object.keys(chat.from ?? {}).join(',') });
    return { ok: false };
  }
  const verdict =
    meeting.access === 'admins-only'
      ? canRunAdminCommand(profile, controls, senderId)
      : canUseDm(profile, controls, senderId);
  if (!verdict.ok) {
    deny(deps, verdict.reason, senderId);
    return { ok: false };
  }
  // Elevation is a deterministic prefix on the text after the trigger, never
  // an inferred intent: letting the model decide its own permissions is circular.
  let text = question;
  let elevate = false;
  if (meeting.elevateWord) {
    const stripped = matchElevation(meeting.elevateWord);
    if (stripped !== undefined) {
      elevate = true;
      text = stripped;
    }
  }
  if (!text) {
    void deps.session
      .sendMessage(`用法：在触发词后接「${meeting.elevateWord} 你的请求」（提权执行）。`)
      .catch(() => {});
    return { ok: false };
  }
  return { ok: true, question: text, elevate, senderId, reason: verdict.reason };
}
