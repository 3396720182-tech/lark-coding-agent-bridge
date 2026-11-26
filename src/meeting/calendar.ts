import type { LarkChannel } from '@larksuite/channel';
import { log } from '../core/logger';
import { call, VcApiError, type VcRequestClient } from './api';
import {
  DOC_MAX_LINKS,
  docErrorText,
  docLabel,
  ingestDoc,
  meetingDocRefs,
  parseDocLinks,
  type DocLink,
  type DocRef,
} from './docs';
import type { MeetingSession } from './session';
import { meetingDataDir } from './store';

/**
 * Docs linked in the meeting's calendar event are bound on joining, so nobody
 * has to send them to the bot first.
 *
 * Finding the event, with the bot's identity only:
 *   1. the meeting's host (`GET vc/v1/meetings/:id`),
 *   2. the primary calendars of the host and the bot owner
 *      (`POST calendar/v4/calendars/primarys`),
 *   3. their events around the meeting's start (`instance_view`), matched by
 *      the meeting number in the event's video-meeting link (else by topic),
 *   4. Feishu doc links in the event description.
 * Each doc is then read like any pre-meeting doc (the bot must be able to
 * read it). An instant meeting has no event: nothing happens.
 *
 * Needs the app scopes `vc:meeting:readonly` and `calendar:calendar:readonly`
 * (or `calendar:calendar.event:read`); not verified against a live tenant yet.
 */

/** Events this far around the meeting's start are looked at. */
const WINDOW_SEC = 3 * 3600;

export interface CalendarEventDocs {
  summary?: string;
  links: DocLink[];
}

interface InstanceItem {
  summary?: string;
  description?: string;
  description_rich?: string;
  start_time?: { timestamp?: string };
  end_time?: { timestamp?: string };
  vchat?: { meeting_url?: string };
}

/**
 * The calendar event behind a meeting and the doc links in it; `undefined`
 * when no event matches (an instant meeting). Throws on API errors.
 */
export async function findCalendarEventDocs(
  client: VcRequestClient,
  meeting: { meetingId: string; meetingNo: string; topic?: string },
  extraUsers: string[],
  now = Date.now(),
): Promise<CalendarEventDocs | undefined> {
  const info = await call<{ meeting?: { host_user?: { id?: string }; start_time?: string; topic?: string } }>(client, {
    method: 'GET',
    url: `/open-apis/vc/v1/meetings/${meeting.meetingId}`,
    params: { user_id_type: 'open_id' },
  });
  const host = info.meeting?.host_user?.id;
  const users = [...new Set([host, ...extraUsers].filter((u): u is string => Boolean(u)))];
  if (users.length === 0) return undefined;
  const startSec = Number(info.meeting?.start_time) || Math.floor(now / 1000);
  const topic = meeting.topic ?? info.meeting?.topic;

  const prim = await call<{ calendars?: { calendar?: { calendar_id?: string } }[] }>(client, {
    method: 'POST',
    url: '/open-apis/calendar/v4/calendars/primarys',
    params: { user_id_type: 'open_id' },
    data: { user_ids: users },
  });
  const calendars = [
    ...new Set((prim.calendars ?? []).map((c) => c.calendar?.calendar_id).filter((c): c is string => Boolean(c))),
  ];
  let byTopic: InstanceItem | undefined;
  for (const calendarId of calendars) {
    const view = await call<{ items?: InstanceItem[] }>(client, {
      method: 'GET',
      url: `/open-apis/calendar/v4/calendars/${encodeURIComponent(calendarId)}/events/instance_view`,
      params: {
        start_time: String(startSec - WINDOW_SEC),
        end_time: String(startSec + WINDOW_SEC),
        user_id_type: 'open_id',
      },
    });
    for (const item of view.items ?? []) {
      if (item.vchat?.meeting_url?.replace(/\D+/g, ' ').split(' ').includes(meeting.meetingNo)) {
        return eventDocs(item);
      }
      if (!byTopic && topic && item.summary?.trim() === topic.trim() && covers(item, startSec)) byTopic = item;
    }
  }
  return byTopic ? eventDocs(byTopic) : undefined;
}

function covers(item: InstanceItem, sec: number): boolean {
  const s = Number(item.start_time?.timestamp);
  const e = Number(item.end_time?.timestamp);
  return Number.isFinite(s) && Number.isFinite(e) && s - 1800 <= sec && sec <= e;
}

function eventDocs(item: InstanceItem): CalendarEventDocs {
  const text = [item.description_rich, item.description].filter(Boolean).join('\n');
  return { ...(item.summary ? { summary: item.summary } : {}), links: parseDocLinks(text).slice(0, DOC_MAX_LINKS) };
}

let toldOwner = false;

interface CalendarDeps {
  session: MeetingSession;
  channel: LarkChannel;
  controls: { configPath: string; profile: string; botOwnerId?: string; profileConfig: { meeting: { calendarDocs: boolean } } };
}

/**
 * On a fresh join (not a resume), bind the docs of the meeting's calendar
 * event and tell the chat that started the join (else the owner) what was
 * taken and what could not be read. Never throws; failures to read the
 * calendar are logged, and the owner hears about them once per process.
 */
export function attachCalendarDocs(deps: CalendarDeps): Promise<void> | undefined {
  const { session, controls } = deps;
  if (!controls.profileConfig.meeting.calendarDocs || session.resumedAt) return undefined;
  const client = deps.channel.rawClient as unknown as VcRequestClient;
  return (async () => {
    let found: CalendarEventDocs | undefined;
    try {
      found = await findCalendarEventDocs(
        client,
        session,
        controls.botOwnerId ? [controls.botOwnerId] : [],
      );
    } catch (err) {
      const reason = err instanceof VcApiError ? `错误码 ${err.code}：${err.message}` : String(err);
      log.warn('meeting', 'calendar-docs-failed', { meetingId: session.meetingId, err: reason });
      if (!toldOwner && controls.botOwnerId) {
        toldOwner = true;
        void deps.channel
          .send(controls.botOwnerId, {
            markdown:
              `⚠️ **会议智能体** · 没能从会议日程里读会前资料（${reason}）。\n` +
              '需要给应用开通「vc:meeting:readonly」和「calendar:calendar:readonly」，并且机器人能读到主持人的日历。' +
              '不影响会中问答；也可以照旧私聊发「会前资料 <链接>」。本次运行只提醒这一次。',
          })
          .catch((e) => log.warn('meeting', 'owner-notify-failed', { kind: 'calendar-docs', err: String(e) }));
      }
      return;
    }
    if (!found || found.links.length === 0) {
      log.info('meeting', 'calendar-docs-none', { meetingId: session.meetingId, event: Boolean(found) });
      return;
    }
    const dir = meetingDataDir(controls);
    const already = new Set(meetingDocRefs(dir, session.meetingId).flatMap((r) => [r.url, r.token]));
    const bound: DocRef[] = [];
    const failed: string[] = [];
    for (const link of found.links) {
      if (already.has(link.url) || already.has(link.token)) continue;
      try {
        bound.push(await ingestDoc({ client, dir, link, target: session, by: 'calendar' }));
      } catch (err) {
        failed.push(`${link.url}\n   ${docErrorText(err)}`);
      }
    }
    log.info('meeting', 'calendar-docs', { meetingId: session.meetingId, bound: bound.length, failed: failed.length });
    const to = session.originChatId ?? controls.botOwnerId;
    if (!to || (bound.length === 0 && failed.length === 0)) return;
    const where = `会议「${session.topic ?? session.meetingNo}」（${session.meetingNo}）`;
    const lines = [
      ...(bound.length
        ? [`📎 已从日程${found.summary ? `「${found.summary}」` : ''}带上会前资料，进了${where}：${bound.map(docLabel).join('、')}`]
        : []),
      ...(failed.length ? [`日程里有 ${failed.length} 份文档读不到，没带上：`, ...failed.map((f) => `- ${f}`)] : []),
    ];
    await deps.channel
      .send(to, { markdown: lines.join('\n') })
      .then(() => undefined)
      .catch((err) => log.warn('meeting', 'calendar-docs-notice-failed', { err: String(err) }));
  })().catch((err) => log.warn('meeting', 'calendar-docs-failed', { meetingId: session.meetingId, err: String(err) }));
}

/** Test hook: forget that the owner was told. */
export function resetCalendarNotice(): void {
  toldOwner = false;
}
