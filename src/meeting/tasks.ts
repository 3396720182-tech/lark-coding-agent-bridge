import type { LarkChannel } from '@larksuite/channel';
import { log } from '../core/logger';
import { call, VcApiError, type VcRequestClient } from './api';
import {
  meetingArchiveFile,
  patchArchiveTodo,
  TO_CONFIRM,
  todoWhoWhen,
  type MeetingTodo,
  type MinutesRecord,
} from './minutes';
import { meetingDataDir, readMeetingJson } from './store';

/**
 * "建成飞书任务" on the minutes card: one todo → one Feishu task (task v2),
 * created with the bot's identity. The person the meeting named as owner is
 * the assignee when the name maps to exactly one participant; whoever clicked
 * follows the task so it shows up in their list. A deadline the meeting pinned
 * to a day becomes the due date; one it did not make clear stays off the task
 * and the description says 待确认.
 *
 * Needs the app scope `task:task:write` (or `task:task:writeonly`).
 */

export interface MeetingTaskContext {
  controls: { configPath: string; profile: string };
  channel: Pick<LarkChannel, 'rawClient'>;
  senderId: string;
}

export interface CreatedTask {
  guid: string;
  url?: string;
}

/** The body of `POST /open-apis/task/v2/tasks` for one todo. */
export function meetingTaskBody(
  rec: Pick<MinutesRecord, 'id' | 'topic' | 'meetingNo' | 'endedAt'>,
  todo: MeetingTodo,
  index: number,
  clickerId: string,
): Record<string, unknown> {
  const topic = rec.topic ?? rec.meetingNo;
  const members: { id: string; type: 'user'; role: 'assignee' | 'follower' }[] = [];
  if (todo.ownerId) members.push({ id: todo.ownerId, type: 'user', role: 'assignee' });
  if (clickerId && clickerId !== todo.ownerId) members.push({ id: clickerId, type: 'user', role: 'follower' });
  const description = [
    `来自会议「${topic}」（${rec.meetingNo}，${rec.endedAt.slice(0, 10)}）的纪要。`,
    todoWhoWhen(todo),
    ...(todo.owner && !todo.ownerId ? [`负责人「${todo.owner}」没能对上参会人，没有自动指派，请手动指派。`] : []),
    ...(!todo.owner ? [`负责人会上没说清（${TO_CONFIRM}），请确认后指派。`] : []),
    `纪要存档 ${rec.id} 第 ${index} 条。`,
  ].join('\n');
  return {
    summary: todo.task.slice(0, 3000),
    description,
    // Same todo, same task: a second click (or a retry after a timeout) does not make another.
    client_token: `meeting-${rec.id}-${index}`,
    ...(todo.dueDate
      ? // Local noon: the same calendar day whether Feishu reads the date part in UTC or in the tenant's zone.
        { due: { timestamp: String(new Date(`${todo.dueDate}T12:00:00`).getTime()), is_all_day: true } }
      : {}),
    ...(members.length ? { members } : {}),
  };
}

export async function createMeetingTask(client: VcRequestClient, body: Record<string, unknown>): Promise<CreatedTask> {
  const data = await call<{ task?: { guid?: string; url?: string } }>(client, {
    method: 'POST',
    url: '/open-apis/task/v2/tasks',
    params: { user_id_type: 'open_id' },
    data: body,
  });
  const guid = data.task?.guid;
  if (!guid) throw new Error('任务接口没有返回任务 ID');
  return { guid, ...(data.task?.url ? { url: data.task.url } : {}) };
}

/** A user-facing reason the task could not be created. */
export function taskErrorText(err: unknown): string {
  if (err instanceof VcApiError) {
    if (/task:task|scope|permission|权限/i.test(err.message) || [99991672, 99991679, 1470403].includes(err.code)) {
      return `应用还没有飞书任务的权限（错误码 ${err.code}）。请在开放平台给应用开通「task:task:write」并发布新版本后再点。`;
    }
    return `飞书任务接口报错（错误码 ${err.code}：${err.message}）`;
  }
  return String((err as Error)?.message ?? err);
}

const inflight = new Set<string>();

/**
 * `/meeting task <archive>#<n>`: create the Feishu task for one todo and
 * record it in the archive. A todo that already has a task is not created
 * again; a double click while the request runs is ignored.
 */
export async function handleMeetingTask(
  ctx: MeetingTaskContext,
  arg: string,
  reply: (markdown: string) => Promise<void>,
): Promise<void> {
  const m = /^(\d{8}-\d{4}-\d{9})#(\d{1,2})$/.exec(arg.trim());
  if (!m) {
    await reply('用法：在会议纪要卡片上点「建成飞书任务」，或 `/meeting task <存档号>#<第几条>`');
    return;
  }
  const archiveId = m[1]!;
  const index = Number(m[2]);
  const file = meetingArchiveFile(meetingDataDir(ctx.controls), archiveId);
  const rec = readMeetingJson<MinutesRecord | null>(file, null);
  const todo = rec?.todos?.[index - 1];
  if (!rec || !todo) {
    await reply(`找不到这条待办（存档 ${archiveId} 第 ${index} 条）。`);
    return;
  }
  if (todo.taskGuid) {
    await reply(`这条待办已经建过飞书任务${todo.taskUrl ? `：${todo.taskUrl}` : '。'}`);
    return;
  }
  const key = `${archiveId}#${index}`;
  if (inflight.has(key)) return;
  inflight.add(key);
  try {
    const client = ctx.channel.rawClient as unknown as VcRequestClient;
    const task = await createMeetingTask(client, meetingTaskBody(rec, todo, index, ctx.senderId));
    todo.taskGuid = task.guid;
    if (task.url) todo.taskUrl = task.url;
    todo.taskAt = new Date().toISOString();
    try {
      // Re-read and change only this todo's task fields: 派成工单 may have written meanwhile.
      patchArchiveTodo(file, index, {
        taskGuid: task.guid,
        ...(task.url ? { taskUrl: task.url } : {}),
        taskAt: todo.taskAt,
      });
    } catch (err) {
      // The task exists; the client_token stops a second click from making another.
      log.warn('meeting', 'task-record-failed', { archive: archiveId, err: String(err) });
    }
    log.info('meeting', 'task-created', { archive: archiveId, index, assigned: Boolean(todo.ownerId), due: Boolean(todo.dueDate) });
    const notes = [
      todo.ownerId ? `负责人：${todo.owner}` : todo.owner ? `负责人「${todo.owner}」没对上参会人，没有指派` : `负责人${TO_CONFIRM}，没有指派`,
      todo.dueDate ? `截止：${todo.dueDate}` : todo.due ? `截止「${todo.due}」没换算成日期，没有设` : `截止${TO_CONFIRM}，没有设`,
    ];
    await reply(
      `✅ 已建成飞书任务：${todo.task}${task.url ? `\n${task.url}` : ''}\n${notes.join(' · ')}。${ctx.senderId === todo.ownerId ? '你是负责人' : '你是关注人'}，可以在任务里改。`,
    );
  } catch (err) {
    log.warn('meeting', 'task-create-failed', { archive: archiveId, err: String(err) });
    await reply(`建飞书任务失败：${taskErrorText(err)}`);
  } finally {
    inflight.delete(key);
  }
}
