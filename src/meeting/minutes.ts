import { execFile } from 'node:child_process';
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import type { LarkChannel } from '@larksuite/channel';
import { log } from '../core/logger';
import { meetingDataDir, readMeetingJson, writeMeetingJson } from './store';

/**
 * Structured meeting minutes: the end-of-meeting prompt, splitting the answer
 * into prose + todos, the archive record, the minutes card, and turning one
 * todo into a Larkyard ticket (`/meeting todo <archive>#<n>`, admin-only via
 * the `/meeting` gate).
 *
 * The archive keeps the minutes and the todos; the verbatim record is the
 * transcript file (`transcript-log.ts`), referenced by path.
 */

/** Longer minutes are cut in the card; the archive keeps the full text. */
export const MINUTES_MAX_CHARS = 8000;
export const MAX_TODOS = 20;

/** Shown for a todo owner or deadline the meeting did not make clear. */
export const TO_CONFIRM = '待确认';

/**
 * What counts as a todo, and who owns it. Without these the model turns
 * progress reports ("监控面板我把崩溃率放到首屏") and "大家待命" into todos, leaves
 * "日志采样率我先调到 5%" without an owner, and names departments as owners
 * nobody can be assigned to.
 */
export const MINUTES_TODO_RULES = [
  '只把会后还要去做、能落到人的事算作待办：已经做完的进度汇报（「我已经…」「我把…放到了…」）、一般性的提醒、「大家都待命」这类不算，写进讨论要点就行。',
  '说「我来」「我把」「我去」「我先」的人就是负责人，照写他的名字，不要写待确认。',
  `负责人是部门或团队（「市场部」「产品那边」）时照写部门名，后面加「（${TO_CONFIRM}到人）」，如「市场部（${TO_CONFIRM}到人）」。`,
  '同一件事只写一条待办，不要拆成好几条。',
];

/**
 * The end-of-meeting ask. `fromSegments`: the context is per-segment summaries
 * of a long meeting rather than the record itself.
 */
export function meetingMinutesQuestion(fromSegments = false, meetingDate?: string): string {
  return [
    fromSegments
      ? '会议已结束。请把以上分段摘要合并成整场会议纪要，分三部分：**讨论要点**、**结论**、**待办**。各段重复的内容合并，前后矛盾的以后面的为准。'
      : '会议已结束。请基于以上会议记录（字幕和弹幕）写一份会议纪要，分三部分：**讨论要点**、**结论**、**待办**。',
    `待办写清做什么、负责人、截止时间。负责人和截止时间只写会上明确说出来的；没说清的（没点名、只说「尽快」「回头」、只能靠猜的）一律写「${TO_CONFIRM}」，不要自己补。`,
    ...MINUTES_TODO_RULES,
    '纪要正文写完后，另起一行输出一个 ```json 代码块，只放待办，格式：',
    `{"todos":[{"task":"做什么","owner":"会上明确说的负责人，没说清就空字符串","due":"会上明确说的截止时间（原话），没说清就空字符串","due_date":"截止时间能确定到具体哪天时写 YYYY-MM-DD，否则空字符串"}]}`,
    ...(meetingDate ? [`会议日期是 ${meetingDate}，「周五」「下周一」这类说法按它换算成 due_date。`] : []),
    '没有待办就输出 {"todos":[]}。代码块之后不要再写任何内容。',
  ].join('\n');
}

/** Values that mean "not said" — the model sometimes writes them instead of leaving the field empty. */
const UNSAID = /^(未明确|待确认|待定|未定|未知|不明|不详|无|没有|暂无|未提及|未说明|n\/?a|none|null|tbd|unknown|-+|—+|\?+|？+)$/i;

/** A todo field as said in the meeting, or '' when it was not made clear. */
export function saidOrEmpty(v: string): string {
  const t = v.trim();
  return UNSAID.test(t) ? '' : t;
}

export const MEETING_MINUTES_QUESTION = meetingMinutesQuestion(false);

export interface MeetingTodo {
  task: string;
  /** As said in the meeting; '' = not made clear (shown as 待确认). */
  owner: string;
  due: string;
  /** `YYYY-MM-DD` when the deadline pins down a day. */
  dueDate?: string;
  /** Owner's open_id, when the name matches exactly one person heard in the meeting. */
  ownerId?: string;
  /** Ticket id once dispatched (`T0061`). */
  ticket?: string;
  ticketAt?: string;
  /** Feishu task once created. */
  taskGuid?: string;
  taskUrl?: string;
  taskAt?: string;
}

/** A pre-meeting document as listed on the minutes (see `docs.ts`). */
export interface MinutesDocRef {
  title: string;
  revision: number | string;
  fetchedAt: string;
}

export interface MinutesRecord {
  id: string;
  meetingId: string;
  meetingNo: string;
  topic?: string;
  endedAt: string;
  originChatId?: string;
  minutes: string;
  todosParsed: boolean;
  todos: MeetingTodo[];
  /** Documents bound to this meeting (empty when none). */
  docs?: MinutesDocRef[];
  /** The readable full transcript on this machine, when it was saved. */
  transcript?: string;
  /** Names of the people heard in the meeting (speech or chat). */
  participants?: string[];
}

/**
 * Split the minutes answer into prose and todos. The todos come from the
 * **last** ```json block (an earlier one may be quoted code). Without a usable
 * block the prose is kept verbatim and `parsed` is false.
 */
export function splitMinutesTodos(answer: string): {
  minutes: string;
  todos: MeetingTodo[];
  parsed: boolean;
} {
  const text = String(answer ?? '');
  let last: RegExpMatchArray | undefined;
  for (const m of text.matchAll(/```json\s*([\s\S]*?)```/gi)) last = m;
  const verbatim = { minutes: text.trim(), todos: [], parsed: false };
  if (!last || last.index === undefined) return verbatim;
  let arr: unknown;
  try {
    const j = JSON.parse(last[1] ?? '') as unknown;
    arr = Array.isArray(j) ? j : (j as { todos?: unknown } | null)?.todos;
  } catch {
    return verbatim;
  }
  if (!Array.isArray(arr)) return verbatim;
  const field = (t: unknown, k: string, max: number): string =>
    String((t as Record<string, unknown> | null)?.[k] ?? '')
      .trim()
      .slice(0, max);
  const todos = arr
    .map((t): MeetingTodo => {
      const dueDate = field(t, 'due_date', 10);
      return {
        task: field(t, 'task', 200),
        owner: saidOrEmpty(field(t, 'owner', 40)),
        due: saidOrEmpty(field(t, 'due', 40)),
        ...(isIsoDate(dueDate) ? { dueDate } : {}),
      };
    })
    .filter((t) => t.task)
    .slice(0, MAX_TODOS);
  const minutes = (text.slice(0, last.index) + text.slice(last.index + last[0].length)).trim();
  return { minutes, todos, parsed: true };
}

function isIsoDate(v: string): boolean {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(v)) return false;
  const t = new Date(`${v}T00:00:00`).getTime();
  // Rejects 2026-02-30 and the like: they roll over to another day.
  return !Number.isNaN(t) && localDate(t) === v;
}

/** `YYYY-MM-DD`, local time. */
export function localDate(now = Date.now()): string {
  const d = new Date(now);
  const p = (n: number) => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`;
}

/** Owner / deadline for display: as said, or 待确认. */
export function todoWhoWhen(t: Pick<MeetingTodo, 'owner' | 'due'>): string {
  return `负责人：${t.owner || TO_CONFIRM} · 截止：${t.due || TO_CONFIRM}`;
}

/** `YYYYMMDD-HHMM-<meeting no>`, local time. */
export function meetingArchiveId(meetingNo: string, now = Date.now()): string {
  const d = new Date(now);
  const p = (n: number) => String(n).padStart(2, '0');
  return `${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}-${p(d.getHours())}${p(d.getMinutes())}-${meetingNo}`;
}

export function meetingArchiveFile(dataDir: string, id: string): string {
  return join(dataDir, 'archive', `${id}.json`);
}

/** `MM-DD HH:mm`, local time; `?` for an unparsable value. */
export function shortTime(iso: string): string {
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return '?';
  const p = (n: number) => String(n).padStart(2, '0');
  return `${p(d.getMonth() + 1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}`;
}

/**
 * The minutes card: prose, the todo list (one "派成工单" button each when the
 * archive was written), and a footer with the notes, bound documents and the
 * archive id.
 */
export function meetingMinutesCard(rec: MinutesRecord, notes: string[], archived = true): object {
  const title = `会议纪要 · ${rec.topic ?? rec.meetingNo}`;
  const minutes =
    rec.minutes.length > MINUTES_MAX_CHARS
      ? rec.minutes.slice(0, MINUTES_MAX_CHARS) + '\n…（太长已截断，全文见存档）'
      : rec.minutes;
  const elements: object[] = [{ tag: 'markdown', content: minutes }, { tag: 'hr' }];
  if (rec.todos.length) {
    elements.push({
      tag: 'markdown',
      content:
        `**待办 ${rec.todos.length} 条**` +
        (archived ? ' · 点「派成工单」交给 Larkyard 去做，点「建成飞书任务」进飞书任务' : '') +
        (rec.todos.some((t) => !t.owner || !t.due) ? `\n<font color='orange'>「${TO_CONFIRM}」是会上没说清的，请人工确认</font>` : ''),
    });
    const said = (v: string) => (v ? `<font color='grey'>${v}</font>` : `<font color='orange'>${TO_CONFIRM}</font>`);
    rec.todos.forEach((t, i) => {
      elements.push({
        tag: 'markdown',
        content: `${i + 1}. ${t.task}\n<font color='grey'>负责人：</font>${said(t.owner)}<font color='grey'> · 截止：</font>${said(t.due)}`,
      });
      if (archived) {
        const arg = `${rec.id}#${i + 1}`;
        const button = (content: string, cmd: string) => ({
          tag: 'column',
          width: 'auto',
          elements: [
            {
              tag: 'button',
              text: { tag: 'plain_text', content },
              type: 'default',
              size: 'small',
              behaviors: [{ type: 'callback', value: { cmd, arg } }],
            },
          ],
        });
        elements.push({
          tag: 'column_set',
          horizontal_spacing: '8px',
          columns: [button('派成工单', 'meeting.todo'), button('建成飞书任务', 'meeting.task')],
        });
      }
    });
  } else {
    elements.push({
      tag: 'markdown',
      content: rec.todosParsed ? '没有识别出明确的待办。' : '待办没能结构化，以纪要正文为准。',
    });
  }
  const foot = [...notes];
  if (rec.docs?.length) {
    foot.push(
      '会前资料：' +
        rec.docs.map((r) => `《${r.title}》r${r.revision}（${shortTime(r.fetchedAt)} 抓取）`).join('、'),
    );
  }
  if (rec.transcript) {
    // Only the tail: the full path names the user's home directory.
    foot.push(`全量逐字稿已存本机：${rec.transcript.split(/[\\/]/).slice(-3).join('/')}`);
  }
  if (archived) foot.push(`存档 ${rec.id}`);
  if (foot.length) elements.push({ tag: 'markdown', content: foot.join('\n'), text_size: 'notation' });
  return {
    schema: '2.0',
    config: { summary: { content: title } },
    header: { title: { tag: 'plain_text', content: title }, template: 'green' },
    body: { elements },
  };
}

// ────────────── todo → Larkyard ticket ──────────────

export interface TicketNewInput {
  chatId: string;
  messageId: string;
  senderId: string;
  title: string;
  request: string;
  ack: string;
  note: string;
}

export interface TicketNewResult {
  code: number;
  stdout: string;
  stderr: string;
}

export type TicketNewRunner = (input: TicketNewInput) => Promise<TicketNewResult>;

/**
 * `ticket new` from Larkyard. The binary is `$LARKYARD_TICKET_BIN` (Larkyard
 * sets it on the bridge service it manages), else `ticket` on PATH. No shell;
 * the request goes through a temp file that is removed afterwards.
 */
export const runTicketNew: TicketNewRunner = async (input) => {
  const bin = process.env.LARKYARD_TICKET_BIN || 'ticket';
  const work = mkdtempSync(join(tmpdir(), 'meeting-todo-'));
  const requestFile = join(work, 'request.md');
  writeFileSync(requestFile, input.request);
  const args = [
    'new',
    '--chat', input.chatId,
    '--message', input.messageId,
    '--sender', input.senderId,
    '--title', input.title,
    '--slug', 'meeting-todo',
    '--request-file', requestFile,
    '--ack', input.ack,
    '--note', input.note,
  ];
  try {
    return await new Promise<TicketNewResult>((resolve) => {
      execFile(bin, args, { timeout: 60_000 }, (err, stdout, stderr) => {
        const code = err ? (typeof err.code === 'number' ? err.code : 1) : 0;
        const missing = (err as NodeJS.ErrnoException | null)?.code === 'ENOENT';
        resolve({
          code,
          stdout: String(stdout ?? ''),
          stderr: missing
            ? `找不到 ticket 命令（${bin}）。bridge 要由 Larkyard 管理，或设置 LARKYARD_TICKET_BIN。`
            : String(stderr ?? '') || (err && !stdout ? String(err.message) : ''),
        });
      });
    });
  } finally {
    rmSync(work, { recursive: true, force: true });
  }
};

/**
 * Set fields on one todo of an archive, reading the file fresh: the two card
 * buttons (ticket, Feishu task) run independently and each takes seconds, so
 * writing back an earlier copy would drop what the other one recorded.
 */
export function patchArchiveTodo(file: string, index: number, patch: Partial<MeetingTodo>): void {
  const rec = readMeetingJson<MinutesRecord | null>(file, null);
  const todo = rec?.todos?.[index - 1];
  if (!rec || !todo) throw new Error(`archive todo ${index} not found`);
  Object.assign(todo, patch);
  writeMeetingJson(file, rec);
}

const inflight = new Set<string>();

export interface MeetingTodoContext {
  controls: { configPath: string; profile: string };
  channel: Pick<LarkChannel, 'send'>;
  chatId: string;
  senderId: string;
}

/**
 * `/meeting todo <archive>#<n>`: post a note describing the todo in the
 * current chat, open a ticket on that note (ticketd wants one message per
 * ticket), record the ticket id in the archive and relay ticketd's receipt.
 * A todo that already has a ticket is not dispatched again; a double click
 * while `ticket new` runs is ignored. Failures record nothing, so a retry works.
 */
export async function handleMeetingTodo(
  ctx: MeetingTodoContext,
  arg: string,
  reply: (markdown: string) => Promise<void>,
  run: TicketNewRunner = runTicketNew,
): Promise<void> {
  const m = /^(\d{8}-\d{4}-\d{9})#(\d{1,2})$/.exec(arg.trim());
  if (!m) {
    await reply('用法：在会议纪要卡片上点「派成工单」，或 `/meeting todo <存档号>#<第几条>`');
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
  if (todo.ticket) {
    await reply(`这条待办已经派过单：${todo.ticket}`);
    return;
  }
  const key = `${archiveId}#${index}`;
  if (inflight.has(key)) return;
  inflight.add(key);
  try {
    const topic = rec.topic ?? rec.meetingNo;
    const text = [
      `📌 会议待办转工单（来自「${topic}」的纪要）`,
      todo.task,
      todoWhoWhen(todo),
    ].join('\n');
    const sent = await ctx.channel.send(ctx.chatId, { markdown: text });
    const out = await run({
      chatId: ctx.chatId,
      messageId: sent.messageId,
      senderId: ctx.senderId,
      title: todo.task.replace(/\s+/g, ' ').slice(0, 15),
      request: [text, '', '下面是这场会的纪要，供理解上下文：', '', rec.minutes].join('\n'),
      ack: `开工：会议待办「${todo.task.slice(0, 20)}」`,
      note: `会议纪要·${topic}`.slice(0, 30),
    });
    const ticketId = /\bT\d{4,}\b/.exec(out.stdout)?.[0];
    if (out.code === 0 && ticketId) {
      todo.ticket = ticketId;
      todo.ticketAt = new Date().toISOString();
      patchArchiveTodo(file, index, { ticket: ticketId, ticketAt: todo.ticketAt });
    }
    log.info('meeting', 'todo-dispatched', { archive: archiveId, index, ticket: ticketId ?? null, code: out.code });
    await reply(
      out.code === 0
        ? out.stdout.trim() || `已派单 ${ticketId ?? ''}`
        : `派单失败：${(out.stderr || out.stdout).trim().slice(0, 300)}`,
    );
  } catch (err) {
    log.warn('meeting', 'todo-dispatch-failed', { archive: archiveId, err: String(err) });
    await reply(`派单失败：${String((err as Error)?.message ?? err)}`);
  } finally {
    inflight.delete(key);
  }
}
