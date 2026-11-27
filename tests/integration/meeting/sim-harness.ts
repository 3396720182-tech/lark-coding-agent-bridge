import { mkdtemp, mkdir, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import type { CardActionEvent, NormalizedMessage } from '@larksuite/channel';
import type { AgentAdapter, AgentEvent, AgentRun, AgentRunOptions } from '../../../src/agent/types';
import { ActiveRuns } from '../../../src/bot/active-runs';
import type { ChatModeCache } from '../../../src/bot/chat-mode-cache';
import { PendingQueue } from '../../../src/bot/pending-queue';
import { ProcessPool } from '../../../src/bot/process-pool';
import { handleCardAction } from '../../../src/card/dispatcher';
import { tryHandleCommand, type CommandContext, type Controls } from '../../../src/commands/index';
import { createDefaultProfileConfig, type MeetingConfig, type ProfileConfig } from '../../../src/config/profile-schema';
import { resetCalendarNotice } from '../../../src/meeting/calendar';
import { MeetingManager, VC_BOT_EVENTS } from '../../../src/meeting/manager';
import { attachMeetingAgent, summarizeEndedMeeting } from '../../../src/meeting/orchestrator';
import { meetingDataDir } from '../../../src/meeting/store';
import { transcriptDir } from '../../../src/meeting/transcript-log';
import type { VcRequestClient } from '../../../src/meeting/api';
import { RunExecutor } from '../../../src/runtime/run-executor';
import { SessionStore } from '../../../src/session/store';
import { WorkspaceStore } from '../../../src/workspace/store';
import { createFakeChannel } from '../../helpers/fake-channel';

/**
 * A simulated meeting environment for the meeting agent: a fake Feishu (VC bot
 * API, calendar, docx, task v2, IM) behind the same request client the bridge
 * uses, events pushed through `onRawEvent` exactly as the long connection
 * delivers them, and the real wiring of `bot/channel.ts` on top — manager,
 * orchestrator, RunExecutor with a process pool, ActiveRuns, the command layer
 * and the card dispatcher. Only the agent is swappable: a scripted one for
 * deterministic checks, or the real Claude CLI.
 */

export const MEETING_ID = '7001000000000000001';
export const MEETING_NO = '123456789';
export const TOPIC = '发布评审周会';
export const BOT = { openId: 'ou_bot', name: '小会' };
export const OWNER = 'ou_owner';
export const PEOPLE = {
  zhang: { id: 'ou_zhang', name: '张三' },
  li: { id: 'ou_li', name: '李四' },
  wang: { id: 'ou_wang', name: '王五' },
  owner: { id: OWNER, name: '老板' },
} as const;
export const DOC = {
  token: 'DocTokenRelease001',
  url: 'https://example.feishu.cn/docx/DocTokenRelease001',
  title: '发布计划',
  revision: 7,
  text: [
    '一、发布窗口',
    '2.4 版本计划在 12 月 4 日（周五）晚上 8 点发布，灰度比例先开 10%。',
    '',
    '二、回滚条件',
    '崩溃率超过 0.3% 或者支付成功率下降 1 个百分点，立即回滚，回滚负责人是值班的后端同学。',
    '',
    '三、联系人',
    '测试负责人是赵敏，运维窗口由运维组老周确认。',
  ].join('\n'),
};
export const SECRET_DOC = 'https://example.feishu.cn/docx/DocTokenSecret999';

export interface FeishuCall {
  method: string;
  url: string;
  params?: Record<string, unknown>;
  data?: unknown;
}

export interface SentMessage {
  to: string;
  markdown?: string;
  card?: unknown;
}

/** What a run was asked and when it ran (for concurrency checks). */
export interface RunRecord {
  prompt: string;
  startedAt: number;
  endedAt?: number;
  stopped?: boolean;
  output?: string;
}

/** Put in a scripted reply to make the run call a tool at that point. */
export const TOOL_CALL = '\u0000TOOL\u0000';

/** Scripted agent: a responder per prompt, optional delay, honours stop(). */
export class ScriptedAgent implements AgentAdapter {
  readonly id = 'claude';
  readonly displayName = 'Scripted';
  readonly runs: RunRecord[] = [];
  constructor(
    private respond: (prompt: string) => string,
    private delayMs: (prompt: string) => number = () => 0,
  ) {}
  async isAvailable(): Promise<boolean> {
    return true;
  }
  setBotIdentity(): void {}
  run(opts: AgentRunOptions): AgentRun {
    const rec: RunRecord = { prompt: opts.prompt, startedAt: Date.now() };
    this.runs.push(rec);
    let stopped = false;
    let wake: (() => void) | undefined;
    const delay = this.delayMs(opts.prompt);
    const respond = this.respond;
    async function* events(): AsyncIterable<AgentEvent> {
      yield { type: 'system', sessionId: `sess-${opts.runId}` };
      if (delay > 0) {
        await new Promise<void>((r) => {
          wake = r;
          setTimeout(r, delay);
        });
      }
      if (stopped) {
        rec.stopped = true;
        rec.endedAt = Date.now();
        yield { type: 'done', terminationReason: 'interrupted' };
        return;
      }
      const text = respond(opts.prompt);
      rec.output = text;
      rec.endedAt = Date.now();
      // TOOL_CALL in the reply stands for a tool call between two pieces of text.
      const pieces = text.split(TOOL_CALL);
      for (const [i, piece] of pieces.entries()) {
        if (i > 0) yield { type: 'tool_use', id: `tool-${i}`, name: 'Read', input: {} };
        if (piece) yield { type: 'text', delta: piece };
      }
      yield { type: 'done', sessionId: `sess-${opts.runId}`, terminationReason: 'normal' };
    }
    return {
      runId: opts.runId,
      events: events(),
      async stop() {
        stopped = true;
        wake?.();
      },
      async waitForExit() {
        return true;
      },
    } as AgentRun;
  }
}

export interface SimOptions {
  agent: AgentAdapter;
  meeting?: Partial<MeetingConfig>;
  /** Make task creation fail as if the app lacked the scope. */
  taskScopeMissing?: boolean;
  /** Make the calendar unreadable (no scope). */
  calendarDenied?: boolean;
  processSlots?: number;
}

export interface Sim {
  root: string;
  feishu: FeishuCall[];
  meetingChat: string[];
  sent: SentMessage[];
  controls: Controls & { meeting: MeetingManager };
  manager: MeetingManager;
  profileConfig: ProfileConfig;
  /** Push one `vc.bot.meeting_activity_v1` item. */
  say(who: keyof typeof PEOPLE, text: string, at: Date): void;
  chat(who: keyof typeof PEOPLE, text: string, at?: Date): void;
  /** A DM / command from someone, through the bridge's command layer; true when it was handled there. */
  dm(from: string, text: string): Promise<boolean>;
  /** Join through `/meeting join` from the owner's DM. */
  join(): Promise<void>;
  /** Push `vc.bot.meeting_ended_v1` and wait for the minutes. */
  end(): Promise<void>;
  /** Click a button on a card sent to `to`. */
  click(value: Record<string, unknown>, operator?: string): Promise<void>;
  waitFor(cond: () => boolean, ms?: number, what?: string): Promise<void>;
  cleanup(): Promise<void>;
}

const OK = (data: unknown = {}) => ({ code: 0, msg: 'success', data });

export async function createSim(opts: SimOptions): Promise<Sim> {
  resetCalendarNotice();
  const root = await mkdtemp(join(tmpdir(), 'meeting-sim-'));
  const workspace = join(root, 'workspace');
  const profileDir = join(root, 'profile');
  await Promise.all([mkdir(workspace, { recursive: true }), mkdir(profileDir, { recursive: true })]);

  const feishu: FeishuCall[] = [];
  const meetingChat: string[] = [];
  const sent: SentMessage[] = [];
  let taskSeq = 0;
  /** Feishu sends the bot's own meeting messages back as chat events; set once the push lane exists. */
  let echoBot: (text: string) => void = () => {};
  const meetingStartSec = Math.floor(Date.now() / 1000);

  const request = async (payload: FeishuCall): Promise<unknown> => {
    feishu.push(payload);
    const { method, url } = payload;
    if (method === 'POST' && url === '/open-apis/vc/v1/bots/join') {
      return OK({ meeting: { id: MEETING_ID, topic: TOPIC } });
    }
    if (method === 'POST' && url === '/open-apis/vc/v1/bots/leave') return OK();
    if (method === 'GET' && url === '/open-apis/vc/v1/bots/events') return OK({ events: [], has_more: false });
    if (method === 'POST' && url === '/open-apis/vc/v1/bots/message') {
      const d = payload.data as { msg_type?: string; content?: string };
      if (d.msg_type === 'text') {
        meetingChat.push(String(d.content));
        const text = String(d.content);
        setTimeout(() => echoBot(text), 5);
      }
      return OK();
    }
    if (method === 'GET' && url === `/open-apis/vc/v1/meetings/${MEETING_ID}`) {
      if (opts.calendarDenied) return { code: 99991672, msg: 'Access denied. scope vc:meeting:readonly required' };
      return OK({ meeting: { topic: TOPIC, host_user: { id: PEOPLE.zhang.id }, start_time: String(meetingStartSec) } });
    }
    if (method === 'POST' && url === '/open-apis/calendar/v4/calendars/primarys') {
      return OK({
        calendars: [
          { calendar: { calendar_id: 'cal_zhang' }, user_id: PEOPLE.zhang.id },
          { calendar: { calendar_id: 'cal_owner' }, user_id: OWNER },
        ],
      });
    }
    if (method === 'GET' && url.startsWith('/open-apis/calendar/v4/calendars/cal_zhang/events/instance_view')) {
      return OK({
        items: [
          // Another team's meeting with the same name at the same time, on the host's calendar but a different number.
          {
            summary: '别的会',
            description: 'https://example.feishu.cn/docx/DocTokenOther000',
            vchat: { meeting_url: 'https://vc.feishu.cn/j/987654321' },
            start_time: { timestamp: String(meetingStartSec) },
            end_time: { timestamp: String(meetingStartSec + 3600) },
          },
          {
            summary: TOPIC,
            description: `会前请看：${DOC.url}\n内部材料：${SECRET_DOC}`,
            vchat: { meeting_url: `https://vc.feishu.cn/j/${MEETING_NO}` },
            start_time: { timestamp: String(meetingStartSec) },
            end_time: { timestamp: String(meetingStartSec + 3600) },
          },
        ],
      });
    }
    if (method === 'GET' && url.startsWith('/open-apis/calendar/v4/calendars/cal_owner/events/instance_view')) {
      return OK({ items: [] });
    }
    if (method === 'GET' && url === `/open-apis/docx/v1/documents/${DOC.token}`) {
      return OK({ document: { title: DOC.title, revision_id: DOC.revision } });
    }
    if (method === 'GET' && url === `/open-apis/docx/v1/documents/${DOC.token}/raw_content`) {
      return OK({ content: DOC.text });
    }
    if (method === 'GET' && url.startsWith('/open-apis/docx/v1/documents/DocTokenSecret999')) {
      return { code: 1770032, msg: 'forbidden' };
    }
    if (method === 'POST' && url === '/open-apis/task/v2/tasks') {
      if (opts.taskScopeMissing) {
        return { code: 99991672, msg: 'Access denied. One of the following scopes is required: [task:task:write]' };
      }
      taskSeq += 1;
      return OK({ task: { guid: `task-guid-${taskSeq}`, url: `https://applink.feishu.cn/client/todo/detail?guid=task-guid-${taskSeq}` } });
    }
    return { code: 404, msg: `fake feishu: no route for ${method} ${url}` };
  };

  const base = createFakeChannel();
  const handlers = new Map<string, Set<(p: unknown) => void | Promise<void>>>();
  let msgSeq = 0;
  const channel = {
    ...base,
    botIdentity: { openId: BOT.openId, name: BOT.name },
    rawClient: { ...base.rawClient, request: request as VcRequestClient['request'] },
    async send(to: string, content: { markdown?: string; card?: unknown; cardId?: string }) {
      sent.push({ to, ...(content.markdown !== undefined ? { markdown: content.markdown } : {}), ...(content.card ? { card: content.card } : {}) });
      return { messageId: `om_sim_${++msgSeq}` };
    },
    onRawEvent(type: string, handler: (p: unknown) => void | Promise<void>) {
      const set = handlers.get(type) ?? new Set();
      set.add(handler);
      handlers.set(type, set);
      return () => set.delete(handler);
    },
  };
  const emit = (type: string, payload: unknown) => {
    for (const h of handlers.get(type) ?? []) void h(payload);
  };
  echoBot = (text) =>
    emit(VC_BOT_EVENTS.activity, {
      meeting_activity_items: [
        {
          event_id: `evt-echo-${++eventSeq}`,
          meeting: { id: MEETING_ID },
          activity_event_type: 'chat_received',
          chat_received_items: [
            { content: text, message_id: `vc_bot_${eventSeq}`, message_type: 1, operator: { id: BOT.openId, user_name: BOT.name }, send_time: new Date().toISOString() },
          ],
        },
      ],
    });

  const profileConfig = createDefaultProfileConfig({
    agentKind: 'claude',
    accounts: { app: { id: 'cli_sim', secret: 'secret', tenant: 'feishu' } },
    access: { allowedUsers: [PEOPLE.zhang.id, PEOPLE.li.id] },
  });
  profileConfig.workspaces = { ...profileConfig.workspaces, default: workspace };
  profileConfig.meeting = {
    ...profileConfig.meeting,
    enabled: true,
    summaryOnEnd: true,
    pollIntervalMs: 60_000,
    ...opts.meeting,
  };

  const activeRuns = new ActiveRuns();
  let runSeq = 0;
  const pool = new ProcessPool(() => opts.processSlots ?? 4);
  const executor = new RunExecutor({
    agent: opts.agent,
    pool,
    activeRuns,
    createRunId: () => `run-${++runSeq}`,
  });
  const sessions = new SessionStore(join(profileDir, 'sessions.json'));
  const workspaces = new WorkspaceStore(join(profileDir, 'workspaces.json'));

  const controlsBase = {
    profile: 'claude',
    profileConfig,
    cfg: profileConfig,
    botOwnerId: OWNER,
    ownerRefreshState: 'ok',
    async refreshOwner() {},
    async restart() {},
    async exit() {},
    configPath: join(profileDir, 'config.json'),
    processId: 'proc-sim',
  } as unknown as Controls;
  const controls = controlsBase as Controls & { meeting: MeetingManager };

  let ended: Promise<void> | undefined;
  const runDeps = { channel: channel as never, controls: controls as never, executor, activeRuns, sessions, workspaces };
  const manager = new MeetingManager({
    client: channel.rawClient as unknown as VcRequestClient,
    config: () => controls.profileConfig.meeting,
    botOpenId: () => BOT.openId,
    channel,
    transcriptDir: transcriptDir(meetingDataDir(controls)),
    onEnded: (session) => {
      ended = summarizeEndedMeeting({ session, ...runDeps });
    },
    onSession: (session) => attachMeetingAgent({ session, ...runDeps }),
  });
  manager.attachPush();
  controls.meeting = manager;

  let eventSeq = 0;
  let sentenceSeq = 0;
  const item = (type: string, items: unknown[]) => ({
    meeting_activity_items: [
      { event_id: `evt-${++eventSeq}`, meeting: { id: MEETING_ID }, activity_event_type: type, [`${type}_items`]: items },
    ],
  });

  const dmCtx = (from: string, text: string, chatType = 'p2p'): CommandContext =>
    ({
      channel,
      msg: {
        messageId: `om_in_${++msgSeq}`,
        chatId: from === OWNER ? 'oc_owner_dm' : `oc_dm_${from}`,
        chatType,
        senderId: from,
        content: text,
        resources: [],
        mentions: [],
        mentionedBot: false,
      } as unknown as NormalizedMessage,
      scope: from === OWNER ? 'oc_owner_dm' : `oc_dm_${from}`,
      chatMode: 'p2p',
      controls,
      runExecutor: executor,
      sessions,
      workspaces,
      activeRuns,
    }) as unknown as CommandContext;

  const sim: Sim = {
    root,
    feishu,
    meetingChat,
    sent,
    controls,
    manager,
    profileConfig,
    say(who, text, at) {
      const p = PEOPLE[who];
      emit(
        VC_BOT_EVENTS.activity,
        item('transcript_received', [
          { sentence_id: String(++sentenceSeq), text, start_time_ms: String(at.getTime()), speaker: { open_id: p.id, name: p.name } },
        ]),
      );
    },
    chat(who, text, at = new Date()) {
      const p = PEOPLE[who];
      emit(
        VC_BOT_EVENTS.activity,
        item('chat_received', [
          {
            content: text,
            message_id: `vc_msg_${++eventSeq}`,
            message_type: 1,
            operator: { id: p.id, user_name: p.name },
            send_time: at.toISOString(),
          },
        ]),
      );
    },
    dm: (from, text) => tryHandleCommand(dmCtx(from, text)),
    async join() {
      await tryHandleCommand(dmCtx(OWNER, `/meeting join ${MEETING_NO}`));
    },
    async end() {
      emit(VC_BOT_EVENTS.ended, { meeting: { id: MEETING_ID } });
      await sim.waitFor(() => ended !== undefined, 2000, 'meeting end handled');
      await ended;
    },
    async click(value, operator = OWNER) {
      await handleCardAction({
        channel: channel as never,
        evt: {
          action: { value },
          operator: { openId: operator },
          chatId: 'oc_owner_dm',
          messageId: 'om_card',
        } as unknown as CardActionEvent,
        sessions,
        workspaces,
        activeRuns,
        agent: opts.agent,
        runExecutor: executor,
        controls,
        pending: new PendingQueue(60_000, () => {}),
        chatModeCache: { resolve: async () => 'p2p' } as unknown as ChatModeCache,
      });
    },
    async waitFor(cond, ms = 15_000, what = 'condition') {
      const until = Date.now() + ms;
      while (!cond()) {
        if (Date.now() > until) {
          throw new Error(`timed out waiting for ${what}\nmeeting chat:\n${meetingChat.join('\n--\n')}`);
        }
        await new Promise((r) => setTimeout(r, 20));
      }
    },
    async cleanup() {
      manager.dispose();
      await Promise.all([sessions.flush(), workspaces.flush()]);
      await rm(root, { recursive: true, force: true, maxRetries: 5, retryDelay: 25 });
    },
  };
  return sim;
}

/** Plain text of everything sent to `to` (markdown and card JSON). */
export function textTo(sim: Sim, to: string): string {
  return sim.sent
    .filter((m) => m.to === to)
    .map((m) => m.markdown ?? JSON.stringify(m.card))
    .join('\n---\n');
}

/** Every button value on cards sent so far. */
export function cardButtons(sim: Sim): { text: string; value: Record<string, unknown> }[] {
  const out: { text: string; value: Record<string, unknown> }[] = [];
  const walk = (n: unknown): void => {
    if (Array.isArray(n)) return n.forEach(walk);
    if (!n || typeof n !== 'object') return;
    const o = n as Record<string, unknown>;
    if (o.tag === 'button') {
      const text = String((o.text as { content?: string } | undefined)?.content ?? '');
      const behaviors = o.behaviors as { value?: Record<string, unknown> }[] | undefined;
      const value = (o.value as Record<string, unknown> | undefined) ?? behaviors?.find((b) => b.value)?.value;
      if (value) out.push({ text, value });
    }
    Object.values(o).forEach(walk);
  };
  for (const m of sim.sent) if (m.card) walk(m.card);
  return out;
}

/**
 * Clock times "HH:MM" at local time, for transcript timestamps. Placed on
 * tomorrow's date so they always come after everything stamped with the real
 * clock (typed chat, the bot's own messages): otherwise a run after 14:00 sees
 * the bot's earlier replies as said "after" a 14:43 question, and the
 * interjection check thinks someone answered it.
 */
export function at(hhmm: string): Date {
  const [h, m] = hhmm.split(':').map(Number);
  const d = new Date();
  d.setDate(d.getDate() + 1);
  d.setHours(h!, m!, 0, 0);
  return d;
}
