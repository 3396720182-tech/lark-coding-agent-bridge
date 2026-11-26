import { mkdtempSync, readdirSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

const mocks = vi.hoisted(() => ({ startRunFlow: vi.fn() }));
vi.mock('../../../src/bot/run-flow', () => ({ startRunFlow: mocks.startRunFlow }));

const {
  answerInMeeting,
  attachMeetingAgent,
  buildMeetingPrompt,
  interruptMeetingRuns,
  meetingAskScopeId,
  meetingJoinNotice,
  splitAnswer,
  summarizeEndedMeeting,
  transcriptCitations,
  FULL_TEXT_MARK,
} = await import('../../../src/meeting/orchestrator');
const { MeetingSession } = await import('../../../src/meeting/session');
const { MEETING_DEFAULTS, createDefaultProfileConfig, normalizeProfileConfig } = await import(
  '../../../src/config/profile-schema'
);
const { meetingDataDir, writeMeetingJson } = await import('../../../src/meeting/store');
const { bindDoc, docSnapshotFile, meetingDocRefs } = await import('../../../src/meeting/docs');
const { meetingArchiveFile, meetingMinutesCard, splitMinutesTodos, TO_CONFIRM } = await import(
  '../../../src/meeting/minutes'
);
const { handleMeetingTask, meetingTaskBody } = await import('../../../src/meeting/tasks');
const { matchWakeName, VOICE_SETTLE_MS, VOICE_FOLLOW_MS } = await import('../../../src/meeting/voice');
const { looksLikeQuestion } = await import('../../../src/meeting/proactive');
const { answerMeetingHistory, historyWindow, looksLikeHistoryQuestion, pickArchives } = await import(
  '../../../src/meeting/history'
);
const { attachCalendarDocs, findCalendarEventDocs, resetCalendarNotice } = await import('../../../src/meeting/calendar');
const { parseMeetingBody } = await import('../../../src/ui/api');

import type { MeetingConfig } from '../../../src/config/profile-schema';
import type { RawActivityItem } from '../../../src/meeting/api';
import type { MinutesRecord } from '../../../src/meeting/minutes';

function fakeRun(text: string) {
  return {
    ok: true as const,
    execution: {
      subscribe: () => ({
        async *[Symbol.asyncIterator]() {
          yield { type: 'text', delta: text };
          yield { type: 'done' };
        },
      }),
    },
    policy: {},
    cwdRealpath: '/repo',
  };
}

const busy = {
  ok: false as const,
  rejectReason: { code: 'run-already-active', userVisible: '当前会话已有运行在执行' },
};

let root: string;
beforeEach(() => {
  vi.clearAllMocks();
  resetCalendarNotice();
  root = mkdtempSync(join(tmpdir(), 'meeting-b23-'));
});
afterEach(() => {
  vi.useRealTimers();
  rmSync(root, { recursive: true, force: true });
});

const T0 = new Date(2026, 10, 27, 14, 30, 0).getTime();

function cfg(over: Partial<MeetingConfig> = {}): MeetingConfig {
  return {
    ...MEETING_DEFAULTS,
    enabled: true,
    announceOnJoin: false,
    calendarDocs: false,
    ...over,
    transcript: { ...MEETING_DEFAULTS.transcript, ...over.transcript },
  };
}

function speech(eid: string, sid: number | string, text: string, who = '张三', id = 'ou_zhang', at = T0): RawActivityItem {
  return {
    event_id: eid,
    activity_event_type: 'transcript_received',
    transcript_received_items: [{ sentence_id: sid, text, speaker: { open_id: id, name: who }, start_time_ms: at }],
  } as never;
}

function chat(eid: string, text: string, who = '李四', id = 'ou_li', at = T0): RawActivityItem {
  return {
    event_id: eid,
    activity_event_type: 'chat_received',
    chat_received_items: [{ content: text, message_type: 1, operator: { id, user_name: who }, time: at }],
  } as never;
}

function setup(
  over: Partial<MeetingConfig> = {},
  opts: { failDm?: boolean; mode?: 'personal' | 'team'; allowed?: string[]; admins?: string[]; client?: unknown } = {},
) {
  const meeting = cfg(over);
  const profileConfig = createDefaultProfileConfig({
    agentKind: 'claude',
    accounts: { app: { id: 'cli_test', secret: '${APP_SECRET}', tenant: 'feishu' } },
  });
  profileConfig.meeting = meeting;
  if (opts.mode) profileConfig.mode = opts.mode;
  profileConfig.access.allowedUsers = opts.allowed ?? ['ou_zhang', 'ou_li'];
  profileConfig.access.admins = opts.admins ?? [];
  let now = T0;
  const session = new MeetingSession({
    client: { request: vi.fn(async () => ({ code: 0, data: {} }) as never) },
    meetingId: 'm1',
    meetingNo: '123456789',
    topic: '周会',
    config: meeting,
    originChatId: 'oc_team',
    now: () => now,
  });
  const meetingMsgs: string[] = [];
  vi.spyOn(session, 'sendMessage').mockImplementation(async (t: string) => {
    meetingMsgs.push(t);
  });
  const sent: { to: string; input: { markdown?: string; card?: object } }[] = [];
  const scopes = new Set<string>();
  const activeRuns = {
    interrupt: vi.fn((scope: string) => scopes.delete(scope)),
    scopes: () => [...scopes],
  };
  const args = {
    session,
    channel: {
      botIdentity: { name: '三儿', openId: 'ou_bot' },
      rawClient: opts.client,
      send: vi.fn(async (to: string, input: { markdown?: string; card?: object }) => {
        if (opts.failDm && to.startsWith('ou_') && to !== 'ou_owner') throw new Error('no dm');
        sent.push({ to, input });
        return { messageId: 'om_1' };
      }),
    },
    controls: { configPath: join(root, 'config.json'), profile: 'claude', profileConfig, botOwnerId: 'ou_owner' },
    executor: {},
    activeRuns,
    sessions: {},
    workspaces: {},
  } as never;
  return {
    session,
    args,
    sent,
    meetingMsgs,
    scopes,
    profileConfig,
    tick: (ms: number) => {
      now += ms;
    },
  };
}

const flush = () => new Promise((r) => setTimeout(r, 0));

// ───────────────────────── batch 2 ─────────────────────────

describe('several people asking at once', () => {
  it('each asker runs in their own scope, so two questions do not queue', async () => {
    let release!: () => void;
    const gate = new Promise<void>((r) => (release = r));
    mocks.startRunFlow.mockImplementation(async () => {
      await gate;
      return fakeRun('好的');
    });
    const { session, args, meetingMsgs } = setup();
    attachMeetingAgent(args);
    session.ingest(chat('c1', '@三儿 第一个问题', '张三', 'ou_zhang'));
    session.ingest(chat('c2', '@三儿 第二个问题', '李四', 'ou_li'));
    await flush();
    // Both started before either finished.
    expect(mocks.startRunFlow).toHaveBeenCalledTimes(2);
    const scopeIds = mocks.startRunFlow.mock.calls.map((c) => c[0].scopeId);
    expect(scopeIds).toEqual([meetingAskScopeId('m1', 'ou_zhang'), meetingAskScopeId('m1', 'ou_li')]);
    release();
    await flush();
    await flush();
    expect(meetingMsgs).toEqual(['好的', '好的']);
  });

  it('the same person asking again while their answer runs is told to wait (not "someone else")', async () => {
    mocks.startRunFlow.mockResolvedValue(busy);
    const { session, args, meetingMsgs } = setup();
    attachMeetingAgent(args);
    session.ingest(chat('c1', '@三儿 再问一次', '张三', 'ou_zhang'));
    await flush();
    await flush();
    expect(meetingMsgs[0]).toContain('你上一个问题还在回答');
    expect(meetingMsgs[0]).toContain('@三儿 stop');
  });

  it('stop in the meeting stops only the asker\'s own question; /meeting stop stops all but the minutes', () => {
    const { args, scopes } = setup();
    const a = meetingAskScopeId('m1', 'ou_zhang');
    const b = meetingAskScopeId('m1', 'ou_li');
    for (const s of [a, b, 'meeting:m1:minutes', 'meeting:m2:ask:ou_x', 'meeting:m1']) scopes.add(s);
    const runs = (args as { activeRuns: Parameters<typeof interruptMeetingRuns>[0] }).activeRuns;
    expect(interruptMeetingRuns(runs, 'm1', 'ou_zhang')).toBe(1);
    expect([...scopes]).not.toContain(a);
    expect(interruptMeetingRuns(runs, 'm1')).toBe(2);
    expect([...scopes].sort()).toEqual(['meeting:m1:minutes', 'meeting:m2:ask:ou_x']);
  });
});

describe('long answers: points in the meeting, the full text by DM', () => {
  it('splits on the marker and keeps the full text for the DM', () => {
    const full = '完整回答。'.repeat(80);
    const out = splitAnswer(`要点一；要点二\n${FULL_TEXT_MARK}\n${full}`, 300);
    expect(out.points).toBe('要点一；要点二');
    expect(out.full).toBe(full);
  });

  it('a long answer without the marker is cut at a sentence end; short ones and max 0 are not split', () => {
    const long = '第一句话说明情况。'.repeat(60);
    const out = splitAnswer(long, 300);
    expect(out.full).toBe(long);
    expect(out.points!.endsWith('。…')).toBe(true);
    expect(out.points!.length).toBeLessThanOrEqual(201);
    expect(splitAnswer('短回答', 300)).toEqual({ full: '短回答' });
    expect(splitAnswer(long, 0)).toEqual({ full: long });
    // A marker on a short answer is dropped, the full part kept.
    expect(splitAnswer(`要点\n${FULL_TEXT_MARK}\n全文`, 300)).toEqual({ full: '全文' });
  });

  it('DMs the asker the full answer as a card and posts only the points in the meeting', async () => {
    const full = '这是完整的回答。'.repeat(60);
    mocks.startRunFlow.mockResolvedValue(fakeRun(`三点要点\n${FULL_TEXT_MARK}\n${full}`));
    const { args, sent, meetingMsgs } = setup();
    await answerInMeeting(args, '怎么发布', { deliver: 'broadcast', askedBy: '张三', askerId: 'ou_zhang' });
    expect(sent).toHaveLength(1);
    expect(sent[0]!.to).toBe('ou_zhang');
    expect(JSON.stringify(sent[0]!.input.card)).toContain('这是完整的回答');
    expect(meetingMsgs).toEqual(['三点要点\n（完整回答已私聊发给 张三）']);
    expect(String(mocks.startRunFlow.mock.calls[0]![0].prompt)).toContain(FULL_TEXT_MARK);
  });

  it('when the DM fails the meeting gets the whole answer instead', async () => {
    const full = '这是完整的回答。'.repeat(60);
    mocks.startRunFlow.mockResolvedValue(fakeRun(`要点\n${FULL_TEXT_MARK}\n${full}`));
    const { args, meetingMsgs } = setup({}, { failDm: true });
    await answerInMeeting(args, '怎么发布', { deliver: 'broadcast', askedBy: '张三', askerId: 'ou_zhang' });
    expect(meetingMsgs).toEqual([full]);
  });

  it('/meeting ask (caller) and respondIn=im are never split, and the prompt says so', async () => {
    const long = '长'.repeat(500);
    mocks.startRunFlow.mockResolvedValue(fakeRun(long));
    const { args, sent, meetingMsgs } = setup();
    expect(await answerInMeeting(args, 'q', { deliver: 'caller', askerId: 'ou_zhang' })).toBe(long);
    expect(String(mocks.startRunFlow.mock.calls[0]![0].prompt)).toContain('私聊');
    expect(String(mocks.startRunFlow.mock.calls[0]![0].prompt)).not.toContain(FULL_TEXT_MARK);
    expect(sent).toHaveLength(0);
    expect(meetingMsgs).toHaveLength(0);
  });
});

describe('citing what was said, and saying so when nothing was', () => {
  it('in-meeting answers get time-labelled lines and the citation / no-invention rules', async () => {
    mocks.startRunFlow.mockResolvedValue(fakeRun('周五发布〔14:30 张三〕'));
    const { session, args } = setup();
    session.ingest(speech('e1', 1, '我们周五发布'));
    await answerInMeeting(args, '什么时候发布', { deliver: 'broadcast', askerId: 'ou_li' });
    const prompt = String(mocks.startRunFlow.mock.calls[0]![0].prompt);
    expect(prompt).toContain('[14:30:00] 张三: 我们周五发布');
    expect(prompt).toContain('〔14:32 张三〕');
    expect(prompt).toContain('都没有提到');
    expect(prompt).toContain('不要编');
  });

  it('a citation matching a line given is kept; a made-up one is marked', () => {
    const lines = ['[14:30:05] 张三: 我们周五发布', '[14:31:10] 李四（弹幕）: 好'];
    const out = transcriptCitations('周五发布〔14:30 张三〕，李四同意〔14:31 李四〕，王五反对〔15:02 王五〕', lines);
    expect(out.cited).toEqual(['14:30 张三', '14:31 李四']);
    expect(out.unverified).toEqual(['15:02 王五']);
    expect(out.text).toContain('〔14:30 张三〕');
    expect(out.text).toContain('〔15:02 王五（未核对到原话）〕');
    // Doc citations are not transcript citations.
    expect(transcriptCitations('见〔《手册》r3 §2〕', lines).unverified).toEqual([]);
  });

  it('the answer posted carries the mark, and the Q&A log records both kinds', async () => {
    mocks.startRunFlow.mockResolvedValue(fakeRun('王五说过〔09:00 王五〕'));
    const { args, meetingMsgs } = setup();
    await answerInMeeting(args, 'q', { deliver: 'broadcast', askerId: 'ou_li' });
    expect(meetingMsgs[0]).toContain('（未核对到原话）');
    const qa = readFileSync(join(meetingDataDir((args as { controls: never }).controls), 'qa', 'm1.jsonl'), 'utf8');
    expect(JSON.parse(qa).said).toEqual({ cited: [], unverified: ['09:00 王五'] });
  });

  it('the no-invention rule is there even without docs', () => {
    expect(buildMeetingPrompt({ question: 'q', transcript: [] })).toContain('会上都没有提到');
  });
});

describe('todos: an owner or deadline not made clear is 待确认', () => {
  it('placeholders the model writes count as not said; due_date only when it is a real day', () => {
    const out = splitMinutesTodos(
      '纪要\n```json\n{"todos":[{"task":"a","owner":"未明确","due":"待定","due_date":"2026-02-30"},{"task":"b","owner":"甲","due":"周五","due_date":"2026-11-28"}]}\n```',
    );
    expect(out.todos).toEqual([
      { task: 'a', owner: '', due: '' },
      { task: 'b', owner: '甲', due: '周五', dueDate: '2026-11-28' },
    ]);
  });

  it('the card shows 待确认 in orange and says what it means', () => {
    const rec = {
      id: '20261127-1500-123456789',
      meetingId: 'm1',
      meetingNo: '123456789',
      endedAt: new Date(T0).toISOString(),
      minutes: 'x',
      todosParsed: true,
      todos: [{ task: 'a', owner: '', due: '周五' }],
    };
    const json = JSON.stringify(meetingMinutesCard(rec, []));
    expect(json).toContain(`<font color='orange'>${TO_CONFIRM}</font>`);
    expect(json).toContain('会上没说清的，请人工确认');
  });

  it('the archive maps a named owner to the one participant with that name', async () => {
    mocks.startRunFlow.mockResolvedValue(
      fakeRun('纪要\n```json\n{"todos":[{"task":"写说明","owner":"张三","due":""},{"task":"测试","owner":"王五","due":""}]}\n```'),
    );
    const { session, args } = setup({ summaryOnEnd: true });
    session.ingest(speech('e1', 1, '我来写说明'));
    session.ingest(chat('c1', '好的', '李四', 'ou_li'));
    session.markEnded();
    await summarizeEndedMeeting(args);
    const dir = join(meetingDataDir((args as { controls: never }).controls), 'archive');
    const rec = JSON.parse(readFileSync(join(dir, readdirSync(dir)[0]!), 'utf8')) as MinutesRecord;
    expect(rec.todos[0]).toMatchObject({ owner: '张三', ownerId: 'ou_zhang' });
    expect(rec.todos[1]!.ownerId).toBeUndefined();
    expect(rec.participants?.sort()).toEqual(['张三', '李四'].sort());
    expect(String(mocks.startRunFlow.mock.calls[0]![0].prompt)).toContain('会议日期是 2026-11-27');
  });
});

describe('建成飞书任务', () => {
  const rec: MinutesRecord = {
    id: '20261127-1500-123456789',
    meetingId: 'm1',
    meetingNo: '123456789',
    topic: '周会',
    endedAt: '2026-11-27T07:00:00.000Z',
    minutes: 'x',
    todosParsed: true,
    todos: [
      { task: '写发布说明', owner: '张三', ownerId: 'ou_zhang', due: '周五', dueDate: '2026-11-28' },
      { task: '回归测试', owner: '', due: '' },
    ],
  };

  it('assignee is the owner when known, the clicker follows, the day becomes an all-day due', () => {
    const body = meetingTaskBody(rec, rec.todos[0]!, 1, 'ou_admin');
    expect(body).toMatchObject({
      summary: '写发布说明',
      client_token: 'meeting-20261127-1500-123456789-1',
      due: { timestamp: String(new Date('2026-11-28T00:00:00').getTime()), is_all_day: true },
      members: [
        { id: 'ou_zhang', type: 'user', role: 'assignee' },
        { id: 'ou_admin', type: 'user', role: 'follower' },
      ],
    });
    const unclear = meetingTaskBody(rec, rec.todos[1]!, 2, 'ou_admin');
    expect(unclear.due).toBeUndefined();
    expect(unclear.members).toEqual([{ id: 'ou_admin', type: 'user', role: 'follower' }]);
    expect(String(unclear.description)).toContain('待确认');
  });

  it('creates the task once, records it, and a second click only points at it', async () => {
    const controls = { configPath: join(root, 'config.json'), profile: 'claude' };
    const file = meetingArchiveFile(meetingDataDir(controls), rec.id);
    writeMeetingJson(file, rec);
    const request = vi.fn(async () => ({ code: 0, data: { task: { guid: 'g1', url: 'https://applink/t/g1' } } }));
    const replies: string[] = [];
    const ctx = { controls, channel: { rawClient: { request } } as never, senderId: 'ou_admin' };
    await handleMeetingTask(ctx, `${rec.id}#1`, async (m) => void replies.push(m));
    expect(request).toHaveBeenCalledWith(
      expect.objectContaining({ method: 'POST', url: '/open-apis/task/v2/tasks', params: { user_id_type: 'open_id' } }),
    );
    expect(replies[0]).toContain('已建成飞书任务');
    expect(replies[0]).toContain('https://applink/t/g1');
    const saved = JSON.parse(readFileSync(file, 'utf8')) as MinutesRecord;
    expect(saved.todos[0]).toMatchObject({ taskGuid: 'g1', taskUrl: 'https://applink/t/g1' });
    await handleMeetingTask(ctx, `${rec.id}#1`, async (m) => void replies.push(m));
    expect(request).toHaveBeenCalledTimes(1);
    expect(replies[1]).toContain('已经建过');
  });

  it('a missing scope is explained instead of dumping the raw error', async () => {
    const controls = { configPath: join(root, 'config.json'), profile: 'claude' };
    writeMeetingJson(meetingArchiveFile(meetingDataDir(controls), rec.id), rec);
    const request = vi.fn(async () => ({ code: 99991672, msg: 'Access denied. One of the following scopes is required: [task:task:write]' }));
    const replies: string[] = [];
    await handleMeetingTask(
      { controls, channel: { rawClient: { request } } as never, senderId: 'ou_admin' },
      `${rec.id}#2`,
      async (m) => void replies.push(m),
    );
    expect(replies[0]).toContain('task:task:write');
    expect(replies[0]).toContain('建飞书任务失败');
  });
});

// ───────────────────────── batch 3 ─────────────────────────

describe('voice wake', () => {
  it('matches a name at the start only, after an interjection, tolerant of spaces and case', () => {
    expect(matchWakeName('小助手，帮我查一下上线时间', ['小助手'])).toBe('帮我查一下上线时间');
    expect(matchWakeName('嘿 小 助手 帮我查', ['小助手'])).toBe('帮我查');
    expect(matchWakeName('Hey Sanr, what is the date', ['sanr'])).toBe('what is the date');
    expect(matchWakeName('小助手', ['小助手'])).toBe('');
    expect(matchWakeName('刚才小助手说的不对', ['小助手'])).toBeUndefined();
    expect(matchWakeName('小助手们都很好', ['小助手'])).toBeUndefined();
    expect(matchWakeName('sanrio 不错', ['sanr'])).toBeUndefined();
  });

  it('an allowlisted speaker saying the name asks, once the sentence settles; others are ignored silently', async () => {
    vi.useFakeTimers();
    mocks.startRunFlow.mockResolvedValue(fakeRun('下周三'));
    const { session, args, meetingMsgs } = setup({ voiceWake: { enabled: true, names: ['小助手'] } });
    attachMeetingAgent(args);
    session.ingest(speech('e1', 1, '小助手，上线', '张三', 'ou_zhang'));
    session.ingest(speech('e2', 1, '小助手，上线时间是哪天', '张三', 'ou_zhang'));
    session.ingest(speech('e3', 2, '小助手，帮我改代码', '王五', 'ou_wang')); // not on the allowlist
    await vi.advanceTimersByTimeAsync(VOICE_SETTLE_MS + 10);
    expect(mocks.startRunFlow).toHaveBeenCalledTimes(1);
    const call = mocks.startRunFlow.mock.calls[0]![0];
    expect(call.prompt).toContain('上线时间是哪天');
    expect(call.prompt).toContain('口头叫你');
    expect(call.scopeId).toBe(meetingAskScopeId('m1', 'ou_zhang'));
    expect(meetingMsgs[0]).toContain('听到 张三 叫我');
    expect(meetingMsgs.join('\n')).not.toContain('授权'); // no refusal posted for 王五
  });

  it('the name alone waits for the same speaker\'s next sentence', async () => {
    vi.useFakeTimers();
    mocks.startRunFlow.mockResolvedValue(fakeRun('ok'));
    const { session, args } = setup({ voiceWake: { enabled: true, names: ['小助手'] } });
    attachMeetingAgent(args);
    session.ingest(speech('e1', 1, '小助手。', '张三', 'ou_zhang'));
    await vi.advanceTimersByTimeAsync(VOICE_SETTLE_MS + 10);
    expect(mocks.startRunFlow).not.toHaveBeenCalled();
    session.ingest(speech('e2', 2, '这周谁值班', '张三', 'ou_zhang'));
    await vi.advanceTimersByTimeAsync(VOICE_SETTLE_MS + 10);
    expect(mocks.startRunFlow).toHaveBeenCalledTimes(1);
    expect(mocks.startRunFlow.mock.calls[0]![0].prompt).toContain('这周谁值班');
    // Too late: a sentence after the follow-up window is not a question.
    session.ingest(speech('e3', 3, '小助手', '张三', 'ou_zhang'));
    await vi.advanceTimersByTimeAsync(VOICE_SETTLE_MS + VOICE_FOLLOW_MS + 10);
    session.ingest(speech('e4', 4, '随便说说', '张三', 'ou_zhang'));
    await vi.advanceTimersByTimeAsync(VOICE_SETTLE_MS + 10);
    expect(mocks.startRunFlow).toHaveBeenCalledTimes(1);
  });

  it('team mode lets everyone type, but voice needs the allowlist; off by default', async () => {
    vi.useFakeTimers();
    mocks.startRunFlow.mockResolvedValue(fakeRun('ok'));
    const team = setup({ voiceWake: { enabled: true, names: ['小助手'] } }, { mode: 'team', allowed: [] });
    attachMeetingAgent(team.args);
    team.session.ingest(speech('e1', 1, '小助手，几点了', '路人', 'ou_x'));
    await vi.advanceTimersByTimeAsync(VOICE_SETTLE_MS + 10);
    expect(mocks.startRunFlow).not.toHaveBeenCalled();

    const off = setup();
    attachMeetingAgent(off.args);
    off.session.ingest(speech('e1', 1, '三儿，几点了', '张三', 'ou_zhang'));
    await vi.advanceTimersByTimeAsync(VOICE_SETTLE_MS + 10);
    expect(mocks.startRunFlow).not.toHaveBeenCalled();
  });
});

describe('speaking up unasked', () => {
  function withDocs(args: unknown) {
    const controls = (args as { controls: { configPath: string; profile: string } }).controls;
    const dir = meetingDataDir(controls);
    const ref = {
      token: 'doxcn1',
      url: 'https://x.feishu.cn/docx/doxcn1',
      title: '发布手册',
      revision: 3,
      fetchedAt: new Date(T0).toISOString(),
      chunks: 1,
      chars: 20,
      boundAt: new Date(T0).toISOString(),
    };
    writeMeetingJson(docSnapshotFile(dir, ref.token, ref.revision), {
      ...ref,
      truncated: false,
      chunks: [{ i: 1, text: '发布窗口是每周三晚上八点。' }],
    });
    bindDoc(dir, { meetingId: 'm1', meetingNo: '123456789' }, ref);
    expect(meetingDocRefs(dir, 'm1')).toHaveLength(1);
  }

  it('recognizes plain questions only', () => {
    expect(looksLikeQuestion('发布窗口是哪天？')).toBe(true);
    expect(looksLikeQuestion('这个能周三发吗')).toBe(true);
    expect(looksLikeQuestion('请问发布窗口在什么时候')).toBe(true);
    expect(looksLikeQuestion('我们周三发布。')).toBe(false);
    expect(looksLikeQuestion('好？')).toBe(false);
  });

  it('an unanswered question the docs answer gets one cited line; later ones get nothing', async () => {
    vi.useFakeTimers();
    mocks.startRunFlow.mockResolvedValue(fakeRun('每周三晚上八点〔《发布手册》r3 §1〕'));
    const { session, args, meetingMsgs } = setup({ proactive: { enabled: true, waitMs: 20_000, maxPerMeeting: 1 } });
    withDocs(args);
    attachMeetingAgent(args);
    session.ingest(speech('e1', 1, '发布窗口是什么时候？'));
    await vi.advanceTimersByTimeAsync(20_010);
    expect(mocks.startRunFlow).toHaveBeenCalledTimes(1);
    const call = mocks.startRunFlow.mock.calls[0]![0];
    expect(call.scopeId).toBe('meeting:m1:proactive');
    expect(call.prompt).toContain('SKIP');
    expect(meetingMsgs).toEqual(['💡 刚才 张三 问「发布窗口是什么时候？」，会前资料里有：每周三晚上八点〔《发布手册》r3 §1〕']);
    session.ingest(speech('e2', 2, '那回滚窗口是什么时候？'));
    await vi.advanceTimersByTimeAsync(20_010);
    expect(mocks.startRunFlow).toHaveBeenCalledTimes(1);
  });

  it('stays quiet when someone else answered, when the model passes, or without a citation', async () => {
    vi.useFakeTimers();
    const { session, args, meetingMsgs } = setup({ proactive: { enabled: true, waitMs: 20_000, maxPerMeeting: 3 } });
    withDocs(args);
    attachMeetingAgent(args);
    session.ingest(speech('e1', 1, '发布窗口是什么时候？'));
    session.ingest(speech('e2', 2, '周三晚上', '李四', 'ou_li'));
    await vi.advanceTimersByTimeAsync(20_010);
    expect(mocks.startRunFlow).not.toHaveBeenCalled();

    mocks.startRunFlow.mockResolvedValueOnce(fakeRun('SKIP'));
    session.ingest(speech('e3', 3, '谁来主持下次会？'));
    await vi.advanceTimersByTimeAsync(20_010);
    mocks.startRunFlow.mockResolvedValueOnce(fakeRun('大概周三吧'));
    session.ingest(speech('e4', 4, '回滚要多久？'));
    await vi.advanceTimersByTimeAsync(20_010);
    expect(mocks.startRunFlow).toHaveBeenCalledTimes(2);
    expect(meetingMsgs).toEqual([]);
  });

  it('is off by default', async () => {
    vi.useFakeTimers();
    const { session, args } = setup();
    withDocs(args);
    attachMeetingAgent(args);
    session.ingest(speech('e1', 1, '发布窗口是什么时候？'));
    await vi.advanceTimersByTimeAsync(60_000);
    expect(mocks.startRunFlow).not.toHaveBeenCalled();
  });
});

describe('asking about past meetings', () => {
  const NOW = new Date(2026, 10, 27, 10, 0, 0).getTime(); // Friday

  function rec(id: string, endedAt: Date, topic: string, minutes: string, participants: string[] = []): MinutesRecord {
    return {
      id,
      meetingId: id,
      meetingNo: id.slice(-9),
      topic,
      endedAt: endedAt.toISOString(),
      minutes,
      todosParsed: true,
      todos: [{ task: '写说明', owner: '', due: '' }],
      participants,
    };
  }

  it('recognizes recall questions about meetings, not ordinary requests', () => {
    for (const q of ['上周那个会定了什么', '和设计组的会结论是什么', '昨天周会谁负责发布', '把上周的会议纪要发我']) {
      expect(looksLikeHistoryQuestion(q)).toBe(true);
    }
    for (const q of ['帮我写个会议纪要模板', '我们的会员系统谁负责', '定了什么规矩', '/meeting status', '开个会吧']) {
      expect(looksLikeHistoryQuestion(q)).toBe(false);
    }
  });

  it('reads the time the question names', () => {
    expect(historyWindow('上周那个会', NOW).label).toBe('上周');
    const w = historyWindow('上周那个会', NOW);
    expect(new Date(w.from).getDay()).toBe(1);
    expect(w.to - w.from).toBe(7 * 24 * 3600 * 1000);
    expect(historyWindow('昨天的会', NOW).label).toBe('昨天');
    expect(historyWindow('11月20日的会', NOW).label).toBe('11月20日');
    expect(historyWindow('最近3天的会', NOW).label).toBe('最近 3 天');
    expect(historyWindow('那个会', NOW).label).toBe('最近 30 天');
  });

  it('picks the meetings in the window that share words, newest first', () => {
    const all = [
      rec('20261126-1500-111111111', new Date(2026, 10, 26, 15), '周会', '定了周五发布'),
      rec('20261120-1500-222222222', new Date(2026, 10, 20, 15), '设计评审', '定了新配色', ['王设计']),
      rec('20261119-1500-333333333', new Date(2026, 10, 19, 15), '周会', '定了周三发布'),
    ];
    expect(pickArchives(all, '上周的设计评审定了什么', NOW).picked.map((r) => r.topic)).toEqual(['设计评审']);
    expect(pickArchives(all, '上周那个会定了什么', NOW).picked.map((r) => r.id)).toEqual([
      '20261120-1500-222222222',
      '20261119-1500-333333333',
    ]);
    expect(pickArchives(all, '昨天的周会定了什么', NOW).picked.map((r) => r.id)).toEqual(['20261126-1500-111111111']);
  });

  it('answers from the picked minutes with sources, and says so plainly when there are none', async () => {
    const { args } = setup();
    const deps = { ...(args as object), chatId: 'oc_dm' } as never;
    expect(await answerMeetingHistory(deps, '上周那个会定了什么', NOW)).toContain('还没有会议纪要存档');

    const controls = (args as { controls: { configPath: string; profile: string } }).controls;
    const dir = meetingDataDir(controls);
    const r = rec('20261120-1500-222222222', new Date(2026, 10, 20, 15), '设计评审', '定了新配色');
    writeMeetingJson(meetingArchiveFile(dir, r.id), r);
    expect(await answerMeetingHistory(deps, '昨天的会定了什么', NOW)).toContain('没找到昨天的会议纪要');
    expect(mocks.startRunFlow).not.toHaveBeenCalled();

    mocks.startRunFlow.mockResolvedValue(fakeRun('定了新配色〔11-20 设计评审〕'));
    const text = await answerMeetingHistory(deps, '上周那个会定了什么', NOW);
    expect(text).toContain('定了新配色〔11-20 设计评审〕');
    expect(text).toContain('存档 20261120-1500-222222222');
    const call = mocks.startRunFlow.mock.calls[0]![0];
    expect(call.scopeId).toBe('meeting-history:oc_dm');
    expect(call.prompt).toContain('〔11-20 设计评审〕');
    expect(call.prompt).toContain('纪要里没有记录');
    expect(call.prompt).toContain('负责人：待确认');
  });
});

describe('docs from the calendar event', () => {
  function calendarClient(over: { event?: object; fail?: boolean; doc?: boolean } = {}) {
    const calls: string[] = [];
    const request = vi.fn(async (p: { method: string; url: string; params?: Record<string, unknown>; data?: unknown }) => {
      calls.push(`${p.method} ${p.url}`);
      if (p.url === '/open-apis/vc/v1/meetings/m1') {
        if (over.fail) return { code: 99991672, msg: 'Access denied [vc:meeting:readonly]' };
        return { code: 0, data: { meeting: { host_user: { id: 'ou_host' }, start_time: String(T0 / 1000) } } };
      }
      if (p.url === '/open-apis/calendar/v4/calendars/primarys') {
        return { code: 0, data: { calendars: [{ calendar: { calendar_id: 'cal_host' } }] } };
      }
      if (p.url.endsWith('/events/instance_view')) {
        return {
          code: 0,
          data: {
            items: [
              { summary: '别的会', vchat: { meeting_url: 'https://vc.feishu.cn/j/999999999' }, description: 'https://x.feishu.cn/docx/AAAAAAAAAAAAAAAAAA' },
              over.event ?? {
                summary: '周会',
                vchat: { meeting_url: 'https://vc.feishu.cn/j/123456789' },
                description: '资料：https://x.feishu.cn/docx/BBBBBBBBBBBBBBBBBB 和 https://x.feishu.cn/docx/CCCCCCCCCCCCCCCCCC',
              },
            ],
          },
        };
      }
      if (p.url.startsWith('/open-apis/docx/v1/documents/')) {
        const token = p.url.split('/')[5]!;
        if (token === 'CCCCCCCCCCCCCCCCCC') return { code: 1770032, msg: 'forbidden' };
        if (p.url.endsWith('/raw_content')) return { code: 0, data: { content: '发布窗口是周三。' } };
        return { code: 0, data: { document: { title: '发布手册', revision_id: 7 } } };
      }
      return { code: 0, data: {} };
    });
    return { request, calls };
  }

  it('finds the event by the meeting number in its video link and returns its doc links', async () => {
    const { request, calls } = calendarClient();
    const found = await findCalendarEventDocs({ request } as never, { meetingId: 'm1', meetingNo: '123456789' }, ['ou_owner'], T0);
    expect(found?.summary).toBe('周会');
    expect(found?.links.map((l) => l.token)).toEqual(['BBBBBBBBBBBBBBBBBB', 'CCCCCCCCCCCCCCCCCC']);
    expect(calls).toEqual([
      'GET /open-apis/vc/v1/meetings/m1',
      'POST /open-apis/calendar/v4/calendars/primarys',
      'GET /open-apis/calendar/v4/calendars/cal_host/events/instance_view',
    ]);
    const prim = request.mock.calls[1]![0] as { data: { user_ids: string[] } };
    expect(prim.data.user_ids).toEqual(['ou_host', 'ou_owner']);
  });

  it('binds the readable ones on join and tells the origin chat what could not be read', async () => {
    const { request } = calendarClient();
    const { args, sent } = setup({ calendarDocs: true }, { client: { request } });
    await attachCalendarDocs(args);
    const dir = meetingDataDir((args as { controls: never }).controls);
    expect(meetingDocRefs(dir, 'm1').map((r) => r.token)).toEqual(['BBBBBBBBBBBBBBBBBB']);
    expect(sent).toHaveLength(1);
    expect(sent[0]!.to).toBe('oc_team');
    expect(sent[0]!.input.markdown).toContain('已从日程「周会」带上会前资料');
    expect(sent[0]!.input.markdown).toContain('1 份文档读不到');
  });

  it('an unreadable calendar is logged and the owner told once; off and resumed meetings do nothing', async () => {
    const { request } = calendarClient({ fail: true });
    const a = setup({ calendarDocs: true }, { client: { request } });
    await attachCalendarDocs(a.args);
    await attachCalendarDocs(a.args);
    expect(a.sent.filter((m) => m.to === 'ou_owner')).toHaveLength(1);
    expect(a.sent[0]!.input.markdown).toContain('calendar:calendar:readonly');

    const off = setup({ calendarDocs: false }, { client: { request } });
    expect(attachCalendarDocs(off.args)).toBeUndefined();
    const resumed = setup({ calendarDocs: true }, { client: { request } });
    resumed.session.resumedAt = Date.now();
    expect(attachCalendarDocs(resumed.args)).toBeUndefined();
  });
});

describe('config', () => {
  it('defaults: split at 300, calendar docs and DM history on, proactive and voice off', () => {
    const p = normalizeProfileConfig({
      schemaVersion: 2,
      agentKind: 'claude',
      accounts: { app: { id: 'cli_x', secret: 's', tenant: 'feishu' } },
    } as never);
    expect(p.meeting).toMatchObject({
      longAnswerChars: 300,
      calendarDocs: true,
      historyInDm: true,
      proactive: { enabled: false, waitMs: 20_000, maxPerMeeting: 1 },
      voiceWake: { enabled: false, names: [] },
    });
  });

  it('clamps and cleans what is configured; the console form keeps it', () => {
    const p = normalizeProfileConfig({
      schemaVersion: 2,
      agentKind: 'claude',
      accounts: { app: { id: 'cli_x', secret: 's', tenant: 'feishu' } },
      meeting: {
        longAnswerChars: 99999,
        proactive: { enabled: true, waitMs: 1, maxPerMeeting: 50 },
        voiceWake: { enabled: true, names: [' 小助手 ', 'x', '小助手', 3] },
      },
    } as never);
    expect(p.meeting.longAnswerChars).toBe(5000);
    expect(p.meeting.proactive).toEqual({ enabled: true, waitMs: 5000, maxPerMeeting: 10 });
    expect(p.meeting.voiceWake).toEqual({ enabled: true, names: ['小助手'] });
    const kept = parseMeetingBody({ enabled: true }, p.meeting);
    expect(kept.proactive).toEqual(p.meeting.proactive);
    expect(kept.voiceWake).toEqual(p.meeting.voiceWake);
    expect(kept.longAnswerChars).toBe(5000);
  });

  it('the join notice mentions DM for long answers, and voice / proactive only when on', () => {
    const profile = createDefaultProfileConfig({
      agentKind: 'claude',
      accounts: { app: { id: 'cli_x', secret: 's', tenant: 'feishu' } },
    });
    profile.meeting = cfg();
    const plain = meetingJoinNotice(profile, '三儿');
    expect(plain).toContain('全文私聊发给提问的人');
    expect(plain).not.toContain('开口叫我');
    expect(plain).not.toContain('插一句');
    profile.meeting = cfg({ voiceWake: { enabled: true, names: [] }, proactive: { enabled: true, waitMs: 20000, maxPerMeeting: 1 } });
    const on = meetingJoinNotice(profile, '三儿');
    expect(on).toContain('开口叫我：句子开头说「三儿」');
    expect(on).toContain('每场最多 1 次');
  });
});
