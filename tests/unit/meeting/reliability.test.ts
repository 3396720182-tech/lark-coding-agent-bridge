import { existsSync, mkdtempSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

const mocks = vi.hoisted(() => ({ startRunFlow: vi.fn() }));
vi.mock('../../../src/bot/run-flow', () => ({ startRunFlow: mocks.startRunFlow }));

const {
  attachMeetingAgent,
  meetingJoinNotice,
  meetingMinutesScopeId,
  splitMeetingRecord,
  summarizeEndedMeeting,
  summarizeMeetingRecord,
  MINUTES_SINGLE_PASS_CHARS,
} = await import('../../../src/meeting/orchestrator');
const { MeetingSession, GAP_FILL_INTERVAL_MS, LEAVE_BACKFILL_DELAY_MS } = await import('../../../src/meeting/session');
const { MeetingManager, VC_BOT_EVENTS } = await import('../../../src/meeting/manager');
const { MEETING_DEFAULTS, createDefaultProfileConfig } = await import('../../../src/config/profile-schema');
const { loadTranscript, transcriptFile, transcriptTextFile } = await import('../../../src/meeting/transcript-log');
const { epochMs, unpackActivity } = await import('../../../src/meeting/types');
const { pushStatusLine } = await import('../../../src/commands/index');

import type { MeetingConfig } from '../../../src/config/profile-schema';
import type { RawActivityItem, VcRequestClient } from '../../../src/meeting/api';

let root: string;
beforeEach(() => {
  vi.clearAllMocks();
  root = mkdtempSync(join(tmpdir(), 'meeting-reliability-'));
});
afterEach(() => {
  vi.useRealTimers();
  rmSync(root, { recursive: true, force: true });
});

function cfg(over: Partial<MeetingConfig> = {}): MeetingConfig {
  return {
    ...MEETING_DEFAULTS,
    enabled: true,
    summaryOnEnd: true,
    ...over,
    transcript: { ...MEETING_DEFAULTS.transcript, ...over.transcript },
  };
}

const T0 = new Date(2026, 10, 27, 14, 0, 0).getTime();

function speech(eid: string, sid: number | string, text: string, who = '甲', atSec?: number): RawActivityItem {
  return {
    event_id: eid,
    activity_event_type: 'transcript_received',
    transcript_received_items: [
      { sentence_id: sid, text, speaker: { open_id: `ou_${who}`, name: who }, ...(atSec ? { start_time_ms: atSec * 1000 } : {}) },
    ],
  };
}

function chat(eid: string, text: string, who = '乙', createSec?: number): RawActivityItem {
  return {
    event_id: eid,
    activity_event_type: 'chat_received',
    chat_received_items: [
      { content: text, message_type: 1, operator: { id: `ou_${who}`, user_name: who }, ...(createSec ? { create_time: String(createSec) } : {}) },
    ],
  };
}

function left(eid: string, who = 'ou_甲'): RawActivityItem {
  return { event_id: eid, activity_event_type: 'participant_left', participant_left_items: [{ user: { open_id: who } }] };
}

/** A client whose `bots/events` answers from `pages` (a fresh read starts at page 0). */
function pagedClient(pages: RawActivityItem[][], extra?: (url: string) => unknown) {
  const calls: { url: string; params?: Record<string, unknown> }[] = [];
  const client: VcRequestClient = {
    request: vi.fn(async (p: { url: string; params?: Record<string, unknown> }) => {
      calls.push({ url: p.url, ...(p.params ? { params: p.params } : {}) });
      const other = extra?.(p.url);
      if (other) return other as never;
      if (p.url.endsWith('/events')) {
        const i = p.params?.page_token ? Number(p.params.page_token) : 0;
        const more = i + 1 < pages.length;
        return { code: 0, data: { events: pages[i] ?? [], has_more: more, page_token: String(i + 1) } } as never;
      }
      if (p.url.endsWith('/join')) return { code: 0, data: { meeting: { id: 'm1', topic: '周会' } } } as never;
      return { code: 0, data: {} } as never;
    }),
  };
  return { client, calls };
}

function makeSession(
  over: Partial<MeetingConfig> = {},
  client?: VcRequestClient,
  dir = join(root, 'transcripts'),
  start = T0,
) {
  let now = start;
  const s = new MeetingSession({
    client: client ?? pagedClient([]).client,
    meetingId: 'm1',
    meetingNo: '123456789',
    topic: '周会',
    config: cfg(over),
    botOpenId: 'ou_bot',
    transcriptDir: dir,
    now: () => (now += 1000),
  });
  return s;
}

describe('event time', () => {
  it('reads seconds and milliseconds, ignores relative offsets', () => {
    expect(epochMs('1700000000')).toBe(1_700_000_000_000);
    expect(epochMs(1_700_000_000_123)).toBe(1_700_000_000_123);
    expect(epochMs(4500)).toBeUndefined(); // ms since meeting start
    expect(epochMs('abc')).toBeUndefined();
  });

  it('carries the event time on transcript and chat events', () => {
    const [t] = unpackActivity(speech('e1', 1, '你好', '甲', 1_700_000_000));
    const [c] = unpackActivity(chat('c1', '收到', '乙', 1_700_000_100));
    expect(t).toMatchObject({ atMs: 1_700_000_000_000 });
    expect(c).toMatchObject({ atMs: 1_700_000_100_000 });
  });
});

describe('full transcript on disk', () => {
  it('keeps every line (not just transcript.keep), with speaker and time', () => {
    const s = makeSession({ transcript: { ...MEETING_DEFAULTS.transcript, keep: 10 } });
    for (let i = 1; i <= 50; i++) s.ingest(speech(`e${i}`, i, `第${i}句`));
    s.ingest(chat('c1', '弹幕一句'));
    const full = s.fullContext();
    expect(full).toHaveLength(51);
    expect(full[0]).toMatch(/^\[\d\d:\d\d:\d\d\] 甲: 第1句$/);
    expect(full[50]).toContain('乙（弹幕）: 弹幕一句');
    // Answers still get only the last `keep` lines.
    expect(s.recentContext(10)).toHaveLength(10);

    const { meta, records } = loadTranscript(transcriptFile(join(root, 'transcripts'), 'm1'));
    expect(meta).toMatchObject({ k: 'meta', meetingId: 'm1', meetingNo: '123456789', topic: '周会' });
    expect(records).toHaveLength(51);
    expect(records[0]).toMatchObject({ k: 'speech', sid: '1', who: '甲', whoId: 'ou_甲', text: '第1句', eid: 'e1' });
    expect(records[50]).toMatchObject({ k: 'chat', who: '乙', text: '弹幕一句' });
    expect(typeof (records[0] as { at: number }).at).toBe('number');
  });

  it('a restarted bridge continues from the file, and the replay is not stored twice', async () => {
    const dir = join(root, 'transcripts');
    const first = makeSession({}, undefined, dir);
    first.ingest(speech('e1', 1, '今天'));
    first.ingest(speech('e2', 1, '今天来讨论发布'));
    first.ingest(chat('c1', '收到'));
    first.dispose();

    // The events API replays everything from the start on catch-up.
    const { client } = pagedClient([[speech('e1', 1, '今天'), speech('e2', 1, '今天来讨论发布'), chat('c1', '收到')], [speech('e3', 2, '新的一句')]]);
    const again = makeSession({}, client, dir, T0 + 3_600_000);
    expect(again.fullContext()).toHaveLength(2);
    await again.catchUp();
    expect(again.recentContext()).toEqual(['甲: 今天来讨论发布', '乙（弹幕）: 收到', '甲: 新的一句']);
  });

  it('finalize compacts the file and writes a readable copy', () => {
    const dir = join(root, 'transcripts');
    const s = makeSession({}, undefined, dir);
    s.ingest(speech('e1', 1, '今天'));
    s.ingest(speech('e2', 1, '今天来讨论'));
    s.ingest(chat('c1', '好的'));
    const txt = s.finalizeTranscript();
    expect(txt).toBe(transcriptTextFile(dir, 'm1'));
    const body = readFileSync(txt!, 'utf8');
    expect(body).toContain('会议逐字稿 · 周会');
    expect(body).toMatch(/\[\d\d:\d\d:\d\d\] 甲：今天来讨论\n/);
    expect(body).toMatch(/\[\d\d:\d\d:\d\d\] 乙（弹幕）：好的\n/);
    expect(body).not.toMatch(/甲：今天\n/);
    const recs = loadTranscript(transcriptFile(dir, 'm1')).records;
    expect(recs.filter((r) => r.k !== 'seen')).toHaveLength(2);
    // Every id seen is kept, for a later re-join of the same meeting.
    expect(recs.find((r) => r.k === 'seen')).toMatchObject({ seen: expect.arrayContaining(['e1', 'e2', 'c1']) });
    expect(s.status().transcriptText).toBe(txt);
  });

  it('transcript.save: false writes nothing', () => {
    const dir = join(root, 'transcripts');
    const s = makeSession({ transcript: { ...MEETING_DEFAULTS.transcript, save: false } }, undefined, dir);
    s.ingest(speech('e1', 1, '不落盘'));
    expect(s.finalizeTranscript()).toBeUndefined();
    expect(existsSync(dir)).toBe(false);
    expect(s.status().transcriptFile).toBeUndefined();
  });

  it('orders by the event time, so a line recovered late lands where it was said', () => {
    const s = makeSession();
    s.ingest(speech('e2', 2, '后说的', '甲', 1_800_000_100));
    s.ingest(chat('c1', '先发的', '乙', 1_800_000_000), { quiet: true });
    expect(s.recentContext()).toEqual(['乙（弹幕）: 先发的', '甲: 后说的']);
  });
});

describe('full re-read (backfill)', () => {
  it('reads from the first page, recovers what was missed, tells no subscriber', async () => {
    const pages = [
      [speech('e1', 1, '第一句'), chat('c1', '@bot 老问题')],
      [speech('e2', 2, '漏掉的一句')],
    ];
    const { client, calls } = pagedClient(pages);
    const s = makeSession({}, client);
    s.ingest(speech('e1', 1, '第一句'));
    const heard: string[] = [];
    s.on('chat', () => heard.push('chat'));
    s.on('transcript', () => heard.push('transcript'));

    const added = await s.backfill('test');
    expect(added).toBe(2);
    expect(heard).toEqual([]); // no answer to an old question
    expect(s.recentContext()).toEqual(['甲: 第一句', '乙（弹幕）: @bot 老问题', '甲: 漏掉的一句']);
    expect(calls.filter((c) => c.url.endsWith('/events')).map((c) => c.params?.page_token)).toEqual([undefined, '1']);
    expect(s.status()).toMatchObject({ backfills: 1, backfilled: 2 });
    // A second pass finds nothing new.
    expect(await s.backfill('again')).toBe(0);
  });

  it('never lets a replayed early prefix overwrite the final sentence, but takes a longer one', async () => {
    const { client } = pagedClient([[speech('x1', 1, '今天'), speech('x2', 2, '散会'), speech('x3', 2, '散会前再确认一下')]]);
    const s = makeSession({}, client);
    s.ingest(speech('e1', 1, '今天来讨论'));
    s.ingest(speech('e2', 2, '散会'));
    await s.backfill('test');
    expect(s.recentTranscript()).toEqual(['甲: 今天来讨论', '甲: 散会前再确认一下']);
  });

  it('does not store a chat line twice when the re-read carries another dedup key', async () => {
    const noId = { ...chat('x', '同一句话', '乙', 1_800_000_000) };
    delete (noId as { event_id?: string }).event_id;
    const { client } = pagedClient([[chat('real-id', '同一句话', '乙', 1_800_000_000)]]);
    const s = makeSession({}, client);
    s.ingest(noId);
    await s.backfill('test');
    expect(s.recentContext()).toEqual(['乙（弹幕）: 同一句话']);
  });

  it('a 122001 during the re-read ends the session like the poller would', async () => {
    const client: VcRequestClient = {
      request: vi.fn(async () => {
        throw Object.assign(new Error('400'), { response: { status: 400, data: { code: 122001, msg: 'meeting status unexpected' } } });
      }),
    };
    const s = makeSession({}, client);
    const gone = vi.fn();
    s.onGone = gone;
    await s.backfill('test');
    expect(s.ended).toBe(true);
    expect(gone).toHaveBeenCalledTimes(1);
  });

  it('someone leaving triggers a re-read shortly after (the meeting may be ending)', async () => {
    vi.useFakeTimers();
    const { client, calls } = pagedClient([[speech('e9', 9, '补回来的')]]);
    const s = makeSession({}, client);
    s.ingest(left('p1'));
    expect(calls).toHaveLength(0);
    await vi.advanceTimersByTimeAsync(LEAVE_BACKFILL_DELAY_MS + 10);
    expect(calls.filter((c) => c.url.endsWith('/events'))).toHaveLength(1);
    expect(s.recentTranscript()).toEqual(['甲: 补回来的']);
    s.dispose();
  });

  it('push does not stop the poller; it slows to a gap-fill', async () => {
    vi.useFakeTimers();
    const { client, calls } = pagedClient([[]]);
    const s = makeSession({}, client);
    s.startPolling();
    await vi.advanceTimersByTimeAsync(0);
    s.markPushActive();
    // Let the tick already scheduled at poll speed run.
    await vi.advanceTimersByTimeAsync(10_000);
    const before = calls.length;
    await vi.advanceTimersByTimeAsync(GAP_FILL_INTERVAL_MS - 15_000);
    expect(calls.length).toBe(before); // not at poll speed any more
    await vi.advanceTimersByTimeAsync(20_000);
    expect(calls.length).toBe(before + 1); // but not stopped either
    s.dispose();
  });
});

describe('manager: leave and end', () => {
  it('reads the whole meeting once more before leaving, then writes the transcript', async () => {
    const { client, calls } = pagedClient([[speech('e1', 1, '最后一句')]]);
    const mgr = new MeetingManager({ client, config: () => cfg(), transcriptDir: join(root, 'transcripts') });
    const s = await mgr.join('123456789');
    s.dispose(); // stop the poller; this test drives the calls
    calls.length = 0;
    await mgr.leave(s.meetingId);
    const urls = calls.map((c) => c.url.split('/').pop());
    expect(urls).toEqual(['events', 'leave']);
    expect(readFileSync(transcriptTextFile(join(root, 'transcripts'), 'm1'), 'utf8')).toContain('甲：最后一句');
  });

  it('a pushed end writes the transcript before the minutes run', async () => {
    const handlers = new Map<string, (d: unknown) => void>();
    const channel = { onRawEvent: (t: string, h: (d: unknown) => void) => (handlers.set(t, h), () => {}) };
    const seen: (string | undefined)[] = [];
    const { client } = pagedClient([[]]);
    const mgr = new MeetingManager({
      client,
      config: () => cfg(),
      channel,
      transcriptDir: join(root, 'transcripts'),
      onEnded: (s) => seen.push(s.status().transcriptText),
    });
    mgr.attachPush();
    const s = await mgr.join('123456789');
    s.ingest(speech('e1', 1, '收尾'));
    handlers.get(VC_BOT_EVENTS.ended)!({ meeting: { id: 'm1' } });
    expect(seen[0]).toBe(transcriptTextFile(join(root, 'transcripts'), 'm1'));
  });
});

describe('resume after the meeting ended while the bridge was down', () => {
  it('still writes the readable transcript of what was recorded', async () => {
    const dir = join(root, 'transcripts');
    const before = makeSession({}, undefined, dir);
    before.ingest(speech('e1', 1, '重启前说的'));
    before.dispose();
    const stateFile = join(root, 'meetings.json');
    const { writeFileSync } = await import('node:fs');
    writeFileSync(stateFile, JSON.stringify({ meetings: [{ meetingId: 'm1', meetingNo: '123456789', startedAt: new Date().toISOString() }] }));
    const client: VcRequestClient = {
      request: vi.fn(async () => {
        throw Object.assign(new Error('400'), { response: { status: 400, data: { code: 122001, msg: 'gone' } } });
      }),
    };
    const ended = vi.fn();
    const mgr = new MeetingManager({ client, config: () => cfg(), stateFile, transcriptDir: dir, onResumeEnded: ended });
    await mgr.resume();
    expect(ended).toHaveBeenCalledTimes(1);
    expect(readFileSync(transcriptTextFile(dir, 'm1'), 'utf8')).toContain('甲：重启前说的');
  });
});

describe('push diagnosis (/meeting status)', () => {
  const hookedChannel = { onRawEvent: () => () => {} };
  const appInfo = (events: string[] | Error) => (url: string) => {
    if (!url.includes('/application/v6/applications/')) return undefined;
    if (events instanceof Error) throw events;
    return { code: 0, data: { app: { event: { subscribed_events: events } } } };
  };

  it('not hooked → polling only', async () => {
    const mgr = new MeetingManager({ client: pagedClient([]).client, config: () => cfg() });
    mgr.attachPush();
    const d = await mgr.diagnosePush();
    expect(d.state).toBe('unhooked');
    expect(pushStatusLine(d)).toContain('只靠轮询');
  });

  it('the Open Platform says the activity event is not subscribed', async () => {
    const { client } = pagedClient([], appInfo(['im.message.receive_v1']));
    const mgr = new MeetingManager({ client, config: () => cfg(), channel: hookedChannel, appId: 'cli_x' });
    mgr.attachPush();
    const d = await mgr.diagnosePush();
    expect(d).toMatchObject({ state: 'not-subscribed', evidence: 'console' });
    expect(pushStatusLine(d)).toContain('没订阅事件');
    expect(pushStatusLine(d)).toContain(VC_BOT_EVENTS.activity);
  });

  it('subscribed and no meeting yet is not an error', async () => {
    const { client } = pagedClient([], appInfo(Object.values(VC_BOT_EVENTS)));
    const mgr = new MeetingManager({ client, config: () => cfg(), channel: hookedChannel, appId: 'cli_x' });
    mgr.attachPush();
    const d = await mgr.diagnosePush();
    expect(d.state).toBe('subscribed-idle');
    expect(pushStatusLine(d)).toContain('还没开过会');
  });

  it('subscription unreadable and no meeting: says it cannot tell', async () => {
    const { client } = pagedClient([], appInfo(new Error('no permission')));
    const mgr = new MeetingManager({ client, config: () => cfg(), channel: hookedChannel, appId: 'cli_x' });
    mgr.attachPush();
    const d = await mgr.diagnosePush();
    expect(d.state).toBe('unknown');
    expect(pushStatusLine(d)).toContain('判断不了');
  });

  it('events polled but no push at all → most likely not subscribed', async () => {
    vi.useFakeTimers();
    const { client } = pagedClient([[speech('e1', 1, '有人说话')]], appInfo(new Error('no permission')));
    const mgr = new MeetingManager({ client, config: () => cfg(), channel: hookedChannel, appId: 'cli_x' });
    mgr.attachPush();
    const s = await mgr.join('123456789');
    await vi.advanceTimersByTimeAsync(0);
    // Right after joining, the history pull alone proves nothing.
    expect((await mgr.diagnosePush()).state).toBe('waiting');
    await vi.advanceTimersByTimeAsync(130_000);
    const d = await mgr.diagnosePush();
    expect(d).toMatchObject({ state: 'not-subscribed', evidence: 'polled', polled: 1 });
    expect(pushStatusLine(d)).toContain('多半没订阅');
    s.dispose();
  });

  it('pushes arrived → ok', async () => {
    const handlers = new Map<string, (d: unknown) => void>();
    const channel = { onRawEvent: (t: string, h: (d: unknown) => void) => (handlers.set(t, h), () => {}) };
    const mgr = new MeetingManager({ client: pagedClient([]).client, config: () => cfg(), channel });
    mgr.attachPush();
    handlers.get(VC_BOT_EVENTS.activity)!({ meeting_activity_items: [] });
    const d = await mgr.diagnosePush();
    expect(d).toMatchObject({ state: 'ok', received: 1 });
    expect(pushStatusLine(d)).toContain('推送：正常');
  });
});

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

function agentDeps(session: InstanceType<typeof MeetingSession>, meeting: MeetingConfig, botName = '三儿') {
  const profileConfig = createDefaultProfileConfig({
    agentKind: 'claude',
    accounts: { app: { id: 'cli_test', secret: '${APP_SECRET}', tenant: 'feishu' } },
  });
  profileConfig.meeting = meeting;
  const sent: { to: string; input: { markdown?: string; card?: object } }[] = [];
  const args = {
    session,
    channel: {
      botIdentity: { name: botName, openId: 'ou_bot' },
      send: vi.fn(async (to: string, input: { markdown?: string; card?: object }) => {
        sent.push({ to, input });
        return { messageId: 'om_1' };
      }),
    },
    controls: { configPath: join(root, 'config.json'), profile: 'claude', profileConfig, botOwnerId: 'ou_owner' },
    executor: {},
    activeRuns: { interrupt: vi.fn() },
    sessions: {},
    workspaces: {},
  } as never;
  return { args, sent, profileConfig };
}

const MINUTES = '**讨论要点**\n- 发布\n```json\n{"todos":[]}\n```';

describe('minutes from the whole meeting', () => {
  it('includes the first half of a long meeting (the old 200-line window lost it)', async () => {
    mocks.startRunFlow.mockResolvedValue(fakeRun(MINUTES));
    const meeting = cfg({ transcript: { ...MEETING_DEFAULTS.transcript, keep: 10 } });
    const s = makeSession(meeting);
    for (let i = 1; i <= 300; i++) s.ingest(speech(`e${i}`, i, `第${i}句`));
    s.markEnded();
    await summarizeEndedMeeting(agentDeps(s, meeting).args);
    expect(mocks.startRunFlow).toHaveBeenCalledTimes(1);
    const call = mocks.startRunFlow.mock.calls[0]![0];
    expect(call.prompt).toContain('甲: 第1句\n');
    expect(call.prompt).toContain('甲: 第300句');
    expect(call.prompt).toContain('整场');
    expect(call.scopeId).toBe(meetingMinutesScopeId('m1'));
  });

  it('splits on line boundaries', () => {
    expect(splitMeetingRecord(['aaaa', 'bbbb', 'cccc'], 10)).toEqual([['aaaa', 'bbbb'], ['cccc']]);
    expect(splitMeetingRecord(['x'.repeat(30)], 10)[0]![0]).toHaveLength(10);
    expect(splitMeetingRecord([], 10)).toEqual([]);
  });

  it('summarizes a long meeting segment by segment, then merges', async () => {
    mocks.startRunFlow.mockImplementation(async (input: { prompt: string }) =>
      fakeRun(input.prompt.includes('分段摘要（按时间顺序）') ? MINUTES : '这一段的摘要'),
    );
    const meeting = cfg();
    const s = makeSession(meeting);
    const long = '很长的发言'.repeat(100); // ~500 chars per line
    const n = Math.ceil((MINUTES_SINGLE_PASS_CHARS * 1.5) / long.length);
    for (let i = 1; i <= n; i++) s.ingest(speech(`e${i}`, i, `${i}-${long}`));
    const { args, sent } = agentDeps(s, meeting);
    s.markEnded();
    await summarizeEndedMeeting(args);
    const prompts = mocks.startRunFlow.mock.calls.map((c) => String(c[0].prompt));
    expect(prompts.length).toBeGreaterThanOrEqual(3);
    const segs = prompts.slice(0, -1);
    expect(segs[0]).toContain(`这是第 1 段`);
    expect(segs[0]).toContain('1-很长的发言');
    expect(segs[segs.length - 1]).toContain(`${n}-很长的发言`);
    const merge = prompts[prompts.length - 1]!;
    expect(merge).toContain('【第 1 段');
    expect(merge).toContain('这一段的摘要');
    expect(merge).toContain('合并成整场会议纪要');
    expect(merge).not.toContain('1-很长的发言');
    expect(sent).toHaveLength(1);
    expect(JSON.stringify(sent[0]!.input.card)).toContain('发布');
  });

  it('a rejected minutes run is a failure for the owner, never posted as minutes', async () => {
    mocks.startRunFlow.mockResolvedValue({
      ok: false,
      rejectReason: { code: 'run-already-active', userVisible: '当前会话已有运行在执行' },
    });
    const meeting = cfg();
    const s = makeSession(meeting);
    s.ingest(speech('e1', 1, '一句话'));
    s.markEnded();
    const { args, sent } = agentDeps(s, meeting);
    await summarizeEndedMeeting(args);
    expect(sent).toHaveLength(1);
    expect(sent[0]!.to).toBe('ou_owner');
    expect(sent[0]!.input.markdown).toContain('生成纪要失败');
    expect(sent[0]!.input.card).toBeUndefined();
  });

  it('/meeting notes covers the whole meeting too', async () => {
    mocks.startRunFlow.mockResolvedValue(fakeRun('小结'));
    const meeting = cfg({ transcript: { ...MEETING_DEFAULTS.transcript, keep: 10 } });
    const s = makeSession(meeting);
    for (let i = 1; i <= 30; i++) s.ingest(speech(`e${i}`, i, `第${i}句`));
    const answer = await summarizeMeetingRecord(agentDeps(s, meeting).args, 'notes');
    expect(answer).toBe('小结');
    expect(String(mocks.startRunFlow.mock.calls[0]![0].prompt)).toContain('甲: 第1句\n');
  });

  it('the minutes card says where the full transcript is', async () => {
    mocks.startRunFlow.mockResolvedValue(fakeRun(MINUTES));
    const meeting = cfg();
    const s = makeSession(meeting);
    s.ingest(speech('e1', 1, '一句话'));
    s.markEnded();
    s.finalizeTranscript();
    const { args, sent } = agentDeps(s, meeting);
    await summarizeEndedMeeting(args);
    expect(JSON.stringify(sent[0]!.input.card)).toContain('全量逐字稿已存本机：meeting-reliability-');
  });
});

describe('join notice', () => {
  it('says who, that it records, how to ask and who may', () => {
    const profile = createDefaultProfileConfig({
      agentKind: 'claude',
      accounts: { app: { id: 'cli_test', secret: '${APP_SECRET}', tenant: 'feishu' } },
    });
    profile.meeting = cfg();
    const text = meetingJoinNotice(profile, '三儿');
    expect(text).toContain('我是「三儿」');
    expect(text).toContain('在记录');
    expect(text).toContain('「@三儿 你的问题」');
    expect(text).toContain('主人和授权名单');
    expect(text).toContain('散会后整理成纪要');
    profile.meeting = cfg({ access: 'admins-only', summaryOnEnd: false });
    expect(meetingJoinNotice(profile, '三儿')).toContain('只有机器人的主人和管理员');
    expect(meetingJoinNotice(profile, '三儿')).not.toContain('纪要');
    profile.mode = 'team';
    profile.meeting = cfg();
    const team = meetingJoinNotice(profile);
    expect(team).toContain('会里的每个人都可以问我');
    expect(team).not.toContain('不会执行');
    expect(team).toContain('我是 AI 会议助手');
  });

  it('is posted once on joining, not when a restart re-adopts the meeting, and can be turned off', async () => {
    const said: string[] = [];
    const client: VcRequestClient = {
      request: vi.fn(async (p: { url: string; data?: unknown }) => {
        if (p.url.endsWith('/message')) said.push(String((p.data as { content?: string } | undefined)?.content));
        return { code: 0, data: {} } as never;
      }),
    };
    const fresh = makeSession({}, client);
    attachMeetingAgent(agentDeps(fresh, cfg()).args);
    const resumed = makeSession({}, client);
    resumed.resumedAt = Date.now();
    attachMeetingAgent(agentDeps(resumed, cfg()).args);
    const off = makeSession({ announceOnJoin: false }, client);
    attachMeetingAgent(agentDeps(off, cfg({ announceOnJoin: false })).args);
    await new Promise((r) => setTimeout(r, 0));
    expect(said).toHaveLength(1);
    expect(said[0]).toContain('我是「三儿」');
  });
});

describe('review fixes', () => {
  it('the gap-fill poller never rolls a sentence back to an older revision', () => {
    const s = makeSession();
    const heard: string[] = [];
    s.on('transcript', (e) => heard.push(e.kind === 'transcript' ? e.text : ''));
    s.ingest(speech('e1', 1, '今天'));
    s.ingest(speech('e3', 1, '今天来讨论发布')); // e2 dropped by push
    s.ingest(speech('e2', 1, '今天来讨论'), { gapFill: true }); // the poller, 60 s behind
    expect(s.recentTranscript()).toEqual(['甲: 今天来讨论发布']);
    expect(heard).toEqual(['今天', '今天来讨论发布']);
    // A longer version from the gap-fill is news and is heard.
    s.ingest(speech('e4', 1, '今天来讨论发布的时间'), { gapFill: true });
    expect(s.recentTranscript()).toEqual(['甲: 今天来讨论发布的时间']);
    expect(heard).toHaveLength(3);
  });

  it('a chat message arriving on both lanes under different keys is stored and heard once', () => {
    const s = makeSession();
    const heard: string[] = [];
    s.on('chat', (e) => heard.push(e.kind === 'chat' ? e.content : ''));
    const noId = chat('x', '@bot 进度如何', '乙', 1_800_000_000);
    delete (noId as { event_id?: string }).event_id;
    s.ingest(noId); // push without event_id
    s.ingest(chat('real-id', '@bot 进度如何', '乙', 1_800_000_000), { gapFill: true });
    expect(heard).toEqual(['@bot 进度如何']);
    expect(s.recentContext()).toHaveLength(1);
  });

  it('the same text typed twice (two event ids) is two messages', () => {
    const s = makeSession();
    s.ingest(chat('c1', '+1', '乙', 1_800_000_000));
    s.ingest(chat('c2', '+1', '乙', 1_800_000_030));
    expect(s.recentContext()).toEqual(['乙（弹幕）: +1', '乙（弹幕）: +1']);
  });

  it('with stabilizeMs, a re-read does not cancel the newer sentence still settling', async () => {
    vi.useFakeTimers();
    const { client } = pagedClient([[speech('e1', 1, '今天')]]);
    const s = makeSession({ transcript: { ...MEETING_DEFAULTS.transcript, stabilizeMs: 800 } }, client);
    const heard: string[] = [];
    s.on('transcript', (e) => heard.push(e.kind === 'transcript' ? e.text : ''));
    s.ingest(speech('e2', 1, '今天来讨论'));
    await s.backfill('test');
    await vi.advanceTimersByTimeAsync(1000);
    expect(s.recentTranscript()).toEqual(['甲: 今天来讨论']);
    expect(heard).toEqual(['今天来讨论']);
  });

  it('a sentence still settling when the meeting ends makes it into the record', () => {
    vi.useFakeTimers();
    const s = makeSession({ transcript: { ...MEETING_DEFAULTS.transcript, stabilizeMs: 800 } });
    s.ingest(speech('e1', 1, '最后一句话'));
    s.markEnded();
    expect(s.recentTranscript()).toEqual(['甲: 最后一句话']);
  });

  it('joining the same meeting again after it was written out does not replay old chat', () => {
    const dir = join(root, 'transcripts');
    const first = makeSession({}, undefined, dir);
    first.ingest(chat('c1', '@bot 老问题'));
    first.ingest(speech('e1', 1, '今天'));
    first.ingest(speech('e2', 1, '今天来讨论'));
    first.markEnded();
    first.finalizeTranscript();

    const again = makeSession({}, undefined, dir, T0 + 3_600_000);
    const heard: string[] = [];
    again.on('chat', () => heard.push('chat'));
    // A new session's poller reads from the first page, non-quietly.
    again.ingest(chat('c1', '@bot 老问题'));
    again.ingest(speech('e1', 1, '今天'));
    expect(heard).toEqual([]);
    expect(again.recentContext()).toEqual(['乙（弹幕）: @bot 老问题', '甲: 今天来讨论']);
  });

  it('a torn last line from a crash does not swallow the next record', async () => {
    const dir = join(root, 'transcripts');
    const first = makeSession({}, undefined, dir);
    first.ingest(speech('e1', 1, '第一句'));
    first.dispose();
    const { appendFileSync } = await import('node:fs');
    appendFileSync(transcriptFile(dir, 'm1'), '{"k":"speech","sid":"2","who":"甲","te');
    const again = makeSession({}, undefined, dir, T0 + 3_600_000);
    again.ingest(speech('e3', 3, '重启后'));
    again.dispose();
    const lines = loadTranscript(transcriptFile(dir, 'm1')).records.filter((r) => r.k !== 'seen');
    expect(lines.map((r) => (r as { text: string }).text)).toEqual(['第一句', '重启后']);
  });

  it('status trusts the Open Platform over the "polled, no push" heuristic', async () => {
    vi.useFakeTimers();
    const appInfo = (url: string) =>
      url.includes('/application/v6/applications/')
        ? { code: 0, data: { app: { event: { subscribed_events: Object.values(VC_BOT_EVENTS) } } } }
        : undefined;
    const { client } = pagedClient([[speech('e1', 1, '有人说话')]], appInfo);
    const mgr = new MeetingManager({ client, config: () => cfg(), channel: { onRawEvent: () => () => {} }, appId: 'cli_x' });
    mgr.attachPush();
    const s = await mgr.join('123456789');
    await vi.advanceTimersByTimeAsync(130_000);
    const d = await mgr.diagnosePush();
    expect(d).toMatchObject({ state: 'subscribed-no-push', polled: 1 });
    expect(pushStatusLine(d)).toContain('已订阅');
    s.dispose();
  });

  it('a minutes run that errors is a failure, not half the minutes', async () => {
    mocks.startRunFlow.mockResolvedValue({
      ok: true,
      execution: {
        subscribe: () => ({
          async *[Symbol.asyncIterator]() {
            yield { type: 'text', delta: '**讨论要点**\n- 半截' };
            yield { type: 'error' };
          },
        }),
      },
    });
    const meeting = cfg();
    const s = makeSession(meeting);
    s.ingest(speech('e1', 1, '一句话'));
    s.markEnded();
    const { args, sent } = agentDeps(s, meeting);
    await summarizeEndedMeeting(args);
    expect(sent).toHaveLength(1);
    expect(sent[0]!.to).toBe('ou_owner');
    expect(sent[0]!.input.markdown).toContain('生成纪要失败');
  });

  it('leave keeps the meeting live during the last read, and a second leave is a no-op', async () => {
    let release!: () => void;
    const gate = new Promise<void>((r) => (release = r));
    const calls: string[] = [];
    const client: VcRequestClient = {
      request: vi.fn(async (p: { url: string }) => {
        calls.push(p.url.split('/').pop()!);
        if (p.url.endsWith('/join')) return { code: 0, data: { meeting: { id: 'm1' } } } as never;
        if (p.url.endsWith('/events')) {
          await gate;
          return { code: 0, data: { events: [], has_more: false } } as never;
        }
        return { code: 0, data: {} } as never;
      }),
    };
    const mgr = new MeetingManager({ client, config: () => cfg(), transcriptDir: join(root, 'transcripts') });
    const s = await mgr.join('123456789');
    s.dispose();
    calls.length = 0;
    const leaving = mgr.leave('m1');
    await new Promise((r) => setTimeout(r, 0));
    expect(mgr.get('m1')).toBe(s); // still routed while reading
    expect(await mgr.leave('m1')).toBe(false);
    release();
    expect(await leaving).toBe(true);
    expect(mgr.get('m1')).toBeUndefined();
    expect(calls).toEqual(['events', 'leave']);
  });
});
