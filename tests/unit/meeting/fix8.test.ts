import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

/**
 * The eight fixes found by the T0099 edge-case probe: the bot knowing its own
 * earlier answers, long meetings, homophone / mid-sentence wake names, todos,
 * the interjection wording, "刚才那个会", typed-ask feedback, and small talk
 * not costing a run.
 */

const mocks = vi.hoisted(() => ({ startRunFlow: vi.fn() }));
vi.mock('../../../src/bot/run-flow', () => ({ startRunFlow: mocks.startRunFlow }));

const { answerInMeeting, attachMeetingAgent, buildMeetingPrompt, relevantEarlierLines, TYPED_ACK_MS } = await import(
  '../../../src/meeting/orchestrator'
);
const { MeetingSession } = await import('../../../src/meeting/session');
const { MEETING_DEFAULTS, createDefaultProfileConfig } = await import('../../../src/config/profile-schema');
const { meetingDataDir, writeMeetingJson } = await import('../../../src/meeting/store');
const { bindDoc, docSnapshotFile } = await import('../../../src/meeting/docs');
const { meetingMinutesQuestion } = await import('../../../src/meeting/minutes');
const { matchWakeName, VOICE_SETTLE_MS } = await import('../../../src/meeting/voice');
const { buildProactivePrompt, looksLikeQuestion, tidyInterjection } = await import('../../../src/meeting/proactive');
const { pickArchives } = await import('../../../src/meeting/history');
const { sameSound } = await import('../../../src/meeting/pinyin');

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

let root: string;
beforeEach(() => {
  vi.clearAllMocks();
  root = mkdtempSync(join(tmpdir(), 'meeting-fix8-'));
});
afterEach(() => {
  vi.useRealTimers();
  rmSync(root, { recursive: true, force: true });
});

const T0 = new Date(2026, 10, 27, 14, 0, 0).getTime();
const MIN = 60_000;

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

function setup(over: Partial<MeetingConfig> = {}) {
  const meeting = cfg(over);
  const profileConfig = createDefaultProfileConfig({
    agentKind: 'claude',
    accounts: { app: { id: 'cli_test', secret: '${APP_SECRET}', tenant: 'feishu' } },
  });
  profileConfig.meeting = meeting;
  profileConfig.access.allowedUsers = ['ou_zhang', 'ou_li'];
  const session = new MeetingSession({
    client: { request: vi.fn(async () => ({ code: 0, data: {} }) as never) },
    meetingId: 'm1',
    meetingNo: '123456789',
    topic: '周会',
    config: meeting,
    originChatId: 'oc_team',
    now: () => T0,
  });
  const meetingMsgs: string[] = [];
  vi.spyOn(session, 'sendMessage').mockImplementation(async (t: string) => {
    meetingMsgs.push(t);
  });
  const args = {
    session,
    channel: {
      botIdentity: { name: '小会', openId: 'ou_bot' },
      send: vi.fn(async () => ({ messageId: 'om_1' })),
    },
    controls: { configPath: join(root, 'config.json'), profile: 'claude', profileConfig, botOwnerId: 'ou_owner' },
    executor: {},
    activeRuns: { interrupt: vi.fn(), scopes: () => [] },
    sessions: {},
    workspaces: {},
  } as never;
  return { session, args, meetingMsgs };
}

const promptOf = (i = 0) => String(mocks.startRunFlow.mock.calls[i]![0].prompt);
const flush = () => new Promise((r) => setTimeout(r, 0));

// ── 1. The bot knows what it said ──

describe('#1 the bot recognizes its own earlier answers', () => {
  it('the prompt says which name in the record is the bot', () => {
    const p = buildMeetingPrompt({ question: '你刚才说的时间是谁定的', transcript: ['[14:20:00] 小会（弹幕）: 下周四'], botName: '小会' });
    expect(p).toContain('「小会」');
    expect(p).toMatch(/「小会」发的.*就是你/);
  });

  it('an in-meeting answer is told the bot name; so is the interjection run', async () => {
    mocks.startRunFlow.mockResolvedValue(fakeRun('是张三定的〔14:20 张三〕'));
    const { session, args } = setup();
    session.ingest(speech('e1', 1, '发布改到下周四', '张三', 'ou_zhang', T0 + 20 * MIN));
    session.ingest(chat('c1', '下周四晚上八点〔14:20 张三〕', '小会', 'ou_bot', T0 + 21 * MIN));
    await answerInMeeting(args, '你刚才说的那个时间，是谁拍板的？', { deliver: 'broadcast', askerId: 'ou_zhang', askedBy: '张三' });
    expect(promptOf()).toMatch(/「小会」发的.*就是你/);
    expect(buildProactivePrompt({ question: 'q', recent: [], materials: [], botName: '小会' })).toContain('「小会」');
  });

  it('no name known: no such line', () => {
    expect(buildMeetingPrompt({ question: 'q', transcript: [] })).not.toContain('就是你之前');
  });
});

// ── 2. Long meetings ──

describe('#2 a meeting longer than the in-meeting window', () => {
  it('① the prompt says the record is only the latest part, and not to claim nobody said it', () => {
    const p = buildMeetingPrompt({ question: '预算多少', transcript: ['[14:30:00] 李四: 好'], earlierOmitted: true });
    expect(p).toContain('不是从开会开始');
    expect(p).toContain('最近的记录里没有');
    expect(p).not.toContain('直接说「会上都没有提到」');
    // The whole record (minutes) or a short meeting keeps the plain rule.
    expect(buildMeetingPrompt({ question: 'q', transcript: [] })).toContain('会上都没有提到');
  });

  it('② picks the earlier lines that share words with the question, in meeting order', () => {
    const earlier = [
      '[14:00:00] 张三: 先说预算：这次发布活动预算定 8 万，超了要找财务审批。',
      '[14:01:00] 张三: 发布先定下周三晚上八点。',
      '[14:02:00] 李四: 回归用例这次加了支付链路的二十个场景。',
      '[14:03:00] 王五: 监控面板我把崩溃率放到首屏。',
    ];
    const picked = relevantEarlierLines(earlier, '会议一开始说的活动预算是多少？');
    expect(picked[0]).toContain('预算定 8 万');
    expect(picked.join('\n')).not.toContain('监控面板');
    expect(relevantEarlierLines(earlier, '？？')).toEqual([]);
  });

  it('an in-meeting answer in a long meeting gets the relevant earlier lines, and citing them is verified', async () => {
    mocks.startRunFlow.mockResolvedValue(fakeRun('预算是 8 万〔14:00 张三〕'));
    const { session, args, meetingMsgs } = setup({ transcript: { keep: 10, stabilizeMs: 0, save: false } });
    session.ingest(speech('e0', 0, '先说预算：这次发布活动预算定 8 万，超了要找财务审批。', '张三', 'ou_zhang', T0));
    for (let i = 1; i <= 14; i++) {
      session.ingest(speech(`e${i}`, i, `第 ${i} 件事：回归测试和监控都在推进。`, '李四', 'ou_li', T0 + i * MIN));
    }
    await answerInMeeting(args, '会议一开始说的活动预算是多少？', { deliver: 'broadcast', askerId: 'ou_li', askedBy: '李四' });
    const p = promptOf();
    expect(p).toContain('不是从开会开始');
    expect(p).toContain('[14:00:00] 张三: 先说预算');
    expect(meetingMsgs[0]).toBe('预算是 8 万〔14:00 张三〕'); // not marked 未核对到原话
  });

  it('a short meeting gives everything and no earlier block', async () => {
    mocks.startRunFlow.mockResolvedValue(fakeRun('8 万'));
    const { session, args } = setup();
    session.ingest(speech('e0', 0, '预算定 8 万'));
    await answerInMeeting(args, '预算多少？', { deliver: 'broadcast', askerId: 'ou_li' });
    expect(promptOf()).not.toContain('不是从开会开始');
    expect(promptOf()).not.toContain('较早');
  });
});

// ── 3. Wake names as speech recognition writes them ──

describe('#3 voice wake: homophones and the name mid-sentence', () => {
  it('same-sound characters match, tones ignored', () => {
    expect(sameSound('会', '惠')).toBe(true);
    expect(sameSound('小', '晓')).toBe(true);
    expect(sameSound('会', '黑')).toBe(false);
    expect(sameSound('a', 'a')).toBe(true);
  });

  it('a name written with other characters of the same sound still wakes', () => {
    expect(matchWakeName('小惠，回滚条件是什么？', ['小会'])).toBe('回滚条件是什么？');
    expect(matchWakeName('晓会，回滚条件是什么？', ['小会'])).toBe('回滚条件是什么？');
    expect(matchWakeName('小黑，回滚条件是什么？', ['小会'])).toBeUndefined();
  });

  it('the name in mid-sentence counts when it is clearly addressed', () => {
    expect(matchWakeName('这个问题小会你查一下，回滚条件是什么？', ['小会'])).toBe('这个问题你查一下，回滚条件是什么？');
    expect(matchWakeName('回滚那块，小惠，你看下条件是什么', ['小会'])).toBe('回滚那块，你看下条件是什么');
    // The everyday word is not a call.
    expect(matchWakeName('我们待会开个小会讨论一下', ['小会'])).toBeUndefined();
    expect(matchWakeName('开完小会，再对一下回滚条件', ['小会'])).toBeUndefined();
    expect(matchWakeName('上次小会你也在吧', ['小会'])).toBeUndefined();
  });

  it('a homophone call is answered by voice and never used as the interjection', async () => {
    vi.useFakeTimers();
    mocks.startRunFlow.mockResolvedValue(fakeRun('崩溃率超 0.3% 就回滚'));
    const { session, args, meetingMsgs } = setup({
      voiceWake: { enabled: true, names: ['小会'] },
      proactive: { enabled: true, waitMs: 3000, maxPerMeeting: 1 },
    });
    attachMeetingAgent(args);
    session.ingest(speech('e1', 1, '小惠，回滚条件是什么？', '李四', 'ou_li'));
    await vi.advanceTimersByTimeAsync(VOICE_SETTLE_MS + 3010);
    expect(meetingMsgs[0]).toContain('听到 李四 叫我');
    expect(mocks.startRunFlow.mock.calls.some((c) => String(c[0].prompt).includes('SKIP'))).toBe(false);
  });

  it('a same-sound spelling that is someone in the meeting calls that person, not the bot', async () => {
    vi.useFakeTimers();
    mocks.startRunFlow.mockResolvedValue(fakeRun('ok'));
    const { session, args, meetingMsgs } = setup({ voiceWake: { enabled: true, names: ['小会'] } });
    attachMeetingAgent(args);
    session.ingest(speech('e0', 0, '我这边测试都过了。', '小慧', 'ou_xh'));
    session.ingest(speech('e1', 1, '小慧，你那边进度怎么样？', '李四', 'ou_li'));
    await vi.advanceTimersByTimeAsync(VOICE_SETTLE_MS + 10);
    expect(meetingMsgs).toEqual([]);
    // The name as configured still wakes.
    session.ingest(speech('e2', 2, '小会，回滚条件是什么？', '李四', 'ou_li'));
    await vi.advanceTimersByTimeAsync(VOICE_SETTLE_MS + 10);
    expect(meetingMsgs[0]).toContain('听到 李四 叫我');
  });

  it('a sentence that opens like the name but is not it is no interjection candidate either', async () => {
    vi.useFakeTimers();
    mocks.startRunFlow.mockResolvedValue(fakeRun('SKIP'));
    const { session, args } = setup({
      voiceWake: { enabled: true, names: ['小会'] },
      proactive: { enabled: true, waitMs: 3000, maxPerMeeting: 1 },
    });
    bindTestDoc(args);
    attachMeetingAgent(args);
    session.ingest(speech('e1', 1, '小黑，回滚条件是什么？', '李四', 'ou_li'));
    await vi.advanceTimersByTimeAsync(VOICE_SETTLE_MS + 3010);
    expect(mocks.startRunFlow).not.toHaveBeenCalled();
  });
});

function bindTestDoc(args: unknown) {
  const controls = (args as { controls: { configPath: string; profile: string } }).controls;
  const dir = meetingDataDir(controls);
  const ref = {
    token: 'doxcn1',
    url: 'https://x.feishu.cn/docx/doxcn1',
    title: '发布手册',
    revision: 3,
    fetchedAt: new Date(T0).toISOString(),
    chunks: 1,
    chars: 40,
    boundAt: new Date(T0).toISOString(),
  };
  writeMeetingJson(docSnapshotFile(dir, ref.token, ref.revision), {
    ...ref,
    truncated: false,
    chunks: [{ i: 1, text: '回滚条件：崩溃率超过 0.3% 立即回滚。发布窗口是每周三晚上八点。' }],
  });
  bindDoc(dir, { meetingId: 'm1', meetingNo: '123456789' }, ref);
}

// ── 4. Todos ──

describe('#4 todos: only things to do, the "我来" speaker owns it, departments to confirm', () => {
  it('the minutes ask spells out the three rules', () => {
    for (const q of [meetingMinutesQuestion(false, '2026-11-27'), meetingMinutesQuestion(true)]) {
      expect(q).toContain('进度汇报');
      expect(q).toMatch(/「我来」「我把」/);
      expect(q).toContain('部门');
    }
  });
});

// ── 5. Interjection wording ──

describe('#5 the interjection does not say "资料里" twice', () => {
  it('drops the model\'s own lead-in', () => {
    expect(tidyInterjection('资料里写的是：每周三晚上八点〔《发布手册》r3 §1〕')).toBe('每周三晚上八点〔《发布手册》r3 §1〕');
    expect(tidyInterjection('会前资料显示，每周三〔x〕')).toBe('每周三〔x〕');
    expect(tidyInterjection('根据资料，每周三〔x〕')).toBe('每周三〔x〕');
    expect(tidyInterjection('每周三〔x〕')).toBe('每周三〔x〕');
  });

  it('the posted line reads once', async () => {
    vi.useFakeTimers();
    mocks.startRunFlow.mockResolvedValue(fakeRun('资料里写的是：每周三晚上八点〔《发布手册》r3 §1〕'));
    const { session, args, meetingMsgs } = setup({ proactive: { enabled: true, waitMs: 3000, maxPerMeeting: 1 } });
    bindTestDoc(args);
    attachMeetingAgent(args);
    session.ingest(speech('e1', 1, '发布窗口是什么时候？'));
    await vi.advanceTimersByTimeAsync(3010);
    expect(meetingMsgs).toEqual(['💡 刚才 张三 问「发布窗口是什么时候？」，会前资料里有：每周三晚上八点〔《发布手册》r3 §1〕']);
  });
});

// ── 6. "刚才那个会" ──

describe('#6 "刚才那个会" is the latest meeting only', () => {
  const NOW = new Date(2026, 10, 27, 16, 0, 0).getTime();
  const rec = (id: string, endedAt: Date, topic: string): MinutesRecord => ({
    id,
    meetingId: id,
    meetingNo: id.slice(-9),
    topic,
    endedAt: endedAt.toISOString(),
    minutes: '定了周五发布',
    todosParsed: true,
    todos: [],
  });
  const all = [
    rec('20261127-1500-111111111', new Date(2026, 10, 27, 15), '发布评审周会'),
    rec('20261120-1500-222222222', new Date(2026, 10, 20, 15), '发布评审周会'),
  ];

  it('刚才 / 刚刚 → one meeting, the newest', () => {
    for (const q of ['刚才那个会定了什么', '刚刚的会说了什么', '刚开完的会结论是什么']) {
      expect(pickArchives(all, q, NOW).picked.map((r) => r.id)).toEqual(['20261127-1500-111111111']);
    }
    // Without it, the window still holds both.
    expect(pickArchives(all, '那个会定了什么', NOW).picked).toHaveLength(2);
  });
});

// ── 7. Feedback while a typed ask runs ──

describe('#7 a typed ask that takes a while gets a "收到" first', () => {
  it('after the delay with no answer yet, one line in the meeting; then the answer', async () => {
    vi.useFakeTimers();
    let release!: () => void;
    const gate = new Promise<void>((r) => (release = r));
    mocks.startRunFlow.mockImplementation(async () => {
      await gate;
      return fakeRun('周四晚上八点');
    });
    const { session, args, meetingMsgs } = setup();
    attachMeetingAgent(args);
    session.ingest(chat('c1', '@小会 最后定的发布时间是什么时候？', '张三', 'ou_zhang'));
    await vi.advanceTimersByTimeAsync(TYPED_ACK_MS - 100);
    expect(meetingMsgs).toEqual([]);
    await vi.advanceTimersByTimeAsync(200);
    expect(meetingMsgs).toEqual(['⏳ 收到 张三 的问题，正在查…']);
    release();
    await vi.advanceTimersByTimeAsync(10);
    expect(meetingMsgs).toEqual(['⏳ 收到 张三 的问题，正在查…', '周四晚上八点']);
  });

  it('a quick answer comes alone', async () => {
    mocks.startRunFlow.mockResolvedValue(fakeRun('周四晚上八点'));
    const { session, args, meetingMsgs } = setup();
    attachMeetingAgent(args);
    session.ingest(chat('c1', '@小会 发布时间？', '张三', 'ou_zhang'));
    await flush();
    await flush();
    await new Promise((r) => setTimeout(r, 20));
    expect(meetingMsgs).toEqual(['周四晚上八点']);
  });
});

// ── 8. Small talk ──

describe('#8 small-talk questions are not interjection candidates', () => {
  it('room questions and tag questions are not questions for the docs', () => {
    for (const q of [
      '大家还有别的问题吗？',
      '还有没有其他补充？',
      '大家觉得怎么样？',
      '就这么定了，对吧？',
      '周四发，行不行？',
      '能听到我说话吗？',
      '这样可以吗？',
    ]) {
      expect(looksLikeQuestion(q), q).toBe(false);
    }
    for (const q of ['灰度比例一开始开多少来着？', '回滚条件是什么？', '这个能周三发吗', '大家知道回滚负责人是谁吗？']) {
      expect(looksLikeQuestion(q), q).toBe(true);
    }
  });
});

// ── Review findings ──

describe('review: wake names do not fire on everyday words', () => {
  it('a near-miss needs every other character to start with the same consonant', async () => {
    const { looksLikeWakeAttempt } = await import('../../../src/meeting/voice');
    expect(looksLikeWakeAttempt('小黑，回滚负责人是谁？', ['小会'])).toBe(true);
    for (const t of ['不会，回滚条件是什么？', '小李，上线时间定了吗？', '小程序，上线时间定了吗？', '销量，上个月是多少？', '一会，谁负责灰度？', '效果，怎么样？']) {
      expect(looksLikeWakeAttempt(t, ['小会']), t).toBe(false);
    }
  });

  it('a same-sound spelling only counts when followed by a pause or "你…"', () => {
    for (const t of ['销毁数据的流程是什么？', '校徽设计稿谁在跟？', '小汇报一下进度', '小回顾一下', '小慧那边进度怎么样？', '晓辉明天来吗']) {
      expect(matchWakeName(t, ['小会']), t).toBeUndefined();
    }
    expect(matchWakeName('小惠你查一下回滚条件', ['小会'])).toBe('你查一下回滚条件');
    expect(matchWakeName('小惠。', ['小会'])).toBe('');
  });

  it('in mid-sentence: "你/您" plus a verb after it, and not "那个/开个 小会"', () => {
    for (const t of ['明天那个小会你帮我约一下会议室', '这个小会帮我理清了思路', '例会，小会，复盘都要准备', '下午开个小会你看行吗']) {
      expect(matchWakeName(t, ['小会']), t).toBeUndefined();
    }
    expect(matchWakeName('这个问题小会你查一下，回滚条件是什么？', ['小会'])).toBe('这个问题你查一下，回滚条件是什么？');
  });
});

describe('review: small-talk filter keeps factual questions', () => {
  it('questions that only look like small talk are still questions', () => {
    for (const q of [
      '这个日志怎么看？',
      '监控面板怎么看崩溃率？',
      '现在还有问题的工单有多少？',
      '上线日期是周五对吧？',
      '回滚条件是错误率超过5%是吧？',
      '大家都知道上线日期是哪天吗？',
      '大家看一下预算是多少？',
    ]) {
      expect(looksLikeQuestion(q), q).toBe(true);
    }
    for (const q of ['这个方案大家怎么看？', '还有问题吗？']) expect(looksLikeQuestion(q), q).toBe(false);
  });
});

describe('review: "刚结束 / 刚散会 / 上一场" do not match minutes by their own words', () => {
  it('picks the newest', () => {
    const NOW = new Date(2026, 10, 27, 16, 0, 0).getTime();
    const mk = (id: string, endedAt: Date, minutes: string): MinutesRecord => ({
      id, meetingId: id, meetingNo: id.slice(-9), topic: '周会', endedAt: endedAt.toISOString(), minutes, todosParsed: true, todos: [],
    });
    const all = [
      mk('20261127-1500-111111111', new Date(2026, 10, 27, 15), '定了周五发布'),
      mk('20261126-1500-222222222', new Date(2026, 10, 26, 15), '项目结束时间待定，上一版方案作废，散会'),
    ];
    for (const q of ['刚结束的会定了什么', '上一场会说了什么', '刚散会的那个会结论是什么']) {
      expect(pickArchives(all, q, NOW).picked.map((r) => r.id), q).toEqual(['20261127-1500-111111111']);
    }
  });
});

describe('review: the interjection keeps answers that merely start with 资料', () => {
  it('only a lead-in with a verb or a pause is dropped', () => {
    for (const t of ['资料库里的地址是 x', '资料说明了回滚条件', '资料中心的负责人是张三', '资料上线时间是周五']) {
      expect(tidyInterjection(t), t).toBe(t);
    }
    expect(tidyInterjection('资料里写的是：每周三')).toBe('每周三');
    expect(tidyInterjection('资料，每周三')).toBe('每周三');
  });
});

describe('review: earlier lines by speaker, without breaking words', () => {
  it('a question naming a speaker finds what they said', () => {
    const earlier = ['[14:00:00] 张三: 预算定 8 万。', '[14:01:00] 李四: 回归要两天。', '[14:02:00] 王五: 我看下会议室。'];
    expect(relevantEarlierLines(earlier, '张三一开始说了什么？')[0]).toContain('张三');
    expect(relevantEarlierLines(earlier, '会议室订了吗？').join('\n')).toContain('会议室');
  });
});

describe('review: the "收到" line never lands after the answer', () => {
  it('the answer waits for an acknowledgement still being sent', async () => {
    vi.useFakeTimers();
    let release!: () => void;
    const gate = new Promise<void>((r) => (release = r));
    mocks.startRunFlow.mockImplementation(async () => {
      await gate;
      return fakeRun('周四');
    });
    const { session, args } = setup();
    const order: string[] = [];
    let finishAck!: () => void;
    vi.spyOn(session, 'sendMessage').mockImplementation(async (t: string) => {
      if (t.startsWith('⏳')) await new Promise<void>((r) => (finishAck = r));
      order.push(t);
    });
    attachMeetingAgent(args);
    session.ingest(chat('c1', '@小会 发布时间？', '张三', 'ou_zhang'));
    await vi.advanceTimersByTimeAsync(TYPED_ACK_MS + 10); // ack started, not delivered yet
    release();
    await vi.advanceTimersByTimeAsync(10);
    finishAck();
    await vi.advanceTimersByTimeAsync(10);
    expect(order).toEqual(['⏳ 收到 张三 的问题，正在查…', '周四']);
  });
});

describe('asking again right after an answer', () => {
  const busy = { ok: false as const, rejectReason: { code: 'run-already-active', userVisible: '当前会话已有运行在执行' } };

  it('the run that just answered still holding the scope a moment is waited out, not reported as "上一个问题还在回答"', async () => {
    mocks.startRunFlow.mockResolvedValueOnce(busy).mockResolvedValueOnce(busy).mockResolvedValue(fakeRun('周四'));
    const { session, args, meetingMsgs } = setup();
    attachMeetingAgent(args);
    session.ingest(chat('c1', '@小会 发布时间再确认一下？', '张三', 'ou_zhang'));
    await new Promise((r) => setTimeout(r, 600));
    expect(meetingMsgs).toEqual(['周四']);
  });

  it('a second question while the first is really running is told to wait at once', async () => {
    let release!: () => void;
    const gate = new Promise<void>((r) => (release = r));
    mocks.startRunFlow.mockImplementation(async () => {
      await gate;
      return fakeRun('周四');
    });
    const { session, args, meetingMsgs } = setup();
    attachMeetingAgent(args);
    session.ingest(chat('c1', '@小会 发布时间再确认一下？', '张三', 'ou_zhang'));
    session.ingest(chat('c2', '@小会 还有别的吗？', '张三', 'ou_zhang'));
    await flush();
    await flush();
    expect(mocks.startRunFlow).toHaveBeenCalledTimes(1);
    expect(meetingMsgs).toEqual(['你上一个问题还在回答，答完再问；发「@小会 stop」可以中断它。']);
    release();
    await new Promise((r) => setTimeout(r, 20));
    expect(meetingMsgs[1]).toBe('周四');
  });
});
