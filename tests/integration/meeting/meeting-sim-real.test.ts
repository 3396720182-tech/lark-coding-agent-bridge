import { readFileSync, readdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { afterAll, describe, expect, it } from 'vitest';
import { ClaudeAdapter } from '../../../src/agent/claude/adapter';
import type { AgentAdapter, AgentEvent, AgentRun, AgentRunOptions } from '../../../src/agent/types';
import { meetingArchiveFile, type MinutesRecord } from '../../../src/meeting/minutes';
import { meetingDataDir, writeMeetingJson } from '../../../src/meeting/store';
import { at, cardButtons, createSim, MEETING_NO, OWNER, PEOPLE, textTo, type RunRecord, type Sim } from './sim-harness';

/**
 * The simulated meeting of `meeting-sim.test.ts`, answered by the real Claude
 * CLI instead of a script: does the model actually follow the prompts (cite
 * 〔时:分 发言人〕, split long answers, say "not mentioned", write 待确认, pass
 * with SKIP, convert 「下周三」 to a date)? Opt-in and slow (minutes, real
 * tokens): `MEETING_SIM_REAL=1 MEETING_SIM_REPORT=<file.md> pnpm vitest run
 * tests/integration/meeting/meeting-sim-real.test.ts`. Writes what happened to
 * the report file; the asserts only cover what must hold whatever the model says.
 */

const OWNER_DM = 'oc_owner_dm';
const LONG = 240_000;

/** Wraps an adapter to record each run's prompt, timing and text. */
class Recording implements AgentAdapter {
  readonly runs: RunRecord[] = [];
  constructor(private inner: AgentAdapter) {}
  get id() {
    return this.inner.id;
  }
  get displayName() {
    return this.inner.displayName;
  }
  isAvailable() {
    return this.inner.isAvailable();
  }
  setBotIdentity(i: Parameters<NonNullable<AgentAdapter['setBotIdentity']>>[0]) {
    this.inner.setBotIdentity?.(i);
  }
  run(opts: AgentRunOptions): AgentRun {
    const rec: RunRecord = { prompt: opts.prompt, startedAt: Date.now(), output: '' };
    this.runs.push(rec);
    const run = this.inner.run(opts);
    async function* tap(): AsyncIterable<AgentEvent> {
      for await (const e of run.events) {
        if (e.type === 'text') rec.output += e.delta;
        if (e.type === 'done' || e.type === 'error') rec.endedAt = Date.now();
        yield e;
      }
      rec.endedAt ??= Date.now();
    }
    return { ...run, runId: run.runId, events: tap(), stop: () => run.stop(), waitForExit: (ms: number) => run.waitForExit(ms) } as AgentRun;
  }
}

const report: string[] = [];
const section = (title: string, body: string) => report.push(`## ${title}\n\n${body.trim()}\n`);
const quote = (s: string) => s.split('\n').map((l) => `> ${l}`).join('\n');

let sim: Sim | undefined;
afterAll(async () => {
  const file = process.env.MEETING_SIM_REPORT;
  if (file && report.length) writeFileSync(file, `# 模拟会议实测（真模型）\n\n${report.join('\n')}`);
  await sim?.cleanup();
});

describe.runIf(process.env.MEETING_SIM_REAL === '1')('meeting agent batch 2/3 — simulated meeting, real Claude', () => {
  it('plays the meeting', async () => {
    const agent = new Recording(new ClaudeAdapter());
    sim = await createSim({
      agent,
      meeting: {
        proactive: { enabled: true, waitMs: 3000, maxPerMeeting: 1 },
        voiceWake: { enabled: true, names: ['小会'] },
      },
    });
    const s = sim;
    const t0 = Date.now();
    /** Markdown of the answer cards DMed to someone. */
    const dmText = (to: string) =>
      s.sent
        .filter((m) => m.to === to)
        .map((m) => m.markdown ?? JSON.stringify(m.card).match(/"content":"((?:[^"\\]|\\.)*)"/g)?.map((c) => JSON.parse(`{${c}}`).content).join('\n') ?? '')
        .join('\n---\n') || '（无）';
    const lastWeek = new Date(Date.now() - 7 * 86400_000);
    writeMeetingJson(meetingArchiveFile(meetingDataDir(s.controls), '20261120-1000-555555555'), {
      id: '20261120-1000-555555555',
      meetingId: 'm-old',
      meetingNo: '555555555',
      topic: '设计评审',
      endedAt: lastWeek.toISOString(),
      minutes: '**结论**\n- 首页改版用方案 B\n- 下周三前给出高保真稿\n**待办**\n1. 高保真稿 — 小林，下周三',
      todosParsed: true,
      todos: [{ task: '出首页高保真稿', owner: '小林', due: '下周三' }],
      participants: ['小林', '张三'],
    } satisfies MinutesRecord);

    await s.join();
    await s.waitFor(() => textTo(s, OWNER_DM).includes('已从日程'), 10_000, 'calendar docs');
    section('入会', `会中公告：\n\n${quote(s.meetingChat[0] ?? '（无）')}\n\n发给发起人的：\n\n${quote(textTo(s, OWNER_DM))}\n\n发给主人的：\n\n${quote(textTo(s, OWNER))}`);

    s.say('zhang', '今天主要定一下 2.4 版本的发布时间和分工。', at('14:30'));
    s.say('li', '资料里写的回滚条件是崩溃率超过千分之三，或者支付成功率掉一个点。', at('14:31'));
    s.say('zhang', '那发布就定在下周三晚上八点，李四负责发布前的回归测试，周二下班前给结论。', at('14:32'));
    s.say('wang', '监控面板我来看看吧，尽快弄。', at('14:33'));
    s.say('zhang', '运维那边谁去通知一下发布窗口的事？', at('14:34'));
    s.say('li', '这个回头再说吧。', at('14:34'));

    /** Meeting messages after `from` that are not the "收到，正在查…" line a slow typed ask gets. */
    const answersSince = (from: number) => s.meetingChat.slice(from).filter((m) => !m.startsWith('⏳'));

    // Two at once, one short, one that needs a long answer.
    const n0 = s.meetingChat.length;
    s.chat('zhang', '@小会 刚才定的发布时间和负责人是什么？');
    s.chat('li', '@小会 详细列一下出问题时的回滚流程，分步骤，越细越好');
    await s.waitFor(() => answersSince(n0).length >= 2, LONG, 'two answers');
    const asks = agent.runs.filter((r) => r.prompt.includes('在会中向你提问'));
    const overlap = Math.max(asks[0]!.startedAt, asks[1]!.startedAt) < Math.min(asks[0]!.endedAt ?? Infinity, asks[1]!.endedAt ?? Infinity);
    section(
      '两人同时问',
      `两个问题同时在跑：${overlap ? '是' : '否'}（各自用时 ${asks.map((r) => `${(((r.endedAt ?? Date.now()) - r.startedAt) / 1000).toFixed(0)} 秒`).join('、')}）\n\n` +
        `会中消息：\n\n${s.meetingChat.slice(n0).map(quote).join('\n\n')}\n\n` +
        `张三收到的私聊：\n\n${quote(dmText(PEOPLE.zhang.id))}\n\n李四收到的私聊：\n\n${quote(dmText(PEOPLE.li.id))}`,
    );
    expect(overlap).toBe(true);

    const n1 = s.meetingChat.length;
    s.chat('zhang', '@小会 这次发布的预算是多少？');
    await s.waitFor(() => answersSince(n1).length >= 1, LONG, 'budget');
    section('问资料和会上都没有的', quote(s.meetingChat.slice(n1).join('\n')));

    const n2 = s.meetingChat.length;
    s.say('wang', '小会，帮我查一下回滚负责人是谁', at('14:40'));
    s.say('li', '小会，回滚负责人是谁？', at('14:41'));
    await s.waitFor(() => s.meetingChat.length >= n2 + 2, LONG, 'voice');
    section('语音点名（王五不在名单，李四在）', s.meetingChat.slice(n2).map(quote).join('\n\n'));

    const n3 = s.meetingChat.length;
    const runsBefore = agent.runs.length;
    // Let the echo of the bot's last answer arrive first: a bot message after the question counts as "answered".
    await new Promise((r) => setTimeout(r, 1000));
    s.say('zhang', '灰度比例一开始开多少来着？', at('14:43'));
    // Before the fix a question to the bot by name had used up the one interjection, and no run came.
    await s
      .waitFor(() => agent.runs.length > runsBefore && agent.runs.slice(runsBefore).every((r) => r.endedAt), 90_000, 'proactive run')
      .catch(() => undefined);
    await new Promise((r) => setTimeout(r, 500));
    const earlier = s.meetingChat.slice(0, n3).filter((m) => m.startsWith('💡'));
    if (earlier.length) section('（问题）插话额度被提前用掉', earlier.map(quote).join('\n\n'));
    const pro = agent.runs.slice(runsBefore).find((r) => r.prompt.includes('只输出 SKIP'));
    section(
      '没人答的问题（资料里有答案）',
      `模型原话：\n\n${quote(pro?.output ?? '（没有运行）')}\n\n会中消息：\n\n${quote(s.meetingChat.slice(n3).join('\n') || '（没插话）')}`,
    );

    await s.waitFor(() => agent.runs.every((r) => r.endedAt), LONG, 'all runs');
    await s.end();
    const card = s.sent.find((m) => m.to === OWNER_DM && m.card);
    const dir = meetingDataDir(s.controls);
    const archiveFile = readdirSync(join(dir, 'archive')).find((f) => f.includes(MEETING_NO))!;
    const rec = JSON.parse(readFileSync(join(dir, 'archive', archiveFile), 'utf8')) as MinutesRecord;
    section(
      '纪要',
      `${rec.minutes}\n\n待办（存档里的结构化结果）：\n\n` +
        rec.todos.map((t, i) => `${i + 1}. ${t.task} · 负责人「${t.owner || '待确认'}」${t.ownerId ? `（${t.ownerId}）` : ''} · 截止「${t.due || '待确认'}」${t.dueDate ? ` → ${t.dueDate}` : ''}`).join('\n') +
        `\n\n卡片按钮：${cardButtons(s).map((b) => b.text).join('、')}`,
    );
    expect(card).toBeDefined();

    const taskButtons = cardButtons(s).filter((b) => b.text === '建成飞书任务');
    for (const b of taskButtons) await s.click(b.value);
    const bodies = s.feishu.filter((c) => c.url === '/open-apis/task/v2/tasks').map((c) => c.data as Record<string, unknown>);
    section(
      '点「建成飞书任务」',
      bodies
        .map((b) => {
          const due = b.due as { timestamp: string } | undefined;
          return `- ${String(b.summary)} · 成员 ${JSON.stringify(b.members ?? [])} · 截止 ${due ? new Date(Number(due.timestamp)).toLocaleString('zh-CN') : '无'}\n  描述：${String(b.description).replace(/\n/g, ' / ')}`;
        })
        .join('\n') + `\n\n回复：\n\n${quote(s.sent.filter((m) => m.to === OWNER_DM && m.markdown?.includes('飞书任务')).map((m) => m.markdown).join('\n'))}`,
    );

    const h0 = s.sent.length;
    expect(await s.dm(OWNER, '上周那个会定了什么')).toBe(true);
    await s.waitFor(() => s.sent.slice(h0).some((m) => m.markdown?.includes('查的是本机存档')), LONG, 'history');
    const h1 = s.sent.length;
    expect(await s.dm(OWNER, '刚才那个会定了什么')).toBe(true);
    await s.waitFor(() => s.sent.slice(h1).some((m) => m.markdown?.includes('查的是本机存档')), LONG, 'history 2');
    const notIntercepted = await s.dm(OWNER, '帮我把今天的会议纪要翻译成英文');
    section(
      '私聊问过去的会',
      `「上周那个会定了什么」：\n\n${quote(s.sent.slice(h0, h1).map((m) => m.markdown).join('\n'))}\n\n` +
        `「刚才那个会定了什么」：\n\n${quote(s.sent.slice(h1).map((m) => m.markdown).join('\n'))}\n\n` +
        `「帮我把今天的会议纪要翻译成英文」被当成查纪要：${notIntercepted ? '是（不对）' : '否（照常交给 agent）'}`,
    );
    expect(notIntercepted).toBe(false);
    section('用时', `整场 ${((Date.now() - t0) / 1000).toFixed(0)} 秒，agent 共运行 ${agent.runs.length} 次。`);
  }, 1_800_000);
});
