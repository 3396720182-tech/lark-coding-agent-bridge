import { readFileSync, readdirSync } from 'node:fs';
import { join } from 'node:path';
import { afterEach, describe, expect, it } from 'vitest';
import { meetingArchiveFile, type MinutesRecord } from '../../../src/meeting/minutes';
import { meetingDataDir, writeMeetingJson } from '../../../src/meeting/store';
import {
  at,
  cardButtons,
  createSim,
  DOC,
  MEETING_NO,
  OWNER,
  PEOPLE,
  ScriptedAgent,
  textTo,
  TOOL_CALL,
  TOPIC,
  type Sim,
} from './sim-harness';

/**
 * Batch 2/3 of the meeting agent, end to end in a simulated meeting: the real
 * bridge wiring (manager → orchestrator → RunExecutor/process pool → command
 * layer → card dispatcher) against a fake Feishu, with a scripted agent so the
 * checks are deterministic. `meeting-sim-real.test.ts` runs the same meeting
 * against the real Claude CLI.
 */

const OWNER_DM = 'oc_owner_dm';
const LONG_TAIL =
  '回滚分三步：第一步，值班后端在监控面板确认崩溃率或支付成功率触发阈值；第二步，在发布平台选择上一个稳定版本一键回滚，并把灰度比例降到 0；' +
  '第三步，在群里同步回滚原因和影响范围，测试负责人赵敏复核核心链路。回滚后 30 分钟内持续观察，确认指标恢复再关闭事故单。' +
  '如果是支付相关问题，还要通知财务对账，必要时冻结当天的退款操作，等对账完成后再恢复。' +
  '回滚完成后的第二天上午，由发布负责人组织一次复盘，写清触发原因、发现时间、处理时长和后续改进项，复盘文档挂到发布计划下面，下一次发布评审时逐条过一遍改进项是否落实。' +
  '另外，回滚期间客服要提前拿到统一口径，对外只说「服务优化中」，不透露具体故障细节；涉及大客户的，由客户成功经理逐个电话说明情况，并在当天下班前把沟通结果汇总到发布群里。';

/** The question line of an in-meeting prompt. */
const askedOf = (prompt: string): string => /(?:向你提问|收到的请求)[^\n]*：\n(.*)/.exec(prompt)?.[1] ?? '';

/** Answers like a model that follows the prompts; slow enough that overlapping runs show. */
function scripted(): ScriptedAgent {
  return new ScriptedAgent(
    (prompt) => {
      if (prompt.includes('```json 代码块')) {
        return [
          '**讨论要点**',
          '- 2.4 版本发布时间和回归测试安排〔14:32 张三〕',
          '**结论**',
          '- 周五晚上八点发布',
          '**待办**',
          '1. 发布前回归测试 — 李四，周五前',
          '2. 看监控面板 — 王五，待确认',
          '3. 通知运维组确认窗口 — 待确认',
          '```json',
          JSON.stringify({
            todos: [
              { task: '发布前回归测试', owner: '李四', due: '周五', due_date: '2026-12-04' },
              { task: '检查监控面板', owner: '王五', due: '尽快', due_date: '' },
              { task: '通知运维组确认发布窗口', owner: '', due: '', due_date: '' },
            ],
          }),
          '```',
        ].join('\n');
      }
      if (prompt.includes('只输出 SKIP')) {
        // Cite the chunk the bridge gave that has the answer.
        const label = /〔(《发布计划》r7 §\d+)〕\n[^〔]*灰度/.exec(prompt)?.[1];
        return label ? `一开始灰度先开 10%〔${label}〕。` : 'SKIP';
      }
      if (prompt.includes('私聊里有人问')) {
        const label = /=== 〔([^〕]+)〕/.exec(prompt)?.[1] ?? '?';
        return `定了 2.4 周五晚上八点发布〔${label}〕。`;
      }
      const asked = askedOf(prompt);
      if (asked.includes('回滚')) {
        return `要点：崩溃率超 0.3% 或支付成功率降 1 个点就回滚〔14:31 李四〕。\n===全文===\n${LONG_TAIL}`;
      }
      if (asked.includes('发布时间')) {
        return '定的是周五晚上八点发布〔14:32 张三〕。赵六说过要推迟〔09:00 赵六〕。';
      }
      if (asked.includes('预算')) return `我先翻一下会前资料。${TOOL_CALL}会上和资料里都没有提到预算，我不确定。`;
      return '收到。';
    },
    (prompt) => (prompt.includes('```json 代码块') || prompt.includes('只输出 SKIP') ? 0 : 400),
  );
}

let sim: Sim | undefined;
afterEach(async () => {
  await sim?.cleanup();
  sim = undefined;
});

/** The meeting so far: four people talk about the release. */
function talk(s: Sim): void {
  s.say('zhang', '今天定一下 2.4 的发布时间。', at('14:30'));
  s.say('li', '回滚条件我看资料里写的是崩溃率和支付成功率。', at('14:31'));
  s.say('zhang', '那就定周五晚上八点发布，李四负责发布前回归测试。', at('14:32'));
  s.say('wang', '监控面板我来看看吧，尽快弄。', at('14:33'));
}

describe('meeting agent batch 2/3 — simulated meeting', () => {
  it('runs a whole meeting: calendar docs, parallel asks, long answers, citations, voice, interjection, minutes, tasks, history', async () => {
    const agent = scripted();
    sim = await createSim({
      agent,
      meeting: {
        proactive: { enabled: true, waitMs: 300, maxPerMeeting: 1 },
        voiceWake: { enabled: true, names: ['小会'] },
      },
    });
    const s = sim;
    // A meeting from last week, already in the archive.
    const lastWeek = new Date(Date.now() - 7 * 86400_000);
    writeMeetingJson(meetingArchiveFile(meetingDataDir(s.controls), '20261120-1000-555555555'), {
      id: '20261120-1000-555555555',
      meetingId: 'm-old',
      meetingNo: '555555555',
      topic: '设计评审',
      endedAt: lastWeek.toISOString(),
      minutes: '结论：首页改版用方案 B，下周给出高保真。',
      todosParsed: true,
      todos: [],
    } satisfies MinutesRecord);

    // ── join ──
    await s.join();
    expect(s.feishu.some((c) => c.url === '/open-apis/vc/v1/bots/join')).toBe(true);
    await s.waitFor(() => s.meetingChat.length >= 1, 15_000, 'join notice');
    expect(s.meetingChat[0]).toContain('小会');

    // ── item 3.1: docs on the calendar event are bound on join ──
    await s.waitFor(() => textTo(s, OWNER_DM).includes('已从日程'), 15_000, 'calendar docs notice');
    const notice = textTo(s, OWNER_DM);
    expect(notice).toContain(`已从日程「${TOPIC}」带上会前资料`);
    expect(notice).toContain(`《${DOC.title}》r${DOC.revision}`);
    expect(notice).toContain('读不到'); // the secret doc is reported, not bound
    expect(s.feishu.some((c) => c.url.includes('DocTokenOther000'))).toBe(false); // other meeting's doc untouched

    talk(s);

    // ── item 2.1: two people ask at once, each gets an answer, no queueing ──
    s.chat('zhang', '@小会 刚才定的发布时间是什么？');
    s.chat('li', '@小会 回滚方案是什么？');
    await s.waitFor(() => s.meetingChat.filter((m) => m.includes('定的是') || m.includes('要点：')).length >= 2, 15_000, 'two answers');
    const asks = agent.runs.filter((r) => /发布时间|回滚/.test(askedOf(r.prompt)));
    expect(asks).toHaveLength(2);
    // Overlapping: the second started before the first finished.
    expect(Math.max(asks[0]!.startedAt, asks[1]!.startedAt)).toBeLessThan(Math.min(asks[0]!.endedAt!, asks[1]!.endedAt!));

    // ── item 2.3: citations; one that matches nothing said is marked ──
    const short = s.meetingChat.find((m) => m.includes('定的是'))!;
    expect(short).toContain('〔14:32 张三〕');
    expect(short).toContain('〔09:00 赵六（未核对到原话）〕');

    // ── item 2.2: long answer → points in the meeting, full text by DM ──
    const points = s.meetingChat.find((m) => m.includes('要点：'))!;
    expect(points).toContain('完整回答已私聊发给 李四');
    expect(points).not.toContain('财务对账');
    expect(textTo(s, PEOPLE.li.id)).toContain('财务对账');

    // ── item 2.3: not in the material → says so (prompt carries the rule) ──
    s.chat('zhang', '@小会 这次的预算是多少？');
    await s.waitFor(() => s.meetingChat.some((m) => m.includes('预算')), 15_000, 'budget answer');
    // Narration before a tool call is not part of the answer.
    expect(s.meetingChat.some((m) => m.includes('我先翻一下'))).toBe(false);
    const budgetPrompt = agent.runs.find((r) => r.prompt.includes('预算'))!.prompt;
    expect(budgetPrompt).toMatch(/没有|不要编|不确定/);

    // Same asker twice while the first runs: told to wait, the other asker unaffected.
    s.chat('zhang', '@小会 发布时间再确认一下？');
    s.chat('zhang', '@小会 还有别的吗？');
    await s.waitFor(() => s.meetingChat.some((m) => m.includes('你上一个问题还在回答')), 15_000, 'busy reply');
    await s.waitFor(() => s.meetingChat.filter((m) => m.includes('定的是')).length >= 2, 15_000, 'second time answer');

    // Someone not allowed to use the bot: no run.
    const before = agent.runs.length;
    s.chat('wang', '@小会 帮我把发布时间改成周六');
    await new Promise((r) => setTimeout(r, 200));
    expect(agent.runs.length).toBe(before);

    // ── item 3.4: calling it by name, allowlisted only ──
    s.say('wang', '小会，帮我查一下回滚负责人是谁', at('14:40'));
    await new Promise((r) => setTimeout(r, 200));
    expect(s.meetingChat.some((m) => m.includes('听到 王五'))).toBe(false);
    s.say('li', '小会，回滚负责人是谁？', at('14:41'));
    await s.waitFor(() => s.meetingChat.some((m) => m.includes('听到 李四 叫我')), 15_000, 'voice ack');
    await s.waitFor(() => agent.runs.some((r) => r.prompt.includes('回滚负责人是谁') && r.endedAt), 15_000, 'voice answer');
    const voiceRun = agent.runs.find((r) => r.prompt.includes('回滚负责人是谁'))!;
    expect(voiceRun.prompt).toContain('只读');
    // A question to the bot by name is not a candidate for the interjection.
    await new Promise((r) => setTimeout(r, 500));
    expect(agent.runs.some((r) => r.prompt.includes('只输出 SKIP'))).toBe(false);

    // ── item 3.3: unanswered question with the answer in the docs ──
    // Answered by someone within the wait: the bot stays quiet.
    s.say('zhang', '回滚条件里支付成功率是降多少来着？', at('14:42'));
    s.say('li', '降 1 个百分点。', at('14:42'));
    await new Promise((r) => setTimeout(r, 500));
    expect(s.meetingChat.some((m) => m.startsWith('💡'))).toBe(false);
    // Nobody answers: one line, citing the doc.
    s.say('zhang', '灰度比例一开始开多少来着？', at('14:43'));
    await s.waitFor(() => s.meetingChat.some((m) => m.startsWith('💡')), 15_000, 'interjection');
    const tip = s.meetingChat.find((m) => m.startsWith('💡'))!;
    expect(tip).toContain('10%');
    expect(tip).toContain(`《${DOC.title}》r${DOC.revision}`);
    // At most once per meeting.
    s.say('zhang', '测试负责人是谁来着？', at('14:44'));
    await new Promise((r) => setTimeout(r, 500));
    expect(s.meetingChat.filter((m) => m.startsWith('💡'))).toHaveLength(1);

    // ── meeting ends → minutes card ──
    await s.waitFor(() => agent.runs.every((r) => r.endedAt), 15_000, 'runs to finish');
    await s.end();
    const minutesMsg = s.sent.find((m) => m.to === OWNER_DM && m.card);
    expect(minutesMsg).toBeDefined();
    const card = JSON.stringify(minutesMsg!.card);
    // ── item 2.4: unclear owner / deadline → 待确认 ──
    expect(card).toContain('检查监控面板');
    expect(card.match(/待确认/g)!.length).toBeGreaterThanOrEqual(3); // 王五's due, todo 3's owner and due
    // ── item 2.5: two buttons per todo ──
    const buttons = cardButtons(s);
    expect(buttons.filter((b) => b.text === '派成工单')).toHaveLength(3);
    const taskButtons = buttons.filter((b) => b.text === '建成飞书任务');
    expect(taskButtons).toHaveLength(3);

    // Archive: owner of todo 1 resolved to 李四's open_id.
    const dir = meetingDataDir(s.controls);
    const archiveFile = readdirSync(join(dir, 'archive')).find((f) => f.includes(MEETING_NO))!;
    const rec = JSON.parse(readFileSync(join(dir, 'archive', archiveFile), 'utf8')) as MinutesRecord;
    expect(rec.todos[0]).toMatchObject({ owner: '李四', ownerId: PEOPLE.li.id, dueDate: '2026-12-04' });
    expect(rec.todos[1]).toMatchObject({ owner: '王五', due: '尽快' });
    expect(rec.todos[2]).toMatchObject({ owner: '', due: '' });

    // ── click 建成飞书任务 on todo 1 ──
    await s.click(taskButtons[0]!.value);
    const taskCalls = () => s.feishu.filter((c) => c.url === '/open-apis/task/v2/tasks');
    expect(taskCalls()).toHaveLength(1);
    const body = taskCalls()[0]!.data as Record<string, unknown>;
    expect(body.summary).toBe('发布前回归测试');
    expect(body.members).toEqual([
      { id: PEOPLE.li.id, type: 'user', role: 'assignee' },
      { id: OWNER, type: 'user', role: 'follower' },
    ]);
    const due = body.due as { timestamp: string; is_all_day: boolean };
    expect(due.is_all_day).toBe(true);
    expect(new Date(Number(due.timestamp)).getHours()).toBe(12);
    expect(new Date(Number(due.timestamp)).getDate()).toBe(4);
    expect(String(body.client_token)).toBe(`meeting-${rec.id}-1`);
    expect(textTo(s, OWNER_DM)).toContain('✅ 已建成飞书任务：发布前回归测试');
    // Twice: not created again.
    await s.click(taskButtons[0]!.value);
    expect(taskCalls()).toHaveLength(1);
    expect(textTo(s, OWNER_DM)).toContain('这条待办已经建过飞书任务');
    // Todo 3: no owner → not assigned, description says 待确认.
    await s.click(taskButtons[2]!.value);
    const body3 = taskCalls()[1]!.data as Record<string, unknown>;
    expect(body3.members).toEqual([{ id: OWNER, type: 'user', role: 'follower' }]);
    expect(body3.due).toBeUndefined();
    expect(String(body3.description)).toContain('待确认');
    const after = JSON.parse(readFileSync(join(dir, 'archive', archiveFile), 'utf8')) as MinutesRecord;
    expect(after.todos[0]!.taskGuid).toBe('task-guid-1');
    expect(after.todos[2]!.taskGuid).toBe('task-guid-2');

    // ── item 3.2: DM questions about past meetings ──
    expect(await s.dm(OWNER, '上周那个会定了什么')).toBe(true);
    await s.waitFor(() => textTo(s, OWNER_DM).includes('方案 B') || textTo(s, OWNER_DM).includes('设计评审'), 15_000, 'history answer');
    const hist = agent.runs.find((r) => r.prompt.includes('私聊里有人问：上周那个会定了什么'))!;
    expect(hist.prompt).toContain('设计评审');
    expect(hist.prompt).not.toContain(TOPIC); // only last week's meeting
    expect(await s.dm(OWNER, '刚才那个会定了什么')).toBe(true);
    await s.waitFor(() => agent.runs.some((r) => r.prompt.includes('私聊里有人问：刚才那个会定了什么') && r.endedAt), 15_000, 'history 2');
    expect(agent.runs.find((r) => r.prompt.includes('私聊里有人问：刚才那个会定了什么'))!.prompt).toContain(TOPIC);
    // Not swallowed: a request, and a non-admin asking.
    expect(await s.dm(OWNER, '帮我把今天的会议纪要翻译成英文')).toBe(false);
    expect(await s.dm(PEOPLE.li.id, '上周那个会定了什么')).toBe(false);
  }, 120_000);

  it('says which permission is missing when the app cannot create tasks or read the calendar', async () => {
    sim = await createSim({ agent: scripted(), taskScopeMissing: true, calendarDenied: true });
    const s = sim;
    await s.join();
    // Straight to the owner, not into the chat that started the join.
    await s.waitFor(() => textTo(s, OWNER).includes('没能从会议日程里读会前资料'), 15_000, 'calendar notice');
    expect(textTo(s, OWNER)).toContain('calendar:calendar:readonly');
    talk(s);
    await s.end();
    const task = cardButtons(s).find((b) => b.text === '建成飞书任务')!;
    await s.click(task.value);
    expect(textTo(s, OWNER_DM)).toContain('task:task:write');
    // Nothing recorded as created.
    const dir = meetingDataDir(s.controls);
    const f = readdirSync(join(dir, 'archive'))[0]!;
    expect((JSON.parse(readFileSync(join(dir, 'archive', f), 'utf8')) as MinutesRecord).todos[0]!.taskGuid).toBeUndefined();
  }, 60_000);

  it('a stop from one asker stops only their question', async () => {
    const agent = scripted();
    sim = await createSim({ agent });
    const s = sim;
    await s.join();
    talk(s);
    s.chat('zhang', '@小会 刚才定的发布时间是什么？');
    s.chat('li', '@小会 回滚方案是什么？');
    await s.waitFor(() => agent.runs.filter((r) => askedOf(r.prompt)).length >= 2, 15_000, 'both running');
    s.chat('zhang', '@小会 stop');
    await s.waitFor(() => s.meetingChat.some((m) => m.includes('已中断你的问题')), 15_000, 'stop reply');
    await s.waitFor(() => s.meetingChat.some((m) => m.includes('要点：')), 15_000, "li's answer");
    expect(s.meetingChat.some((m) => m.includes('定的是'))).toBe(false);
  }, 60_000);
});
