import { readdirSync } from 'node:fs';
import { join } from 'node:path';
import { log } from '../core/logger';
import { docTerms } from './docs';
import { shortTime, todoWhoWhen, type MinutesRecord } from './minutes';
import { runReadOnlyPrompt, type RunDeps } from './orchestrator';
import { meetingDataDir, readMeetingJson } from './store';

/**
 * Questions about past meetings ("上周那个会定了什么"), answered from the
 * minutes archive (`meeting-data/archive/*.json`) — never from memory. The
 * archive is narrowed by the time the question names and ranked by the words
 * it shares with each meeting; only those minutes go to a read-only run, which
 * must cite the meeting each point comes from. When nothing matches, the
 * answer says so without running the agent.
 */

const DAY = 24 * 3600 * 1000;
/** Without a time in the question, look this far back. */
const DEFAULT_LOOKBACK_DAYS = 30;
/** Minutes given to one run, newest-and-most-relevant first, and their total size. */
const MAX_MEETINGS = 6;
const BUDGET_CHARS = 40_000;

const MEETING_WORD = /(会议|例会|周会|晨会|早会|夕会|评审会|复盘会|启动会|站会|开会|的会(?=[\s，,。？?上里定说聊讨有都结谁怎最的]|$)|那个会|那场会|这个会|上次会|会上|纪要)/;
/** Asking what a meeting settled: a recall on its own. */
const STRONG_RECALL = /(定了什么|定了啥|定的什么|定的啥|怎么定的|决定了什么|决定了啥|结论是什么|结论是啥|什么结论|啥结论|有什么结论)/;
/** Could be about a past meeting or about one being planned: needs a past anchor. */
const WEAK_RECALL =
  /(说了什么|说了啥|聊了什么|聊了啥|讨论了什么|讨论了啥|讲了什么|讲了啥|有什么待办|有哪些待办|待办是什么|待办有哪些|谁负责|安排了什么|纪要)/;
/** Points at a meeting that already happened. */
const PAST =
  /(上上周|上周|上个?星期|上次|上回|昨天|前天|今天上午|上个月|这周|本周|那天|最近|之前|刚才|刚刚|那个会|那场会|那次|(?<![\w.])\d{1,2}\s*(?:月|\/|-)\s*\d{1,2})/;
/** Planning ahead or asking for work: not a recall, whatever else the sentence says. */
const NOT_RECALL =
  /(明天|后天|下周|下个月|待会|一会儿|等会|要开|将要|准备开|打算|翻译|设计一下|帮我写|写个|写一份|写一下|列个|列一下|整理成|生成|润色|改一下|模板)/;

/**
 * Is this DM a question about a past meeting? Needs a meeting word and a
 * recall phrase. "定了什么 / 结论是什么" count on their own; "谁负责 / 说了什么
 * / 纪要" only with a past anchor (上周, 昨天, 那个会, 11月20日 …). Plans
 * ("明天的评审会谁负责主持") and requests ("把今天的会议纪要翻译成英文") never count.
 */
export function looksLikeHistoryQuestion(text: string): boolean {
  const t = String(text ?? '').trim();
  if (!t || t.startsWith('/') || t.length > 200) return false;
  if (!MEETING_WORD.test(t) || NOT_RECALL.test(t)) return false;
  if (STRONG_RECALL.test(t)) return true;
  return WEAK_RECALL.test(t) && PAST.test(t);
}

export interface TimeWindow {
  from: number;
  to: number;
  label: string;
}

function startOfDay(ms: number): number {
  const d = new Date(ms);
  d.setHours(0, 0, 0, 0);
  return d.getTime();
}

/** Monday 00:00 of the week `ms` is in. */
function startOfWeek(ms: number): number {
  const d = new Date(startOfDay(ms));
  const dow = (d.getDay() + 6) % 7;
  return d.getTime() - dow * DAY;
}

/** The time a question names (今天 / 昨天 / 上周 / 11月20日 / 最近3天 …), else the last 30 days. */
export function historyWindow(text: string, now = Date.now()): TimeWindow {
  const t = String(text ?? '');
  const today = startOfDay(now);
  const day = (offset: number, label: string): TimeWindow => ({
    from: today + offset * DAY,
    to: today + (offset + 1) * DAY,
    label,
  });
  if (/前天/.test(t)) return day(-2, '前天');
  if (/昨天|昨日/.test(t)) return day(-1, '昨天');
  if (/今天|今日/.test(t)) return day(0, '今天');
  const week = startOfWeek(now);
  if (/上上周|上上个?星期/.test(t)) return { from: week - 14 * DAY, to: week - 7 * DAY, label: '上上周' };
  if (/上周|上个?星期|上个?礼拜/.test(t)) return { from: week - 7 * DAY, to: week, label: '上周' };
  if (/这周|本周|这个?星期|这个?礼拜/.test(t)) return { from: week, to: now + 1, label: '这周' };
  const d = new Date(now);
  if (/上个?月/.test(t)) {
    return {
      from: new Date(d.getFullYear(), d.getMonth() - 1, 1).getTime(),
      to: new Date(d.getFullYear(), d.getMonth(), 1).getTime(),
      label: '上个月',
    };
  }
  if (/这个?月|本月/.test(t)) return { from: new Date(d.getFullYear(), d.getMonth(), 1).getTime(), to: now + 1, label: '这个月' };
  const recent = /(?:最近|近|过去)\s*(\d{1,3})\s*天/.exec(t);
  if (recent) {
    const n = Number(recent[1]);
    return { from: today - (n - 1) * DAY, to: now + 1, label: `最近 ${n} 天` };
  }
  // Not "2.3" in "v2.3版本": a dot is a version, and the number must stand on its own.
  const md = /(?<![\w.])(\d{1,2})\s*(?:月|\/|-)\s*(\d{1,2})(?![\d.])\s*(?:日|号)?/.exec(t);
  if (md) {
    const month = Number(md[1]) - 1;
    const date = Number(md[2]);
    if (month >= 0 && month < 12 && date >= 1 && date <= 31) {
      let y = d.getFullYear();
      // "12月30日" asked in January is last year's.
      if (new Date(y, month, date).getTime() > now + DAY) y -= 1;
      const from = new Date(y, month, date).getTime();
      return { from, to: from + DAY, label: `${month + 1}月${date}日` };
    }
  }
  return { from: today - (DEFAULT_LOOKBACK_DAYS - 1) * DAY, to: now + 1, label: `最近 ${DEFAULT_LOOKBACK_DAYS} 天` };
}

/** Every archived minutes record, newest first. Unreadable files are skipped. */
export function loadArchives(dataDir: string): MinutesRecord[] {
  const dir = join(dataDir, 'archive');
  let files: string[];
  try {
    files = readdirSync(dir).filter((f) => f.endsWith('.json'));
  } catch {
    return [];
  }
  const out: MinutesRecord[] = [];
  for (const f of files) {
    const rec = readMeetingJson<MinutesRecord | null>(join(dir, f), null);
    if (rec && typeof rec.minutes === 'string' && typeof rec.endedAt === 'string') out.push(rec);
  }
  return out.sort((a, b) => Date.parse(b.endedAt) - Date.parse(a.endedAt));
}

/** Points at the meeting that just ended. */
const JUST_NOW = /(刚才|刚刚|刚开完|刚开的|刚结束|刚散会|上一场)/;

/** Words that say *when* or *how to ask*, not *which meeting*: not used for matching. */
const NOISE =
  /(刚才|刚刚|刚开完|刚开的|刚结束|刚散会|散会|上一场|上上周|上周|这周|本周|今天|昨天|前天|上个?月|这个?月|本月|最近|上次|上回|那个|那场|这个|会议|例会|周会|开会|的会|纪要|定了什么|定了啥|决定|结论|说了什么|聊了什么|讨论了什么|待办|谁负责|是什么|什么|一下|帮我|查查|查一下|请问|我们|你们|他们|？|\?|，|。)/g;

/** Pick the minutes a question is about: in its time window, ranked by shared words, then by recency. */
export function pickArchives(
  all: MinutesRecord[],
  question: string,
  now = Date.now(),
): { window: TimeWindow; picked: MinutesRecord[] } {
  const window = historyWindow(question, now);
  const inWindow = all.filter((r) => {
    const t = Date.parse(r.endedAt);
    return t >= window.from && t < window.to;
  });
  const terms = [...docTerms(question.replace(NOISE, ' '))];
  const scored = inWindow.map((r, i) => {
    const hay = docTerms([r.topic ?? '', ...(r.participants ?? []), r.minutes, ...r.todos.map((t) => `${t.task} ${t.owner}`)].join('\n'));
    const topicTerms = docTerms(`${r.topic ?? ''} ${(r.participants ?? []).join(' ')}`);
    let score = 0;
    for (const term of terms) {
      if (topicTerms.has(term)) score += 3;
      else if (hay.has(term)) score += 1;
    }
    return { r, score, i };
  });
  // With words to go by, keep the meetings that share any; if none do, fall back to recency.
  const matched = terms.length ? scored.filter((x) => x.score > 0) : [];
  const pool = matched.length ? matched : scored;
  // "刚才那个会" is one meeting: the newest that fits, not the newest few.
  if (JUST_NOW.test(question)) {
    const newest = [...pool].sort((a, b) => a.i - b.i)[0];
    return { window, picked: newest ? [newest.r] : [] };
  }
  const ranked = pool.sort((a, b) => b.score - a.score || a.i - b.i);
  const picked: MinutesRecord[] = [];
  let size = 0;
  for (const { r } of ranked) {
    if (picked.length >= MAX_MEETINGS) break;
    const n = r.minutes.length + 500;
    if (picked.length && size + n > BUDGET_CHARS) break;
    picked.push(r);
    size += n;
  }
  return { window, picked };
}

/** `11-20 周会` — how the answer names a meeting. */
export function historyLabel(r: Pick<MinutesRecord, 'endedAt' | 'topic' | 'meetingNo'>): string {
  return `${shortTime(r.endedAt).slice(0, 5)} ${r.topic ?? r.meetingNo}`;
}

export function buildHistoryPrompt(question: string, picked: MinutesRecord[], window: TimeWindow): string {
  const parts = [
    `以下是本机存档的会议纪要（${window.label}，共 ${picked.length} 场，新的在前）。`,
    '',
  ];
  for (const r of picked) {
    parts.push(`=== 〔${historyLabel(r)}〕会议号 ${r.meetingNo} · 结束于 ${shortTime(r.endedAt)}${r.participants?.length ? ` · 参会：${r.participants.join('、')}` : ''} ===`);
    parts.push(r.minutes.slice(0, 12_000));
    if (r.todos.length) {
      parts.push('待办：');
      r.todos.forEach((t, i) =>
        parts.push(
          `${i + 1}. ${t.task}（${todoWhoWhen(t)}${t.ticket ? ` · 已派工单 ${t.ticket}` : ''}${t.taskGuid ? ' · 已建飞书任务' : ''}）`,
        ),
      );
    }
    parts.push('=== 纪要结束 ===', '');
  }
  parts.push(
    `私聊里有人问：${question}`,
    '',
    '只依据上面的纪要回答。每个要点后面标出自哪场会，照抄方括号里的会名，如〔11-20 周会〕。',
    '纪要里没有的，直接说「纪要里没有记录」，不要编。问的是哪场会对不上时，列出最接近的几场让对方确认。回答简洁，不要用 markdown 标题。',
  );
  return parts.join('\n');
}

/** Would the archive have anything for this question? (The DM shortcut only takes questions it can answer.) */
export function historyHasMatch(dataDir: string, question: string, now = Date.now()): boolean {
  const all = loadArchives(dataDir);
  return all.length > 0 && pickArchives(all, question, now).picked.length > 0;
}

export interface HistoryDeps extends RunDeps {
  chatId: string;
}

/**
 * Answer a question about past meetings from the archive. Returns the text to
 * send back.
 */
export async function answerMeetingHistory(deps: HistoryDeps, question: string, now = Date.now()): Promise<string> {
  const dataDir = meetingDataDir(deps.controls);
  const all = loadArchives(dataDir);
  if (all.length === 0) {
    return '本机还没有会议纪要存档。会议智能体参加的会散会后会自动存一份，之后就能这样问。';
  }
  const { window, picked } = pickArchives(all, question, now);
  log.info('meeting', 'history-question', { archives: all.length, window: window.label, picked: picked.length });
  if (picked.length === 0) {
    const latest = all[0]!;
    return `没找到${window.label}的会议纪要（本机共 ${all.length} 份，最近一份是 ${historyLabel(latest)}）。换个时间问问看，比如「${historyLabel(latest).slice(0, 5)} 那个会定了什么」。`;
  }
  const { text: answer, refused } = await runReadOnlyPrompt(
    deps,
    buildHistoryPrompt(question, picked, window),
    `meeting-history:${deps.chatId}`,
    deps.chatId,
  );
  if (refused) {
    return refused.code === 'run-already-active' ? '上一个查纪要的问题还在处理，稍后再问。' : `没能查：${refused.userVisible}`;
  }
  if (!answer) return '没有查出结果（agent 没有返回内容）。';
  const sources = picked.map((r) => `${historyLabel(r)}（存档 ${r.id}）`).join('、');
  return `${answer}\n\n> 查的是本机存档的纪要：${sources}`;
}
