import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import type { NormalizedMessage } from '@larksuite/channel';

const mocks = vi.hoisted(() => ({ startRunFlow: vi.fn() }));
vi.mock('../../../src/bot/run-flow', () => ({ startRunFlow: mocks.startRunFlow }));

const { tryHandleCommand } = await import('../../../src/commands/index');
const { createDefaultProfileConfig } = await import('../../../src/config/profile-schema');
const { meetingDataDir, writeMeetingJson } = await import('../../../src/meeting/store');
const { meetingArchiveFile } = await import('../../../src/meeting/minutes');
const { createFakeChannel } = await import('../../helpers/fake-channel');
import type { CommandContext } from '../../../src/commands/index';

/** The DM shortcut for past meetings must never swallow a message it can't answer, and never runs in groups. */

let root: string;
beforeEach(() => {
  vi.clearAllMocks();
  root = mkdtempSync(join(tmpdir(), 'meeting-history-cmd-'));
});
afterEach(() => rmSync(root, { recursive: true, force: true }));

function harness(content: string, opts: { chatType?: string; senderId?: string; resources?: unknown[]; archive?: boolean } = {}) {
  const channel = createFakeChannel();
  const profileConfig = createDefaultProfileConfig({
    agentKind: 'claude',
    accounts: { app: { id: 'app-id', secret: 'secret', tenant: 'feishu' } },
  });
  profileConfig.meeting = { ...profileConfig.meeting, enabled: true };
  const controls = { configPath: join(root, 'config.json'), profile: 'claude', profileConfig, botOwnerId: 'ou-owner', meeting: {} };
  if (opts.archive !== false) {
    const id = '20261126-1500-111111111';
    writeMeetingJson(meetingArchiveFile(meetingDataDir(controls), id), {
      id,
      meetingId: 'm',
      meetingNo: '111111111',
      topic: '周会',
      endedAt: new Date(Date.now() - 3600_000).toISOString(),
      minutes: '定了周五发布',
      todosParsed: true,
      todos: [],
    });
  }
  const ctx = {
    channel,
    msg: {
      messageId: 'om-1',
      chatId: 'oc-dm',
      chatType: opts.chatType ?? 'p2p',
      senderId: opts.senderId ?? 'ou-owner',
      content,
      resources: opts.resources ?? [],
      mentions: [],
      mentionedBot: false,
    } as unknown as NormalizedMessage,
    scope: 'oc-dm',
    chatMode: 'p2p',
    controls,
    runExecutor: {},
    sessions: {},
    workspaces: {},
    activeRuns: {},
  } as unknown as CommandContext;
  const replies = () => channel.sent.map((m) => String((m.content as { markdown?: string }).markdown ?? ''));
  return { ctx, replies };
}

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
  };
}

describe('past-meeting questions in a DM', () => {
  it('answers from the archive when the owner asks and something matches', async () => {
    mocks.startRunFlow.mockResolvedValue(fakeRun('定了周五发布〔11-26 周会〕'));
    const h = harness('最近那个会定了什么');
    expect(await tryHandleCommand(h.ctx)).toBe(true);
    expect(h.replies().join('\n')).toContain('定了周五发布〔11-26 周会〕');
  });

  it('lets the message through to the agent when nothing matches, files are attached, or the sender is no admin', async () => {
    expect(await tryHandleCommand(harness('最近那个会定了什么', { archive: false }).ctx)).toBe(false);
    expect(await tryHandleCommand(harness('最近那个会定了什么', { resources: [{ type: 'file' }] }).ctx)).toBe(false);
    expect(await tryHandleCommand(harness('最近那个会定了什么', { senderId: 'ou-someone' }).ctx)).toBe(false);
    expect(await tryHandleCommand(harness('最近那个会定了什么', { chatType: 'group' }).ctx)).toBe(false);
    expect(mocks.startRunFlow).not.toHaveBeenCalled();
  });

  it('/meeting history refuses to run in a group', async () => {
    const h = harness('/meeting history 上周定了什么', { chatType: 'group' });
    expect(await tryHandleCommand(h.ctx)).toBe(true);
    expect(h.replies().join('\n')).toContain('只能私聊');
    expect(mocks.startRunFlow).not.toHaveBeenCalled();
  });
});
