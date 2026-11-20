import { mkdtemp, readdir, readFile, readlink, mkdir, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import {
  ATTACHMENT_BATCH_DIR,
  landAttachmentBatch,
  pruneAttachmentBatches,
  renderAttachmentNotice,
  safeAttachmentName,
} from '../../../src/media/batch';

const att = (over = {}) => ({
  decision: 'accepted' as const,
  path: '/cache/a.png',
  kind: 'image' as const,
  mime: 'image/png',
  size: 10,
  originalName: 'a.png',
  hash: 'abcdef0123456789',
  ...over,
});

describe('renderAttachmentNotice', () => {
  it('is empty when nothing real was dropped', () => {
    expect(renderAttachmentNotice([att()])).toBe('');
    expect(renderAttachmentNotice([att(), att({ decision: 'skipped', rejectionReason: 'sticker' })])).toBe('');
  });
  it('counts drops by reason', () => {
    const t = renderAttachmentNotice([att(), att({ decision: 'rejected', rejectionReason: 'file-too-large' })]);
    expect(t).toBe('⚠️ 附件：收到 2 个，已处理 1 个；1 个（单个文件过大）未处理');
    expect(renderAttachmentNotice([att({ decision: 'rejected', rejectionReason: 'x' })])).toContain('全部未处理');
  });
});

describe('safeAttachmentName', () => {
  it('keeps only the base name and numbers it', () => {
    expect(safeAttachmentName(0, { originalName: '../../etc/passwd', kind: 'file', mime: '' })).toBe('01-passwd');
    expect(safeAttachmentName(9, { originalName: '..hidden', kind: 'file', mime: '' })).toBe('10-hidden');
    expect(safeAttachmentName(1, { kind: 'image', mime: 'image/png' })).toBe('02-image.png');
    expect(safeAttachmentName(0, { originalName: `${'x'.repeat(100)}.pdf`, kind: 'file', mime: '' })).toHaveLength(83);
  });
});

describe('landAttachmentBatch', () => {
  it('links accepted files in order with a manifest, only from the threshold on', async () => {
    const ws = await mkdtemp(join(tmpdir(), 'batch-'));
    const cfg = { batchThreshold: 2, batchKeep: 5 };
    expect(await landAttachmentBatch([att()], ws, cfg)).toBeUndefined();
    const b = await landAttachmentBatch(
      [att({ originalName: 'one.png' }), att({ decision: 'rejected' }), att({ originalName: 'two.png', path: '/cache/b.png' })],
      ws,
      cfg,
      new Date('2026-11-20T10:00:00Z'),
    );
    expect(b?.dir).toBe(join(ws, ATTACHMENT_BATCH_DIR, '20261120100000-abcdef01'));
    expect((await readdir(b!.dir)).sort()).toEqual(['01-one.png', '02-two.png', 'manifest.json']);
    expect(await readlink(join(b!.dir, '02-two.png'))).toBe('/cache/b.png');
    const manifest = JSON.parse(await readFile(join(b!.dir, 'manifest.json'), 'utf8'));
    expect(manifest.count).toBe(2);
  });

  it('prunes the oldest batches', async () => {
    const root = await mkdtemp(join(tmpdir(), 'batch-'));
    for (const n of ['20260101000000-a', '20260102000000-b', '20260103000000-c', 'keep-me']) {
      await mkdir(join(root, n));
    }
    await writeFile(join(root, 'note.txt'), '');
    await pruneAttachmentBatches(root, 2);
    expect((await readdir(root)).sort()).toEqual(['20260102000000-b', '20260103000000-c', 'keep-me', 'note.txt']);
  });
});
