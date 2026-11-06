import { open } from './peerlessTestHarness';
import { expect, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { act, waitFor } from '@testing-library/react';
import JSZip from 'jszip';

const DECK = new URL('../../../../crates/pptx-render/tests/fixtures/theme-fonts.pptx', import.meta.url);

test('fontFamily reports and writes theme fonts', async () => {
  const value = await open({ file: new Uint8Array(await readFile(DECK)) });
  const content = await value.api.readContent();
  if (!content.ok) throw new Error(content.failure.message);
  const story = (prefix: string) => content.stories.find((candidate) => candidate.text.startsWith(prefix))!;
  const select = async (prefix: string, start: number, end: number) => {
    const target = story(prefix);
    await act(async () => { await value.api.selectText({ slide: 1, shapeId: target.shapeId, storyId: target.storyId, start, end }); });
  };
  const commands = value.api.commands!;

  await select('Major theme reference', 0, 5);
  await waitFor(() => expect(commands.getState('fontFamily').theme).toBe('major'));
  const heading = commands.getState('fontFamily');
  expect(heading.value).toBe('Liberation Serif');
  expect(heading.themeFonts).toEqual({
    major: { latin: 'Liberation Serif', eastAsian: null },
    minor: { latin: 'Liberation Sans', eastAsian: null },
  });
  expect(heading.options?.slice(0, 2).map((option) => option.label))
    .toEqual(['Liberation Serif (Headings)', 'Liberation Sans (Body)']);

  await select('Explicit font', 0, 8);
  await waitFor(() => expect(commands.getState('fontFamily').value).toBe('Arial'));
  expect(commands.getState('fontFamily').theme).toBeNull();
  await act(async () => {
    expect(await commands.execute('fontFamily', { theme: 'minor' })).toMatchObject({ ok: true, status: 'executed' });
  });
  await waitFor(() => expect(commands.getState('fontFamily').theme).toBe('minor'));
  expect(commands.getState('fontFamily').value).toBe('Liberation Sans');

  const saved = await JSZip.loadAsync(await value.api.save());
  const xml = await saved.file('ppt/slides/slide1.xml')!.async('string');
  expect(xml).toContain('<a:latin typeface="+mn-lt"/></a:rPr><a:t>Explicit</a:t>');
  expect(xml).toContain('<a:latin typeface="+mj-lt"/></a:rPr><a:t>Major theme reference</a:t>');
});
