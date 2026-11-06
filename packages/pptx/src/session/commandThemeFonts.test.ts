import { expect, spyOn, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { createSessionClient } from '../../../../shared/office-session/client';
import { createSessionHost } from '../../../../shared/office-session/host';
import { createInProcessPair } from '../../../../shared/office-session/testing/inProcessTransport';
import type { PptxStoryText } from '../edits';
import { themeFontReference, themeFontSlot } from '../textFormatting';
import * as wasm from '../wasm/loader';
import { registerPresentationCommandHandlers } from './commandHost';
import { PRESENTATION_COMMAND_METHODS, type PresentationCommandAck, type PresentationCommandEnvelope,
  type PresentationCommandEvents, type PresentationCommandIntent, type PresentationCommandMethods,
  type PresentationCommandSelection } from './commandProtocol';

const DECK = new URL('../../../../crates/pptx-render/tests/fixtures/theme-fonts.pptx', import.meta.url);
const identity = { session: 'theme', incarnation: 'theme-worker' };
const input = { inputTime: 1, undoGroup: 'theme' };
const THEME_FONTS = {
  major: { latin: 'Liberation Serif', eastAsian: null },
  minor: { latin: 'Liberation Sans', eastAsian: null },
};

function selection(story: PptxStoryText, anchor = 0, focus = anchor): PresentationCommandSelection {
  return { slideId: story.slideId, shapeId: story.shapeId, storyId: story.storyId,
    anchor, focus, affinity: 'downstream' };
}

async function open() {
  const pair = createInProcessPair();
  const registration = registerPresentationCommandHandlers({ initWasm: async () => {
    await wasm.initWasm(await readFile(new URL('../wasm/generated/pptx_wasm_bg.wasm', import.meta.url)));
  } });
  let handle: wasm.PresentationHandle;
  const original = wasm.openPresentation;
  const opened = spyOn(wasm, 'openPresentation').mockImplementation((...args) => {
    handle = original(...args);
    return handle;
  });
  createSessionHost<PresentationCommandMethods, PresentationCommandEvents, null>(pair.host, {
    handlers: registration.handlers, policies: registration.policies, context: null, onDispose: registration.dispose,
  });
  const rpc = createSessionClient<PresentationCommandMethods, PresentationCommandEvents>(pair.client,
    { methods: PRESENTATION_COMMAND_METHODS, silenceMs: 2000 });
  try {
    const font = await readFile(new URL('../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf', import.meta.url));
    const state = await rpc.call.openCommandEditor(new Uint8Array(await readFile(DECK)).buffer,
      { ...identity, fonts: [{ family: 'Liberation Sans', bytes: new Uint8Array(font).buffer }] });
    const result = (await rpc.call.readEditor({ ...identity, seq: 0 })).result;
    if (!result.ok) throw new Error(result.failure.message);
    const story = (prefix: string) => result.stories.find((candidate) => candidate.text.startsWith(prefix))!;
    return { state, story, handle: () => handle!,
      command(state: PresentationCommandAck, seq: number, intent: unknown, afterSeq?: number) {
        return rpc.call.command({ ...identity, seq, basis: { epoch: state.epoch, version: state.version },
          ...(afterSeq === undefined ? {} : { afterSeq }), ...(intent as PresentationCommandIntent) } as PresentationCommandEnvelope);
      },
      async dispose() { try { await rpc.dispose(); } finally { opened.mockRestore(); registration.dispose(); pair.client.close(); } },
    };
  } catch (error) {
    opened.mockRestore();
    registration.dispose();
    pair.client.close();
    throw error;
  }
}

test('theme font references map to slots and back', () => {
  expect(themeFontReference('major')).toBe('+mj-lt');
  expect(themeFontReference('minor')).toBe('+mn-lt');
  expect(themeFontSlot('+MJ-LT')).toBe('major');
  expect(themeFontSlot('+mn-lt')).toBe('minor');
  expect(themeFontSlot('+mn-ea')).toBeNull();
  expect(themeFontSlot('Calibri')).toBeNull();
});

test('a heading run reports the major slot and the slide theme fonts', async () => {
  const session = await open();
  try {
    const title = session.story('Major theme reference');
    const selected = await session.command(session.state, 1, { kind: 'setSelection', payload: {
      ...input, selection: selection(title, 0, 5) } });
    expect(selected.editState.toolbar.formatting).toMatchObject({ fontFamily: 'Liberation Serif', fontTheme: 'major' });
    expect(selected.editState.toolbar.themeFonts).toEqual(THEME_FONTS);
    expect(wasm.presentationThemeFonts(session.handle(), title.slideId)).toEqual(THEME_FONTS);
    const literal = session.story('Explicit font');
    const moved = await session.command(selected, 2, { kind: 'setSelection', payload: {
      ...input, selection: selection(literal, 0, 5) } });
    expect(moved.editState.toolbar.formatting).toMatchObject({ fontFamily: 'Arial', fontTheme: null });
  } finally { await session.dispose(); }
});

test('a run set to the body theme font saves as +mn-lt', async () => {
  const session = await open();
  try {
    const literal = session.story('Explicit font');
    const ack = await session.command(session.state, 1, { kind: 'formatText', payload: { ...input,
      at: { source: 'confirmed', selection: selection(literal, 0, 8) }, patch: { fontFamily: themeFontReference('minor') } } });
    expect(ack.outcomes).toEqual([{ seq: 1, status: 'applied', committed: true }]);
    expect(ack.editState.toolbar.formatting).toMatchObject({ fontFamily: 'Liberation Sans', fontTheme: 'minor' });
    const xml = await (await JSZip.loadAsync(session.handle().save())).file('ppt/slides/slide1.xml')!.async('string');
    expect(xml).toContain('<a:latin typeface="+mn-lt"/></a:rPr><a:t>Explicit</a:t>');
    expect(xml).toContain('<a:latin typeface="Arial"/></a:rPr><a:t> font stays unchanged.</a:t>');
    expect(xml).toContain('<a:latin typeface="+mj-lt"/></a:rPr><a:t>Major theme reference</a:t>');
  } finally { await session.dispose(); }
});

test('a caret mark of a theme font reports the resolved family and its slot', async () => {
  const session = await open();
  try {
    const literal = session.story('Explicit font');
    const marked = await session.command(session.state, 1, { kind: 'formatText', payload: { ...input,
      at: { source: 'confirmed', selection: selection(literal, 3) }, patch: { fontFamily: themeFontReference('major') } } });
    expect(marked.editState.toolbar.storedMarks).toEqual({ fontFamily: '+mj-lt' });
    expect(marked.editState.toolbar.formatting).toMatchObject({ fontFamily: 'Liberation Serif', fontTheme: 'major' });
  } finally { await session.dispose(); }
});
