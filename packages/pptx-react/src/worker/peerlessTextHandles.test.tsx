import { expect, test } from 'bun:test';
import JSZip from 'jszip';
import { decodeCommandDisplayList } from '@betteroffice/pptx/worker';
import type { TextBoxPrimitive } from '@betteroffice/pptx';
import { open, type } from './peerlessTestHarness';

const { act, fireEvent, waitFor } = await import('@testing-library/react');

type Harness = Awaited<ReturnType<typeof open>>;

async function editExisting(h: Harness) {
  const line = h.box.lines[0]!;
  const at = h.point(line.caretStops[2]!.x, line.y + line.height / 2);
  act(() => { fireEvent.pointerDown(h.canvas, at); fireEvent.pointerUp(h.stage, at); });
  await waitFor(() => expect(h.client.state.selection).toMatchObject({ shapeId: h.box.shapeId, anchor: 2, focus: 2 }));
  await waitFor(() => expect(h.view.getAllByTestId('pptx-text-edit-outline')).toHaveLength(1));
}

function shownBox(h: Harness): TextBoxPrimitive {
  return decodeCommandDisplayList(h.client.state.frame!).primitives.find((primitive) =>
    primitive.kind === 'textBox' && primitive.shapeId === h.box.shapeId) as TextBoxPrimitive;
}

test('text editing shows the eight resize handles around the dashed outline', async () => {
  const h = await open();
  await editExisting(h);
  expect(h.client.state.selection?.kind).not.toBe('object');
  const outline = h.view.getByTestId('pptx-text-edit-outline');
  expect(outline.style.border).toBe('1px dashed #7a7a7a');
  const handles = h.view.getAllByTestId(/pptx-resize-/);
  expect(handles).toHaveLength(8);
  const scale = parseFloat(h.canvas.style.width) / h.list.width;
  const se = h.view.getByTestId('pptx-resize-se');
  expect(se.style.border).toBe('1px solid #7a7a7a');
  expect(se.style.width).toBe('8px');
  expect(parseFloat(se.style.left) + 4).toBeCloseTo(parseFloat(outline.style.left) + parseFloat(outline.style.width), 3);
  expect(parseFloat(se.style.top) + 4).toBeCloseTo(parseFloat(outline.style.top) + parseFloat(outline.style.height), 3);
  expect(parseFloat(outline.style.left)).toBeCloseTo(h.box.x * scale, 3);
});

test('no rotation handle is shown: the engine has no rotate command', async () => {
  const h = await open();
  await editExisting(h);
  expect(h.view.queryAllByTestId(/pptx-rotate/)).toHaveLength(0);
  await h.select();
  expect(h.view.queryAllByTestId(/pptx-rotate/)).toHaveLength(0);
});

test('read-only text selection shows the outline without handles', async () => {
  const h = await open();
  await editExisting(h);
  h.view.rerender(h.ui(true));
  await waitFor(() => expect(h.client.state.editState.readOnly).toBe(true));
  await waitFor(() => expect(h.view.queryAllByTestId(/pptx-resize-/)).toHaveLength(0));
});

test('rotated text boxes have no handles while editing, like object selection', async () => {
  const seed = await open();
  const zip = await JSZip.loadAsync(seed.bytes);
  const source = await zip.file('ppt/slides/slide1.xml')!.async('string');
  zip.file('ppt/slides/slide1.xml', source.replace('<a:xfrm>', '<a:xfrm rot="1800000">'));
  seed.view.unmount();
  await seed.client.dispose();
  const h = await open({ file: await zip.generateAsync({ type: 'uint8array' }) });
  await act(async () => { await h.client.dispatch({ kind: 'setSelection', payload: {
    selection: { slideId: h.client.state.slides[0]!.id, shapeId: h.box.shapeId!, storyId: h.box.storyId!,
      anchor: 2, focus: 2, affinity: 'downstream' }, inputTime: 1, undoGroup: 'selection' } }); });
  await waitFor(() => expect(h.view.getAllByTestId('pptx-text-edit-outline')).toHaveLength(1));
  expect(h.view.queryAllByTestId(/pptx-resize-/)).toHaveLength(0);
});

test('dragging a handle while editing resizes the box and keeps the caret', async () => {
  const h = await open();
  await editExisting(h);
  const before = shownBox(h);
  const handle = h.view.getByTestId('pptx-resize-se');
  const start = h.point(before.x + before.w, before.y + before.h);
  const end = { ...start, clientX: start.clientX + 60, clientY: start.clientY + 30, buttons: 1 };
  await act(async () => { fireEvent.pointerDown(handle, start); await h.client.flush(); });
  await waitFor(() => expect(h.client.state.selection?.kind).toBe('object'));
  await act(async () => { fireEvent.pointerMove(h.stage, end); });
  // The drag keeps the text-edit look: a dashed frame, not the solid object border.
  const frame = h.view.getByTestId('pptx-shape-selection');
  expect(frame.style.border).toBe('1px dashed #7a7a7a');
  await act(async () => { fireEvent.pointerUp(h.stage, end); await h.api.flushPendingInput(); });
  expect(h.commands.filter((command) => command.kind === 'setShapeRect')).toHaveLength(1);
  await waitFor(() => expect(h.client.state.selection).toMatchObject({ shapeId: h.box.shapeId, anchor: 2, focus: 2 }));
  expect(h.client.state.selection?.kind).not.toBe('object');
  await waitFor(() => expect(shownBox(h).w).toBeCloseTo(before.w + 60, 0));
  await waitFor(() => expect(h.view.getAllByTestId(/pptx-resize-/)).toHaveLength(8));
  const input = h.view.getByRole('textbox', { name: 'Slide text' });
  await waitFor(() => expect(document.activeElement === input).toBe(true));
  await act(async () => { type(input, 'Z'); await h.api.flushPendingInput(); });
  const read = await h.api.readContent();
  if (!read.ok) throw new Error(read.failure.message);
  const story = read.stories.find((value) => value.storyId === h.box.storyId)!;
  expect(story.text[2]).toBe('Z');
});

test('clicking a handle without dragging leaves text editing as it was', async () => {
  const h = await open();
  await editExisting(h);
  const handle = h.view.getByTestId('pptx-resize-e');
  const at = h.point(h.box.x + h.box.w, h.box.y + h.box.h / 2);
  await act(async () => { fireEvent.pointerDown(handle, at); await h.client.flush(); });
  await act(async () => { fireEvent.pointerUp(h.stage, at); await h.api.flushPendingInput(); });
  expect(h.commands.filter((command) => command.kind === 'setShapeRect')).toHaveLength(0);
  await waitFor(() => expect(h.client.state.selection).toMatchObject({ shapeId: h.box.shapeId, anchor: 2, focus: 2 }));
  expect(h.client.state.selection?.kind).not.toBe('object');
});
