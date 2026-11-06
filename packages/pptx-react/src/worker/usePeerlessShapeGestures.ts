import { useCallback, useEffect, useRef, type KeyboardEvent, type PointerEvent, type RefObject } from 'react';
import type { ShapeRect } from '@betteroffice/pptx';
import type { PresentationCommandFrameTag, PresentationCommandObjectSelection, PresentationCommandSelection } from '@betteroffice/pptx/worker';
import { MIN_SHAPE_SIZE_EMU, passedDragThreshold, RESIZE_HANDLES, resizeCommitDelta, resizedBounds,
  slidePoint, type ResizeHandle, type SlidePoint } from '../interactions';
import { hitConfirmedGeometry, peerlessGestureActive, peerlessInputReference, setPeerlessGestureActive, type PeerlessViewStore } from './peerlessViewStore';
import type { PeerlessPresentation } from './usePeerlessPresentation';

interface ShapeGesture {
  pointerId: number;
  element: HTMLElement;
  group: string;
  frame: PresentationCommandFrameTag;
  start: SlidePoint;
  client: SlidePoint;
  delta: SlidePoint;
  shapeId: string;
  dragThreshold: number;
  handle?: ResizeHandle;
  object: Promise<PresentationCommandObjectSelection | null>;
  selected: PresentationCommandObjectSelection | null;
  dragging: boolean;
  cancelled: boolean;
  /** Set when the drag started on a handle while editing text: the caret to put back afterwards. */
  restore?: PresentationCommandSelection;
  preview?: PeerlessShapePreview;
}

export interface PeerlessShapePreview {
  token: string;
  bounds: ShapeRect;
  /** The resize started from text editing, so the overlay keeps the text-edit chrome. */
  editing?: boolean;
}

export interface PeerlessShapePreviewStore {
  getSnapshot(): PeerlessShapePreview | null;
  subscribe(listener: () => void): () => void;
}

let nextShapeGroup = 0;

export function peerlessGestureRect(object: PresentationCommandObjectSelection, delta: SlidePoint,
  handle?: ResizeHandle): ShapeRect | null {
  if (!object.rect || !object.canMove || handle && !object.canResize) return null;
  const emu = { x: Math.round(delta.x * object.units.x), y: Math.round(delta.y * object.units.y) };
  if (emu.x === 0 && emu.y === 0) return { ...object.rect };
  return handle ? resizedBounds(object.rect, handle, emu, MIN_SHAPE_SIZE_EMU, object.aspectRatioLocked)
    : { ...object.rect, x: object.rect.x + emu.x, y: object.rect.y + emu.y };
}

function previewBounds(object: PresentationCommandObjectSelection, rect: ShapeRect, resize: boolean): ShapeRect {
  return resize ? { x: rect.x / object.units.x, y: rect.y / object.units.y,
    width: rect.width / object.units.x, height: rect.height / object.units.y }
    : { ...object.bounds, x: object.bounds.x + (rect.x - object.rect!.x) / object.units.x,
      y: object.bounds.y + (rect.y - object.rect!.y) / object.units.y };
}

export function usePeerlessShapeGestures(store: PeerlessViewStore | null, {
  canvas, presentation, focus,
}: { canvas: RefObject<HTMLCanvasElement | null>; presentation: PeerlessPresentation | null; focus?: () => void }) {
  const gesture = useRef<ShapeGesture | null>(null);
  const outstanding = useRef(new Set<ShapeGesture>());
  const recent = useRef<{ slideId: string; shapeId: string; x: number; y: number; time: number } | null>(null);
  const previewValue = useRef<PeerlessShapePreview | null>(null);
  const previewListeners = useRef(new Set<() => void>());
  const preview = useRef<PeerlessShapePreviewStore>({
    getSnapshot: () => previewValue.current,
    subscribe: (listener) => {
      previewListeners.current.add(listener);
      return () => { previewListeners.current.delete(listener); };
    },
  });
  const setPreview = useCallback((value: PeerlessShapePreview | null) => {
    if (previewValue.current === value) return;
    previewValue.current = value;
    for (const listener of previewListeners.current) listener();
  }, []);
  const showPreview = useCallback((current: ShapeGesture, value: PeerlessShapePreview) => {
    current.preview = value;
    setPreview(value);
  }, [setPreview]);

  // A handle drag from text editing goes through object selection (the only path that can resize);
  // once it settles, put the caret back so typing continues where it was.
  const restoreText = useCallback((current: ShapeGesture, inputTime: number, work?: Promise<unknown>) => {
    const selection = current.restore;
    if (!store || !selection) return;
    void Promise.all([current.object, work]).catch(() => {}).then(() => {
      const view = store.getSnapshot();
      if (view.selection?.kind !== 'object' || view.selection.slideId !== selection.slideId ||
        view.selection.shapeId !== selection.shapeId) return;
      return store.dispatch({ kind: 'setSelection', payload: { selection, inputTime, undoGroup: current.group } });
    }).catch(() => {}).finally(() => {
      if (current.preview && previewValue.current === current.preview) setPreview(null);
      if (gesture.current === null) focus?.();
    });
  }, [store, setPreview, focus]);

  const cancel = useCallback(() => {
    const current = gesture.current;
    if (store) setPeerlessGestureActive(store, 'shape', false);
    gesture.current = null;
    recent.current = null;
    for (const pending of outstanding.current) pending.cancelled = true;
    outstanding.current.clear();
    if (current) current.element.releasePointerCapture?.(current.pointerId);
    if (current?.restore) restoreText(current, performance.now());
    else setPreview(null);
  }, [store, setPreview, restoreText]);

  useEffect(() => {
    setPreview(null);
    if (!store) return;
    const changed = () => {
      const current = gesture.current;
      const view = store.getSnapshot();
      const frame = view.confirmed?.frame;
      if (current && (view.editState.readOnly || !frame || frame.slideId !== current.frame.slideId ||
        frame.slideRevision !== current.frame.slideRevision ||
        current.selected && (view.selection?.kind !== 'object' || view.selection.token !== current.selected.token))) cancel();
    };
    const unsubscribe = [store.subscribePresentation(changed), store.subscribeEditState(changed)];
    window.addEventListener('blur', cancel);
    return () => {
      window.removeEventListener('blur', cancel);
      for (const release of unsubscribe) release();
      const current = gesture.current;
      setPeerlessGestureActive(store, 'shape', false);
      gesture.current = null;
      for (const pending of outstanding.current) pending.cancelled = true;
      outstanding.current.clear();
      recent.current = null;
      current?.element.releasePointerCapture?.(current.pointerId);
    };
  }, [store, presentation, cancel, setPreview]);

  const position = useCallback((event: PointerEvent<HTMLElement>): SlidePoint | null => {
    const confirmed = presentation?.getVisualSnapshot()?.view.confirmed;
    return canvas.current && confirmed
      ? slidePoint(canvas.current.getBoundingClientRect(), confirmed.displayList, event.clientX, event.clientY) : null;
  }, [canvas, presentation]);

  const paint = useCallback((current: ShapeGesture) => {
    if (gesture.current !== current || current.cancelled || !current.dragging || !current.selected ||
      store?.getSnapshot().editState.readOnly) return;
    const rect = peerlessGestureRect(current.selected, current.delta, current.handle);
    if (rect) showPreview(current, { token: current.selected.token, bounds: previewBounds(current.selected, rect, Boolean(current.handle)),
      ...(current.restore ? { editing: true } : {}) });
  }, [store, showPreview]);

  const onPointerDown = useCallback((event: PointerEvent<HTMLElement>): boolean => {
    if (!store || !presentation?.selectionEnabled || !canvas.current || event.button !== 0 || event.isPrimary === false ||
      store.getSnapshot().composition) return false;
    const visual = presentation.getVisualSnapshot();
    const confirmed = visual?.view.confirmed;
    const start = position(event);
    if (!confirmed || !start) return false;
    const handle = (event.target as HTMLElement).dataset?.pptxResize as ResizeHandle | undefined;
    if (handle && !RESIZE_HANDLES.includes(handle)) return false;
    const selected = store.getSnapshot().selection;
    const scale = canvas.current.getBoundingClientRect().width / confirmed.displayList.width;
    // While the caret is in a text box its handles resize it too, through the same object path.
    const editing = handle && selected && selected.kind !== 'object' && selected.slideId === confirmed.frame.slideId
      ? selected : null;
    if (handle && store.getSnapshot().editState.readOnly) return true;
    if (handle && !editing && (selected?.kind !== 'object' || !selected.canResize)) return true;
    const hit = handle ? null : hitConfirmedGeometry(store, start, scale);
    if (!handle && hit?.kind !== 'shape') return false;
    const shapeId = handle && selected ? selected.shapeId : hit!.shapeId;
    const previous = recent.current;
    const repeated = !handle && previous?.slideId === confirmed.frame.slideId && previous.shapeId === shapeId &&
      event.timeStamp >= previous.time && event.timeStamp - previous.time <= 600 &&
      Math.hypot(event.clientX - previous.x, event.clientY - previous.y) <= 6;
    cancel();
    const { slideId, slideRevision, coversThroughSeq, documentEpoch, incarnation } = confirmed.frame;
    const frame = { slideId, slideRevision, coversThroughSeq, documentEpoch, incarnation };
    const group = `shape-${++nextShapeGroup}`;
    const operation = handle && selected?.kind === 'object' ? Promise.resolve(selected)
      : store.dispatch({ kind: 'selectObject', payload: { frame, ...(editing ? { shapeId } : { x: start.x, y: start.y, scale }),
        inputTime: event.timeStamp, undoGroup: group } }).then((ack) =>
        ack.outcomes.some((outcome) => outcome.status === 'refused' || outcome.status === 'cancelled') ||
        ack.selection?.kind !== 'object' ? null : ack.selection, () => null);
    const current: ShapeGesture = { pointerId: event.pointerId, element: event.currentTarget, frame, group,
      start, client: { x: event.clientX, y: event.clientY }, delta: { x: 0, y: 0 },
      shapeId, dragThreshold: repeated ? 8 : 4,
      handle, object: operation, selected: handle && selected?.kind === 'object' ? selected : null,
      dragging: false, cancelled: false, ...(editing ? { restore: { ...editing } } : {}) };
    gesture.current = current;
    setPeerlessGestureActive(store, 'shape', true);
    outstanding.current.add(current);
    void operation.then((object) => {
      current.selected = object;
      if (current.restore && object && gesture.current === current && !current.cancelled) {
        showPreview(current, { token: object.token, bounds: object.bounds, editing: true });
      }
      paint(current);
    });
    event.preventDefault();
    event.currentTarget.focus();
    event.currentTarget.setPointerCapture?.(event.pointerId);
    return true;
  }, [store, presentation, canvas, cancel, position, paint, showPreview]);

  const hoverCursor = useCallback((event: PointerEvent<HTMLElement>): string | null => {
    if (event.buttons !== 0 || (store && peerlessGestureActive(store))) return null;
    const confirmed = presentation?.getVisualSnapshot()?.view.confirmed;
    const point = position(event);
    if (!store || !presentation?.selectionEnabled || !canvas.current || !confirmed || !point) return 'default';
    const hit = hitConfirmedGeometry(store, point, canvas.current.getBoundingClientRect().width / confirmed.displayList.width);
    return hit?.kind === 'text' ? 'text' : hit?.kind === 'shape' ? 'move' : 'default';
  }, [store, presentation, canvas, position]);

  const onPointerMove = useCallback((event: PointerEvent<HTMLElement>): boolean => {
    const current = gesture.current;
    if (!current || current.pointerId !== event.pointerId) return false;
    const point = position(event);
    if (!point) return true;
    current.dragging ||= passedDragThreshold(current.client.x, current.client.y, event.clientX, event.clientY, current.dragThreshold);
    current.delta = { x: point.x - current.start.x, y: point.y - current.start.y };
    paint(current);
    return true;
  }, [position, paint]);

  const onPointerUp = useCallback((event: PointerEvent<HTMLElement>): boolean => {
    const current = gesture.current;
    if (!store || !current || current.pointerId !== event.pointerId) return false;
    const point = position(event);
    const delta = resizeCommitDelta(current.start, current.delta, point);
    const dragging = current.dragging || passedDragThreshold(current.client.x, current.client.y,
      event.clientX, event.clientY, current.dragThreshold);
    const inputTime = event.timeStamp;
    gesture.current = null;
    setPeerlessGestureActive(store, 'shape', false);
    recent.current = !current.handle && !dragging ? { slideId: current.frame.slideId, shapeId: current.shapeId,
      x: event.clientX, y: event.clientY, time: inputTime } : null;
    current.element.releasePointerCapture?.(current.pointerId);
    if (!current.restore) setPreview(null);
    outstanding.current.delete(current);
    const finish = (work?: Promise<unknown>) => {
      if (current.restore) restoreText(current, inputTime, work);
      return true;
    };
    if (!dragging || delta.x === 0 && delta.y === 0) return finish();
    if (current.cancelled || !presentation?.current || store.getSnapshot().editState.readOnly) return finish();
    const object = current.selected;
    if (!object) {
      const at = peerlessInputReference(store, true);
      if (at?.source === 'afterSeq' && !current.handle) {
        void store.dispatch({ kind: 'moveShape', payload: { at, frame: current.frame,
          shapeId: current.shapeId, delta, inputTime, undoGroup: current.group } }).catch(() => {});
      }
      return finish();
    }
    const rect = peerlessGestureRect(object, delta, current.handle);
    if (!rect || Object.keys(rect).every((key) => rect[key as keyof ShapeRect] === object.rect![key as keyof ShapeRect])) return finish();
    const payload = { frame: current.frame,
      target: { slideId: object.slideId, shapeId: object.shapeId, token: object.token },
      inputTime, undoGroup: current.group };
    return finish(store.dispatch(current.handle ? { kind: 'setShapeRect', payload: { ...payload, rect } }
      : { kind: 'moveShape', payload: { ...payload, x: rect.x, y: rect.y } }).catch(() => {}));
  }, [store, presentation, position, setPreview, restoreText]);

  const onPointerCancel = useCallback((event: PointerEvent<HTMLElement>): boolean => {
    if (gesture.current?.pointerId !== event.pointerId) return false;
    cancel();
    return true;
  }, [cancel]);

  const onKeyDown = useCallback((event: KeyboardEvent<HTMLElement>) => {
    if (event.key !== 'Escape' || !outstanding.current.size) return;
    event.preventDefault();
    event.stopPropagation();
    cancel();
  }, [cancel]);

  return { preview: preview.current, hoverCursor, onPointerDown, onPointerMove, onPointerUp, onPointerCancel, onKeyDown,
    onLostPointerCapture: onPointerCancel };
}
