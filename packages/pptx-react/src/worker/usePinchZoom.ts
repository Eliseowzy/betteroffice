import { useEffect, useRef } from 'react';
import type { RefObject } from 'react';
import { MAX_ZOOM, MIN_ZOOM } from '../commands/evaluate';

/** Wheel distance (in pixels) that scales the slide by a factor of e. */
const PIXELS_PER_E = 100;

/** The next scale for one wheel event, clamped to the zoom range. */
export function pinchScale(base: number, deltaY: number, deltaMode: number): number {
  const pixels = deltaMode === 1 ? deltaY * 16 : deltaMode === 2 ? deltaY * 400 : deltaY;
  return Math.min(MAX_ZOOM, Math.max(MIN_ZOOM, base * Math.exp(-pixels / PIXELS_PER_E)));
}

/**
 * Trackpad pinch (and Ctrl/⌘ + wheel) zooms the stage around the pointer. Chromium reports a
 * pinch as `wheel` events with `ctrlKey`; the listener is non-passive so the page itself does
 * not zoom.
 */
export function usePinchZoom(stage: RefObject<HTMLElement | null>, zoom: number | 'fit',
  fitScale: () => number | null, setZoom: (value: number) => void): void {
  const latest = useRef({ zoom, fitScale, setZoom });
  latest.current = { zoom, fitScale, setZoom };
  // The scale already asked for but maybe not rendered yet, so fast events compound.
  const requested = useRef<number | null>(null);
  useEffect(() => { requested.current = typeof zoom === 'number' ? zoom : null; }, [zoom]);
  // The stage mounts after the editor loads, so bind to whichever element it is now.
  const bound = useRef<{ element: HTMLElement; release(): void } | null>(null);
  useEffect(() => {
    const element = stage.current;
    if (bound.current?.element === element) return;
    bound.current?.release();
    bound.current = null;
    if (!element) return;
    const wheel = (event: WheelEvent) => {
      if (!event.ctrlKey && !event.metaKey) return;
      const { zoom, fitScale, setZoom } = latest.current;
      const base = requested.current ?? (typeof zoom === 'number' ? zoom : fitScale());
      if (!base) return;
      event.preventDefault();
      const next = pinchScale(base, event.deltaY, event.deltaMode);
      if (next === base) return;
      requested.current = next;
      const box = element.getBoundingClientRect();
      const x = event.clientX - box.left;
      const y = event.clientY - box.top;
      const contentX = element.scrollLeft + x;
      const contentY = element.scrollTop + y;
      setZoom(next);
      const ratio = next / base;
      element.ownerDocument.defaultView?.requestAnimationFrame(() => {
        element.scrollLeft = contentX * ratio - x;
        element.scrollTop = contentY * ratio - y;
      });
    };
    element.addEventListener('wheel', wheel, { passive: false });
    bound.current = { element, release: () => element.removeEventListener('wheel', wheel) };
  });
  useEffect(() => () => { bound.current?.release(); bound.current = null; }, []);
}
