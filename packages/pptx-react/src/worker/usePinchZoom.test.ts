import { expect, test } from 'bun:test';
import { MAX_ZOOM, MIN_ZOOM } from '../commands/evaluate';
import { pinchScale } from './usePinchZoom';

test('pinching out grows the scale, pinching in shrinks it, within the zoom range', () => {
  expect(pinchScale(1, -10, 0)).toBeGreaterThan(1);
  expect(pinchScale(1, 10, 0)).toBeLessThan(1);
  expect(pinchScale(1, -10, 0) * pinchScale(1, 10, 0)).toBeCloseTo(1);
  expect(pinchScale(MAX_ZOOM, -500, 0)).toBe(MAX_ZOOM);
  expect(pinchScale(MIN_ZOOM, 500, 0)).toBe(MIN_ZOOM);
  expect(pinchScale(1, -1, 1)).toBeCloseTo(pinchScale(1, -16, 0));
});
