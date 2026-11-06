import type { ShapeRect } from '@betteroffice/pptx';
import type { PresentationCommandObjectSelection } from '@betteroffice/pptx/worker';
import { useSyncExternalStore } from 'react';
import { handleAnchor, RESIZE_HANDLES, resizeCursor } from '../interactions';
import type { PeerlessShapePreviewStore } from './usePeerlessShapeGestures';
import { useTranslation } from '../i18n';

/** The eight PowerPoint-style square resize handles around `bounds` (slide pixels). */
export function PeerlessResizeHandles({ bounds, scale }: { bounds: ShapeRect; scale: number }) {
  const { t } = useTranslation();
  return <>{RESIZE_HANDLES.map((handle) => {
    const point = handleAnchor(bounds, handle);
    return <span key={handle} data-testid={`pptx-resize-${handle}`} data-pptx-resize={handle}
      role="img" aria-label={t(`objects.resize.${handle}`)}
      style={{ position: 'absolute', zIndex: 4, pointerEvents: 'auto', left: point.x * scale - 4, top: point.y * scale - 4,
        width: 8, height: 8, borderRadius: 1, border: '1px solid #7a7a7a', background: '#fff',
        boxSizing: 'border-box', cursor: resizeCursor(handle), touchAction: 'none' }} />;
  })}</>;
}

export function PeerlessShapeOverlay({ selection, preview: previews, scale, readOnly }: {
  selection: PresentationCommandObjectSelection; preview: PeerlessShapePreviewStore; scale: number; readOnly: boolean;
}) {
  const preview = useSyncExternalStore(previews.subscribe, previews.getSnapshot, previews.getSnapshot);
  const original = selection.bounds;
  const current = preview?.token === selection.token ? preview : null;
  const bounds = current ? current.bounds : original;
  // A resize started from inside a text box keeps the text-edit chrome (dashed outline) throughout.
  const editing = current?.editing === true;
  return <>
    <div data-testid="pptx-shape-selection" style={{ position: 'absolute', pointerEvents: 'none',
      left: original.x * scale, top: original.y * scale, width: original.width * scale, height: original.height * scale,
      border: `1px ${editing ? 'dashed' : 'solid'} #7a7a7a`, boxSizing: 'border-box', transformOrigin: '0 0',
      transform: `translate(${(bounds.x - original.x) * scale}px, ${(bounds.y - original.y) * scale}px) scale(${original.width ? bounds.width / original.width : 1}, ${original.height ? bounds.height / original.height : 1})` }} />
    {!readOnly && selection.canResize && <PeerlessResizeHandles bounds={bounds} scale={scale} />}
  </>;
}
