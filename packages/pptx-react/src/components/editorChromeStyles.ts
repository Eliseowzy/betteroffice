import type { CSSProperties } from 'react';
import { editorTheme } from './editorTheme';

export const notesStyles: Record<'panel' | 'label' | 'textarea', CSSProperties> = {
  panel: {
    flex: '0 0 auto',
    display: 'flex',
    flexDirection: 'column',
    gap: 4,
    padding: '8px 14px 10px',
    background: editorTheme.surface,
    borderTop: `1px solid ${editorTheme.line}`,
  },
  label: {
    fontSize: 11,
    fontWeight: 650,
    color: editorTheme.inkMuted,
    textTransform: 'uppercase',
    letterSpacing: '0.02em',
  },
  textarea: {
    width: '100%',
    height: 64,
    resize: 'vertical',
    border: `1px solid ${editorTheme.line}`,
    background: editorTheme.surface,
    borderRadius: 6,
    padding: '6px 8px',
    font: '13px ui-sans-serif, system-ui, sans-serif',
    color: editorTheme.ink,
    boxSizing: 'border-box',
    outline: 'none',
  },
};

export const slideActionStyles: Record<'row' | 'button', CSSProperties> = {
  row: { display: 'flex', flexWrap: 'wrap', gap: 6, paddingTop: 8 },
  button: {
    flex: '1 1 auto',
    height: 26,
    whiteSpace: 'nowrap',
    padding: '0 8px',
    border: `1px solid ${editorTheme.control}`,
    borderRadius: 5,
    background: editorTheme.surface,
    color: editorTheme.controlInk,
    font: '12px ui-sans-serif, system-ui, sans-serif',
    cursor: 'pointer',
  },
};
