/**
 * The editor's chrome palette (stage, slide strip, notes, panels) as CSS custom
 * properties, the way `toolbarTheme` does it for the toolbars. A host gives the
 * editor a dark (or any) look by setting `--bo-editor-*` on an ancestor; the
 * second value is the light default. Slides draw in their own colours.
 */
export const editorTheme = {
  ink: 'var(--bo-editor-ink, #172033)',
  inkMuted: 'var(--bo-editor-ink-muted, #647087)',
  /** Thumbnail captions and secondary labels. */
  inkSoft: 'var(--bo-editor-ink-soft, #39445a)',
  /** Behind the slide. */
  stage: 'var(--bo-editor-stage, #e2e8f0)',
  /** The slide strip. */
  rail: 'var(--bo-editor-rail, #eef1f5)',
  /** Panels, fields and buttons. */
  surface: 'var(--bo-editor-surface, #ffffff)',
  /** Dialogs a step off the surface. */
  raised: 'var(--bo-editor-raised, #f8fafc)',
  line: 'var(--bo-editor-line, #d8dee9)',
  /** Button and field borders. */
  control: 'var(--bo-editor-control, #cbd5e1)',
  controlInk: 'var(--bo-editor-control-ink, #334155)',
  selected: 'var(--bo-editor-selected, #e8f0fe)',
  accent: 'var(--bo-editor-accent, #325ee6)',
  /** Agent proposals: their cards, links and the strip's marks. */
  proposal: 'var(--bo-editor-proposal, #4338ca)',
  proposalSelected: 'var(--bo-editor-proposal-selected, #e0e7ff)',
  slideShadow: 'var(--bo-editor-slide-shadow, 0 8px 32px rgba(27, 39, 61, 0.2))',
  thumbShadow: 'var(--bo-editor-thumb-shadow, 0 1px 4px rgba(20, 31, 50, 0.16))',
} as const;
