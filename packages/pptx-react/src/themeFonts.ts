import type { PptxThemeFontSlot } from '@betteroffice/pptx';

/** The `a:latin` typeface reference that selects a theme font slot. */
export function themeFontReference(slot: PptxThemeFontSlot): '+mj-lt' | '+mn-lt' {
  return slot === 'major' ? '+mj-lt' : '+mn-lt';
}

/** The theme slot a `+mj-lt` / `+mn-lt` typeface refers to; `null` for a literal family. */
export function themeFontSlot(family: string | null | undefined): PptxThemeFontSlot | null {
  const reference = family?.toLowerCase();
  return reference === '+mj-lt' ? 'major' : reference === '+mn-lt' ? 'minor' : null;
}
