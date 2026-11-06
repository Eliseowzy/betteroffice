import { expect, test } from 'bun:test';
import { createT, zhCN, type LocaleStrings } from '@betteroffice/pptx-i18n';
import { evaluatePptxCommand } from './evaluate';
import { testEnvironment } from './testing';

const themeFonts = {
  major: { latin: 'Calibri Light', eastAsian: '等线 Light' },
  minor: { latin: 'Calibri', eastAsian: '等线' },
};

function themed(fontTheme: 'major' | 'minor' | null, eastAsianFontTheme: 'major' | 'minor' | null = null) {
  const base = testEnvironment();
  return testEnvironment({
    themeFonts,
    text: base.text && base.text !== 'unsupported'
      ? { ...base.text, fontFamily: fontTheme ? themeFonts[fontTheme].latin : 'Arial', fontTheme, eastAsianFontTheme }
      : null,
  });
}

test('fontFamily reports the theme slot beside the resolved family', () => {
  const state = evaluatePptxCommand('fontFamily', undefined, themed('major', 'major'));
  expect(state.value).toBe('Calibri Light');
  expect(state.theme).toBe('major');
  expect(state.eastAsianTheme).toBe('major');
  expect(state.themeFonts).toEqual(themeFonts);
  const literal = evaluatePptxCommand('fontFamily', undefined, themed(null));
  expect(literal.value).toBe('Arial');
  expect(literal.theme).toBeNull();
});

test('a mixed selection reports no theme slot', () => {
  const base = testEnvironment();
  const env = testEnvironment({
    themeFonts,
    text: base.text && base.text !== 'unsupported' ? { ...base.text, fontFamily: null, fontTheme: null } : null,
  });
  const state = evaluatePptxCommand('fontFamily', undefined, env);
  expect(state.value).toBeNull();
  expect(state.theme).toBeNull();
});

test('theme fonts lead the options and match by slot', () => {
  const env = themed('minor');
  const state = evaluatePptxCommand('fontFamily', undefined, env);
  expect(state.options?.slice(0, 3).map((option) => [option.args, option.label])).toEqual([
    [{ theme: 'major' }, 'Calibri Light (Headings)'],
    [{ theme: 'minor' }, 'Calibri (Body)'],
    [{ family: 'Arial' }, 'Arial'],
  ]);
  expect(evaluatePptxCommand('fontFamily', { theme: 'minor' }, env).active).toBe(true);
  expect(evaluatePptxCommand('fontFamily', { theme: 'major' }, env).active).toBe(false);
  expect(evaluatePptxCommand('fontFamily', { family: 'Calibri' }, env).active).toBe(true);
});

test('theme font labels are localized', () => {
  const env = testEnvironment({ ...themed('major'), translate: createT(zhCN as LocaleStrings, 'zh-CN') });
  const labels = evaluatePptxCommand('fontFamily', undefined, env).options?.slice(0, 2).map((option) => option.label);
  expect(labels).toEqual(['Calibri Light (标题)', 'Calibri (正文)']);
});

test('an unknown theme slot is invalid', () => {
  const state = evaluatePptxCommand('fontFamily', { theme: 'other' } as never, themed('major'));
  expect(state.enabled ? null : state.disabledReason.code).toBe('invalid-arguments');
});

test('without theme fonts the options stay literal', () => {
  const state = evaluatePptxCommand('fontFamily', undefined, testEnvironment());
  expect(state.options?.every((option) => 'family' in option.args)).toBe(true);
  expect(state.themeFonts).toBeNull();
});
