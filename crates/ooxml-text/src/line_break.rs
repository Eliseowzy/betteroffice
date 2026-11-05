//! UAX-14 line breaks with kinsoku filtering and Unicode scalar boundaries.

use unicode_linebreak::BreakOpportunity as Uax14Opportunity;

/// One line-break opportunity.
///
/// `byte_index` is the UTF-8 byte position where the next line would start
/// (i.e. the break is *before* this index's character), matching
/// `unicode-linebreak` semantics. Always a `char` boundary. Per UAX-14 the
/// end of text is reported as a final mandatory break at `text.len()`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct BreakOpportunity {
    pub byte_index: usize,
    /// `true` for mandatory breaks (LF, paragraph separator, end of text),
    /// `false` for allowed (optional) break opportunities.
    pub mandatory: bool,
}

/// All UAX-14 break opportunities for `text`, in ascending byte order.
pub fn break_opportunities(text: &str) -> Vec<BreakOpportunity> {
    unicode_linebreak::linebreaks(text)
        .filter(|(byte_index, _)| word_kinsoku_allows(text, *byte_index))
        .map(|(byte_index, kind)| BreakOpportunity {
            byte_index,
            mandatory: matches!(kind, Uax14Opportunity::Mandatory),
        })
        .collect()
}

/// Whether a line may break between `before` and `after` when they meet,
/// e.g. at the seam of two runs.
pub fn break_allowed_between(before: char, after: char) -> bool {
    let pair: String = [before, after].iter().collect();
    break_opportunities(&pair)
        .iter()
        .any(|opportunity| opportunity.byte_index == before.len_utf8())
}

/// UAX-14 keeps a hyphen glued to the number after it, so `COVID-19` never
/// wraps. PowerPoint breaks there, as its own render of the corpus shows, so a
/// slide takes this set instead: the UAX-14 opportunities plus one after every
/// hyphen that a digit follows.
pub fn presentation_break_opportunities(text: &str) -> Vec<BreakOpportunity> {
    presentation_break_opportunities_with(text, true)
}

/// [`presentation_break_opportunities`], where `east_asian` says whether the
/// text carries an East Asian `lang`. Without one PowerPoint wraps it by its
/// Latin rules: a line may also end after an ideograph that full-width
/// punctuation follows (`服务于人|，守护`), or after a curly opening quote an
/// ideograph follows (`测试“|文字`), as its export shows.
pub fn presentation_break_opportunities_with(
    text: &str,
    east_asian: bool,
) -> Vec<BreakOpportunity> {
    let mut opportunities = break_opportunities(text);
    let mut extra: Vec<usize> = Vec::new();
    for (byte_index, character) in text.char_indices() {
        if !east_asian {
            let after = byte_index + character.len_utf8();
            let previous = text[..byte_index].chars().next_back();
            let next = text[after..].chars().next();
            if matches!(character, '\u{201C}' | '\u{2018}')
                && previous.is_some_and(is_ideograph)
                && next.is_some_and(is_ideograph)
            {
                extra.push(after);
            }
            if is_full_width_symbol(character) && previous.is_some_and(is_ideograph) {
                extra.push(byte_index);
            }
        }
        if !matches!(character, '-' | '\u{2010}') {
            continue;
        }
        let after = byte_index + character.len_utf8();
        let follows_word = text[..byte_index]
            .chars()
            .next_back()
            .is_some_and(char::is_alphanumeric);
        let leads_digit = text[after..]
            .chars()
            .next()
            .is_some_and(|next| next.is_ascii_digit());
        if follows_word && leads_digit {
            extra.push(after);
        }
    }
    if extra.is_empty() {
        return opportunities;
    }
    extra.retain(|index| !opportunities.iter().any(|value| value.byte_index == *index));
    opportunities.extend(extra.into_iter().map(|byte_index| BreakOpportunity {
        byte_index,
        mandatory: false,
    }));
    opportunities.sort_by_key(|value| value.byte_index);
    opportunities
}

fn is_full_width_symbol(character: char) -> bool {
    matches!(character as u32,
        0xFF01..=0xFF0F | 0xFF1A..=0xFF20 | 0xFF3B..=0xFF40 | 0xFF5B..=0xFF60)
}

fn is_ideograph(character: char) -> bool {
    matches!(character as u32, 0x3400..=0x4DBF | 0x4E00..=0x9FFF | 0xF900..=0xFAFF | 0x2_0000..=0x3_FFFF)
}

/// Word's default East Asian prohibited-start/prohibited-end refinement.
/// UAX-14 supplies the broad opportunity set; kinsoku removes breaks that
/// would strand opening punctuation at line end or closing punctuation at
/// line start. The terminal break is always retained.
fn word_kinsoku_allows(text: &str, byte_index: usize) -> bool {
    if byte_index == 0 || byte_index >= text.len() {
        return true;
    }
    const PROHIBITED_START: &[char] = &[
        '!', '%', ')', ',', '.', ':', ';', '?', ']', '}', '¢', '°', '’', '”', '†', '‡', '…', '‰',
        '′', '″', '℃', '、', '。', '〉', '》', '」', '』', '】', '〕', '〗', '〙', '〛', '゛',
        '゜', 'ゝ', 'ゞ', 'ー', 'ァ', 'ィ', 'ゥ', 'ェ', 'ォ', 'ッ', 'ャ', 'ュ', 'ョ', 'ヮ', 'ヵ',
        'ヶ', '・', 'ヽ', 'ヾ', '！', '％', '）', '，', '．', '：', '；', '？', '］', '｝',
    ];
    const PROHIBITED_END: &[char] = &[
        '(', '[', '{', '£', '¥', '‘', '“', '〈', '《', '「', '『', '【', '〔', '〖', '〘', '〚',
        '（', '［', '｛', '￥',
    ];
    let previous = text[..byte_index].chars().next_back();
    let next = text[byte_index..].chars().next();
    !previous.is_some_and(|ch| PROHIBITED_END.contains(&ch))
        && !next.is_some_and(|ch| PROHIBITED_START.contains(&ch))
}

#[cfg(test)]
mod word_tests {
    use super::*;

    #[test]
    fn a_slide_breaks_after_a_hyphen_that_a_digit_follows() {
        let indexes = |text: &str| {
            presentation_break_opportunities(text)
                .into_iter()
                .map(|value| value.byte_index)
                .collect::<Vec<_>>()
        };
        assert_eq!(indexes("COVID-19 on"), [6, 9, 11]);
        assert_eq!(indexes("2020-2021 x"), [5, 10, 11]);
        assert_eq!(indexes("well-known thing"), [5, 11, 16]);
        assert_eq!(indexes("-19 x"), [4, 5]);
    }

    #[test]
    fn text_without_an_east_asian_lang_wraps_around_ideographs_by_latin_rules() {
        let indexes = |text: &str, east_asian: bool| {
            presentation_break_opportunities_with(text, east_asian)
                .into_iter()
                .map(|value| value.byte_index)
                .collect::<Vec<_>>()
        };
        let quoted = "测试\u{201C}文字";
        assert!(indexes(quoted, false).contains(&"测试\u{201C}".len()));
        assert!(!indexes(quoted, true).contains(&"测试\u{201C}".len()));
        assert!(!indexes("a \u{201C}word", false).contains(&"a \u{201C}".len()));
        let comma = "于人，守护";
        assert!(indexes(comma, false).contains(&"于人".len()));
        assert!(!indexes(comma, true).contains(&"于人".len()));
    }

    #[test]
    fn kinsoku_filters_breaks_before_closing_and_after_opening_punctuation() {
        let text = "漢（字）漢";
        let breaks = break_opportunities(text);
        let indices: Vec<usize> = breaks.iter().map(|item| item.byte_index).collect();
        let after_open = "漢（".len();
        let before_close = "漢（字".len();
        assert!(!indices.contains(&after_open));
        assert!(!indices.contains(&before_close));
        assert_eq!(indices.last().copied(), Some(text.len()));
    }
}
