//! Port of `packages/docx/src/yrs/yrsToDocument.ts` (part).
#![allow(unused_variables, unused_imports, dead_code, clippy::all)]

use std::collections::{HashMap, HashSet};
use std::rc::Rc;
use indexmap::IndexMap;
use crate::js::{Js, Result, SaveError};
use crate::session::{Inheritance, Session};
use crate::{arr, obj, throw};
use super::*;

/// `collectBaseParagraphs` (src/yrs/yrsToDocument.ts:2908).
pub(crate) fn collect_base_paragraphs(stories: &Js) -> Result<Js> {
    todo!("collectBaseParagraphs")
}

/// `collectBaseStories` (src/yrs/yrsToDocument.ts:2926).
pub(crate) fn collect_base_stories(session: &Session, document: &Js) -> Result<Js> {
    todo!("collectBaseStories")
}

/// `commentRanges` (src/yrs/yrsToDocument.ts:2970).
pub(crate) fn comment_ranges(session: &Session, comments: &Js, session_ids: &Js) -> Result<Js> {
    todo!("commentRanges")
}

/// `collectChartDrawings` (src/yrs/yrsToDocument.ts:3008).
pub(crate) fn collect_chart_drawings(blocks: &Js, into: &IndexMap<String, String>) -> Result<()> {
    todo!("collectChartDrawings")
}

/// `paragraphChartRuns` (src/yrs/yrsToDocument.ts:3032).
pub(crate) fn paragraph_chart_runs(content: &Js) -> Result<Js> {
    todo!("paragraphChartRuns")
}

/// `isBlockBookmark` (src/yrs/yrsToDocument.ts:3055).
pub(crate) fn is_block_bookmark(block: &Js) -> Result<bool> {
    todo!("isBlockBookmark")
}

/// `containsBookmark` (src/yrs/yrsToDocument.ts:3059).
pub(crate) fn contains_bookmark(blocks: &Js) -> Result<bool> {
    todo!("containsBookmark")
}

/// `assertCommentBookmarkReplacement` (src/yrs/yrsToDocument.ts:3068).
pub(crate) fn assert_comment_bookmark_replacement(blocks: &Js) -> Result<()> {
    todo!("assertCommentBookmarkReplacement")
}

/// `rawBlockBookmark` (src/yrs/yrsToDocument.ts:3078).
pub(crate) fn raw_block_bookmark(xml: &str) -> Result<Js> {
    todo!("rawBlockBookmark")
}

/// `storyBlockKeys` (src/yrs/yrsToDocument.ts:3110).
pub(crate) fn story_block_keys(blocks: &Js, story_id: &str) -> Result<Js> {
    todo!("storyBlockKeys")
}

/// `projectedBookmarkPositions` (src/yrs/yrsToDocument.ts:3135).
pub(crate) fn projected_bookmark_positions(blocks: &Js) -> Result<Js> {
    todo!("projectedBookmarkPositions")
}

/// `restoreBlockBookmarks` (src/yrs/yrsToDocument.ts:3182).
pub(crate) fn restore_block_bookmarks(projected: &Js, base: &Js, story_id: &str) -> Result<Js> {
    todo!("restoreBlockBookmarks")
}

/// `restoreTextBoxBookmarks` (src/yrs/yrsToDocument.ts:3257).
pub(crate) fn restore_text_box_bookmarks(content: &Js, base: &Js) -> Result<()> {
    todo!("restoreTextBoxBookmarks")
}

/// `restoreRawBlocks` (src/yrs/yrsToDocument.ts:3287).
pub(crate) fn restore_raw_blocks(projected: &Js, base: &Js, story_id: &str) -> Result<Js> {
    todo!("restoreRawBlocks")
}

/// `cloneProjectedParagraph` (src/yrs/yrsToDocument.ts:3341).
pub(crate) fn clone_projected_paragraph(paragraph: &Js, changes: &Js, identity: &Js) -> Result<Js> {
    todo!("cloneProjectedParagraph")
}

/// `projectedSessionKey` (src/yrs/yrsToDocument.ts:3355).
pub(crate) fn projected_session_key(paragraph: &Js) -> Result<Option<String>> {
    todo!("projectedSessionKey")
}

/// `carryParagraphIdentities` (src/yrs/yrsToDocument.ts:3363).
pub(crate) fn carry_paragraph_identities(replacement: &Js, replaced: &Js) -> Result<Js> {
    todo!("carryParagraphIdentities")
}

/// `ownProjectedParagraphs` (src/yrs/yrsToDocument.ts:3382).
pub(crate) fn own_projected_paragraphs(blocks: &Js) -> Result<Js> {
    todo!("ownProjectedParagraphs")
}

/// `sessionProjectionMemo` (src/yrs/yrsToDocument.ts:3407).
pub(crate) fn session_projection_memo(session: &Session) -> Result<Js> {
    todo!("sessionProjectionMemo")
}

/// `sameJson` (src/yrs/yrsToDocument.ts:3423).
pub(crate) fn same_json(a: &Js, b: &Js) -> Result<bool> {
    todo!("sameJson")
}

/// `isBlockBreak` (src/yrs/yrsToDocument.ts:3441).
pub(crate) fn is_block_break(item: &Js) -> Result<bool> {
    todo!("isBlockBreak")
}

/// `blockBreakKind` (src/yrs/yrsToDocument.ts:3446).
pub(crate) fn block_break_kind(embed_kind: &str) -> Result<Js> {
    todo!("blockBreakKind")
}

/// `runFlowTokens` (src/yrs/yrsToDocument.ts:3452).
pub(crate) fn run_flow_tokens(run: &Js, tokens: &Js) -> Result<()> {
    todo!("runFlowTokens")
}

/// `flowTokens` (src/yrs/yrsToDocument.ts:3460).
pub(crate) fn flow_tokens(content: &Js, tokens: &Js) -> Result<Js> {
    todo!("flowTokens")
}

/// `leadingBlockBreaks` (src/yrs/yrsToDocument.ts:3477).
pub(crate) fn leading_block_breaks(paragraph: &Js) -> Result<Js> {
    todo!("leadingBlockBreaks")
}

/// `trailingBlockBreaks` (src/yrs/yrsToDocument.ts:3484).
pub(crate) fn trailing_block_breaks(paragraph: &Js) -> Result<Js> {
    todo!("trailingBlockBreaks")
}

/// `noteYrsStoriesDirty` (src/yrs/yrsToDocument.ts:3499).
pub(crate) fn note_yrs_stories_dirty(session: &Session, stories: &Js) -> Result<()> {
    todo!("noteYrsStoriesDirty")
}

/// `yrsToDocument` (src/yrs/yrsToDocument.ts:4243).
pub(crate) fn yrs_to_document(session: &Session, base: &Js, options: &Js) -> Result<Js> {
    todo!("yrsToDocument")
}

/// `yrsBodyToDocumentWithRevisionIds` (src/yrs/yrsToDocument.ts:4255).
pub(crate) fn yrs_body_to_document_with_revision_ids(session: &Session, base: &Js, revision_ids: &Js) -> Result<Js> {
    todo!("yrsBodyToDocumentWithRevisionIds")
}

/// `projectDocument` (src/yrs/yrsToDocument.ts:4263).
pub(crate) fn project_document(session: &Session, base: &Js, options: &Js, revision_ids: &Js) -> Result<Js> {
    todo!("projectDocument")
}

/// `withHostSectionChanges` (src/yrs/yrsToDocument.ts:4281).
pub(crate) fn with_host_section_changes(authored: &Js, source: &Js, host: &Js) -> Result<Js> {
    todo!("withHostSectionChanges")
}

/// `projectStories` (src/yrs/yrsToDocument.ts:4292).
pub(crate) fn project_stories(context: &crate::yrs_to_document::SaveContext, session: &Session, base: &Js, options: &Js) -> Result<Js> {
    todo!("projectStories")
}

/// `class SaveContext`.
pub(crate) struct SaveContext {
    _todo: (),
}

impl SaveContext {
    /// `SaveContext.sourceTable` (src/yrs/yrsToDocument.ts:3550).
    pub(crate) fn source_table(&self, prefix: &str) -> Result<Js> {
        todo!("SaveContext.sourceTable")
    }

    /// `SaveContext.constructor` (src/yrs/yrsToDocument.ts:3559).
    pub(crate) fn new(session: &Session, base: &Js, comment_ids: &Js, revision_ids: &Js) -> Result<Self> {
        todo!("SaveContext.constructor")
    }

    /// `SaveContext.storyIsClean` (src/yrs/yrsToDocument.ts:3617).
    pub(crate) fn story_is_clean(&self, story_id: &str) -> Result<bool> {
        todo!("SaveContext.storyIsClean")
    }

    /// `SaveContext.sectionChanged` (src/yrs/yrsToDocument.ts:3623).
    pub(crate) fn section_changed(&self, para_id: &str, properties: &Js) -> Result<bool> {
        todo!("SaveContext.sectionChanged")
    }

    /// `SaveContext.inheritance` (src/yrs/yrsToDocument.ts:3627).
    pub(crate) fn inheritance(&self) -> Result<Js> {
        todo!("SaveContext.inheritance")
    }

    /// `SaveContext.prepareTable` (src/yrs/yrsToDocument.ts:3631).
    pub(crate) fn prepare_table(&self, payload: &Js, parent: &str) -> Result<()> {
        todo!("SaveContext.prepareTable")
    }

    /// `SaveContext.applyTableBaseline` (src/yrs/yrsToDocument.ts:3680).
    pub(crate) fn apply_table_baseline(&self, payload: &Js, attrs: &Js) -> Result<()> {
        todo!("SaveContext.applyTableBaseline")
    }

    /// `SaveContext.cellAttrs` (src/yrs/yrsToDocument.ts:3687).
    pub(crate) fn cell_attrs(&self, payload: &Js) -> Result<Js> {
        todo!("SaveContext.cellAttrs")
    }

    /// `SaveContext.cellContents` (src/yrs/yrsToDocument.ts:3709).
    pub(crate) fn cell_contents(&self, payload: &Js, parent: &str) -> Result<Js> {
        todo!("SaveContext.cellContents")
    }

    /// `SaveContext.restoreCommentReferences` (src/yrs/yrsToDocument.ts:3724).
    pub(crate) fn restore_comment_references(&self, paragraph: &Js) -> Result<()> {
        todo!("SaveContext.restoreCommentReferences")
    }

    /// `SaveContext.inheritedFormatting` (src/yrs/yrsToDocument.ts:3762).
    pub(crate) fn inherited_formatting(&self, story_id: &str, properties: &Js) -> Result<Js> {
        todo!("SaveContext.inheritedFormatting")
    }

    /// `SaveContext.copiedParagraphProperties` (src/yrs/yrsToDocument.ts:3798).
    pub(crate) fn copied_paragraph_properties(&self, story: &str, properties: &Js, items: &Js, preceding: &Js) -> Result<Js> {
        todo!("SaveContext.copiedParagraphProperties")
    }

    /// `SaveContext.runFormatter` (src/yrs/yrsToDocument.ts:3871).
    pub(crate) fn run_formatter(&self, story: &str, properties: &Js) -> Result<Js> {
        todo!("SaveContext.runFormatter")
    }

    /// `SaveContext.inheritedRun` (src/yrs/yrsToDocument.ts:3881).
    pub(crate) fn inherited_run(&self, paragraph_style: Option<&str>, run_style: Option<&str>, table_run: &Js) -> Result<Js> {
        todo!("SaveContext.inheritedRun")
    }

    /// `SaveContext.dispose` (src/yrs/yrsToDocument.ts:3892).
    pub(crate) fn dispose(&self) -> Result<()> {
        todo!("SaveContext.dispose")
    }

    /// `SaveContext.recovery` (src/yrs/yrsToDocument.ts:3897).
    pub(crate) fn recovery(&self, story_id: &str) -> Result<Js> {
        todo!("SaveContext.recovery")
    }

    /// `SaveContext.storyToBlocks` (src/yrs/yrsToDocument.ts:3915).
    pub(crate) fn story_to_blocks(&self, story_id: &str) -> Result<Js> {
        todo!("SaveContext.storyToBlocks")
    }

}
