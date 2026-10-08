//! Compact paragraph-to-table context acquired from the shared DOC table grammar.
//! [MS-DOC] 2.4.3 and 2.4.6.6. This module does not apply table styles. Pre-existing
//! [`Properties`] allocations must be admitted by the caller. Owned rows are consumed;
//! borrowed rows remain with the caller. The index retains neither and accounts for
//! its own additional storage.

use super::{
    table::Properties,
    table_structure::{Assembler, FirstParagraph, FrameKey, Payload, RawEvent},
    unsupported, MAX_STORY_CONTROLS,
};
use std::borrow::Borrow;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) struct ParagraphContext {
    pub(super) table_id: usize,
    pub(super) row_index: usize,
    pub(super) source_cell_index: Option<usize>,
    /// Grammar block ordinal within that source cell; TTP marks are not blocks.
    pub(super) source_block_index: Option<usize>,
    pub(super) ttp_id: usize,
    pub(super) table_style: Option<usize>,
    /// Optional table-style flags from the same source TTP.
    pub(super) table_style_options: Option<u16>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) struct LogicalColumn {
    pub(super) ordinal: usize,
    pub(super) count: usize,
}

#[derive(Debug, PartialEq, Eq)]
pub(super) struct TableContext {
    pub(super) rows: Vec<RowContext>,
    pub(super) location: Option<TableLocation>,
}

/// Native grammar ancestry, scoped to one `Index` and its owning story. These
/// postorder table IDs and pre-normalization block ordinals are not canonical
/// DOCX SourceRef paths: merges, guards and break normalization need the shared
/// producer/consumer mapping before such a projection can be admitted. One link
/// per table keeps retained storage O(paragraphs + tables), not O(depth * paragraphs).
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) enum TableLocation {
    Root {
        source_block_index: usize,
    },
    Cell {
        table_id: usize,
        row_index: usize,
        source_cell_index: usize,
        source_block_index: usize,
    },
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) struct RowContext {
    pub(super) ttp_id: usize,
    /// Effective sprmTIstd selection resolved from this row's TTP mark.
    pub(super) table_style: Option<usize>,
    /// Optional table-style flags from the same source TTP.
    pub(super) table_style_options: Option<u16>,
    pub(super) source_cell_count: usize,
    pub(super) header: bool,
    /// Number of preceding rows that are neither the actual top row nor headers.
    pub(super) preceding_body_rows: usize,
    /// Scalar native placement from this row's source TTP; no renderer guess.
    pub(super) position: super::table::Position,
    /// The [MS-DOC] 2.4.3 first-cell first paragraph that decided this row's
    /// frame identity. The emitted-table producer replays the same key and
    /// must name the same paragraph, so both share one segmentation.
    pub(super) first: FirstParagraph,
}

/// Story inputs the grammar index needs beyond table/paragraph properties.
pub(super) trait Resolver {
    /// Effective sprmTIstd selection of a row-owning TTP.
    fn table_style(&mut self, selected: Option<usize>) -> Result<Option<usize>, String>;
    /// [MS-DOC] 2.4.3 frame key of a story paragraph; called once per row.
    fn first_paragraph_frame(&mut self, paragraph: usize) -> Result<FrameKey, String>;
}

#[cfg(test)]
pub(super) struct Unframed<S>(pub(super) S);

#[cfg(test)]
impl<S: FnMut(Option<usize>) -> Result<Option<usize>, String>> Resolver for Unframed<S> {
    fn table_style(&mut self, selected: Option<usize>) -> Result<Option<usize>, String> {
        (self.0)(selected)
    }
    fn first_paragraph_frame(&mut self, _: usize) -> Result<FrameKey, String> {
        Ok(FrameKey::default())
    }
}

pub(super) struct Index {
    paragraphs: Vec<Option<ParagraphContext>>,
    tables: Vec<TableContext>,
}

impl Index {
    #[cfg(test)]
    pub(super) fn build<I, A, R>(
        paragraph_count: usize,
        paragraphs: I,
        admit: &mut A,
    ) -> Result<Self, String>
    where
        I: IntoIterator<Item = (Properties<R>, char)>,
        R: Borrow<super::table::Row>,
        A: FnMut(usize) -> Result<(), String>,
    {
        Self::build_with_styles(paragraph_count, paragraphs, &mut Unframed(Ok), admit)
    }

    pub(super) fn build_with_styles<I, S, A, R>(
        paragraph_count: usize,
        paragraphs: I,
        resolver: &mut S,
        admit: &mut A,
    ) -> Result<Self, String>
    where
        I: IntoIterator<Item = (Properties<R>, char)>,
        R: Borrow<super::table::Row>,
        S: Resolver,
        A: FnMut(usize) -> Result<(), String>,
    {
        if paragraph_count > MAX_STORY_CONTROLS + 1 {
            return Err(unsupported("Word table context paragraph budget exceeded"));
        }
        let mut contexts = allocate_none(paragraph_count, admit)?;
        let mut tables = Vec::new();
        let mut assembler = Assembler::<Ids, R>::new();
        let mut actual = 0usize;
        // The grammar calls the frame and table-event callbacks one at a time;
        // both use the story resolver.
        let resolver = std::cell::RefCell::new(resolver);

        for (properties, mark) in paragraphs {
            if actual >= paragraph_count {
                return Err(unsupported("Word table context paragraph count mismatch"));
            }
            let payload = Ids::one(SourceBlock::Paragraph(actual), admit)?;
            assembler.push_raw(
                properties,
                mark,
                payload,
                Some(actual),
                &mut || resolver.borrow_mut().first_paragraph_frame(actual),
                |event, admit| {
                    index_event(
                        event,
                        &mut contexts,
                        &mut tables,
                        &mut **resolver.borrow_mut(),
                        admit,
                    )
                },
                admit,
            )?;
            actual += 1;
        }
        if actual != paragraph_count {
            return Err(unsupported("Word table context paragraph count mismatch"));
        }
        let root = assembler.finish_raw(
            |event, admit| {
                index_event(
                    event,
                    &mut contexts,
                    &mut tables,
                    &mut **resolver.borrow_mut(),
                    admit,
                )
            },
            admit,
        )?;
        for (source_block_index, block) in root.blocks.into_iter().enumerate() {
            if let SourceBlock::Table(table_id) = block {
                set_location(
                    &mut tables,
                    table_id,
                    TableLocation::Root { source_block_index },
                )?;
            }
        }
        if tables.iter().any(|table| table.location.is_none()) {
            return Err(unsupported("Word table context lacks source ancestry"));
        }
        Ok(Self {
            paragraphs: contexts,
            tables,
        })
    }

    pub(super) fn paragraph(&self, id: usize) -> Result<Option<ParagraphContext>, String> {
        self.paragraphs
            .get(id)
            .copied()
            .ok_or_else(|| unsupported("Word table context paragraph outside index"))
    }

    /// Returns the paragraph's row-local source-cell ordinal and count. Row cell
    /// definitions and cell marks retain the source sequence independently for
    /// each row ([MS-DOC] 2.4.3 and 2.4.4); VertMergeOperand also counts every
    /// source cell, including merged cells ([MS-DOC] 2.9.343). Controlled Word
    /// 16.112.4 direct-DOC tests use the same ordinal/count for wide and ragged
    /// rows, horizontal and vertical merge slots, and RTL rows without reversal.
    /// sprmTMerge and VerticalMergeFlag continuation-content suppression are
    /// separate from this source indexing.
    pub(super) fn logical_column(&self, id: usize) -> Result<Option<LogicalColumn>, String> {
        let Some(context) = self.paragraph(id)? else {
            return Ok(None);
        };
        let Some(ordinal) = context.source_cell_index else {
            return Ok(None);
        };
        let table = self
            .tables
            .get(context.table_id)
            .ok_or_else(|| unsupported("Word table column table outside context index"))?;
        let row = table
            .rows
            .get(context.row_index)
            .ok_or_else(|| unsupported("Word table column row outside context index"))?;
        if row.source_cell_count == 0 || ordinal >= row.source_cell_count {
            return Err(unsupported("invalid Word table source column position"));
        }
        Ok(Some(LogicalColumn {
            ordinal,
            count: row.source_cell_count,
        }))
    }

    pub(super) fn tables(&self) -> &[TableContext] {
        &self.tables
    }

    /// The indexed [MS-DOC] 2.4.3 frame key for a paragraph that a second
    /// pass of the same grammar identifies as its row's first-cell first
    /// paragraph. Any disagreement with this index is a producer defect and
    /// fails closed instead of emitting a different segmentation.
    pub(super) fn first_paragraph_frame(&self, id: usize) -> Result<FrameKey, String> {
        let context = self
            .paragraph(id)?
            .ok_or_else(|| unsupported("Word table grammar disagrees with its context index"))?;
        self.tables
            .get(context.table_id)
            .and_then(|table| table.rows.get(context.row_index))
            .filter(|row| row.first.paragraph == Some(id))
            .map(|row| row.first.frame)
            .ok_or_else(|| unsupported("Word table grammar disagrees with its context index"))
    }

    /// Enclosing depth-1 row owner from the grammar, including nested tables.
    /// Unlike a following-row scan, this follows explicit source cell/row links.
    pub(super) fn outer_row_owner(&self, id: usize) -> Result<Option<usize>, String> {
        let Some(context) = self.paragraph(id)? else {
            return Ok(None);
        };
        let (mut table_id, mut row_index) = (context.table_id, context.row_index);
        for _ in 0..32 {
            let table = self
                .tables
                .get(table_id)
                .ok_or_else(|| unsupported("Word source ancestry table outside index"))?;
            match table.location {
                Some(TableLocation::Cell {
                    table_id: parent,
                    row_index: row,
                    ..
                }) => {
                    // Children are emitted before their parents. Enforce that
                    // ordering, so an invalid link cannot loop or cross stories.
                    if parent <= table_id {
                        return Err(unsupported("invalid Word source ancestry parent"));
                    }
                    table_id = parent;
                    row_index = row;
                }
                Some(TableLocation::Root { .. }) => {
                    return table
                        .rows
                        .get(row_index)
                        .map(|row| Some(row.ttp_id))
                        .ok_or_else(|| unsupported("Word source ancestry row outside index"));
                }
                None => return Err(unsupported("Word table context lacks source ancestry")),
            }
        }
        Err(unsupported("Word source ancestry depth budget exceeded"))
    }

    pub(super) fn outer_row_position(
        &self,
        id: usize,
    ) -> Result<Option<&super::table::Position>, String> {
        let Some(ttp_id) = self.outer_row_owner(id)? else {
            return Ok(None);
        };
        let row = self
            .paragraph(ttp_id)?
            .ok_or_else(|| unsupported("Word source row lacks table context"))?;
        self.tables
            .get(row.table_id)
            .and_then(|table| table.rows.get(row.row_index))
            .map(|row| Some(&row.position))
            .ok_or_else(|| unsupported("Word source row outside table context"))
    }
}

enum SourceBlock {
    Paragraph(usize),
    Table(usize),
}

#[derive(Default)]
struct Ids {
    blocks: Vec<SourceBlock>,
}

impl Ids {
    fn one<A: FnMut(usize) -> Result<(), String>>(
        block: SourceBlock,
        admit: &mut A,
    ) -> Result<Self, String> {
        let mut blocks = Vec::new();
        reserve_additional(&mut blocks, 1, admit)?;
        blocks.push(block);
        Ok(Self { blocks })
    }
}

impl Payload for Ids {
    fn append<A: FnMut(usize) -> Result<(), String>>(
        &mut self,
        mut other: Self,
        admit: &mut A,
    ) -> Result<(), String> {
        reserve_additional(&mut self.blocks, other.blocks.len(), admit)?;
        self.blocks.append(&mut other.blocks);
        Ok(())
    }

    fn is_empty(&self) -> bool {
        self.blocks.is_empty()
    }
}

fn index_event<
    S: Resolver + ?Sized,
    A: FnMut(usize) -> Result<(), String>,
    R: Borrow<super::table::Row>,
>(
    RawEvent(raw_tables): RawEvent<Ids, R>,
    contexts: &mut [Option<ParagraphContext>],
    tables: &mut Vec<TableContext>,
    resolver: &mut S,
    admit: &mut A,
) -> Result<Ids, String> {
    let mut output = Ids::default();
    for raw_table in raw_tables {
        let table_id = tables.len();
        let mut rows = Vec::new();
        let mut header_skipped_rows = 0usize;
        reserve_additional(&mut rows, raw_table.rows.len(), admit)?;
        for (row_index, raw_row) in raw_table.rows.into_iter().enumerate() {
            let ttp_id = raw_row
                .ttp_id
                .ok_or_else(|| unsupported("Word table row lacks source paragraph"))?;
            let first = raw_row.first;
            if first.paragraph.is_none() {
                return Err(unsupported("Word table row lacks source paragraph"));
            }
            // Resolve from the TTP-owned row before assigning the same compact
            // selection to its paragraph contexts.
            let (table_style, table_style_options, source_cell_count, header, position) = {
                let source = raw_row.source.borrow();
                (
                    resolver.table_style(source.table_style)?,
                    source.table_style_options,
                    source.cells.len(),
                    source.header,
                    source.position,
                )
            };
            set_innermost(
                contexts,
                ttp_id,
                ParagraphContext {
                    table_id,
                    row_index,
                    source_cell_index: None,
                    source_block_index: None,
                    ttp_id,
                    table_style,
                    table_style_options,
                },
            )?;
            if raw_row.cells.len() != source_cell_count {
                return Err(unsupported("Word table context cell count mismatch"));
            }
            for (source_cell_index, payload) in raw_row.cells.into_iter().enumerate() {
                for (source_block_index, block) in payload.blocks.into_iter().enumerate() {
                    if let SourceBlock::Table(child) = block {
                        set_location(
                            tables,
                            child,
                            TableLocation::Cell {
                                table_id,
                                row_index,
                                source_cell_index,
                                source_block_index,
                            },
                        )?;
                        continue;
                    }
                    let SourceBlock::Paragraph(id) = block else {
                        unreachable!()
                    };
                    set_innermost(
                        contexts,
                        id,
                        ParagraphContext {
                            table_id,
                            row_index,
                            source_cell_index: Some(source_cell_index),
                            source_block_index: Some(source_block_index),
                            ttp_id,
                            table_style,
                            table_style_options,
                        },
                    )?;
                }
            }
            let preceding_body_rows = header_skipped_rows;
            if row_index != 0 && !header {
                header_skipped_rows = header_skipped_rows
                    .checked_add(1)
                    .ok_or("OUTPUT_TOO_LARGE")?;
            }
            rows.push(RowContext {
                ttp_id,
                table_style,
                table_style_options,
                source_cell_count,
                header,
                preceding_body_rows,
                position,
                first,
            });
        }
        reserve_additional(tables, 1, admit)?;
        tables.push(TableContext {
            rows,
            location: None,
        });
        reserve_additional(&mut output.blocks, 1, admit)?;
        output.blocks.push(SourceBlock::Table(table_id));
    }
    Ok(output)
}

fn set_location(
    tables: &mut [TableContext],
    id: usize,
    location: TableLocation,
) -> Result<(), String> {
    let table = tables
        .get_mut(id)
        .ok_or_else(|| unsupported("Word source ancestry table outside index"))?;
    if table.location.is_some() {
        return Err(unsupported("Word source ancestry assigned twice"));
    }
    table.location = Some(location);
    Ok(())
}

fn set_innermost(
    contexts: &mut [Option<ParagraphContext>],
    id: usize,
    context: ParagraphContext,
) -> Result<(), String> {
    let slot = contexts
        .get_mut(id)
        .ok_or_else(|| unsupported("Word table context paragraph outside index"))?;
    if slot.is_some() {
        return Err(unsupported("Word table context paragraph assigned twice"));
    }
    *slot = Some(context);
    Ok(())
}

fn allocate_none<A: FnMut(usize) -> Result<(), String>>(
    count: usize,
    admit: &mut A,
) -> Result<Vec<Option<ParagraphContext>>, String> {
    let mut values = Vec::new();
    reserve_additional(&mut values, count, admit)?;
    values.resize(count, None);
    Ok(values)
}

fn reserve_additional<T, A: FnMut(usize) -> Result<(), String>>(
    values: &mut Vec<T>,
    additional: usize,
    admit: &mut A,
) -> Result<(), String> {
    let required = values
        .len()
        .checked_add(additional)
        .ok_or("OUTPUT_TOO_LARGE")?;
    let old = values.capacity();
    if required <= old {
        return Ok(());
    }
    let doubled = old.saturating_mul(2);
    let target = required.max(doubled).max(1);
    admit(
        target
            .checked_sub(old)
            .and_then(|growth| growth.checked_mul(std::mem::size_of::<T>()))
            .ok_or("OUTPUT_TOO_LARGE")?,
    )?;
    values
        .try_reserve_exact(target - values.len())
        .map_err(|_| "OUTPUT_TOO_LARGE".to_string())?;
    admit(
        values
            .capacity()
            .saturating_sub(target)
            .checked_mul(std::mem::size_of::<T>())
            .ok_or("OUTPUT_TOO_LARGE")?,
    )
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::doc::table::{Cell, Row};

    fn cell(width: i32, flags: u16) -> Cell {
        Cell {
            width,
            flags,
            ..Cell::default()
        }
    }

    fn properties(depth: u32, row_end: bool, cells: Vec<Cell>) -> Properties {
        let mut value = Properties::default();
        value.apply(0x6649, &depth.to_le_bytes()).unwrap();
        value.row_end = row_end;
        value.row.cells = cells;
        value
    }

    fn build(paragraphs: Vec<(Properties, char)>) -> Index {
        Index::build(paragraphs.len(), paragraphs, &mut |_| Ok(())).unwrap()
    }

    #[test]
    fn indexes_multiple_rows_cells_and_ttp_contexts() {
        let mut first_end = properties(1, true, vec![cell(10, 0), cell(20, 0)]);
        first_end.row.header = true;
        let paragraphs = vec![
            (properties(1, false, vec![]), '\u{7}'),
            (properties(1, false, vec![]), '\u{7}'),
            (first_end, '\u{7}'),
            (properties(1, false, vec![]), '\u{7}'),
            (properties(1, false, vec![]), '\u{7}'),
            (properties(1, true, vec![cell(10, 0), cell(20, 0)]), '\u{7}'),
        ];
        let index = build(paragraphs);
        assert_eq!(index.tables().len(), 1);
        assert_eq!(index.tables()[0].rows.len(), 2);
        assert_eq!(
            index.tables()[0].rows,
            vec![
                RowContext {
                    ttp_id: 2,
                    table_style: None,
                    table_style_options: None,
                    source_cell_count: 2,
                    header: true,
                    preceding_body_rows: 0,
                    position: Default::default(),
                    first: FirstParagraph {
                        paragraph: Some(0),
                        frame: FrameKey::default()
                    },
                },
                RowContext {
                    ttp_id: 5,
                    table_style: None,
                    table_style_options: None,
                    source_cell_count: 2,
                    header: false,
                    preceding_body_rows: 0,
                    position: Default::default(),
                    first: FirstParagraph {
                        paragraph: Some(3),
                        frame: FrameKey::default()
                    },
                },
            ]
        );
        for (id, row, cell) in [
            (0, 0, Some(0)),
            (1, 0, Some(1)),
            (2, 0, None),
            (3, 1, Some(0)),
            (4, 1, Some(1)),
            (5, 1, None),
        ] {
            assert_eq!(
                index.paragraph(id).unwrap(),
                Some(ParagraphContext {
                    table_id: 0,
                    row_index: row,
                    source_cell_index: cell,
                    source_block_index: cell.map(|_| 0),
                    ttp_id: if row == 0 { 2 } else { 5 },
                    table_style: None,
                    table_style_options: None,
                })
            );
        }
    }

    #[test]
    fn indexes_thousands_of_paragraphs_in_one_source_cell() {
        let mut paragraphs = Vec::new();
        for _ in 0..4_096 {
            paragraphs.push((properties(1, false, vec![]), '\r'));
        }
        paragraphs.push((properties(1, false, vec![]), '\u{7}'));
        paragraphs.push((properties(1, true, vec![cell(10, 0)]), '\u{7}'));
        let index = build(paragraphs);

        for id in [0, 4_095, 4_096] {
            let context = index.paragraph(id).unwrap().unwrap();
            assert_eq!(context.source_cell_index, Some(0));
            assert_eq!(context.ttp_id, 4_097);
        }
        assert_eq!(
            index.paragraph(4_097).unwrap().unwrap().source_cell_index,
            None
        );
    }

    #[test]
    fn nested_table_context_is_innermost_and_outer_cell_keeps_placeholder() {
        let mut inner_cell = properties(2, false, vec![]);
        inner_cell.inner_cell = true;
        let mut inner_row = properties(2, false, vec![cell(10, 0)]);
        inner_row.inner_row = true;
        let paragraphs = vec![
            (inner_cell, '\r'),
            (inner_row, '\r'),
            (properties(1, false, vec![]), '\u{7}'),
            (properties(1, true, vec![cell(10, 0)]), '\u{7}'),
        ];
        let index = build(paragraphs);
        assert_eq!(index.tables().len(), 2);
        assert_eq!(index.paragraph(0).unwrap().unwrap().table_id, 0);
        assert_eq!(index.paragraph(1).unwrap().unwrap().table_id, 0);
        assert_eq!(index.paragraph(2).unwrap().unwrap().table_id, 1);
        assert_eq!(index.paragraph(3).unwrap().unwrap().table_id, 1);
        assert_eq!(index.paragraph(0).unwrap().unwrap().ttp_id, 1);
        assert_eq!(index.paragraph(2).unwrap().unwrap().ttp_id, 3);
    }

    #[test]
    fn nested_source_ancestry_keeps_row_cell_and_block_ownership() {
        let nested_cell = || {
            let mut p = properties(2, false, vec![]);
            p.inner_cell = true;
            p
        };
        let nested_row = || {
            let mut p = properties(2, false, vec![cell(10, 0)]);
            p.inner_row = true;
            p
        };
        let index = build(vec![
            (properties(0, false, vec![]), '\r'),
            (properties(1, false, vec![]), '\u{7}'),
            (properties(1, true, vec![cell(10, 0)]), '\u{7}'),
            (properties(1, false, vec![]), '\r'),
            (nested_cell(), '\r'),
            (nested_row(), '\r'),
            (properties(1, false, vec![]), '\r'),
            (nested_cell(), '\r'),
            (nested_row(), '\r'),
            (properties(1, false, vec![]), '\u{7}'),
            (properties(1, false, vec![]), '\u{7}'),
            (properties(1, true, vec![cell(10, 0), cell(10, 0)]), '\u{7}'),
            (properties(0, false, vec![]), '\r'),
        ]);
        let outer = index.paragraph(3).unwrap().unwrap().table_id;
        for (id, block_index) in [(4, 1), (7, 3)] {
            let inner = index.paragraph(id).unwrap().unwrap();
            assert_eq!(
                index.tables()[inner.table_id].location,
                Some(TableLocation::Cell {
                    table_id: outer,
                    row_index: 1,
                    source_cell_index: 0,
                    source_block_index: block_index,
                })
            );
            assert_eq!(inner.source_block_index, Some(0));
            assert_eq!(index.outer_row_owner(id).unwrap(), Some(11));
        }
        assert_eq!(
            index.tables()[outer].location,
            Some(TableLocation::Root {
                source_block_index: 1
            })
        );
        assert_eq!(
            index.paragraph(6).unwrap().unwrap().source_block_index,
            Some(2)
        );
        assert_eq!(index.outer_row_owner(1).unwrap(), Some(2));
        assert_eq!(index.outer_row_owner(0).unwrap(), None);
        assert_eq!(index.outer_row_owner(12).unwrap(), None);
    }

    fn borrowed_fixture() -> Vec<(Properties, char)> {
        let mut inner_cell = properties(2, false, vec![]);
        inner_cell.inner_cell = true;
        let mut inner_row = properties(2, false, vec![cell(10, 0)]);
        inner_row.inner_row = true;
        inner_row.row.header = true;
        let mut outer_row = properties(1, true, vec![cell(10, 2), cell(10, 1)]);
        outer_row
            .row
            .identity
            .insert(0x563a, 7u16.to_le_bytes().to_vec());
        vec![
            (inner_cell, '\r'),
            (inner_row, '\r'),
            (properties(1, false, vec![]), '\u{7}'),
            (properties(1, false, vec![]), '\u{7}'),
            (outer_row, '\u{7}'),
        ]
    }

    #[test]
    fn borrowed_and_owned_rows_produce_identical_context_without_moving_source_heaps() {
        let source = borrowed_fixture();
        let cell_pointer = source[4].0.row.cells.as_ptr();
        let identity_pointer = source[4].0.row.identity[&0x563a].as_ptr();
        let borrowed = Index::build(
            source.len(),
            source
                .iter()
                .map(|(properties, mark)| (properties.borrowed(), *mark)),
            &mut |_| Ok(()),
        )
        .unwrap();
        assert_eq!(source[4].0.row.cells.as_ptr(), cell_pointer);
        assert_eq!(source[4].0.row.identity[&0x563a].as_ptr(), identity_pointer);

        let owned = Index::build(source.len(), source, &mut |_| Ok(())).unwrap();
        assert_eq!(borrowed.tables(), owned.tables());
        for id in 0..5 {
            assert_eq!(borrowed.paragraph(id), owned.paragraph(id));
        }
    }

    #[test]
    fn emitted_grammar_must_name_the_indexed_first_paragraph() {
        struct Frames;
        impl Resolver for Frames {
            fn table_style(&mut self, selected: Option<usize>) -> Result<Option<usize>, String> {
                Ok(selected)
            }
            fn first_paragraph_frame(&mut self, paragraph: usize) -> Result<FrameKey, String> {
                Ok(FrameKey {
                    width: paragraph as u16,
                    ..FrameKey::default()
                })
            }
        }
        let paragraphs = vec![
            (properties(1, false, vec![]), '\r'),
            (properties(1, false, vec![]), '\u{7}'),
            (properties(1, true, vec![cell(10, 0)]), '\u{7}'),
        ];
        let index = Index::build_with_styles(3, paragraphs, &mut Frames, &mut |_| Ok(())).unwrap();
        assert_eq!(
            index.tables()[0].rows[0].first,
            FirstParagraph {
                paragraph: Some(0),
                frame: FrameKey {
                    width: 0,
                    ..FrameKey::default()
                }
            }
        );
        assert_eq!(index.first_paragraph_frame(0), Ok(FrameKey::default()));
        // A later paragraph of the cell and the TTP are not identity sources.
        for other in [1, 2] {
            assert!(index
                .first_paragraph_frame(other)
                .unwrap_err()
                .contains("disagrees"));
        }
        assert!(index.first_paragraph_frame(3).is_err());
    }

    #[test]
    fn resolved_table_style_is_selected_from_the_row_owning_ttp() {
        let mut cell_properties = properties(1, false, vec![]);
        cell_properties.row.table_style = Some(99);
        let mut row_properties = properties(1, true, vec![cell(10, 0)]);
        row_properties.row.table_style = Some(7);
        let source = [(cell_properties, '\u{7}'), (row_properties, '\u{7}')];

        let index = Index::build_with_styles(
            source.len(),
            source
                .iter()
                .map(|(properties, mark)| (properties.borrowed(), *mark)),
            &mut Unframed(|selected: Option<usize>| {
                Ok(selected.map(|id| if id == 7 { 3 } else { 11 }))
            }),
            &mut |_| Ok(()),
        )
        .unwrap();

        assert_eq!(index.paragraph(0).unwrap().unwrap().table_style, Some(3));
        assert_eq!(index.paragraph(1).unwrap().unwrap().table_style, Some(3));
        assert_eq!(index.tables()[0].rows[0].table_style, Some(3));
    }

    #[test]
    fn table_style_options_follow_each_ttp_without_nested_or_absent_leakage() {
        let inner_options = 0x0120;
        let outer_options = 0x0340;
        let mut inner_cell = properties(2, false, vec![]);
        inner_cell.inner_cell = true;
        let mut inner_row = properties(2, false, vec![cell(1_000, 0)]);
        inner_row.inner_row = true;
        inner_row.row.table_style_options = Some(inner_options);
        let mut outer_row = properties(1, true, vec![cell(2_000, 0)]);
        outer_row.row.table_style_options = Some(outer_options);
        let index = build(vec![
            (inner_cell, '\r'),
            (inner_row, '\r'),
            (properties(1, false, vec![]), '\u{7}'),
            (outer_row, '\u{7}'),
            (properties(1, false, vec![]), '\u{7}'),
            (properties(1, true, vec![cell(2_000, 0)]), '\u{7}'),
        ]);

        for (id, table_id, row_index, source_cell_index, ttp_id, options) in [
            (0, 0, 0, Some(0), 1, Some(inner_options)),
            (1, 0, 0, None, 1, Some(inner_options)),
            (2, 1, 0, Some(0), 3, Some(outer_options)),
            (3, 1, 0, None, 3, Some(outer_options)),
            (4, 1, 1, Some(0), 5, None),
            (5, 1, 1, None, 5, None),
        ] {
            assert_eq!(
                index.paragraph(id).unwrap(),
                Some(ParagraphContext {
                    table_id,
                    row_index,
                    source_cell_index,
                    source_block_index: source_cell_index.map(|_| usize::from(id == 2)),
                    ttp_id,
                    table_style: None,
                    table_style_options: options,
                })
            );
        }
        assert_eq!(
            index.tables()[0].rows[0].table_style_options,
            Some(inner_options)
        );
        assert_eq!(
            index.tables()[1].rows[0].table_style_options,
            Some(outer_options)
        );
        assert_eq!(index.tables()[1].rows[1].table_style_options, None);
    }

    #[test]
    fn borrowed_depth_errors_leave_source_reusable() {
        let mut source = properties(33, false, vec![cell(10, 0)]);
        source.row.header = true;
        let pointer = source.row.cells.as_ptr();
        let error = Index::build(
            1,
            std::iter::once((source.borrowed(), '\u{7}')),
            &mut |_| Ok(()),
        )
        .err()
        .unwrap();
        assert!(error.contains("nesting budget"), "{error}");
        assert_eq!(source.row.cells.as_ptr(), pointer);
        assert!(source.row.header);
        assert!(source.borrowed().depth().is_err());
    }

    #[test]
    fn source_merge_continuations_keep_distinct_paragraph_contexts() {
        let index = build(vec![
            (properties(1, false, vec![]), '\u{7}'),
            (properties(1, false, vec![]), '\u{7}'),
            (properties(1, true, vec![cell(10, 2), cell(10, 1)]), '\u{7}'),
        ]);
        assert_eq!(
            index.paragraph(0).unwrap().unwrap().source_cell_index,
            Some(0)
        );
        assert_eq!(
            index.paragraph(1).unwrap().unwrap().source_cell_index,
            Some(1)
        );
        assert_eq!(index.tables()[0].rows[0].source_cell_count, 2);
    }

    #[test]
    fn logical_columns_use_each_rows_source_cell_sequence() {
        let mut first_row = properties(1, true, vec![cell(2_400, 0), cell(300, 0), cell(900, 0)]);
        first_row.row.left = 1_800;
        let second_row = properties(1, true, vec![cell(750, 0), cell(2_250, 0)]);
        let index = build(vec![
            (properties(1, false, vec![]), '\u{7}'),
            (properties(1, false, vec![]), '\u{7}'),
            (properties(1, false, vec![]), '\u{7}'),
            (first_row, '\u{7}'),
            (properties(1, false, vec![]), '\u{7}'),
            (properties(1, false, vec![]), '\u{7}'),
            (second_row, '\u{7}'),
        ]);

        assert_eq!(
            index.logical_column(0).unwrap(),
            Some(LogicalColumn {
                ordinal: 0,
                count: 3,
            })
        );
        assert_eq!(
            index.logical_column(2).unwrap(),
            Some(LogicalColumn {
                ordinal: 2,
                count: 3,
            })
        );
        assert_eq!(
            index.logical_column(4).unwrap(),
            Some(LogicalColumn {
                ordinal: 0,
                count: 2,
            })
        );
        assert_eq!(
            index.logical_column(5).unwrap(),
            Some(LogicalColumn {
                ordinal: 1,
                count: 2,
            })
        );
        assert_eq!(index.logical_column(3).unwrap(), None);
        assert_eq!(index.logical_column(6).unwrap(), None);
    }

    #[test]
    fn nested_paragraph_logical_column_uses_innermost_table() {
        let mut inner_cell = properties(2, false, vec![]);
        inner_cell.inner_cell = true;
        let mut inner_row = properties(2, false, vec![cell(2_400, 0)]);
        inner_row.inner_row = true;
        let index = build(vec![
            (inner_cell, '\r'),
            (inner_row, '\r'),
            (properties(1, false, vec![]), '\u{7}'),
            (properties(1, true, vec![cell(3_000, 0)]), '\u{7}'),
        ]);

        assert_eq!(
            index.logical_column(0).unwrap(),
            Some(LogicalColumn {
                ordinal: 0,
                count: 1,
            })
        );
        assert_eq!(index.logical_column(1).unwrap(), None);
        assert_eq!(
            index.logical_column(2).unwrap(),
            Some(LogicalColumn {
                ordinal: 0,
                count: 1,
            })
        );
        assert_eq!(index.logical_column(3).unwrap(), None);
    }

    #[test]
    fn horizontal_merge_slots_keep_source_ordinals_and_full_row_count() {
        let merged_first = build(vec![
            (properties(1, false, vec![]), '\u{7}'),
            (properties(1, false, vec![]), '\u{7}'),
            (properties(1, false, vec![]), '\u{7}'),
            (
                properties(
                    1,
                    true,
                    vec![cell(1_000, 2), cell(1_000, 1), cell(1_000, 0)],
                ),
                '\u{7}',
            ),
        ]);
        assert_eq!(
            merged_first.logical_column(2).unwrap(),
            Some(LogicalColumn {
                ordinal: 2,
                count: 3,
            })
        );
        assert_eq!(
            merged_first.logical_column(1).unwrap(),
            Some(LogicalColumn {
                ordinal: 1,
                count: 3,
            })
        );

        let merged_last = build(vec![
            (properties(1, false, vec![]), '\u{7}'),
            (properties(1, false, vec![]), '\u{7}'),
            (properties(1, false, vec![]), '\u{7}'),
            (
                properties(
                    1,
                    true,
                    vec![cell(1_000, 0), cell(1_000, 2), cell(1_000, 1)],
                ),
                '\u{7}',
            ),
        ]);
        assert_eq!(
            merged_last.logical_column(1).unwrap(),
            Some(LogicalColumn {
                ordinal: 1,
                count: 3,
            })
        );
        assert_eq!(
            merged_last.logical_column(2).unwrap(),
            Some(LogicalColumn {
                ordinal: 2,
                count: 3,
            })
        );

        let merged_all = build(vec![
            (properties(1, false, vec![]), '\u{7}'),
            (properties(1, false, vec![]), '\u{7}'),
            (properties(1, false, vec![]), '\u{7}'),
            (
                properties(
                    1,
                    true,
                    vec![cell(1_000, 2), cell(1_000, 1), cell(1_000, 1)],
                ),
                '\u{7}',
            ),
        ]);
        assert_eq!(
            merged_all.logical_column(0).unwrap(),
            Some(LogicalColumn {
                ordinal: 0,
                count: 3,
            })
        );
        assert_eq!(
            merged_all.logical_column(1).unwrap(),
            Some(LogicalColumn {
                ordinal: 1,
                count: 3,
            })
        );
        assert_eq!(
            merged_all.logical_column(2).unwrap(),
            Some(LogicalColumn {
                ordinal: 2,
                count: 3,
            })
        );
    }

    #[test]
    fn vertical_merge_slots_keep_source_ordinals_and_full_row_count() {
        let mut paragraphs = Vec::new();
        for vertical in [3, 1, 1] {
            paragraphs.extend([
                (properties(1, false, vec![]), '\u{7}'),
                (properties(1, false, vec![]), '\u{7}'),
                (properties(1, false, vec![]), '\u{7}'),
                (
                    properties(
                        1,
                        true,
                        vec![cell(1_000, 0), cell(1_000, vertical << 5), cell(1_000, 0)],
                    ),
                    '\u{7}',
                ),
            ]);
        }
        let index = build(paragraphs);

        for id in [1, 5, 9] {
            assert_eq!(
                index.logical_column(id).unwrap(),
                Some(LogicalColumn {
                    ordinal: 1,
                    count: 3,
                })
            );
        }
        for (id, ordinal) in [(0, 0), (2, 2), (4, 0), (6, 2), (8, 0), (10, 2)] {
            assert_eq!(
                index.logical_column(id).unwrap(),
                Some(LogicalColumn { ordinal, count: 3 })
            );
        }
        for ttp in [3, 7, 11] {
            assert_eq!(index.logical_column(ttp).unwrap(), None);
        }
    }

    #[test]
    fn bidi_rows_keep_source_order_without_reversal() {
        let mut ordinary = properties(
            1,
            true,
            vec![cell(1_000, 0), cell(1_000, 0), cell(1_000, 0)],
        );
        ordinary.row.bidi = true;
        let mut merged = properties(
            1,
            true,
            vec![cell(1_000, 0), cell(1_000, 2), cell(1_000, 1)],
        );
        merged.row.bidi = true;
        let index = build(vec![
            (properties(1, false, vec![]), '\u{7}'),
            (properties(1, false, vec![]), '\u{7}'),
            (properties(1, false, vec![]), '\u{7}'),
            (ordinary, '\u{7}'),
            (properties(1, false, vec![]), '\u{7}'),
            (properties(1, false, vec![]), '\u{7}'),
            (properties(1, false, vec![]), '\u{7}'),
            (merged, '\u{7}'),
        ]);

        for start in [0, 4] {
            for ordinal in 0..3 {
                assert_eq!(
                    index.logical_column(start + ordinal).unwrap(),
                    Some(LogicalColumn { ordinal, count: 3 })
                );
            }
            assert_eq!(index.logical_column(start + 3).unwrap(), None);
        }
    }

    #[test]
    fn logical_column_checks_paragraph_table_row_and_source_ids() {
        let mut index = build(vec![
            (properties(1, false, vec![]), '\u{7}'),
            (properties(1, true, vec![cell(1_000, 0)]), '\u{7}'),
        ]);
        assert!(index.logical_column(2).is_err());

        let valid = index.paragraphs[0].unwrap();
        index.paragraphs[0] = Some(ParagraphContext {
            table_id: 1,
            ..valid
        });
        assert!(index.logical_column(0).is_err());
        index.paragraphs[0] = Some(ParagraphContext {
            row_index: 1,
            ..valid
        });
        assert!(index.logical_column(0).is_err());
        index.paragraphs[0] = Some(ParagraphContext {
            source_cell_index: Some(1),
            ..valid
        });
        assert!(index.logical_column(0).is_err());
    }

    #[test]
    fn table_identity_changes_create_stable_table_ids() {
        let mut first = properties(1, true, vec![cell(10, 0)]);
        first
            .row
            .identity
            .insert(0x563a, 7u16.to_le_bytes().to_vec());
        let mut second = properties(1, true, vec![cell(10, 0)]);
        second
            .row
            .identity
            .insert(0x563a, 8u16.to_le_bytes().to_vec());
        let index = build(vec![
            (properties(1, false, vec![]), '\u{7}'),
            (first, '\u{7}'),
            (properties(1, false, vec![]), '\u{7}'),
            (second, '\u{7}'),
        ]);
        assert_eq!(index.tables().len(), 2);
        assert_eq!(index.paragraph(0).unwrap().unwrap().table_id, 0);
        assert_eq!(index.paragraph(2).unwrap().unwrap().table_id, 1);
    }

    #[test]
    fn empty_and_non_table_stories_have_no_contexts() {
        let empty = Index::build(0, Vec::<(Properties, char)>::new(), &mut |_| Ok(())).unwrap();
        assert!(empty.tables().is_empty());
        assert!(empty.paragraph(0).is_err());
        let plain = build(vec![(Properties::default(), '\r')]);
        assert!(plain.tables().is_empty());
        assert_eq!(plain.paragraph(0).unwrap(), None);
        assert!(plain.paragraph(1).is_err());
    }

    #[test]
    fn rejects_paragraph_count_and_table_grammar_mismatches() {
        for declared in [0, 2] {
            let error = Index::build(declared, vec![(Properties::default(), '\r')], &mut |_| {
                Ok(())
            })
            .err()
            .unwrap();
            assert!(error.contains("paragraph count mismatch"), "{error}");
        }
        let error = Index::build(
            1,
            vec![(properties(1, true, vec![cell(10, 0)]), '\u{7}')],
            &mut |_| Ok(()),
        )
        .err()
        .unwrap();
        assert!(
            error.contains("row definition does not match cell marks"),
            "{error}"
        );

        let mut inner_cell = properties(2, false, vec![]);
        inner_cell.inner_cell = true;
        let mut inner_row = properties(2, false, vec![cell(10, 0)]);
        inner_row.inner_row = true;
        let error = Index::build(2, vec![(inner_cell, '\r'), (inner_row, '\r')], &mut |_| {
            Ok(())
        })
        .err()
        .unwrap();
        assert!(error.contains("unterminated Word table row"), "{error}");
    }

    #[test]
    fn budget_denial_precedes_index_allocation_and_iterator_consumption() {
        let consumed = std::cell::Cell::new(0);
        let paragraphs = std::iter::from_fn(|| {
            consumed.set(consumed.get() + 1);
            Some((Properties::default(), '\r'))
        });
        assert_eq!(
            Index::build(1, paragraphs, &mut |_| Err("denied".into()))
                .err()
                .as_deref(),
            Some("denied")
        );
        assert_eq!(consumed.get(), 0);
        assert!(Index::build(
            MAX_STORY_CONTROLS + 2,
            Vec::<(Properties, char)>::new(),
            &mut |_| panic!("budget must be rejected before allocation"),
        )
        .err()
        .unwrap()
        .contains("paragraph budget"));
    }

    fn nested(depth: u32) -> Vec<(Properties, char)> {
        let mut paragraphs = Vec::new();
        for level in 1..=depth {
            paragraphs.push((properties(level, false, vec![]), '\r'));
        }
        for level in (1..=depth).rev() {
            let mut cell_end = properties(level, false, vec![]);
            let mut row_end = properties(level, false, vec![cell(10, 0)]);
            let mark = if level == 1 { '\u{7}' } else { '\r' };
            if level == 1 {
                row_end.row_end = true;
            } else {
                cell_end.inner_cell = true;
                row_end.inner_row = true;
            }
            paragraphs.push((cell_end, mark));
            paragraphs.push((row_end, mark));
        }
        paragraphs
    }

    #[test]
    fn deep_nesting_maps_each_id_once_without_ancestor_fanout() {
        let paragraphs = nested(32);
        let thirty_two = Index::build(paragraphs.len(), paragraphs, &mut |_| Ok(())).unwrap();
        assert_eq!(thirty_two.tables().len(), 32);
        assert!((0..96).all(|id| thirty_two.paragraph(id).unwrap().is_some()));
        assert!((0..96).all(|id| thirty_two.outer_row_owner(id).unwrap() == Some(95)));
    }

    #[test]
    fn reserve_growth_is_geometric_for_large_cells_and_table_lists() {
        let mut paragraph_ids = Vec::new();
        let mut id_admissions = 0usize;
        for id in 0..4_096 {
            reserve_additional(&mut paragraph_ids, 1, &mut |bytes| {
                id_admissions += usize::from(bytes != 0);
                Ok(())
            })
            .unwrap();
            paragraph_ids.push(id);
        }
        let mut tables = Vec::new();
        let mut table_admissions = 0usize;
        for _ in 0..4_096 {
            reserve_additional(&mut tables, 1, &mut |bytes| {
                table_admissions += usize::from(bytes != 0);
                Ok(())
            })
            .unwrap();
            tables.push(TableContext {
                rows: Vec::new(),
                location: None,
            });
        }

        // Doubling from one element needs 1 + log2(count) growths. Each growth
        // has at most one requested-capacity and one allocator-excess admission.
        let maximum_admissions = 2 * (1 + 4_096usize.ilog2() as usize);
        assert!(id_admissions <= maximum_admissions);
        assert!(table_admissions <= maximum_admissions);
    }

    #[test]
    fn empty_raw_event_does_not_fabricate_nested_content() {
        let mut contexts = Vec::new();
        let mut tables = Vec::new();
        let payload = index_event(
            RawEvent::<Ids, Row>(Vec::new()),
            &mut contexts,
            &mut tables,
            &mut Unframed(Ok),
            &mut |_| Ok(()),
        )
        .unwrap();
        assert!(payload.is_empty());
        assert!(tables.is_empty());
    }

    #[test]
    fn duplicate_context_assignment_is_rejected() {
        let context = ParagraphContext {
            table_id: 0,
            row_index: 0,
            source_cell_index: Some(0),
            source_block_index: Some(0),
            ttp_id: 1,
            table_style: None,
            table_style_options: None,
        };
        let mut contexts = vec![Some(context)];
        let error = set_innermost(&mut contexts, 0, context).unwrap_err();
        assert!(error.contains("assigned twice"), "{error}");
    }
}
