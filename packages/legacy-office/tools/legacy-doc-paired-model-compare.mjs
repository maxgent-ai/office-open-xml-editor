#!/usr/bin/env node
/**
 * Compare parser-emitted DOCX/DOC meaning, never previous-renderer correctness.
 * This is a diagnostic projection, not a renderer, Office compatibility policy,
 * or an assertion that two Office saves preserve every semantic fact.
 * Unknown private fields remain in acquisitionMetadata; the full raw difference
 * ledger preserves every encoding difference, including unknown public fields.
 * Unexercised domains are explicit.
 */
import { createHash } from 'node:crypto';

import { constants as fsConstants, openSync, closeSync, fstatSync, readSync, writeFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { pathToFileURL } from 'node:url';

const GROUPS = [
  'content',
  'paragraphFormats',
  'runFormats',
  'tables',
  'notes',
  'images',
  'anchors',
  'sections',
  'settings',
  'acquisitionMetadata'
];

const MAX_DEPTH = 128, MAX_NODES = 1000000, MAX_DIFFS = 20000, MAX_FILE_BYTES = 64 * 1024 * 1024;
const stable = value => JSON.stringify(value);
const identityKeys = new Set(['paragraphId', 'styleId', 'effectiveStyleId', 'logicalSequenceId']);

function context(model, resources) {
  if (resources.length > MAX_NODES)
    throw new RangeError('paired resource inventory budget exceeded');

  const byPath = new Map(resources.map(r => [r.path, r])), unavailableResources = new Set();

  if (byPath.size !== resources.length)
    throw new TypeError('duplicate paired resource path');

  for (const resource of resources) if (typeof resource.path !== 'string' || resource.sha256 !== undefined && (!/^[a-f0-9]{64}$/.test(resource.sha256) || !Number.isSafeInteger(resource.bytes) || resource.bytes < 0))
    throw new TypeError('invalid paired resource inventory');

  const styles = new Map(),
        tables = new Map(),
        notes = {
          footnote: new Map(),
          endnote: new Map()
        };

  const knownNotes = {
    footnote: new Map((model.footnotes ?? []).map(n => [n.id, n])),
    endnote: new Map((model.endnotes ?? []).map(n => [n.id, n]))
  };

  for (const kind of ['footnote', 'endnote']) if (knownNotes[kind].size !== (model[kind + 's'] ?? []).length)
    throw new TypeError('duplicate paired note target');

  let work = 0;

  const charge = depth => {
    if (depth > MAX_DEPTH || ++work > MAX_NODES)
      throw new RangeError('paired-model projection budget exceeded');
  };

  const noteTarget = ref => {
    const facts = Object.fromEntries(
      Object.entries(ref).filter(([key, value]) => key !== 'id' && value !== null && value !== undefined)
    );

    if (ref.id === '') return {
      ...facts,

      id: {
        role: 'enclosing-note-number'
      }
    };

    const map = notes[ref.kind];

    if (!map) return {
      ...ref,
      unrecognizedKind: true
    };

    if (!map.has(ref.id))
      map.set(ref.id, map.size);

    return {
      ...facts,

      id: {
        targetOrdinal: map.get(ref.id),
        missingTarget: !knownNotes[ref.kind].has(ref.id)
      }
    };
  };

  // Validate the entire input before recursive story traversal. The active set
  // rejects cycles in API inputs; JSON CLI inputs are acyclic by construction.
  const active = new Set();

  const validate = (value, depth = 0) => {
    charge(depth);

    if (!value || typeof value !== 'object')
      return;

    if (active.has(value))
      throw new TypeError('cyclic paired model');

    active.add(value);

    for (const v of Object.values(value))
      validate(v, depth + 1);

    active.delete(value);
  };

  validate(model);

  const register = (value, depth = 0) => {
    charge(depth);

    if (!value || typeof value !== 'object')
      return;

    if (value.noteRef)
      noteTarget(value.noteRef);

    for (const [key, v] of Object.entries(value)) if (key !== 'noteRef')
      register(v, depth + 1);
  };

  register(model.body ?? []);
  register(model.headers ?? {});
  register(model.footers ?? {});

  const idClass = (key, value) => {
    const map = key === 'logicalSequenceId' ? tables : styles;
    const name = key + ':' + value;

    if (!map.has(name))
      map.set(name, map.size);

    return map.get(name);
  };

  const canonical = (value, depth = 0, interpretReferences = true) => {
    charge(depth);

    if (Array.isArray(value))
      return value.map(v => canonical(v, depth + 1, interpretReferences));

    if (!value || typeof value !== 'object')
      return value;

    const result = Object.create(null);

    for (const key of Object.keys(value).sort()) {
      const v = value[key];

      // Optional parser fields use null/omission for "not specified". Explicit
      // zero/false are retained; raw presence changes are reported separately.
      if (v === undefined || v === null)
        continue;

      if (interpretReferences && key === 'paragraphId')
        continue;

      // Reuse the original key: synthesized sibling names could collide with
      // unknown source fields and silently overwrite their retained evidence.
      if (interpretReferences && (key === 'styleId' || key === 'effectiveStyleId' || key === 'logicalSequenceId')) result[key] = {
        encodingClass: idClass(key, v)
      };
      else if (interpretReferences && (key === 'imagePath' || key === 'picBulletImagePath')) {
        const resource = byPath.get(v);

        if (!resource?.sha256)
          unavailableResources.add(v);

        result[key] = resource?.sha256 ? {
          sha256: resource.sha256,
          bytes: resource.bytes
        } : {
          unavailable: true,
          error: resource?.error ?? 'not extracted'
        };
      } else if (interpretReferences && key === 'noteRef')
        result[key] = noteTarget(v);
      else
        result[key] = canonical(v, depth + 1, interpretReferences);
    }

    return result;
  };

  return {
    canonical,
    charge,
    notes,
    knownNotes,
    unavailableResources
  };
}

function properties(value, excluded, ctx) {
  return ctx.canonical(Object.fromEntries(Object.entries(value).filter(([key]) => !excluded.has(key))));
}

const internals = value => Object.fromEntries(Object.entries(value).filter(([key]) => key.startsWith('__')));
const withoutInternals = value => Object.fromEntries(Object.entries(value).filter(([key]) => !key.startsWith('__')));
const emptyGroups = () => Object.fromEntries(GROUPS.map(name => [name, []]));

/** Source margins are preserved in metadata. Effective margins use the source
 * cell override, then row exception, then resolved table margin (ECMA-376
 * 17.4.68 tcMar /17.4.41 tblPrEx /17.4.42 tblCellMar). No missing value is guessed.
 * This diagnostic supports physical dxa edges; logical/other kinds are retained
 * as unresolved facts rather than silently choosing a geometry rule. */
function effectiveMargins(table, row, cell) {
  const result = {};
  const exception = row.__tableRowLayout?.exception?.cellMargins;

  for (const edge of ['Top', 'Bottom', 'Left', 'Right']) {
    const own = cell['margin' + edge];

    if (own !== null && own !== undefined) {
      result[edge] = own;
      continue;
    }

    const lexical = exception?.[edge.toLowerCase()];

    if (lexical) {
      if (lexical.kind === 'dxa' && /^-?\d+$/.test(lexical.value) && Number.isSafeInteger(Number(lexical.value)))
        result[edge] = Number(lexical.value) / 20;
      else result[edge] = {
        unresolvedRowMargin: lexical
      };
    } else result[edge] = table['cellMargin' + edge] ?? {
      unspecified: true
    };
  }

  if (exception?.start || exception?.end) result.unresolvedLogicalRowMargins = {
    start: exception.start,
    end: exception.end
  };

  return result;
}

export function projectModelMeaning(
  model,
  {
    resources = []
  } = {}
) {
  if (!model || !Array.isArray(model.body))
    throw new TypeError('paired model requires a body array');

  if (model.parseError)
    throw new TypeError('partial parser model cannot establish paired meaning');

  const ctx = context(model, resources),
        groups = emptyGroups(),
        counts = {
          paragraphs: 0,
          tables: 0,
          notes: 0,
          images: 0,
          anchors: 0
        };

  const metadata = (owner, value) => {
    const extra = internals(value);

    if (Object.keys(extra).length) groups.acquisitionMetadata.push({
      owner,
      facts: ctx.canonical(extra, 0, false)
    });
  };

  const runs = (values, owner) => {
    const content = [], formats = [];
    let offset = 0;
    const textParts = [];

    for (const run of values ?? []) {
      ctx.charge(0);
      metadata(owner + '/run:' + offset, run);

      if (run.type === 'text' && !run.noteRef) {
        const text = run.text ?? '';
        textParts.push(text);
        const format = properties(withoutInternals(run), new Set(['type', 'text']), ctx);
        const previous = formats.at(-1);

        // ECMA-37617.3.2.25 r: equal contiguous character properties are
        // independent of serialization run boundaries. No trimming/case folding.
        if (previous?.kind === 'text' && previous.end === offset && stable(previous.format) === stable(format))
          previous.end += text.length;
        else formats.push({
          kind: 'text',
          start: offset,
          end: offset + text.length,
          format
        });

        offset += text.length;
        continue;
      }

      if (textParts.length) {
        content.push({
          type: 'text',
          text: textParts.join('')
        });

        textParts.length = 0;
      }

      const projected = ctx.canonical(run);
      content.push(projected);

      formats.push({
        kind: run.type,
        offset,
        format: properties(withoutInternals(run), new Set(['text', 'fallbackText']), ctx)
      });

      offset += (run.text ?? run.fallbackText ?? '').length;
    }

    if (textParts.length) content.push({
      type: 'text',
      text: textParts.join('')
    });

    groups.runFormats.push({
      owner,
      spans: formats
    });

    return content;
  };

  const blocks = (values, owner) => values.map((block, index) => {
    ctx.charge(0);
    const address = owner + '/' + index;
    metadata(address, block);

    if (block.type === 'paragraph') {
      counts.paragraphs++;
      const format = properties(withoutInternals(block), new Set(['type', 'runs', 'paragraphId']), ctx);

      // Effective cascade has already run in each production parser. If never
      // authored, paragraph snapToGrid is on (ECMA-37617.3.1.32); false survives.
      format.snapToGrid = block.snapToGrid ?? true;

      groups.paragraphFormats.push({
        owner: address,
        format
      });

      if (block.framePr) {
        groups.anchors.push({
          owner: address,
          kind: 'paragraphFrame',
          facts: ctx.canonical(block.framePr)
        });

        counts.anchors++;
      }

      // A picture bullet is a paragraph-format-owned resource (numbering), not
      // an image run: one marker fact per numbered paragraph, by extracted bytes.
      if (block.numbering?.picBulletImagePath !== undefined && block.numbering?.picBulletImagePath !== null) {
        groups.images.push({
          owner: address,
          kind: 'pictureBullet',
          facts: ctx.canonical({
            picBulletImagePath: block.numbering.picBulletImagePath
          })
        });

        counts.images++;
      }

      for (const run of block.runs ?? []) {
        if (run.type === 'image') {
          groups.images.push({
            owner: address,
            facts: ctx.canonical(run)
          });

          counts.images++;
        }

        if (run.anchor || run.__anchor || run.type === 'anchorHost') {
          groups.anchors.push({
            owner: address,
            facts: ctx.canonical(run)
          });

          counts.anchors++;
        }
      }

      return {
        type: 'paragraph',
        runs: runs(block.runs, address)
      };
    }

    if (block.type === 'table') {
      counts.tables++;

      const format = properties(
        withoutInternals(block),
        new Set(['type', 'rows', 'cellMarginTop', 'cellMarginBottom', 'cellMarginLeft', 'cellMarginRight']),
        ctx
      );

      const rows = (block.rows ?? []).map((row, r) => {
        metadata(address + '/row:' + r, row);
        const rowFormat = properties(withoutInternals(row), new Set(['cells']), ctx);

        // Keep native/OOXML constraints unchanged. Missing vs explicit zero and
        // hRule disagreements remain visible, rather than fitting an Office floor.
        const cells = (row.cells ?? []).map((cell, c) => {
          const cellOwner = address + '/row:' + r + '/cell:' + c;
          metadata(cellOwner, cell);

          const cellFormat = properties(
            withoutInternals(cell),
            new Set(['content', 'marginTop', 'marginBottom', 'marginLeft', 'marginRight']),
            ctx
          );

          cellFormat.effectiveMargins = effectiveMargins(block, row, cell);

          return {
            format: cellFormat,
            content: blocks(cell.content ?? [], cellOwner)
          };
        });

        return {
          format: rowFormat,
          cells
        };
      });

      groups.tables.push({
        owner: address,
        format,

        rows: rows.map(row => ({
          format: row.format,
          cells: row.cells.map(cell => cell.format)
        }))
      });

      // Table defaults remain as source/encoding evidence even when every cell
      // overrides them; they are not mistaken for active cell-margin differences.
      groups.acquisitionMetadata.push({
        owner: address,

        tableMarginDefaults: ctx.canonical(
          Object.fromEntries(['Top', 'Bottom', 'Left', 'Right'].map(e => [e, block['cellMargin' + e]]))
        )
      });

      return {
        type: 'table',
        rows: rows.map(row => row.cells.map(cell => cell.content))
      };
    }

    if (block.type === 'sectionBreak') groups.sections.push({
      owner: address,
      facts: ctx.canonical(block)
    });

    return ctx.canonical(block);
  });

  groups.content.push({
    story: 'body',
    blocks: blocks(model.body, 'body')
  });

  for (const kind of ['headers', 'footers']) for (const [instance, story] of Object.entries(model[kind] ?? {})) if (story) {
    groups.content.push({
      story: kind + ':' + instance,
      facts: properties(story, new Set(['body']), ctx),
      blocks: blocks(story.body ?? [], kind + ':' + instance)
    });
  }

  for (const kind of ['footnote', 'endnote']) {
    const entries = ctx.knownNotes[kind], order = ctx.notes[kind];

    for (const id of entries.keys()) if (!order.has(id))
      order.set(id, order.size);

    for (const [id, ordinal] of order) {
      const note = entries.get(id);

      if (!note) {
        groups.notes.push({
          kind,
          ordinal,
          missingTarget: true
        });

        counts.notes++;
        continue;
      }

      const projected = blocks(note.content ?? [], kind + ':' + ordinal);

      groups.notes.push({
        kind,
        ordinal,
        facts: properties(note, new Set(['id', 'content']), ctx),
        blocks: projected
      });

      counts.notes++;
    }
  }

  if (counts.notes && model.__noteLayoutSettings) groups.notes.push({
    settings: ctx.canonical(model.__noteLayoutSettings)
  });

  groups.sections.unshift(ctx.canonical(model.section ?? {}));

  const root = properties(
    withoutInternals(model),
    new Set(['body', 'section', 'headers', 'footers', 'footnotes', 'endnotes']),
    ctx
  );

  groups.settings.push(root);
  metadata('document', model);

  return {
    groups,
    counts,
    unavailableResources: [...ctx.unavailableResources]
  };
}

function differences(a, b, path = '', result = [], depth = 0) {
  if (depth > MAX_DEPTH || result.length >= MAX_DIFFS)
    throw new RangeError('paired difference budget exceeded');

  if (stable(a) === stable(b))
    return result;

  if (Array.isArray(a) && Array.isArray(b)) {
    if (a.length !== b.length) {
      result.push({
        path,
        kind: 'sequence-length-and-alignment',
        beforeLength: a.length,
        afterLength: b.length,
        before: a,
        after: b
      });

      return result;
    }

    for (let i = 0; i < a.length; i++)
      differences(a[i], b[i], path + '/' + i, result, depth + 1);

    return result;
  }

  if (a && b && typeof a === 'object' && typeof b === 'object') {
    for (const key of [...new Set([...Object.keys(a), ...Object.keys(b)])].sort()) differences(
      a[key],
      b[key],
      path + '/' + key.replaceAll('~', '~0').replaceAll('/', '~1'),
      result,
      depth + 1
    );

    return result;
  }

  result.push({
    path,
    kind: a === undefined || b === undefined ? 'presence' : 'value',

    before: a === undefined ? {
      absent: true
    } : a,

    after: b === undefined ? {
      absent: true
    } : b
  });

  return result;
}

export function compareModelMeaning(
  a,
  b,
  {
    resourcesA = [],
    resourcesB = []
  } = {}
) {
  const left = projectModelMeaning(a, {
          resources: resourcesA
        }),
        right = projectModelMeaning(b, {
          resources: resourcesB
        });

  const categories = {};

  for (const name of GROUPS) {
    const diffs = differences(left.groups[name], right.groups[name]);
    const exercised = ['notes', 'images', 'anchors', 'tables'].includes(name) ? left.counts[name] + right.counts[name] > 0 : true;

    categories[name] = {
      equal: !diffs.length,
      status: !exercised ? 'NOT_EXERCISED' : diffs.length ? 'DIFFERENT_FACTS_REQUIRE_CLASSIFICATION' : 'AGREEMENT_FOR_COMPARED_FACTS',
      differences: diffs
    };
  }

  // Matching failures never establish image equivalence, even when both sides
  // report the same extraction error. Retained facts can still be inspected.
  // Paragraph formats depend on resources through picture-bullet numbering.
  if (left.unavailableResources.length || right.unavailableResources.length) for (const name of ['content', 'paragraphFormats', 'runFormats', 'images', 'anchors']) if (categories[name].status !== 'NOT_EXERCISED') {
    categories[name].equal = null;
    categories[name].status = 'INCOMPLETE_RESOURCE_EVIDENCE';
  }

  if ([...left.groups.notes, ...right.groups.notes].some(note => note.missingTarget)) {
    categories.notes.equal = null;
    categories.notes.status = 'INCOMPLETE_NOTE_EVIDENCE';
  }

  // Preserve the entire raw difference ledger: unknown public subtrees may
  // contain names also used for known references. Their diagnostic projection
  // never substitutes for this source-representation evidence.
  const raw = differences(a, b);

  const representationDifferences = raw.filter(d => d.kind === 'presence' || identityKeys.has(d.path.split('/').at(-1)));

  return {
    leftCounts: left.counts,
    rightCounts: right.counts,

    unavailableResources: {
      left: left.unavailableResources,
      right: right.unavailableResources
    },

    categories,
    rawDifferences: raw,
    representationDifferences,
    scope: 'Meaning of retained parser facts. Agreement does not prove source completeness or Office fidelity; unknown acquisition metadata and unused domains remain explicit.'
  };
}

function boundedJson(path) {
  // Non-blocking open also rejects FIFOs before waiting for a writer.
  const fd = openSync(path, fsConstants.O_RDONLY | (fsConstants.O_NONBLOCK ?? 0));

  try {
    const info = fstatSync(fd);

    if (!info.isFile() || info.size > MAX_FILE_BYTES)
      throw new RangeError('paired JSON byte budget exceeded');

    const bytes = Buffer.allocUnsafe(info.size + 1);
    let offset = 0;

    while (offset < bytes.length) {
      const read = readSync(fd, bytes, offset, bytes.length - offset, null);

      if (!read)
        break;

      offset += read;
    }

    if (offset !== info.size || fstatSync(fd).size !== info.size)
      throw new Error('paired JSON changed during bounded read');

    const input = bytes.subarray(0, offset);

    return {
      value: JSON.parse(input),
      sha256: createHash('sha256').update(input).digest('hex')
    };
  } finally {
    closeSync(fd);
  }
}

function unpackInput(value) {
  if (Array.isArray(value?.body)) return {
    document: value,
    resources: []
  };

  if (value?.body !== undefined || !Array.isArray(value?.document?.body) || !Array.isArray(value?.resources) || Object.keys(value).some(key => key !== 'document' && key !== 'resources'))
    throw new TypeError('invalid paired model envelope');

  return value;
}

if (process.argv[1] && pathToFileURL(resolve(process.argv[1])).href === import.meta.url) {
  const [leftFile, rightFile, outFile] = process.argv.slice(2);

  if (!leftFile || !rightFile || !outFile) throw new Error(
    'Usage: legacy-doc-paired-model-compare.mjs <left-model-or-envelope.json> <right-model-or-envelope.json> <report.json>'
  );

  // Envelopes supply extracted-byte inventory: {document, resources:[{path,
  // sha256,bytes}|{path,error}]}. The CLI never opens arbitrary resource paths.
  const left = boundedJson(leftFile),
        right = boundedJson(rightFile),
        a = unpackInput(left.value),
        b = unpackInput(right.value),
        result = compareModelMeaning(a.document, b.document, {
          resourcesA: a.resources,
          resourcesB: b.resources
        });

  const bytes = Buffer.from(JSON.stringify(result, null, 2) + '\n');

  if (bytes.length > 16 * 1024 * 1024)
    throw new RangeError('paired report byte budget exceeded');

  writeFileSync(outFile, bytes);

  console.log(JSON.stringify({
    leftSha256: left.sha256,
    rightSha256: right.sha256,

    categories: Object.fromEntries(Object.entries(result.categories).map(([k, v]) => [k, {
      status: v.status,
      differences: v.differences.length
    }]))
  }));
}
