import { drawingMLGeometryBounds } from '@silurus/ooxml-core';
import type { BodyLayoutInput, BodyParagraphSourceInput } from './body-layout-input.js';
import type { ParagraphAcquisitionInput } from './text.js';
import type { AnchorReferenceFramesInput } from './anchor-frame.js';
import { resolveAnchorPlacement } from './anchor-frame.js';
import { NativeReadingSceneError } from './native-reading-block-scene.js';
import { endnoteIdsInRetainedSlice, footnoteIdsInRetainedSlice } from './note-reference-ownership.js';
import type { acquireParagraphAnchorRunStage } from './paragraph.js';
import type { LayoutRect, PaintResourceRegistry, ParagraphLayout, DrawingLayout, ImagePaintResourceDescriptor } from './types.js';
import { isDeepFrozenPlainDataRoot, snapshotPlainData } from './plain-data.js';
import { translateDrawing, translateParagraphLayout } from './retained-geometry-translation.js';
import { acquireNativeReadingImagePlan, nativeReadingImageFrame } from './native-reading-image-frame.js';

const MAX_RUN_WORK = 16384;
function fail(message: string): never { throw new NativeReadingSceneError('placement', message); }

/** Structural source ownership only: no label/caption/proximity inference.
 * MS-DOC §2.9.253 supplies column/paragraph offsets. Preserving those offsets
 * across this source run with one common translation is a disclosed library
 * reading policy. Empty paragraph marks retain their source advances inside
 * this structural run; real text, inline content, table/section/break entries
 * terminate it. Activation requires a requested vector group; standalone
 * requested pictures retain the existing paragraph reading owner. This is
 * not an overflow fallback. This is not a reconstruction of Word's automatic wrap contour. */
export function nativeReadingAnchorRuns(
  input: BodyLayoutInput, resolve: (source: BodyParagraphSourceInput['source']) => ParagraphAcquisitionInput,
): ReadonlyMap<number, readonly BodyParagraphSourceInput[]> {
  const result = new Map<number, readonly BodyParagraphSourceInput[]>();
  let run: BodyParagraphSourceInput[] = [], requested = false, work = 0;
  const flush = () => {
    if (requested) {
      if (work > MAX_RUN_WORK) throw new NativeReadingSceneError('budget', 'anchor run source allowance exceeded');
      result.set(run[0].source.path[0], Object.freeze(run));
    }
    run = []; requested = false; work = 0;
  };
  for (const entry of input.sequence) {
    if (entry.kind !== 'body-block' || entry.block.kind !== 'paragraph') { flush(); continue; }
    const block = entry.block;
    if (block.source.story !== 'body' || block.source.path.length !== 1 || !Number.isSafeInteger(block.source.path[0]))
      throw new NativeReadingSceneError('source-ownership','anchor run requires body-root paragraph identities');
    const p = resolve(block.source);
    const eligible = !block.pageBreakBefore && !block.continuousSectionRole && !p.framePr && !p.numbering
      && p.runs.every(r => r.type === 'anchorHost'
        || ((r.type === 'shape' || r.type === 'image' || r.type === 'chart' || r.type === 'unavailableDrawing') && !!r.anchorAcquisitionInput));
    if (!eligible) { flush(); continue; }
    run.push(block); work += p.runs.length + 1;
    requested ||= p.runs.some(r=>r.type === 'shape' && r.anchorAcquisitionInput?.group != null
      && r.anchorAcquisitionInput.nativeReadingRelocation === 'completeScene');
  }
  flush(); return result;
}

function union(rects: readonly LayoutRect[]): LayoutRect {
  let left = Infinity, top = Infinity, right = -Infinity, bottom = -Infinity;
  for (const b of rects) {
    if (![b.xPt,b.yPt,b.widthPt,b.heightPt,b.xPt+b.widthPt,b.yPt+b.heightPt].every(Number.isFinite)
      || b.widthPt < 0 || b.heightPt < 0) fail('invalid complete anchor enclosure');
    left = Math.min(left,b.xPt); top = Math.min(top,b.yPt);
    right = Math.max(right,b.xPt+b.widthPt); bottom = Math.max(bottom,b.yPt+b.heightPt);
  }
  if (!Number.isFinite(left)) fail('anchor run has no retained allocation');
  return { xPt:left,yPt:top,widthPt:right-left,heightPt:bottom-top };
}

/** Ordinary peers retain their actual commands. Resolve enclosure through the
 * same pure geometry/crop owners consumed by canonical paint; no dry painting.
 * Unsupported clip, oriented story or command is an atomic class boundary. */
function completeDrawing(d: DrawingLayout, registry: PaintResourceRegistry | undefined, debit: (n: number) => void): DrawingLayout {
  if (d.clip || d.orientation === 'upright-physical') fail('anchor run requires unrotated section coordinates');
  const boxes: LayoutRect[] = [d.flowBounds];
  const commands = d.commands.map(command => {
    debit(1);
    if (command.kind === 'noop') return command;
    if (command.kind === 'drawingml-shape') {
      const units = command.plan.resolvedGeometry?.workUnits;
      if (!Number.isSafeInteger(units) || units === undefined || units < 1) fail('ordinary peer numeric geometry is unresolved');
      debit(units);
      const box = drawingMLGeometryBounds(command.plan,d.transform);
      if (box) boxes.push({xPt:box.x,yPt:box.y,widthPt:box.w,heightPt:box.h});
      return command;
    }
    if (command.kind === 'resource' && command.resourceKind === 'image' && !d.transform) {
      // Image descriptors share a union discriminator with picture bullets;
      // keep the canonical runtime source-kind/key/seal checks after resolving.
      const descriptor = registry?.resolve(command.resourceKey, 'image') as ImagePaintResourceDescriptor | undefined;
      if (!descriptor || !isDeepFrozenPlainDataRoot(descriptor) || descriptor.kind !== 'image'
        || descriptor.resourceKey !== command.resourceKey) fail('anchor run image resource is unavailable');
      const plan = command.nativeImagePlan ?? acquireNativeReadingImagePlan(descriptor,command.rect.widthPt,command.rect.heightPt);
      boxes.push(nativeReadingImageFrame(plan,command.rect));
      return {...command,nativeImagePlan:plan};
    }
    fail('anchor run cannot enclose this retained command');
  });
  return snapshotPlainData({...d,commands,inkBounds:union(boxes)},'complete reading peer',d);
}

export interface AnchorRunMember {
  readonly input: BodyParagraphSourceInput;
  readonly paragraph: ParagraphAcquisitionInput;
  readonly stage: ReturnType<typeof acquireParagraphAnchorRunStage>;
  readonly frames: AnchorReferenceFramesInput;
}
export interface ComposedAnchorRun {
  readonly members: readonly { readonly input: BodyParagraphSourceInput; readonly layout: ParagraphLayout; readonly blockExtentPt: number }[];
  readonly extentPt: number;
}

/** Every stage/story/resource is acquired before this sole composition owner
 * returns publication-ready paragraphs. Marks keep their original advances;
 * only the final source member owns allocation below the original mark run.
 * A new column must acquire all stages again, never translate stale frames. */
export function composeNativeReadingAnchorRun(
  members: readonly AnchorRunMember[], column: LayoutRect,
  externalInk: readonly LayoutRect[], registry?: PaintResourceRegistry,
): ComposedAnchorRun {
  if (!members.length) fail('empty anchor run');
  let work = 0;
  const debit = (n: number) => {
    work += n;
    if (!Number.isSafeInteger(work) || work > MAX_RUN_WORK) throw new NativeReadingSceneError('budget','anchor run geometry allowance exceeded');
  };
  const ownerIds = new Set<string>();
  const start = members[0].stage.acquired.layout.flowBounds.yPt;
  const prepared = members.map(({input,paragraph,stage,frames}) => {
    debit(paragraph.runs.length+1);
    const layout = stage.acquired.layout;
    // This profile contains anchored paint and zero-paint host/empty marks.
    // Paragraph decoration and bar-tab/line-number strokes need their own
    // complete enclosure before they can participate in horizontal rebasing.
    if (layout.borders.length || layout.shading || layout.lineNumbers?.length
      || layout.lines.some(line=>line.barTabRules?.length))
      fail('anchor run cannot rebase decorated paragraph paint');
    // markOnly is reserved for the genuinely run-free measurement path.
    // Anchor hosts use ordinary metric-only lines (paragraph-measure), so the
    // retained source and zero-paint host placements establish this class.
    const retainedAnchorDrawingIds = new Set(layout.drawings.flatMap(d=>d.anchorLayer ? [d.id] : []));
    if (layout.continuation || layout.clipBounds || layout.events.length || layout.resources.length
      || layout.lines.some(line=>line.placements.some(p=> {
        if (p.range.start !== p.range.end) return true;
        if (p.kind === 'anchor-host') return p.noteReference !== undefined || p.bounds.widthPt !== 0;
        // Canonical acquireAnchorOccurrence appends a zero-advance drawing
        // placement to its host line. Its retained occurrence, not merely the
        // placement kind, distinguishes it from visible inline content.
        return p.kind !== 'drawing' || p.advancePt !== 0
          || !retainedAnchorDrawingIds.has(p.drawingId);
      })))
      fail('anchor run must contain only complete anchored marks');
    const sourceOwners = new Set<string>();
    for (const r of paragraph.runs) {
      if (r.type === 'anchorHost') continue;
      if (r.type !== 'shape' && r.type !== 'image') fail('anchor run has an unsupported chart/unavailable drawing peer');
      const a = r.anchorAcquisitionInput;
      if (!a || a.simplePosition.enabled || a.horizontal.relativeFrom !== 'column'
        || a.vertical.relativeFrom !== 'paragraph' || a.horizontal.choice.kind !== 'offset'
        || a.vertical.choice.kind !== 'offset' || a.relativeSize.horizontal !== null || a.relativeSize.vertical !== null)
        fail('anchor run requires resolved column/paragraph offsets');
      // Within this disclosed structural run, each mark uses its canonical
      // source origin without applying earlier members' collision shifts.
      // Only authored overlap-permitted anchors belong to that policy class;
      // non-overlapping anchors retain the ordinary ordered registry owner.
      if (a.behavior.allowOverlapStatus !== 'valid' || a.behavior.allowOverlap !== true)
        fail('anchor run requires valid overlap-permitted source anchors');
      sourceOwners.add(a.occurrenceId);
    }
    const hosts = paragraph.runs.flatMap(r=>r.type === 'anchorHost' ? [r.anchorOccurrenceId] : []);
    if (hosts.length !== sourceOwners.size || new Set(hosts).size !== hosts.length
      || hosts.some(id=>id === undefined || !sourceOwners.has(id))) fail('anchor run host ownership is incomplete');
    for (const owner of sourceOwners) {
      if (ownerIds.has(owner)) fail('duplicate anchor run occurrence');
      ownerIds.add(owner);
    }
    if (layout.anchorFrames?.some(f => f.status !== 'resolved')) fail('ordinary peer anchor acquisition is incomplete');
    for (const frame of layout.anchorFrames ?? []) {
      if (frame.status !== 'resolved') fail('ordinary peer is unresolved');
      const horizontal = frame.axes.horizontal, vertical = frame.axes.vertical;
      if (horizontal.status !== 'resolved' || vertical.status !== 'resolved'
        || frame.geometry.objectFrame.xPt !== horizontal.resolvedOriginPt
        || frame.geometry.objectFrame.yPt !== vertical.resolvedOriginPt)
        fail('anchor run cannot preserve a collision-shifted source origin');
    }
    const drawings = layout.drawings.map(d => {
      const complete = completeDrawing(d,registry,debit);
      if (!complete.anchorLayer || complete.anchorLayer.verticalOwnership !== 'host')
        fail('anchor run ordinary paint has unresolved host ownership');
      // Ordinary column anchors remain page-owned during paragraph projection.
      // This admitted reading cohort instead owns one common relocation of its
      // complete geometry. Retained ownership changes only here; original
      // column/paragraph source references and ordinary acquisition stay intact.
      return {...complete,anchorLayer:{...complete.anchorLayer,horizontalOwnership:'host' as const}};
    });
    const lines = [...layout.lines];
    const requestedIds: string[] = [];
    for (const scene of stage.scenes) {
      debit(scene.observedPaintOperations);
      const first = paragraph.runs.find(r => (r.type === 'shape' || r.type === 'image')
        && r.anchorAcquisitionInput?.occurrenceId === scene.occurrenceId);
      if (!first || (first.type !== 'shape' && first.type !== 'image') || !first.anchorAcquisitionInput) fail('scene source owner is missing');
      const resolved = resolveAnchorPlacement({acquisition:first.anchorAcquisitionInput,
        frames:{...frames,paragraph:layout.flowBounds,line:null,character:null}});
      const h = resolved.axes.horizontal,v = resolved.axes.vertical;
      // Completed scenes already own full paint geometry. This acquisition is
      // solely the authored origin, not ordinary wrap geometry admission.
      if (resolved.status !== 'resolved' || h.status !== 'resolved' || v.status !== 'resolved')
        fail('scene authored origin is unresolved');
      const behavior = first.anchorAcquisitionInput.behavior;
      if (typeof behavior.behindDoc !== 'boolean' || typeof behavior.relativeHeight !== 'number') fail('scene stacking owner is unresolved');
      const d = {...translateDrawing(scene.drawing,{xPt:h.resolvedOriginPt,yPt:v.resolvedOriginPt}),
        anchorLayer:{occurrenceId:scene.occurrenceId,behindDoc:behavior.behindDoc,
          relativeHeight:behavior.relativeHeight,
          sourceOrder:scene.memberRunIndices[0],horizontalOwnership:'host' as const,verticalOwnership:'host' as const}};
      drawings.push(d); requestedIds.push(d.id);
      const end = lines.at(-1)?.range.end ?? 0;
      lines.push({range:{start:end,end},bounds:d.flowBounds,baselinePt:d.flowBounds.yPt+d.flowBounds.heightPt,
        advancePt:0,placements:[{kind:'drawing',drawingId:d.id,range:{start:end,end},bounds:d.inkBounds,advancePt:0}]});
    }
    if (drawings.some(d=>d.diagnostics?.some(diagnostic=>diagnostic.severity==='error' || diagnostic.code==='UNSUPPORTED_FEATURE' || diagnostic.code==='MISSING_RESOURCE'))) fail('anchor run contains unsupported retained paint');
    const covered = new Set(drawings.map(d=>d.anchorLayer?.occurrenceId ?? stage.scenes.find(s=>s.drawing.id===d.id)?.occurrenceId));
    if (covered.size !== sourceOwners.size || [...sourceOwners].some(id=>!covered.has(id))) fail('anchor run omitted a source occurrence');
    const boxIds = new Set(drawings.flatMap(d=>d.textBoxIds ?? []));
    if (boxIds.size !== layout.textBoxes.length || layout.textBoxes.some(b=>!boxIds.has(b.id) || b.verticalMode !== undefined))
      fail('anchor run textbox story ownership is incomplete');
    for (const box of layout.textBoxes) {
      // The canonical story owner retains unsupported source blocks as
      // diagnostics instead of paint nodes. An empty retained block list is
      // therefore not evidence of complete source content acquisition.
      if (box.story.diagnostics.some(d=>d.severity === 'error' || d.code === 'UNSUPPORTED_FEATURE' || d.code === 'MISSING_RESOURCE'))
        fail('anchor run contains unsupported owned textbox content');
      const t = box.transform;
      // The ordinary complete-story owner retains page-point glyph ink for
      // this profile. Rotated/local or nested stories require a corresponding
      // translation/enclosure proof before joining a source-owned run.
      if (t.a !== 1 || t.b !== 0 || t.c !== 0 || t.d !== 1 || t.e !== 0 || t.f !== 0
        || box.story.bandDependent === true || box.story.blocks.some(b=>b.kind !== 'paragraph' || b.drawings.length || b.textBoxes.length || b.resources.length
          || b.borders.length || b.shading || b.lineNumbers?.length || b.lines.some(line=>line.barTabRules?.length)))
        fail('anchor run requires complete horizontal glyph-only textbox stories');
      for (const block of box.story.blocks) {
        if (block.kind !== 'paragraph') fail('anchor run textbox story kind is unsupported');
        if (block.diagnostics?.some(d=>d.severity === 'error' || d.code === 'UNSUPPORTED_FEATURE' || d.code === 'MISSING_RESOURCE'))
          fail('anchor run contains unsupported owned textbox content');
        // Atom admission does not commit notes from independent member stories.
        // Their authority is retained on placements, independently of events.
        if (block.events.length || footnoteIdsInRetainedSlice(block).length
          || endnoteIdsInRetainedSlice(block).length
          || block.lines.some(line=>line.placements.some(p=>p.kind === 'text' && p.dependency !== undefined)))
          fail('anchor run textbox requires independent note/event acquisition');
      }
    }
    const ink = union([layout.inkBounds,...drawings.map(d=>d.inkBounds),...layout.textBoxes.flatMap(b=>[b.inkBounds,b.story.inkBounds])]);
    return {input,sourceOwners,stageLineCount:layout.lines.length,stageInk:layout.inkBounds,
      layout:snapshotPlainData({...layout,drawings,lines,inkBounds:ink,
      ...(layout.anchorCollisions ? {anchorCollisions:layout.anchorCollisions.map(entry=>sourceOwners.has(entry.occurrenceId)
        ? {...entry,horizontalOwnership:'host' as const} : entry)} : {}),
      ...(requestedIds.length ? {nativeReadingRelocations:requestedIds} : {}),
      ...(paragraph.runs.some(r=>r.type === 'image' && [r.anchorAcquisitionInput?.nativePictureMetadata?.inactiveFillCarrier,
        r.anchorAcquisitionInput?.nativePictureMetadata?.inactiveLineCarrier].some(p=>p != null && p.retention !== 'ignoredZeroIndex'))
        ? {nativeReadingInactivePictureData:true} : {})},'composed reading anchor paragraph',layout)};
  });
  // Full-column paragraph flow boxes reserve vertical mark advances, not
  // horizontal paint. Include every authored drawing reserve, canonical paint
  // enclosure and owned glyph story in the independent horizontal fit proof.
  const content = union(prepared.flatMap(m=>[
    ...m.layout.drawings.flatMap(d=>[d.flowBounds,d.inkBounds]),
    ...m.layout.textBoxes.flatMap(b=>[b.inkBounds,b.story.inkBounds]),
  ]));
  // Reserve authored drawing height even when its painted child occupies only
  // part of that frame. Marks and complete paint/reserve share one vertical
  // envelope; their full-column logical widths do not enter horizontal fitting.
  const all = union([...prepared.flatMap(m=>[m.layout.flowBounds,m.layout.inkBounds]),content]);
  // Column rebase is library policy. Preserve relative positions instead of
  // individually normalizing scenes. The nearest-to-zero feasible common
  // translation supports either overhanging edge; an empty interval is a real
  // complete-content overflow and cannot be repaired without changing content.
  const minimumDx = column.xPt-content.xPt;
  const maximumDx = column.xPt+column.widthPt-(content.xPt+content.widthPt);
  if (minimumDx > maximumDx) fail('anchor run exceeds column width');
  const dx = Math.max(minimumDx,Math.min(0,maximumDx));
  let dy = Math.max(0,start-all.yPt);
  const ordered = externalInk.map(b=>{debit(1);union([b]);return b;}).sort((a,b)=>(a.yPt+a.heightPt)-(b.yPt+b.heightPt));
  for (const b of ordered) {
    if (content.xPt+dx < b.xPt+b.widthPt && content.xPt+dx+content.widthPt > b.xPt
      && all.yPt+dy < b.yPt+b.heightPt && all.yPt+dy+all.heightPt > b.yPt)
      dy = b.yPt+b.heightPt-all.yPt;
  }
  const extent = all.yPt+dy+all.heightPt-start;
  if (!Number.isFinite(extent) || extent < 0 || extent > column.heightPt) fail('complete anchor run exceeds a full column');
  const placed = prepared.map(m=> {
    const translated = translateParagraphLayout(m.layout,{xPt:dx,yPt:dy});
    // Canonical translation moves all commands, glyphs, owned stories and
    // anchor diagnostics together. The logical paragraph column allocation
    // remains horizontally authoritative; only its vertical mark origin moves.
    const lines = translated.lines.map((line,i)=> {
      if (i >= m.stageLineCount) return line; // Appended scene lines own paint geometry.
      const original = m.layout.lines[i];
      // Stage metric lines and zero-paint hosts retain the logical column
      // origin; their drawing placements still follow the common paint shift.
      return {...line,bounds:{...line.bounds,xPt:original.bounds.xPt},
        ...(line.wrapBounds && original.wrapBounds ? {wrapBounds:{...line.wrapBounds,xPt:original.wrapBounds.xPt}} : {}),
        ...(line.wrapFragments && original.wrapFragments ? {wrapFragments:line.wrapFragments.map((b,j)=>
          ({...b,xPt:original.wrapFragments![j].xPt}))} : {}),
        placements:line.placements.map((p,j)=> {
          if (p.kind !== 'anchor-host') return p;
          const originalHost = original.placements[j];
          if (originalHost.kind !== 'anchor-host') fail('anchor run metric host ownership changed');
          return {...p,bounds:{...p.bounds,xPt:originalHost.bounds.xPt}};
        })};
    });
    const layout = snapshotPlainData({...translated,lines,
      inkBounds:union([{...m.stageInk,yPt:m.stageInk.yPt+dy},
        ...translated.drawings.map(d=>d.inkBounds),...translated.textBoxes.flatMap(b=>[b.inkBounds,b.story.inkBounds])]),
      flowBounds:{...translated.flowBounds,
      xPt:m.layout.flowBounds.xPt,widthPt:m.layout.flowBounds.widthPt},
      ...(translated.paragraphMark && m.layout.paragraphMark ? {paragraphMark:{...translated.paragraphMark,
        bounds:{...translated.paragraphMark.bounds,xPt:m.layout.paragraphMark.bounds.xPt},
        ...(translated.paragraphMark.wrapBounds && m.layout.paragraphMark.wrapBounds ? {wrapBounds:{
          ...translated.paragraphMark.wrapBounds,xPt:m.layout.paragraphMark.wrapBounds.xPt}} : {})}} : {}),
      // Inherited registry authority belongs to its own previously admitted
      // occurrence. Only this member's owned wrap/collision geometry moves.
      exclusions:translated.exclusions.map((entry,i)=> {
        const original = m.layout.exclusions[i];
        return original.anchorOccurrenceId && m.sourceOwners.has(original.anchorOccurrenceId) ? entry : original;
      }),
      ...(translated.anchorCollisions ? {anchorCollisions:translated.anchorCollisions.map((entry,i)=>
        m.sourceOwners.has(entry.occurrenceId) ? entry : m.layout.anchorCollisions![i])} : {}),
    },'placed complete reading anchor paragraph',m.layout);
    return {input:m.input,layout,blockExtentPt:m.layout.advancePt};
  });
  // The final member owns the remaining translated envelope allocation once.
  // No inter-member page/keep/wrap retry changes the acquired source origins.
  const last = placed.at(-1)!;
  const tail = Math.max(0,extent-placed.reduce((n,m)=>n+m.layout.advancePt,0));
  last.blockExtentPt += tail;
  return Object.freeze({members:Object.freeze(placed),extentPt:extent});
}
