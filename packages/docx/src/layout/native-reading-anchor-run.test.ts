import { describe, expect, it } from 'vitest';
import type { BodyElement, DocRun, SectionProps, ShapeRun } from '../types.js';
import type { InternalShapeRun } from '../parser-model.js';
import { layoutBodyModel } from '../test-support/document-layout.test-support.js';
import type { AnchorAcquisitionInput, AnchorEdgesInput, InternalAnchorRunWire } from './anchor-input.js';
import type { DrawingLayout, DrawingPlacement, ParagraphLayout } from './types.js';
import { resolveAnchorFrame, resolveAnchorPlacement } from './anchor-frame.js';

const edges: AnchorEdgesInput = {
  topPt: null, topStatus: 'missing', rightPt: null, rightStatus: 'missing',
  bottomPt: null, bottomStatus: 'missing', leftPt: null, leftStatus: 'missing',
};
function anchor(id: string, x: number, y: number, w: number, h: number): AnchorAcquisitionInput {
  return {
    occurrenceId: id,
    simplePosition: { enabled: false, status: 'valid', xPt: 0, xStatus: 'valid', yPt: 0, yStatus: 'valid' },
    horizontal: { relativeFrom: 'column', relativeFromStatus: 'valid', choice: { kind: 'offset', valuePt: x } },
    vertical: { relativeFrom: 'paragraph', relativeFromStatus: 'valid', choice: { kind: 'offset', valuePt: y } },
    extent: { widthPt: w, widthStatus: 'valid', heightPt: h, heightStatus: 'valid' },
    parentEffectExtent: edges, anchorDistances: edges, relativeSize: { horizontal: null, vertical: null },
    wrap: { kind: 'none', authoredKinds: ['wrapNone'], side: 'bothSides', distances: edges, effectExtent: null, polygon: null },
    behavior: { behindDoc: false, behindDocStatus: 'valid', relativeHeight: 3, relativeHeightStatus: 'valid',
      locked: false, lockedStatus: 'valid', allowOverlap: true, allowOverlapStatus: 'valid', layoutInCell: true, layoutInCellStatus: 'valid' },
    group: null,
  };
}
type AuthoredParagraph = Extract<BodyElement,{type:'paragraph'}>;
function paragraph(runs: DocRun[] = []): AuthoredParagraph {
  return { type: 'paragraph', alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: { rule: 'exact', value: 10 }, numbering: null,
    tabStops: [], defaultFontSize: 10, defaultFontFamily: 'Arial', widowControl: false, runs };
}
function text(value: string): Extract<DocRun,{type:'text'}> {
  return { type: 'text', text: value, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: 8, color: null, fontFamily: 'Arial', fontFamilyEastAsia: '', isLink: false,
    background: null, vertAlign: null, hyperlink: null };
}
function shape(acquisition: AnchorAcquisitionInput, label?: string): Extract<DocRun, { type: 'shape' }> & InternalAnchorRunWire {
  const value: InternalShapeRun & InternalAnchorRunWire = {
    widthPt: acquisition.extent.widthPt!, heightPt: acquisition.extent.heightPt!,
    anchorXPt: acquisition.horizontal.choice.kind === 'offset' ? acquisition.horizontal.choice.valuePt : 0,
    anchorYPt: acquisition.vertical.choice.kind === 'offset' ? acquisition.vertical.choice.valuePt : 0,
    anchorXFromMargin: true, anchorYFromPara: true, zOrder: 3,
    presetGeometry: 'rect', subpaths: [], fill: { fillType: 'solid', color: '00AA00' }, stroke: null,
    __anchorAcquisition: acquisition,
    ...(label ? { textBoxContent: [paragraph([text(label)])],
      textInsetL: 0, textInsetR: 0, textInsetT: 0, textInsetB: 0 } : {}),
  };
  return { type: 'shape', ...value };
}
function hosted(id: string, payload: DocRun): DocRun[] {
  const host: Extract<DocRun, { type: 'anchorHost' }> & { __anchorOccurrenceId: string } = {
    type: 'anchorHost', fontSize: 10, fontFamily: 'Arial', __anchorOccurrenceId: id,
  };
  return [host, payload];
}
function measureContext(): CanvasRenderingContext2D {
  let font = '10px Arial';
  return {
    get font() { return font; }, set font(v: string) { font = v; },
    letterSpacing: '0px', fontKerning: 'none',
    save() {}, restore() {},
    measureText(value: string) {
      const size = Number.parseFloat(/([\d.]+)px/u.exec(font)?.[1] ?? '10');
      return { width: [...value].length * size / 2, fontBoundingBoxAscent: size * .8,
        fontBoundingBoxDescent: size * .2, actualBoundingBoxAscent: size * .8, actualBoundingBoxDescent: size * .2 };
    },
  } as unknown as CanvasRenderingContext2D;
}
const section: SectionProps = {
  pageWidth: 220, pageHeight: 140, marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
  headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
};

function anchorRunBody(reading = true, wrapSide: 'bothSides' | null = 'bothSides'): AuthoredParagraph[] {
    const group: AnchorAcquisitionInput = {
      ...anchor('group', 50, 10, 60, 40), ...(reading ? {nativeReadingRelocation:'completeScene' as const} : {}),
      wrap: { kind: 'through', authoredKinds: ['wrapThrough'], side: wrapSide, distances: edges, effectExtent: null, polygon: null },
      group: { childSourceId: 'rectangle', sourceIndex: 0, sourceCount: 1, transformChain: [], childTransform: null,
        resolvedChildFrame: { offsetXPt: 0, offsetYPt: 0, widthPt: 60, heightPt: 40, rotationDeg: 0, flipH: false, flipV: false } },
    };
    const arrowShape: ShapeRun & InternalAnchorRunWire = {
      ...shape(anchor('arrow', 54, 0, 20, 2)), presetGeometry: 'straightConnector1', fill: null,
      stroke: '000000', strokeWidth: .5, tailEnd: { type: 'triangle', w: 'med', len: 'med' },
    };
    return [
      paragraph([...hosted('label-a', shape(anchor('label-a', 8, 0, 20, 10), 'A')),
        ...hosted('group', shape(group))]),
      paragraph([...hosted('arrow', { type: 'shape', ...arrowShape }),
        ...hosted('label-b', shape(anchor('label-b', 120, 10, 20, 10), 'B'))]),
      paragraph(), paragraph([text('After')]),
    ];
}

describe('source-owned native reading anchor runs', () => {
  it('fits a complete left-overhanging paint enclosure without translating full-column mark allocation', () => {
    const body = anchorRunBody();
    const leftPeer: AnchorAcquisitionInput = {...anchor('label-a', -8, 0, 20, 10),
      wrap:{kind:'square',authoredKinds:['wrapSquare'],side:'bothSides',distances:edges,effectExtent:null,polygon:null}};
    body[0].runs[1] = shape(leftPeer, 'A');
    const original = JSON.stringify(body);
    const layout = layoutBodyModel(body, section, measureContext());
    const nodes = layout.pages.flatMap(p=>p.layers.body).filter((n):n is ParagraphLayout=>n.kind==='paragraph');
    const first = nodes.find(n=>n.source.path[0]===0)!, next = nodes.find(n=>n.source.path[0]===1)!;
    const group = first.drawings.find(d=>first.nativeReadingRelocations?.includes(d.id))!;
    const label = first.drawings.find(d=>d.source.path.at(-1)===1)!;
    const arrow = next.drawings.find(d=>d.source.path.at(-1)===1)!;
    expect(label.flowBounds.xPt).toBe(10);
    expect(group.flowBounds.xPt-label.flowBounds.xPt).toBe(58);
    expect(group.flowBounds.yPt-label.flowBounds.yPt).toBe(10);
    expect(arrow.flowBounds.xPt-group.flowBounds.xPt).toBe(4);
    expect(arrow.flowBounds.yPt-group.flowBounds.yPt).toBe(0);
    const frame = first.anchorFrames!.find(f=>f.occurrenceId===label.anchorLayer!.occurrenceId)!;
    expect(frame.status).toBe('resolved');
    if (frame.status === 'resolved') expect(frame.geometry.objectFrame.xPt).toBe(label.flowBounds.xPt);
    const collision = first.anchorCollisions!.find(c=>c.occurrenceId===label.anchorLayer!.occurrenceId)!;
    expect(collision.bounds.xPt).toBe(label.flowBounds.xPt);
    const exclusion = first.exclusions.find(e=>e.anchorOccurrenceId===label.anchorLayer!.occurrenceId)!;
    expect(exclusion.bounds.xPt).toBe(label.flowBounds.xPt);
    const hostDrawing = first.lines.flatMap(l=>l.placements).find((p):p is DrawingPlacement=>p.kind==='drawing' && p.drawingId===label.id)!;
    expect(hostDrawing.bounds.xPt).toBe(label.flowBounds.xPt);
    expect(nodes.flatMap(p=>p.drawings)).toHaveLength(4);
    for (const p of nodes.filter(p=>p.source.path[0] < 3)) {
      expect(p.flowBounds.xPt).toBe(10);
      expect(p.flowBounds.widthPt).toBe(200);
    }
    // This genuinely empty source mark is below the square-wrapped peer.
    // Earlier host lines may retain a valid wrap offset, not the column edge.
    expect(first.lines[0].bounds.xPt).toBe(22); // Original square-wrap metric offset.
    const empty = nodes.find(p=>p.source.path[0]===2)!;
    expect(empty.paragraphMark!.bounds.xPt).toBe(10);
    expect(nodes.flatMap(p=>p.textBoxes).flatMap(b=>b.story.blocks).flatMap(b=>b.kind==='paragraph'
      ? b.lines.flatMap(l=>l.placements).flatMap(p=>p.kind==='text'?[p.text]:[]) : [])).toEqual(['A','B']);
    const labelText = first.textBoxes[0].story.blocks[0];
    expect(labelText.kind).toBe('paragraph');
    if (labelText.kind === 'paragraph') {
      const glyph = labelText.lines.flatMap(l=>l.placements).find(p=>p.kind==='text')!;
      expect(glyph.bounds.xPt).toBe(label.flowBounds.xPt);
      expect(glyph.bounds.xPt+glyph.bounds.widthPt).toBeLessThanOrEqual(label.flowBounds.xPt+label.flowBounds.widthPt);
    }
    expect(JSON.stringify(body)).toBe(original);
  });

  it('atomically refuses an actual complete paint enclosure wider than the column', () => {
    const body = anchorRunBody();
    body[0].runs[1] = shape(anchor('label-a', -8, 0, 20, 10), 'A');
    body[1].runs[3] = shape(anchor('label-b', 220, 10, 20, 10), 'B');
    const original = JSON.stringify(body);
    expect(()=>layoutBodyModel(body, section, measureContext())).toThrow('anchor run exceeds column width');
    expect(JSON.stringify(body)).toBe(original);
  });

  it('rebases a fitting right-overhanging complete scene left without shrinking its paint', () => {
    const body = anchorRunBody();
    body[0].runs[1] = shape(anchor('label-a', 50, 0, 20, 10), 'A');
    body[1].runs[3] = shape(anchor('label-b', 190, 10, 20, 10), 'B');
    const original = JSON.stringify(body);
    const layout = layoutBodyModel(body, section, measureContext());
    const nodes = layout.pages.flatMap(p=>p.layers.body).filter((n):n is ParagraphLayout=>n.kind==='paragraph');
    const first = nodes.find(n=>n.source.path[0]===0)!, next = nodes.find(n=>n.source.path[0]===1)!;
    const group = first.drawings.find(d=>first.nativeReadingRelocations?.includes(d.id))!;
    const right = next.drawings.find(d=>d.source.path.at(-1)===3)!;
    expect(right.flowBounds.xPt+right.flowBounds.widthPt).toBe(210);
    expect(group.flowBounds.xPt).toBe(50);
    expect(group.flowBounds.widthPt).toBe(60);
    expect(right.flowBounds.xPt-group.flowBounds.xPt).toBe(140);
    expect(first.flowBounds.xPt).toBe(10);
    expect(first.flowBounds.widthPt).toBe(200);
    expect(JSON.stringify(body)).toBe(original);
  });

  it('refuses real bar-tab paragraph paint before rebasing an anchored run', () => {
    const body = anchorRunBody();
    body[0].tabStops = [{pos:40,alignment:'bar',leader:'none'}];
    const original = JSON.stringify(body);
    expect(()=>layoutBodyModel(body, section, measureContext())).toThrow('anchor run cannot rebase decorated paragraph paint');
    expect(JSON.stringify(body)).toBe(original);
  });

  it('reserves the complete authored group height when its painted child is thin', () => {
    const body = anchorRunBody();
    const group: AnchorAcquisitionInput = {
      ...anchor('group', 50, 10, 60, 80), nativeReadingRelocation:'completeScene',
      wrap:{kind:'through',authoredKinds:['wrapThrough'],side:'bothSides',distances:edges,effectExtent:null,polygon:null},
      group:{childSourceId:'thin-child',sourceIndex:0,sourceCount:1,transformChain:[],childTransform:null,
        resolvedChildFrame:{offsetXPt:0,offsetYPt:0,widthPt:60,heightPt:1,rotationDeg:0,flipH:false,flipV:false}},
    };
    body[0].runs[3] = shape(group);
    const original = JSON.stringify(body);
    const layout = layoutBodyModel(body, section, measureContext());
    const nodes = layout.pages.flatMap(p=>p.layers.body).filter((n):n is ParagraphLayout=>n.kind==='paragraph');
    const first = nodes.find(n=>n.source.path[0]===0)!;
    const scene = first.drawings.find(d=>first.nativeReadingRelocations?.includes(d.id))!;
    const following = nodes.find(n=>n.source.path[0]===3)!;
    expect(scene.flowBounds.heightPt).toBe(80);
    expect(scene.inkBounds.heightPt).toBe(1);
    expect(following.flowBounds.yPt).toBeGreaterThanOrEqual(scene.flowBounds.yPt+scene.flowBounds.heightPt);
    expect(JSON.stringify(body)).toBe(original);
  });

  it('acquires canonical source origins without requiring a side for a disclosed missing-contour scene', () => {
    const body = anchorRunBody(true,null);
    const original = JSON.stringify(body);
    const layout = layoutBodyModel(body,section,measureContext());
    const nodes = layout.pages.flatMap(p=>p.layers.body).filter((n):n is ParagraphLayout=>n.kind==='paragraph');
    const first = nodes.find(n=>n.source.path[0]===0)!, next = nodes.find(n=>n.source.path[0]===1)!;
    const group = first.drawings.find(d=>first.nativeReadingRelocations?.includes(d.id))!;
    const label = first.drawings.find(d=>d.source.path.at(-1)===1)!;
    const arrow = next.drawings.find(d=>d.source.path.at(-1)===1)!;
    expect(group.flowBounds.xPt-label.flowBounds.xPt).toBe(42);
    expect(group.flowBounds.yPt-label.flowBounds.yPt).toBe(10);
    expect(arrow.flowBounds.xPt-group.flowBounds.xPt).toBe(4);
    expect(arrow.flowBounds.yPt-group.flowBounds.yPt).toBe(0);
    expect(nodes.flatMap(p=>p.drawings)).toHaveLength(4);
    expect(nodes.flatMap(p=>p.textBoxes).flatMap(b=>b.story.blocks).flatMap(b=>b.kind==='paragraph'
      ? b.lines.flatMap(l=>l.placements).flatMap(p=>p.kind==='text'?[p.text]:[]) : [])).toEqual(['A','B']);
    expect(JSON.stringify(body)).toBe(original);
    // Ordinary wrap acquisition still requires a side. Position-only reading
    // acquisition must retain the same canonical rejection of invalid axes.
    const authored = {...anchor('placement-owner',50,10,60,40),
      wrap:{kind:'through' as const,authoredKinds:['wrapThrough'],side:null,distances:edges,effectExtent:null,polygon:null}};
    const frames = {page:{xPt:0,yPt:0,widthPt:220,heightPt:140},margin:{xPt:10,yPt:10,widthPt:200,heightPt:120},
      column:{xPt:10,yPt:10,widthPt:200,heightPt:120},paragraph:{xPt:10,yPt:10,widthPt:200,heightPt:10},
      line:null,character:null,pageParity:'odd' as const};
    expect(resolveAnchorFrame({acquisition:authored,frames})).toMatchObject({status:'unsupported',
      issues:[{code:'invalid-wrap-side',path:'wrap.side'}]});
    expect(resolveAnchorPlacement({acquisition:{...authored,
      horizontal:{...authored.horizontal,choice:{kind:'invalid'}}},frames})).toMatchObject({status:'unsupported',
      issues:[{code:'invalid-axis-choice'}]});
  });

  it('keeps sibling labels and following paragraph anchors colocated with a complete reading group', () => {
    const body = anchorRunBody();
    const original = JSON.stringify(body);
    const layout = layoutBodyModel(body, section, measureContext());
    const nodes = layout.pages.flatMap(page => page.layers.body)
      .filter((node): node is ParagraphLayout => node.kind === 'paragraph');
    const first = nodes.find(node => node.source.path[0] === 0)!;
    const next = nodes.find(node => node.source.path[0] === 1)!;
    const requested = first.drawings.find(drawing => first.nativeReadingRelocations?.includes(drawing.id))!;
    const byRun = (node: ParagraphLayout, index: number): DrawingLayout => {
      const result = node.drawings.find(drawing => drawing.source.path.at(-1) === index);
      if (!result) throw new Error('Source-owned drawing was omitted');
      return result;
    };
    const labelA = byRun(first, 1), arrow = byRun(next, 1), labelB = byRun(next, 3);
    // Independent authored offsets plus the first paragraph's exact 10pt mark.
    // One common translation may change the block origin, never these relations.
    expect(requested.flowBounds.xPt - labelA.flowBounds.xPt).toBe(42);
    expect(requested.flowBounds.yPt - labelA.flowBounds.yPt).toBe(10);
    expect(arrow.flowBounds.xPt - requested.flowBounds.xPt).toBe(4);
    expect(arrow.flowBounds.yPt - requested.flowBounds.yPt).toBe(0);
    expect(labelB.flowBounds.xPt - requested.flowBounds.xPt).toBe(70);
    expect(labelB.flowBounds.yPt - requested.flowBounds.yPt).toBe(10);
    expect(first.textBoxes.flatMap(box => box.story.blocks).flatMap(block =>
      block.kind === 'paragraph' ? block.lines.flatMap(line => line.placements)
        .flatMap(p => p.kind === 'text' ? [p.text] : []) : [])).toEqual(['A']);
    expect(next.textBoxes.flatMap(box => box.story.blocks).flatMap(block =>
      block.kind === 'paragraph' ? block.lines.flatMap(line => line.placements)
        .flatMap(p => p.kind === 'text' ? [p.text] : []) : [])).toEqual(['B']);
    expect(nodes.reduce((n, p) => n + p.drawings.length, 0)).toBe(4);
    const after = nodes.find(node => node.source.path[0] === 3)!;
    expect(after.flowBounds.yPt).toBe(Math.max(...[requested, labelA, arrow, labelB]
      .map(drawing => drawing.inkBounds.yPt + drawing.inkBounds.heightPt)));
    expect(layout.pages).toHaveLength(1);
    expect(JSON.stringify(body)).toBe(original);
  });
  it('reacquires every source member together after a fresh-page admission retry', () => {
    const body = [...Array.from({length:9},()=>paragraph([text('Before')])),...anchorRunBody()];
    const layout = layoutBodyModel(body,section,measureContext());
    const page = layout.pages.find(p=>p.layers.body.some(n=>n.source.path[0]===9))!;
    expect(layout.pages.indexOf(page)).toBe(1);
    const first = page.layers.body.find((n): n is ParagraphLayout=>n.kind==='paragraph' && n.source.path[0]===9)!;
    const next = page.layers.body.find((n): n is ParagraphLayout=>n.kind==='paragraph' && n.source.path[0]===10)!;
    expect(first.lines[0].bounds.yPt).toBe(10);
    expect(next.lines[0].bounds.yPt-first.lines[0].bounds.yPt).toBe(10);
    const group = first.drawings.find(d=>first.nativeReadingRelocations?.includes(d.id))!;
    const arrow = next.drawings.find(d=>d.source.path.at(-1)===1)!;
    expect(arrow.flowBounds.xPt-group.flowBounds.xPt).toBe(4);
    expect(arrow.flowBounds.yPt-group.flowBounds.yPt).toBe(0);
    expect(layout.pages[0].layers.body.some(n=>n.source.path[0]===9 || n.source.path[0]===10)).toBe(false);
  });

  it('fails the whole closure when a later owned anchor uses an unsupported reference frame', () => {
    const body = anchorRunBody();
    const bad = anchor('unsupported-peer',4,0,10,10);
    body[2] = paragraph(hosted('unsupported-peer',shape({...bad,vertical:{...bad.vertical,relativeFrom:'page'}})));
    const original = JSON.stringify(body);
    expect(()=>layoutBodyModel(body,section,measureContext())).toThrow('resolved column/paragraph offsets');
    expect(JSON.stringify(body)).toBe(original);
  });

  it('keeps default missing-contour acquisition unsupported without a reading request', () => {
    const body = anchorRunBody(false);
    const layout = layoutBodyModel(body,section,measureContext());
    const paragraphs = layout.pages.flatMap(p=>p.layers.body).filter((n):n is ParagraphLayout=>n.kind==='paragraph');
    expect(paragraphs.flatMap(p=>p.nativeReadingRelocations ?? [])).toEqual([]);
    expect(paragraphs.flatMap(p=>p.drawings)).toHaveLength(3);
    expect(paragraphs.flatMap(p=>p.anchorFrames ?? []).some(f=>f.status==='unsupported'
      && f.issues.some(i=>i.code==='invalid-wrap-polygon'))).toBe(true);
  });

  it('rejects spacing-folding formatting atomically instead of duplicating its ordinary owner', () => {
    const body = anchorRunBody();
    body[1].spaceBefore = 4;
    expect(()=>layoutBodyModel(body,section,measureContext())).toThrow('ordinary spacing/keep acquisition');
  });

  it('rejects a later source anchor that requires ordered non-overlap acquisition', () => {
    const body = anchorRunBody();
    const peer = anchor('non-overlap-peer',50,0,20,20);
    body[2] = paragraph(hosted(peer.occurrenceId,shape({...peer,
      behavior:{...peer.behavior,allowOverlap:false}})));
    const original = JSON.stringify(body);
    expect(()=>layoutBodyModel(body,section,measureContext())).toThrow('valid overlap-permitted source anchors');
    expect(JSON.stringify(body)).toBe(original);
  });

  it('rejects independently owned textbox notes before accepting any atom member', () => {
    for (const kind of ['footnote','endnote'] as const) {
      const body = anchorRunBody();
      const note: DocRun = {...text('1'),noteRef:{kind,id:'1'}};
      const payload: InternalShapeRun & InternalAnchorRunWire = {
        ...shape(anchor('note-story',4,0,20,20)),textBoxContent:[paragraph([note])],
      };
      body[2] = paragraph(hosted('note-story',{type:'shape',...payload}));
      const original = JSON.stringify(body);
      expect(()=>layoutBodyModel(body,section,measureContext())).toThrow('independent note/event acquisition');
      expect(JSON.stringify(body)).toBe(original);
    }
  });

  it('retains the canonical document-grid mark advance inside the same acquired atom', () => {
    const body = anchorRunBody();
    for (const p of body) { p.lineSpacing = {rule:'auto',value:1,explicit:false}; p.snapToGrid = true; }
    const gridSection = {...section,docGridType:'lines',docGridLinePitch:20};
    const layout = layoutBodyModel(body,gridSection,measureContext());
    const nodes = layout.pages.flatMap(p=>p.layers.body).filter((n):n is ParagraphLayout=>n.kind==='paragraph');
    const first = nodes.find(n=>n.source.path[0]===0)!, next = nodes.find(n=>n.source.path[0]===1)!;
    expect(first.lines[0].advancePt).toBe(20);
    expect(next.lines[0].bounds.yPt-first.lines[0].bounds.yPt).toBe(20);
    const group = first.drawings.find(d=>first.nativeReadingRelocations?.includes(d.id))!;
    const arrow = next.drawings.find(d=>d.source.path.at(-1)===1)!;
    expect(arrow.flowBounds.yPt-group.flowBounds.yPt).toBe(10);
    expect(nodes.flatMap(n=>n.textBoxes).flatMap(b=>b.story.blocks).flatMap(b=>b.kind==='paragraph'
      ? b.lines.flatMap(l=>l.placements).flatMap(p=>p.kind==='text'?[p.text]:[]) : [])).toEqual(['A','B']);
  });

  it('terminates the source run at real text and authored hard page breaks', () => {
    const body = anchorRunBody();
    body.splice(1,0,paragraph([text('Boundary')]));
    const layout = layoutBodyModel(body,section,measureContext());
    const nodes = layout.pages.flatMap(p=>p.layers.body).filter((n):n is ParagraphLayout=>n.kind==='paragraph');
    const first = nodes.find(n=>n.source.path[0]===0)!, following = nodes.find(n=>n.source.path[0]===2)!;
    const group = first.drawings.find(d=>first.nativeReadingRelocations?.includes(d.id))!;
    const arrow = following.drawings.find(d=>d.source.path.at(-1)===1)!;
    expect(arrow.flowBounds.yPt).toBeGreaterThanOrEqual(group.inkBounds.yPt+group.inkBounds.heightPt);
    expect(nodes.find(n=>n.source.path[0]===1)!.lines.flatMap(l=>l.placements).flatMap(p=>p.kind==='text'?[p.text]:[])).toEqual(['Boundary']);
    const broken = anchorRunBody(); broken[1].pageBreakBefore = true;
    const pages = layoutBodyModel(broken,section,measureContext()).pages;
    expect(pages.findIndex(p=>p.layers.body.some(n=>n.source.path[0]===0))).toBe(0);
    expect(pages.findIndex(p=>p.layers.body.some(n=>n.source.path[0]===1))).toBe(1);
  });

  it('does not publish a valid scene prefix when a later owned textbox cannot be acquired', () => {
    const body = anchorRunBody();
    const payload: InternalShapeRun & InternalAnchorRunWire = {
      ...shape(anchor('unsupported-story',4,0,10,10)),
      textBoxContent:[{type:'unsupportedTextBoxBlock',qName:'w:altChunk',sourcePath:[0]}],
    };
    body[2] = paragraph(hosted('unsupported-story',{type:'shape',...payload}));
    const original = JSON.stringify(body);
    expect(()=>layoutBodyModel(body,section,measureContext())).toThrow('unsupported owned textbox content');
    expect(JSON.stringify(body)).toBe(original);
  });

});
