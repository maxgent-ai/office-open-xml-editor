import { expect, it } from 'vitest';
import { storedZip } from '@silurus/ooxml-core/testing';
import { assertBlipPixelEffectsBudget } from '../../core/src/image/blip-effects.js';
import { serializeWorkerError, deserializeWorkerError } from '@silurus/ooxml-core/worker';
import { materializePptxPresentation } from './pptx.js';

const a = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const p = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const rels = (body: string) => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${body}</Relationships>`;
const imageFill = (effects: string, prefix = 'a') => `<${prefix}:blipFill><a:blip r:embed="img">${effects}</a:blip><a:stretch/></${prefix}:blipFill>`;
function deck(effects: string, inherited: boolean) {
  const picture = `<p:pic><p:nvPicPr><p:cNvPr id="2" name="Picture"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>${imageFill(effects, 'p')}<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm></p:spPr></p:pic>`;
  return storedZip([
    ['[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/></Types>'],
    ['ppt/presentation.xml', `<p:presentation xmlns:p="${p}" xmlns:r="${r}"><p:sldIdLst><p:sldId id="256" r:id="slide"/></p:sldIdLst><p:sldSz cx="914400" cy="914400"/></p:presentation>`],
    ['ppt/_rels/presentation.xml.rels', rels(`<Relationship Id="slide" Type="${r}/slide" Target="slides/slide1.xml"/>`)],
    ['ppt/slides/slide1.xml', `<p:sld xmlns:p="${p}" xmlns:a="${a}" xmlns:r="${r}"><p:cSld><p:spTree>${inherited ? `<p:grpSp><p:grpSpPr>${imageFill(effects)}</p:grpSpPr>${picture}</p:grpSp>` : picture}</p:spTree></p:cSld></p:sld>`],
    ['ppt/slides/_rels/slide1.xml.rels', rels(`<Relationship Id="img" Type="${r}/image" Target="../media/image.png"/>`)],
    ['ppt/media/image.png', 'Source bytes are never decoded by effect preflight'],
  ]);
}

it.each([
  ['inherited picture fill cycle', '<a:fillOverlay blend="over"><a:grpFill/></a:fillOverlay>', true, 'unsupported-image-effect'],
  ['deep inline alphaMod', '<a:alphaMod><a:cont>'.repeat(40) + '<a:grayscl/>' + '</a:cont></a:alphaMod>'.repeat(40), false, 'ooxml-decoded-image-limit'],
] as const)('real WASM safely transports %s as a typed failure', async (_name, effects, inherited, code) => {
  const model = await materializePptxPresentation(deck(effects, inherited));
  const picture = model.slides[0].elements.find(e => e.type === 'picture');
  expect(picture?.type).toBe('picture');
  if (picture?.type !== 'picture') throw new Error('missing picture');
  let failure: unknown;
  try { assertBlipPixelEffectsBudget({ effects: picture.blipEffects ?? [] }, 0); }
  catch (error) { failure = error; }
  expect(failure).toMatchObject({ code });
  expect(deserializeWorkerError(serializeWorkerError(failure))).toMatchObject({ code });
});

it('real WASM reads a malformed Unicode color without panicking (channel 0 policy)', async () => {
  const model = await materializePptxPresentation(deck('<a:clrRepl><a:srgbClr val="A😀B"/></a:clrRepl>', false));
  const picture = model.slides[0].elements.find(e => e.type === 'picture');
  if (picture?.type !== 'picture') throw new Error('missing picture');
  const [effect] = picture.blipEffects ?? [];
  expect(effect?.type).toBe('colorReplace');
  expect(effect?.type === 'colorReplace' && /^[0-9A-F]{6}$/.test(effect.color)).toBe(true);
});
