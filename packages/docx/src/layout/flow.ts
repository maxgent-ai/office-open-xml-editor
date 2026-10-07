import type {
  BlockLayoutAlgorithms,
  FlowLayout,
  FlowLayoutInput,
  LayoutRect,
  LayoutServices,
  ParagraphLayout,
  TableLayout,
} from './types.js';
import { LayoutInvariantError } from './diagnostics.js';

export class FlowCapacityExceededError extends LayoutInvariantError {
  constructor(
    readonly containerId: string,
    readonly layoutId: string,
  ) {
    super('INVALID_GEOMETRY', `${layoutId} exceeds the available flow capacity`);
    this.name = 'FlowCapacityExceededError';
  }
}

function unionBounds(bounds: readonly LayoutRect[], fallback: LayoutRect): LayoutRect {
  if (bounds.length === 0) {
    return { xPt: fallback.xPt, yPt: fallback.yPt, widthPt: 0, heightPt: 0 };
  }
  const left = Math.min(...bounds.map((rect) => rect.xPt));
  const top = Math.min(...bounds.map((rect) => rect.yPt));
  const right = Math.max(...bounds.map((rect) => rect.xPt + rect.widthPt));
  const bottom = Math.max(...bounds.map((rect) => rect.yPt + rect.heightPt));
  return { xPt: left, yPt: top, widthPt: right - left, heightPt: bottom - top };
}

export function layoutFlowBlocks(
  input: FlowLayoutInput,
  services: LayoutServices,
  algorithms: BlockLayoutAlgorithms,
): FlowLayout {
  const blocks: Array<ParagraphLayout | TableLayout> = [];
  let cursor = input.cursor;
  const bounds = input.container.bounds;
  if (![bounds.xPt, bounds.yPt, bounds.widthPt, bounds.heightPt].every(Number.isFinite)
    || bounds.widthPt < 0
    || bounds.heightPt < 0) {
    throw new LayoutInvariantError('INVALID_GEOMETRY', `${input.container.id} has invalid bounds`);
  }
  const containerBottom = input.container.bounds.yPt + input.container.bounds.heightPt;
  const capacityBottom = input.container.capacity === 'unbounded'
    ? Number.MAX_SAFE_INTEGER
    : containerBottom;
  const containerRight = input.container.bounds.xPt + input.container.bounds.widthPt;
  if (!Number.isFinite(cursor.xPt)
    || !Number.isFinite(cursor.yPt)
    || cursor.xPt < bounds.xPt
    || cursor.xPt > containerRight
    || cursor.yPt < bounds.yPt
    || cursor.yPt > containerBottom) {
    throw new LayoutInvariantError('INVALID_GEOMETRY', `${input.container.id} has an invalid initial flow cursor`);
  }

  for (const block of input.blocks) {
    const placement = {
      container: input.container,
      cursor,
      availableBounds: {
        xPt: input.container.bounds.xPt,
        yPt: cursor.yPt,
        widthPt: input.container.bounds.widthPt,
        heightPt: Math.max(0, capacityBottom - cursor.yPt),
      },
    };
    const result = block.kind === 'paragraph'
      ? algorithms.layoutParagraph(block, placement, services)
      : algorithms.layoutTable(block, placement, services);
    if (result.layout.flowDomainId !== input.container.id) {
      throw new LayoutInvariantError(
        'INVALID_REFERENCE',
        `${result.layout.id} belongs to ${result.layout.flowDomainId}, not ${input.container.id}`,
      );
    }
    if (input.container.capacity !== 'unbounded'
      && Number.isFinite(result.nextCursor.yPt)
      && result.nextCursor.yPt > containerBottom) {
      throw new FlowCapacityExceededError(input.container.id, result.layout.id);
    }
    if (!Number.isFinite(result.nextCursor.xPt)
      || !Number.isFinite(result.nextCursor.yPt)
      || result.nextCursor.xPt < input.container.bounds.xPt
      || result.nextCursor.xPt > containerRight
      || result.nextCursor.yPt < cursor.yPt) {
      throw new LayoutInvariantError('INVALID_GEOMETRY', `${result.layout.id} returned an invalid flow cursor`);
    }
    blocks.push(result.layout);
    cursor = result.nextCursor;
  }

  // A block with page-owned axes (a header/footer story-root cell-owner host
  // placed by a page/margin frame) is not in this flow's coordinates, so it
  // neither extends nor positions the flow's extent.
  const flowOwned = blocks.filter((block) => block.kind !== 'table' || !block.ownerHostPageAxes);
  return {
    source: input.source,
    container: input.container,
    blocks,
    nextCursor: cursor,
    flowDomainId: input.container.id,
    flowBounds: unionBounds(flowOwned.map((block) => block.flowBounds), input.container.bounds),
    inkBounds: unionBounds(flowOwned.map((block) => block.inkBounds), input.container.bounds),
    ...(input.container.capacity === 'unbounded'
      ? {}
      : { clipBounds: input.container.bounds }),
    advancePt: cursor.yPt - input.cursor.yPt,
    ordinaryFlow: true,
  };
}
