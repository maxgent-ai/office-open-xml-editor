import type { OoxmlResourceUsageSnapshot, WorkerBridgeTransport } from '@silurus/ooxml-core';
import {
  BoundedPullSession,
  HARD_MAX_PPTX_SLIDE_JSON_BYTES,
  PULL_SESSION_PROTOCOL,
  pullWithCreditRetry,
  type PullSessionIdentity,
  type PullSessionResponse,
} from '@silurus/ooxml-core/worker';
import type { Slide } from './types.js';

export const PPTX_INITIAL_SLIDE_PULL_BYTES = 1024 * 1024;

export interface PptxSlidePullClientOptions {
  readonly slideCount: number;
  readonly generation?: number;
  readonly transport: WorkerBridgeTransport<PullSessionResponse<ArrayBuffer, number>>;
  readonly open: (
    slideIndex: number,
    identity: PullSessionIdentity<number>,
    timeoutMs?: number,
  ) => Promise<void>;
  readonly onUsage?: (usage: OoxmlResourceUsageSnapshot) => void;
}

/** Main-realm owner of PPTX's one-slide transferred pull sessions. */
export class PptxSlidePullClient {
  private readonly active = new Set<BoundedPullSession<ArrayBuffer, number>>();
  private nextSessionId = 1;

  constructor(private readonly options: PptxSlidePullClientOptions) {
    if (!Number.isSafeInteger(options.slideCount) || options.slideCount < 0) {
      throw new TypeError('slideCount must be a non-negative safe integer');
    }
    if (
      options.generation !== undefined &&
      (!Number.isSafeInteger(options.generation) || options.generation <= 0)
    ) {
      throw new TypeError('generation must be a positive safe integer');
    }
  }

  /** Pull, optionally decode, and ACK exactly one complete slide. */
  async load(
    slideIndex: number,
    decode = true,
    timeoutMs?: number,
  ): Promise<Slide | undefined> {
    this.assertSlideIndex(slideIndex);
    const sessionId = this.nextSessionId++;
    const identity = {
      sessionId,
      operationId: sessionId,
      generation: this.options.generation ?? 1,
    };
    const session = new BoundedPullSession(this.options.transport, {
      ...identity,
      maxByteCredit: HARD_MAX_PPTX_SLIDE_JSON_BYTES,
      timeoutMs,
    });
    this.active.add(session);
    try {
      await this.options.open(slideIndex, identity, timeoutMs);
      const chunk = await pullWithCreditRetry(
        session,
        PPTX_INITIAL_SLIDE_PULL_BYTES,
        HARD_MAX_PPTX_SLIDE_JSON_BYTES,
      );
      try {
        const usage = chunk.usage ?? session.usageCheckpoint;
        if (usage) this.options.onUsage?.(usage);
        // JSON.parse is consumer acceptance for rendering pulls. Preflight pulls
        // deliberately skip it because the worker-side builder already parsed and
        // prepared the same transferred unit before this ACK.
        const slide = decode
          ? JSON.parse(new TextDecoder().decode(new Uint8Array(chunk.payload))) as Slide
          : undefined;
        await chunk.ack();
        return slide;
      } finally {
        // ACK releases producer staging; this separately drops the main-realm
        // transfer reference immediately instead of relying on a later GC.
        chunk.disposeTransferred();
      }
    } catch (error) {
      await session.cancel('request-error').catch(() => undefined);
      throw error;
    } finally {
      this.active.delete(session);
    }
  }

  /** Best-effort convergence used before terminating the owning worker. */
  cancelAll(): void {
    for (const session of this.active) {
      void session.cancel('closed').catch(() => undefined);
    }
    this.active.clear();
  }

  private assertSlideIndex(slideIndex: number): void {
    if (
      !Number.isSafeInteger(slideIndex) ||
      slideIndex < 0 ||
      slideIndex >= this.options.slideCount
    ) {
      throw new RangeError(
        `Slide index ${slideIndex} out of range (count: ${this.options.slideCount})`,
      );
    }
  }
}

export function isPptxSlidePullResponse(
  value: unknown,
): value is PullSessionResponse<ArrayBuffer, number> {
  return !!value && typeof value === 'object' &&
    (value as { protocol?: unknown }).protocol === PULL_SESSION_PROTOCOL;
}
