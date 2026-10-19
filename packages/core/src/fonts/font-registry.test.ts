import { vi, beforeEach, describe, it, expect } from 'vitest';
let { retainFace, releaseFaces } = await import('./font-registry.js');

beforeEach(async () => {
  vi.resetModules();
  ({ retainFace, releaseFaces } = await import('./font-registry.js'));
});

// A minimal FontFace / FontFaceSet stand-in. The registry only needs identity
// (===) on faces and a `delete()` on the set, so these are intentionally tiny.
interface FakeFace {
  family: string;
}
function makeSet() {
  const faces: FakeFace[] = [];
  const set = {
    add: (f: FakeFace) => {
      faces.push(f);
    },
    delete: (f: FakeFace) => {
      const i = faces.indexOf(f);
      if (i >= 0) faces.splice(i, 1);
      return i >= 0;
    },
  } as unknown as FontFaceSet;
  return { set, faces };
}
/** A create() that builds a fresh fake face, adds it to the set, and records how
 *  many times it ran (retain must call it ONLY on the first-holder path). */
function creator(set: FontFaceSet, family: string) {
  let calls = 0;
  const create = (): FontFace => {
    calls++;
    const f = { family } as unknown as FontFace;
    (set as unknown as { add: (f: FontFace) => void }).add(f);
    return f;
  };
  return { create, calls: () => calls };
}

describe('retainFace', () => {
  it('first holder creates + adds the face and reports isNew', () => {
    const { set, faces } = makeSet();
    const c = creator(set, 'A');
    const { face, isNew } = retainFace('sig-a', set, c.create);
    expect(isNew).toBe(true);
    expect(c.calls()).toBe(1);
    expect(faces).toContain(face as unknown as FakeFace);
  });

  it('a later holder of the same signature reuses the face and does NOT re-create', () => {
    const { set, faces } = makeSet();
    const c = creator(set, 'A');
    const first = retainFace('sig-a', set, c.create);
    const second = retainFace('sig-a', set, c.create);
    expect(second.isNew).toBe(false);
    expect(second.face).toBe(first.face); // same shared FontFace
    expect(c.calls()).toBe(1); // create ran once
    expect(faces).toHaveLength(1); // added once
  });

  it('distinct signatures create distinct faces', () => {
    const { set, faces } = makeSet();
    const a = creator(set, 'A');
    const b = creator(set, 'B');
    const ra = retainFace('sig-a', set, a.create);
    const rb = retainFace('sig-b', set, b.create);
    expect(ra.face).not.toBe(rb.face);
    expect(faces).toHaveLength(2);
  });

  it('the same signature in a DIFFERENT set is treated as absent (creates anew)', () => {
    const { set: setA } = makeSet();
    const { set: setB, faces: facesB } = makeSet();
    const ca = creator(setA, 'A');
    const cb = creator(setB, 'A');
    const ra = retainFace('sig-a', setA, ca.create);
    // Same signature, different set → must NOT hand back setA's face.
    const rb = retainFace('sig-a', setB, cb.create);
    expect(rb.isNew).toBe(true);
    expect(rb.face).not.toBe(ra.face);
    expect(facesB).toHaveLength(1);
  });

  it('keeps same-signature registrations independent across document sets', () => {
    const { set: setA, faces: facesA } = makeSet();
    const { set: setB, faces: facesB } = makeSet();
    const faceA = retainFace('sig-a', setA, creator(setA, 'A').create).face;
    const faceB = retainFace('sig-a', setB, creator(setB, 'A').create).face;

    releaseFaces([faceB]);
    expect(facesB).toHaveLength(0);
    expect(facesA).toEqual([faceA as unknown as FakeFace]);

    releaseFaces([faceA]);
    expect(facesA).toHaveLength(0);
  });
});

describe('releaseFaces', () => {
  it('removes the face from its set when the last (only) holder releases', () => {
    const { set, faces } = makeSet();
    const c = creator(set, 'A');
    const { face } = retainFace('sig-a', set, c.create);
    expect(faces).toHaveLength(1);
    releaseFaces([face]);
    expect(faces).toHaveLength(0);
  });

  it('keeps the face until the LAST holder releases (refcount)', () => {
    const { set, faces } = makeSet();
    const c = creator(set, 'A');
    const a = retainFace('sig-a', set, c.create).face; // holder 1
    const b = retainFace('sig-a', set, c.create).face; // holder 2 (same face)
    expect(a).toBe(b);
    expect(faces).toHaveLength(1);

    releaseFaces([a]); // holder 1 gone
    expect(faces).toHaveLength(1); // holder 2 keeps it

    releaseFaces([b]); // holder 2 gone
    expect(faces).toHaveLength(0);
  });

  it('the same face passed twice in ONE call is decremented at most once', () => {
    const { set, faces } = makeSet();
    const c = creator(set, 'A');
    const a = retainFace('sig-a', set, c.create).face; // refs 1
    retainFace('sig-a', set, c.create); // refs 2 (second holder)
    expect(faces).toHaveLength(1);

    // Duplicate in one release must collapse to a single decrement (2 → 1),
    // NOT evict the face the second holder still uses.
    releaseFaces([a, a]);
    expect(faces).toHaveLength(1);
  });

  it('re-releasing a fully-released face is a no-op (no negative refs, no cross-hit)', () => {
    const { set, faces } = makeSet();
    const c = creator(set, 'A');
    const a = retainFace('sig-a', set, c.create).face;
    releaseFaces([a]); // fully released, entry removed
    expect(faces).toHaveLength(0);

    // A fresh registration of the same signature (distinct FontFace object).
    const b = retainFace('sig-a', set, c.create).face;
    expect(b).not.toBe(a);
    expect(faces).toHaveLength(1);

    // A stray re-release of the OLD face must not disturb b's registration.
    releaseFaces([a]);
    expect(faces).toHaveLength(1);
  });

  it('is a no-op for a face the registry never saw', () => {
    const { set, faces } = makeSet();
    creator(set, 'A'); // nothing retained
    const stray = { family: 'Stray' } as unknown as FontFace;
    expect(() => releaseFaces([stray])).not.toThrow();
    expect(faces).toHaveLength(0);
  });

  it('tolerates a set whose delete() throws (older shim) and still drops the entry', () => {
    const faces: FakeFace[] = [];
    const set = {
      add: (f: FakeFace) => {
        faces.push(f);
      },
      delete: () => {
        throw new Error('no delete on this shim');
      },
    } as unknown as FontFaceSet;
    const c = creator(set, 'A');
    const { face } = retainFace('sig-a', set, c.create);
    // delete() throws, but the release must not throw and must drop the entry so
    // a later retain of the same signature creates a fresh registration.
    expect(() => releaseFaces([face])).not.toThrow();
    const again = retainFace('sig-a', set, c.create);
    expect(again.isNew).toBe(true); // entry was dropped despite the throwing delete
  });
});
