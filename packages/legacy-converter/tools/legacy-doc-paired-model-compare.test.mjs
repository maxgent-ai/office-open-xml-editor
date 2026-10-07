import { test } from 'node:test';
import assert from 'node:assert/strict';
import { projectModelMeaning, compareModelMeaning } from './legacy-doc-paired-model-compare.mjs';

const paragraph = (runs, properties = {}) => ({
  type: 'paragraph',
  runs,
  ...properties
});

const text = (value, properties = {}) => ({
  type: 'text',
  text: value,
  bold: false,
  ...properties
});

const document = (body, properties = {}) => ({
  body,
  section: {},
  ...properties
});

const category = (a, b, name) => compareModelMeaning(a, b).categories[name];

test('compares text and styled offsets independently of incidental run splitting', () => {
  const a = document([paragraph([text('ab')])]), b = document([paragraph([text('a'), text('b')])]);
  assert.equal(category(a, b, 'content').equal, true);
  assert.equal(category(a, b, 'runFormats').equal, true);

  const changed = document([paragraph([text('a'), text('b', {
    bold: true
  })])]);

  assert.equal(category(a, changed, 'content').equal, true);
  assert.equal(category(a, changed, 'runFormats').equal, false);
});

test('preserves paragraph style equality relationships while renaming encoding IDs', () => {
  const a = document([paragraph([text('a')], {
    styleId: 'Normal',
    contextualSpacing: true
  }), paragraph([text('b')], {
    styleId: 'Normal'
  })]);

  const b = document([paragraph([text('a')], {
    styleId: '0',
    contextualSpacing: true
  }), paragraph([text('b')], {
    styleId: '0'
  })]);

  assert.equal(category(a, b, 'paragraphFormats').equal, true);
  b.body[1].styleId = '1';
  assert.equal(category(a, b, 'paragraphFormats').equal, false);
});

test(
  'resolves image references by extracted bytes and retains geometry and missing resources',
  () => {
    const a = document([paragraph([{
      type: 'image',
      imagePath: 'word/media/a.png',
      widthPt: 12,
      heightPt: 6
    }])]);

    const b = document([paragraph([{
      type: 'image',
      imagePath: 'native:7',
      widthPt: 12,
      heightPt: 6
    }])]);

    const resourcesA = [{
            path: 'word/media/a.png',
            sha256: 'a'.repeat(64),
            bytes: 11
          }],
          resourcesB = [{
            path: 'native:7',
            sha256: 'a'.repeat(64),
            bytes: 11
          }];

    assert.equal(compareModelMeaning(a, b, {
      resourcesA,
      resourcesB
    }).categories.images.equal, true);

    resourcesB[0].sha256 = 'b'.repeat(64);

    assert.equal(compareModelMeaning(a, b, {
      resourcesA,
      resourcesB
    }).categories.images.equal, false);

    assert.equal(compareModelMeaning(a, b, {
      resourcesA,
      resourcesB: []
    }).categories.images.equal, null);
  }
);

// A picture bullet is an image reference held by paragraph numbering, not by
// an image run. Two numbered paragraphs (one marker each) and one plain one.
const pictureBullets = path => document([
  paragraph([text('one')], {
    numbering: {
      format: 'bullet',
      text: '',
      picBulletImagePath: path
    }
  }),
  paragraph([text('two')], {
    numbering: {
      format: 'bullet',
      text: '',
      picBulletImagePath: path
    }
  }),
  paragraph([text('plain')])
]);

const bulletResource = (path, sha256) => [{
  path,
  sha256,
  bytes: 11
}];

test('compares picture-bullet markers by extracted bytes, not resource paths', () => {
  const a = pictureBullets('word/media/bullet.png'), b = pictureBullets('native:3');

  const compared = compareModelMeaning(a, b, {
    resourcesA: bulletResource('word/media/bullet.png', 'a'.repeat(64)),
    resourcesB: bulletResource('native:3', 'a'.repeat(64))
  });

  // Image-only numbering exercises images: one marker fact per numbered paragraph.
  assert.deepEqual([compared.leftCounts.images, compared.rightCounts.images], [2, 2]);
  assert.equal(compared.categories.images.status, 'AGREEMENT_FOR_COMPARED_FACTS');
  assert.equal(compared.categories.images.equal, true);
  assert.equal(compared.categories.paragraphFormats.equal, true);

  // The renamed encoding path remains in the raw ledger.
  assert.ok(compared.rawDifferences.some(d => d.path.endsWith('/numbering/picBulletImagePath')));
});

test('reports a picture-bullet byte difference in images and paragraph formats', () => {
  const compared = compareModelMeaning(pictureBullets('word/media/bullet.png'), pictureBullets('native:3'), {
    resourcesA: bulletResource('word/media/bullet.png', 'a'.repeat(64)),
    resourcesB: bulletResource('native:3', 'b'.repeat(64))
  });

  assert.equal(compared.categories.images.equal, false);
  assert.equal(compared.categories.paragraphFormats.equal, false);
});

test('does not claim picture-bullet equivalence from an unextracted resource', () => {
  // Matching opaque paths are not evidence of equal bytes.
  const marker = pictureBullets('native:3');
  const compared = compareModelMeaning(marker, marker);

  assert.deepEqual(compared.unavailableResources.left, ['native:3']);

  for (const name of ['images', 'paragraphFormats']) {
    assert.equal(compared.categories[name].equal, null, name);
    assert.equal(compared.categories[name].status, 'INCOMPLETE_RESOURCE_EVIDENCE', name);
  }
});

test(
  'renames note IDs by referenced targets while preserving duplicate and distinct ownership',
  () => {
    const note = id => text('', {
      noteRef: {
        kind: 'footnote',
        id
      }
    });

    const a = document([paragraph([note('4'), note('4')])], {
      footnotes: [{
        id: '4',
        content: [paragraph([text('note')])]
      }]
    });

    const b = document([paragraph([note('9'), note('9')])], {
      footnotes: [{
        id: '9',
        content: [paragraph([text('note')])]
      }]
    });

    assert.equal(category(a, b, 'notes').equal, true);
    assert.equal(category(a, b, 'content').equal, true);
    b.body[0].runs[1].noteRef.id = '10';

    b.footnotes.push({
      id: '10',
      content: [paragraph([text('note')])]
    });

    assert.equal(category(a, b, 'content').equal, false);
    assert.equal(category(a, b, 'notes').equal, false);
  }
);

test('cell margins honor inheritance, row exceptions and authored zero', () => {
  const table = margin => ({
    type: 'table',
    cellMarginTop: 6,

    rows: [{
      cells: [{
        content: [],
        marginTop: margin
      }],

      __tableRowLayout: {
        exception: {
          cellMargins: {
            top: {
              kind: 'dxa',
              value: '160'
            }
          }
        }
      }
    }]
  });

  const a = document([table(null)]), b = document([table(8)]);
  assert.equal(category(a, b, 'tables').equal, true);
  b.body[0].rows[0].cells[0].marginTop = 0;
  assert.equal(category(a, b, 'tables').equal, false);
});

test('normative snap-to-grid default is distinct from false and absence is recorded', () => {
  const a = document([paragraph([text('a')])]),
        b = document([paragraph([text('a')], {
          snapToGrid: true
        })]);

  const compared = compareModelMeaning(a, b);
  assert.equal(compared.categories.paragraphFormats.equal, true);
  assert.ok(compared.representationDifferences.length);
  b.body[0].snapToGrid = false;
  assert.equal(category(a, b, 'paragraphFormats').equal, false);
});

test(
  'empty feature domains are reported unexercised; unknown facts and ordered breaks are retained',
  () => {
    const a = document([paragraph([text('a')]), {
      type: 'pageBreak'
    }, paragraph([], {
      spaceAfter: 0
    })]);

    const b = document([paragraph([text('a')]), paragraph([], {
      spaceAfter: 0
    }), {
      type: 'pageBreak'
    }]);

    assert.equal(category(a, b, 'content').equal, false);
    assert.equal(category(a, a, 'notes').status, 'NOT_EXERCISED');

    a.body[0].__futureFact = {
      value: 1
    };

    b.body[0].__futureFact = {
      value: 2
    };

    assert.equal(category(a, b, 'acquisitionMetadata').equal, false);

    assert.throws(() => projectModelMeaning({
      body: [],
      parseError: 'partial'
    }), /partial/);
  }
);

test('does not claim equivalence from missing image bytes or duplicate note targets', () => {
  const image = document([paragraph([{
    type: 'image',
    imagePath: 'missing'
  }])]);

  assert.equal(category(image, image, 'images').status, 'INCOMPLETE_RESOURCE_EVIDENCE');

  const missing = document([paragraph([text('', {
    noteRef: {
      kind: 'footnote',
      id: '4'
    }
  })])]);

  assert.equal(category(missing, missing, 'notes').status, 'INCOMPLETE_NOTE_EVIDENCE');
  assert.equal(compareModelMeaning(missing, missing).leftCounts.notes, 1);

  assert.throws(() => projectModelMeaning(document([], {
    footnotes: [{
      id: '4',
      content: []
    }, {
      id: '4',
      content: []
    }]
  })), /duplicate/);
});

test('retains unknown note and story facts and rejects unbounded recursive API inputs', () => {
  const a = document([], {
          headers: {
            default: {
              body: [],
              future: 1
            }
          },

          footnotes: [{
            id: '4',
            content: [],
            future: 1
          }]
        }),
        b = structuredClone(a);

  b.headers.default.future = 2;
  b.footnotes[0].future = 2;
  assert.equal(category(a, b, 'content').equal, false);
  assert.equal(category(a, b, 'notes').equal, false);
  const cyclic = document([]);
  cyclic.future = cyclic;
  assert.throws(() => projectModelMeaning(cyclic), /cyclic/);
  const deep = document([]);
  let value = deep;

  for (let i = 0; i < 130; i++) {
    value.future = {};
    value = value.future;
  }

  assert.throws(() => projectModelMeaning(deep), /budget/);
});

test('preserves custom note numbering and unknown reference facts when remapping IDs', () => {
  const a = document([paragraph([text('', {
    noteRef: {
      kind: 'footnote',
      id: '4',
      customMarkFollows: false,
      future: 1
    }
  })])], {
    footnotes: [{
      id: '4',
      content: []
    }]
  });

  const b = structuredClone(a);
  b.body[0].runs[0].noteRef.customMarkFollows = true;
  assert.equal(category(a, b, 'content').equal, false);
  assert.equal(category(a, b, 'runFormats').equal, false);
  b.body[0].runs[0].noteRef.customMarkFollows = false;
  b.body[0].runs[0].noteRef.future = 2;
  assert.equal(category(a, b, 'content').equal, false);
});

test('unknown keys cannot overwrite canonical identity or hide prototype-named facts', () => {
  const a = document([paragraph([], {
    styleId: 'Normal',
    styleIdClass: 0
  }), paragraph([], {
    styleId: 'Normal',
    styleIdClass: 0
  })]);

  const b = structuredClone(a);
  b.body[0].styleId = '0';
  b.body[1].styleId = '1';
  assert.equal(category(a, b, 'paragraphFormats').equal, false);

  const image = path => document([paragraph([{
    type: 'image',
    imagePath: path,

    imageResource: {
      future: 'same'
    }
  }])]);

  assert.equal(compareModelMeaning(image('a'), image('b'), {
    resourcesA: [{
      path: 'a',
      sha256: 'a'.repeat(64),
      bytes: 1
    }],

    resourcesB: [{
      path: 'b',
      sha256: 'b'.repeat(64),
      bytes: 1
    }]
  }).categories.images.equal, false);

  const unknown = value => document([paragraph([], {
    __futureFact: JSON.parse('{"__proto__":{"value":' + value + '}}')
  })]);

  assert.equal(category(unknown(1), unknown(2), 'acquisitionMetadata').equal, false);
});

test(
  'CLI binds parsed envelope bytes and refuses oversized or non-regular JSON inputs',
  async () => {
    const {
      mkdtempSync,
      writeFileSync,
      openSync,
      ftruncateSync,
      closeSync,
      rmSync,
      readFileSync
    } = await import('node:fs');

    const {
      tmpdir
    } = await import('node:os');

    const {
      join
    } = await import('node:path');

    const {
      spawnSync
    } = await import('node:child_process');

    const {
      createHash
    } = await import('node:crypto');

    const root = mkdtempSync(join(tmpdir(), 'legacy-doc-pair-cli-'));
    const tool = new URL('./legacy-doc-paired-model-compare.mjs', import.meta.url);

    const invoke = (left, right) => spawnSync(process.execPath, [tool.pathname, left, right, join(root, 'report.json')], {
      encoding: 'utf8',
      timeout: 10000
    });

    try {
      const path = join(root, 'input.json');

      writeFileSync(path, JSON.stringify({
        document: document([]),
        resources: []
      }));

      const good = invoke(path, path);
      assert.equal(good.status, 0, good.stderr);

      assert.equal(
        JSON.parse(good.stdout).leftSha256,
        createHash('sha256').update(readFileSync(path)).digest('hex')
      );

      const left = join(root, 'raw-left.json'), right = join(root, 'raw-right.json');

      writeFileSync(left, JSON.stringify(document([paragraph([text('a')])], {
        document: document([]),
        resources: 'future'
      })));

      writeFileSync(right, JSON.stringify(document([paragraph([text('b')])], {
        document: document([]),
        resources: 'future'
      })));

      const distinct = invoke(left, right);
      assert.equal(distinct.status, 0, distinct.stderr);

      assert.equal(
        JSON.parse(readFileSync(join(root, 'report.json'), 'utf8')).categories.content.equal,
        false
      );

      const huge = join(root, 'huge.json'), fd = openSync(huge, 'wx');

      try {
        ftruncateSync(fd, 64 * 1024 * 1024 + 1);
      } finally {
        closeSync(fd);
      }

      const oversized = invoke(huge, path);
      assert.notEqual(oversized.status, 0);
      assert.match(oversized.stderr, /byte budget/);
      if (process.platform !== 'win32') {
        const fifo = join(root, 'fifo');
        const created = spawnSync('mkfifo', [fifo], { encoding: 'utf8', timeout: 10000 });
        assert.equal(created.status, 0, created.stderr);
        const pipe = invoke(fifo, path);
        assert.notEqual(pipe.status, 0);
        assert.match(pipe.stderr, /byte budget/);
      }
      const special = invoke(process.platform === 'win32' ? 'NUL' : '/dev/null', path);
      assert.notEqual(special.status, 0);
      assert.match(special.stderr, /byte budget/);
    } finally {
      rmSync(root, {
        recursive: true,
        force: true
      });
    }
  }
);
