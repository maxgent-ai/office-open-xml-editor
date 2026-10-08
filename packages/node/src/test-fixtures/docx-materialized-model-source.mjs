// Test-only DOCX model source module that serves a self-authored shared model
// through the materialized document cursor, without the parser WASM. The Node
// session still acquires, seals and paginates it with production layout.
//
// The model is noteContinuationBoundaryModel (exact 10pt lines, one note); it
// has no images, so extract_image is never a valid request.
import { MaterializedDocumentCursorArchive } from '../../../docx/src/document-pull-worker.ts';
import { noteContinuationBoundaryModel } from '../../../docx/src/testing/note-continuation-model.ts';

export async function openModelSource() {
  const cursor = new MaterializedDocumentCursorArchive(noteContinuationBoundaryModel());
  return {
    archive: {
      open_document_cursor: (...args) => cursor.open_document_cursor(...args),
      pull_document_chunk: (...args) => cursor.pull_document_chunk(...args),
      document_chunk_done: () => cursor.document_chunk_done(),
      acknowledge_document_chunk: (...args) => cursor.acknowledge_document_chunk(...args),
      cancel_document_cursor: () => cursor.cancel_document_cursor(),
      close_document_session: () => cursor.close_document_session(),
      assert_healthy: () => undefined,
      extract_image: () => { throw new Error('materialized test model has no images'); },
    },
    viewDefaults: {},
    close() {},
  };
}
