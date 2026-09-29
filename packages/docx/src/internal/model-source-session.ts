/** Optional Node model-source acquisition; imported only for configured sources. */
export {
  validateDocxModelSourceArchive,
  validateDocxModelSourceViewDefaults,
  type DocxModelSourceArchive,
  type DocxModelSourceViewDefaults,
} from './worker-document-source.js';
export {
  acquireDocxSessionFromArchive,
  type DocxOwnedArchiveSource,
} from './node-model-source-acquisition.js';
