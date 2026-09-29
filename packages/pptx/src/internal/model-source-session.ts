/** Optional Node model-source acquisition; imported only for configured sources. */
export {
  validatePptxModelSourceArchive,
  validatePptxModelSourceViewDefaults,
  type PptxModelSourceArchive,
} from './worker-presentation-source.js';
export {
  acquirePptxSessionFromArchive,
  type PptxNodeSessionArchive,
  type PptxOwnedArchiveSource,
} from './node-session-acquisition.js';
