// Types only: the differential config resolves these literal imports to the
// selected clean main checkout. No candidate implementation is a fallback.
declare module '@docx-tab-baseline/document-layout' {
  export const layoutDocument: typeof import('../../src/document-layout.js').layoutDocument;
}
declare module '@docx-tab-baseline/layout-runtime' {
  export const createLayoutServices: typeof import('../../src/layout-runtime.js').createLayoutServices;
}
