# @silurus/ooxml-legacy-office

Experimental DOC, XLS and PPT model sources for `@silurus/ooxml`. Both packages
must have the same version.

```sh
npm i @silurus/ooxml @silurus/ooxml-legacy-office
```

```ts
import { DocxViewer } from '@silurus/ooxml/docx';
import { legacyDoc } from '@silurus/ooxml-legacy-office/doc';

const canvas = document.getElementById('viewer') as HTMLCanvasElement;
const viewer = new DocxViewer(canvas, { legacyDoc });
```

Use `/xls` (`legacyXls`) and `/ppt` (`legacyPpt`) with the corresponding
viewers or Node sessions, passed as `{ legacyXls }` and `{ legacyPpt }`.
Importing a reader loads nothing: its source module and WASM load only when a
matching file is opened. When an asset pipeline relocates the bundled files,
create the reader with `createLegacyDoc({ wasmUrl, moduleUrl })` (likewise
`createLegacyXls` and `createLegacyPpt`) and pass it as the same option.
The package ships three external WASM assets and self-contained source modules.
See the [usage and support guide](https://github.com/yukiyokotani/office-open-xml-viewer/blob/main/docs/legacy-office-conversion.md).
