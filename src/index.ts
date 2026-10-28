export * as pptx from './pptx';
export * as xlsx from './xlsx';
export * as docx from './docx';
// Font sources for the shared `fonts` load option. The format entries export
// the same objects; they are flat here so `import { googleFonts } from
// '@silurus/ooxml'` works as the `useGoogleFonts` deprecation message says.
export {
  cssApi,
  fontFaceCss,
  fontFiles,
  fontResolver,
  googleFonts,
  officeFontSubstitutions,
  type CssApiOptions,
  type CssApiSource,
  type FontDemand,
  type FontFaceCssOptions,
  type FontFaceCssSource,
  type FontFaceDeclaration,
  type FontFilesOptions,
  type FontFilesSource,
  type FontResolveContext,
  type FontResolveFn,
  type FontResolverOptions,
  type FontResolverSource,
  type FontSource,
  type FontSourceCredentials,
  type FontSubstituteScript,
  type FontSubstitution,
  type GoogleFontsOptions,
} from './docx';
