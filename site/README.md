# @silurus/ooxml — showcase site

A standalone marketing / showcase site for [`@silurus/ooxml`](https://www.npmjs.com/package/@silurus/ooxml).
It embeds the **real** viewers (Rust/WASM + Canvas) rendering live sample
documents, alongside framework integration snippets.

This is separate from Storybook (`pnpm storybook`), which stays the test +
component workbench.

## Develop

```bash
pnpm build:wasm                      # build the parsers first (once)
pnpm --filter @silurus/ooxml-site dev
```

`predev` / `prebuild` copy the redistributable demo samples from
`packages/*/public/demo` into `public/samples/` (office files are not committed
under `site/`).

Dev-only `/preview/*` routes render individual sections in isolation (useful for
visual checks); they are not linked from the site.

## Build

```bash
pnpm --filter @silurus/ooxml-site build      # → site/dist/
SITE_BASE=/office-open-xml-viewer/ pnpm --filter @silurus/ooxml-site build  # project-pages base
```

## Deploy (GitHub Pages)

GitHub Pages serves one site per repo, so `.github/workflows/deploy-pages.yml`
builds **both** projects and merges them into a single artifact:

```
site/dist/            → https://ooxml.silurus.dev/
site/dist/storybook/  → https://ooxml.silurus.dev/storybook/
```

Steps: build WASM → build the site (base `/`) → build Storybook with
`STORYBOOK_BASE=/storybook/` → copy `storybook-static/` into
`site/dist/storybook/` → write `CNAME` → deploy. Storybook's Vite `base` reads
`STORYBOOK_BASE` (default `/`, so local `pnpm storybook` and standalone builds
are unaffected).

Triggers on `v*` tags and `workflow_dispatch`. This replaces the old
`deploy-site.yml` + `deploy-storybook.yml`. Running it makes the showcase site
the front door at the custom domain, with Storybook nested under `/storybook/`.
