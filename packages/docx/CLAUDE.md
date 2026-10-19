# @silurus/ooxml-docx — DOCX セッション用

ルートの `AGENTS.md`（正）と `CLAUDE.md` を先に読むこと。ワークフロー、
ブランチ、VRT、参照画像のルールはそちらに従う。このファイルは DOCX 固有の
事実だけを記録する。

## 編集範囲
- 主に `packages/docx/**`（`parser/`、`model/` の Rust を含む）。
- `packages/core` / `packages/ooxml-common` や他フォーマットへの変更は、
  `AGENTS.md` の Cross-Package Integration に沿った一貫した修正の場合のみ。

## ディレクトリ構成（主要なもの）

```
packages/docx/
├── parser/src/          Rust パーサー（lib.rs, parser.rs, styles.rs, numbering.rs, …）
├── model/src/           Rust モデル crate
├── src/
│   ├── index.ts         公開 API
│   ├── document.ts      DocxDocument
│   ├── viewer.ts        DocxViewer
│   ├── types.ts         モデル型（Rust 出力と 1:1）
│   ├── layout/          ページネーション・レイアウト
│   ├── line-breaker/    行分割
│   ├── paint/           Canvas 描画
│   ├── render-worker.ts レンダーワーカー
│   └── wasm/            wasm-pack 出力（生成物・git 管理対象外）
├── public/demo/         追跡対象の demo サンプル（sample-1.docx）
├── public/private/docx/ ローカル専用の private コーパス（git 管理対象外）
└── tests/visual/        Playwright spec（visual.spec.ts, conformance.spec.ts, …）
    └── references/demo/<name>/page-N.png + scores.json
                         採用済みレンダラー出力（Word export ではない）。自動更新禁止。
```

## WASM ビルド

```bash
pnpm --filter @silurus/ooxml-docx wasm
```

`parser/` で `wasm-pack build --target web --out-dir ../src/wasm` を実行し、
続けて `scripts/append-wasm-reinit.mjs` を適用する。

## テスト
- ユニットテスト: ルートで `pnpm test`（vitest）。
- VRT: `pnpm --filter @silurus/ooxml-docx vrt` / `vrt:private` / `vrt:fidelity` /
  `conformance`。ベースライン用の環境変数はルート `AGENTS.md` を参照。
- Storybook はルートで `pnpm storybook`。

## ヒューリスティック禁止（ルート `AGENTS.md` の OOXML Implementation Policy と整合）

DOCX 描画では特に anchor / wgp / docGrid / ruby 周りでヒューリスティックを入れたくなる場面が多い。以下を遵守すること:

- **anchor 位置**: `positionH/V` の `relativeFrom` と `posOffset` / `align` / `pctPos` の意味は ECMA-376 §20.4.3.x で完全に定義されている。「sample-N に合わせるための定数 / 特殊分岐」は禁止。
- **wgp グループ変換**: `lIns/tIns/rIns/bIns` は page-space EMU（グループ変換後の寸法で測る）。`anchorXPt/anchorYPt` は子座標系の offset に sx/sy を掛けて page-relative にする。`widthPt/heightPt` も cx/cy × sx/sy。これら全てが一貫したスケーリングを受けていないと位置ズレが積み重なる。
- **フォント分類**: 名前パターンマッチではなく `word/fontTable.xml` の `<w:family w:val=…/>` (§17.8.3.10) を使う。パターンマッチはフォントテーブルにないケースへのフォールバックとしてのみ残す。
- **改ページヒント**: `<w:lastRenderedPageBreak/>` (§17.3.1.20) は Word のレイアウトキャッシュ。「自前計算が壊れている特定パスでだけ信用する」のは禁止。自前計算を直すか、常に無視するかのいずれかを選ぶ。混在させない。
- **仕様デフォルト値**: 属性省略時の `unwrap_or(0.0)` は多くの場合誤り。§21.1.2.1.1 の lIns=91440 EMU のように ECMA-376 にデフォルト値が定義されているものは必ずその値を使う。
