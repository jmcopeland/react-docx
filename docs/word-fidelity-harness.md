# Microsoft Word fidelity oracle

This harness turns a Microsoft-rendered PDF into a frozen, auditable reference
artifact. Validation and rasterization are completely offline: provider access is
only needed when a reference PDF is first captured.

The oracle answers a narrow question: “How did this exact DOCX render in this
recorded Microsoft environment?” It deliberately does not claim that Microsoft
Word Desktop, Word for the web, and Microsoft Graph always render identically.
Each provider capture is a distinct reference.

## Files

- `scripts/word-oracle/manifest.schema.json` is the versioned JSON contract.
- `scripts/word-oracle/record-reference.mjs` inventories a local DOCX/PDF pair.
- `scripts/word-oracle/validate-manifest.mjs` validates metadata, hashes, page
  count, and page geometry.
- `scripts/word-oracle/materialize-references.mjs` rasterizes verified PDFs into
  deterministic page names and writes `reference-pages.json`.
- `scripts/word-oracle/comparison-contract.mjs` validates viewer-page,
  threshold, and metric-output contracts.
- `scripts/word-oracle/compare-pages.mjs` is the page-count, hash, pixel, and
  structure comparison gate.
- `scripts/word-oracle/capture-viewer-pages.mjs` captures settled viewer pages
  using the same case and reference identifiers as the Word oracle.
- `scripts/word-oracle/edit-actions.schema.json` and
  `edit-action-contract.mjs` define the allowlisted editor-action contract.
- `scripts/word-oracle/run-edit-roundtrip.mjs` replays edits, exports, reopens,
  and produces a hash-addressed candidate for a new Word reference.
- `scripts/word-oracle/requirements.txt` pins the live pixel-comparison stack.
- `tests/fidelity/word-oracle-contract.test.mjs` covers the dependency-free
  runtime validator and `pdfinfo` parser.
- `tests/fidelity/word-oracle-comparison.test.mjs` covers joining, integrity,
  metric invocation/consumption, threshold failures, and gate ordering.
- `tests/fidelity/edit-action-contract.test.mjs` covers action validation,
  source binding, structural summary comparison, and result generation.

## Deterministic capture contract

A reference is valid only when all of the following are frozen in its manifest:

1. The original DOCX bytes, identified by a lowercase SHA-256 digest.
2. The unmodified Microsoft-produced PDF bytes, also identified by SHA-256.
3. A stable case ID, reference ID, provider ID, and capture-profile ID.
4. The PDF export choices: print/screen optimization, final/markup view, field
   update policy, and PDF/A policy.
5. The renderer name/version and, when available, build and update channel.
6. Locale, timezone, platform, and the font policy. A pinned font set requires a
   digest of its normalized inventory. Cloud renderers may explicitly declare
   `provider-managed` fonts instead of pretending their font environment is
   reproducible.
7. Capture time, PDF page count, size in points, and rotation for every page.
8. The rasterization settings used by local comparisons. The materialized
   report records the actual `pdftoppm` version as well.

The PDF is the durable oracle. PNG pages are derived artifacts and should be
regenerated with a pinned Poppler build in CI. Replacing either source file
requires a new digest; changing the Word build, font set, provider, export
settings, or field-update policy should create a new reference ID or corpus
revision.

The pixel comparator requires Python 3.12 or newer. Install its pinned image
stack before running a live comparison:

```sh
python3 -m pip install -r scripts/word-oracle/requirements.txt
```

Use the same Python patch release, Poppler build, and browser image for baseline
creation and enforcement. The dependency-free contract tests do not require
these Python packages.

For a pinned font set, run `pnpm word-oracle:font-fingerprint`. It asks
Fontconfig for every face, hashes the actual font-file bytes, and normalizes
UTF-8 lines of
`face-identifier<TAB>font-version<TAB>SHA-256-of-font-bytes`. The face
identifier is the PostScript name, or a deterministic family/style/index
fallback when a face has no PostScript name. Lines are sorted by raw byte
order, LF-terminated, and hashed as one byte sequence. The normalized output
does not contain machine paths. Font files and licensing-sensitive metadata do
not belong in the manifest.

Do not include access tokens, tenant IDs, drive item IDs, signed URLs, user
names, or machine paths in the manifest.

## Capture a reference

Place the DOCX and generated PDF under the directory that will contain the
manifest. Then generate a case fragment:

```sh
node scripts/word-oracle/record-reference.mjs \
  --root tests/fidelity-corpus \
  --source tests/fidelity-corpus/sources/sections.docx \
  --pdf tests/fidelity-corpus/oracles/sections-word-win.pdf \
  --case-id sections \
  --reference-id word-win-16.0.19127 \
  --provider-id microsoft-word-desktop \
  --provider-name "Microsoft Word Desktop" \
  --profile-id final-print-v1 \
  --captured-at 2026-07-09T16:00:00Z \
  --renderer-name "Microsoft Word" \
  --renderer-version 16.0 \
  --renderer-build 19127.20154 \
  --renderer-channel current \
  --platform-name "Windows 11" \
  --platform-version 24H2 \
  --platform-architecture x64 \
  --locale en-US \
  --timezone America/New_York \
  --font-policy pinned \
  --font-fingerprint 0123456789abcdef0123456789abcdef0123456789abcdef0123456789abcdef \
  --output /tmp/sections-case.json
```

The generated fragment contains computed source/PDF hashes and page geometry.
Add it to the manifest’s `cases` array. Define the referenced capture profile at
the top of the manifest:

```json
{
  "$schema": "../../scripts/word-oracle/manifest.schema.json",
  "schemaVersion": 1,
  "corpus": {
    "id": "word-fidelity",
    "revision": "2026-07-09"
  },
  "captureProfiles": [
    {
      "id": "final-print-v1",
      "pdfExport": {
        "mode": "print",
        "markup": "final",
        "updateFields": false,
        "pdfa": false
      },
      "rasterization": {
        "dpi": 144,
        "pageBox": "crop",
        "format": "png",
        "background": "#ffffff"
      }
    }
  ],
  "cases": [
    {
      "id": "sections",
      "source": {
        "path": "sources/sections.docx",
        "sha256": "...64 lowercase hexadecimal characters..."
      },
      "references": [
        {
          "id": "word-win-16.0.19127",
          "provider": {
            "id": "microsoft-word-desktop",
            "displayName": "Microsoft Word Desktop"
          },
          "captureProfileId": "final-print-v1",
          "capturedAt": "2026-07-09T16:00:00Z",
          "artifact": {
            "path": "oracles/sections-word-win.pdf",
            "sha256": "...64 lowercase hexadecimal characters...",
            "mediaType": "application/pdf",
            "pageCount": 1,
            "pages": [
              {
                "widthPoints": 612,
                "heightPoints": 792,
                "rotation": 0
              }
            ]
          },
          "environment": {
            "renderer": {
              "name": "Microsoft Word",
              "version": "16.0",
              "build": "19127.20154",
              "channel": "current"
            },
            "platform": {
              "name": "Windows 11",
              "version": "24H2",
              "architecture": "x64"
            },
            "locale": "en-US",
            "timezone": "America/New_York",
            "fontSet": {
              "policy": "pinned",
              "fingerprintSha256": "0123456789abcdef0123456789abcdef0123456789abcdef0123456789abcdef"
            }
          }
        }
      ]
    }
  ]
}
```

The ellipses above are explanatory placeholders, not valid digests. Prefer
`record-reference.mjs` over entering artifact metadata manually.

## Validate and consume references

Metadata-only validation needs only Node:

```sh
node scripts/word-oracle/validate-manifest.mjs \
  --manifest tests/fidelity-corpus/word-oracle.json \
  --metadata-only
```

Full validation also reads both files, verifies their hashes, and uses
`pdfinfo` to confirm every page:

```sh
node scripts/word-oracle/validate-manifest.mjs \
  --manifest tests/fidelity-corpus/word-oracle.json
```

Materialize selected or all PDF pages with Poppler:

```sh
node scripts/word-oracle/materialize-references.mjs \
  --manifest tests/fidelity-corpus/word-oracle.json \
  --out-dir .artifacts/word-reference-pages
```

The result is stable by case/reference ID:

```text
.artifacts/word-reference-pages/
  reference-pages.json
  sections/
    word-win-16.0.19127/
      page-0001.png
```

`reference-pages.json` is the handoff to screenshot comparison: join each
`imagePath` with the viewer-rendered page for the same case and page number,
then feed those pairs to `scripts/measure_png_visual_diff.py`. The immutable PDF
hash and each derived PNG hash stay in the report so a page can always be traced
back to its oracle and rasterizer version.

For a quick contract test independent of the repository’s Vitest include list:

```sh
node --test tests/fidelity/word-oracle-contract.test.mjs
```

## Provider acquisition

The harness intentionally does not fetch from a provider. Capture adapters may
write a local PDF, after which the same recording/validation flow applies.

### Microsoft Word Desktop

Use an interactive, licensed Word installation with a clean capture profile.
Open the frozen DOCX without editing, apply the recorded final/markup and field
update choices, export as PDF, and preserve the PDF bytes. Record the precise
Office build and installed-font inventory.

Do not implement this as unattended server-side Office automation. Besides
being an unreliable fidelity environment, Microsoft does not recommend or
support that deployment model. A pinned, interactive Windows capture worker is
appropriate if desktop references need periodic refreshes.

### Microsoft Graph

An authorized external adapter can download a PDF using the documented Drive
item conversion endpoint (`GET .../content?format=pdf`). Store the response as a
local artifact and record the provider as `microsoft-graph-drive-conversion`.
Use a renderer version such as `v1.0-service-managed`, omit the platform, and
declare the font policy `provider-managed` unless Microsoft exposes stronger
environment guarantees.

Graph documentation:
<https://learn.microsoft.com/en-us/graph/api/driveitem-get-content-format>

Provider credentials belong in the adapter’s secret store. They must never be
written to this repository, the capture manifest, or the generated report.

## Suggested gates

A small generated/open corpus can run on every pull request; a larger licensed
or private corpus can run from protected artifact storage on a schedule. In
both cases, gate in this order:

1. Validate schema, source/PDF hashes, page count, and page geometry.
2. Rasterize references with a pinned Poppler version.
3. Render react-docx pages with a pinned browser, viewport, DPR, and font set.
4. Fail on page-count mismatch before pixel comparison.
5. Compare page raster, ink structure, and eventually text line/bounding-box
   geometry.
6. For editing cases, export the mutated DOCX and repeat the Microsoft capture
   as a distinct round-trip reference.

The repository should retain the manifest and any legally redistributable
sources. Large or private binary corpora can live in immutable artifact storage
as long as the paths are materialized beside the manifest before validation.
All input paths are canonicalized with `realpath`; a manifest-relative symlink
may point elsewhere inside the corpus tree, but a symlink that resolves outside
that tree is rejected before hashing, rasterizing, or browser upload.

### Opt-in private-corpus CI

`.github/workflows/word-differential.yml` runs the complete
validate/materialize/capture/compare chain on trusted `main`, scheduled, and
manual runs. It is skipped unless both repository variables are configured:

- `WORD_ORACLE_PRIVATE_ARTIFACT_NAME` is the immutable GitHub Actions artifact
  containing `word-oracle.json`, `thresholds.json`, `font-fingerprint.txt`,
  `poppler-version.txt` (the first line of `pdftoppm -v`), and every source/PDF
  at its manifest-relative path.
- `WORD_ORACLE_PRIVATE_ARTIFACT_RUN_ID` is the trusted workflow run that
  produced that artifact.

The private job deliberately does not run on pull requests, where untrusted
changes could read or transmit the corpus. The ordinary generated smoke and
contract jobs remain credential-free. The differential job downloads the
pinned artifact, recomputes the installed font fingerprint and compares it to
`font-fingerprint.txt`, runs the local loopback playground, invokes the live
pinned comparator, and retains only the comparison report and playground log.

## Gate viewer pages against Word

`compare-pages.mjs` is the offline comparison gate. It joins materialized Word
pages and viewer pages by the complete `(caseId, referenceId, pageNumber)` key,
not by array position or file name. It also binds every metric result to the
SHA-256 digests from both page manifests.

The viewer capture must write a manifest next to its page PNGs:

```json
{
  "schemaVersion": 1,
  "corpus": {
    "id": "word-fidelity",
    "revision": "2026-07-09"
  },
  "sourceManifestSha256": "...same value as reference-pages.json...",
  "renderer": {
    "name": "react-docx",
    "version": "0.8.0",
    "browser": "Mozilla/5.0 (...) HeadlessChrome/138.0.7204.0 Safari/537.36",
    "browserVersion": "138.0.7204.0",
    "platform": "Linux x86_64",
    "hostPlatform": "linux",
    "hostArchitecture": "x64",
    "viewport": "816x1056",
    "deviceScaleFactor": 1,
    "fontSetFingerprintSha256": "...64 lowercase hexadecimal characters...",
    "locale": "en-US",
    "timezone": "UTC"
  },
  "references": [
    {
      "caseId": "sections",
      "referenceId": "word-win-16.0.19127",
      "pages": [
        {
          "pageNumber": 1,
          "widthPoints": 612,
          "heightPoints": 792,
          "imagePath": "sections/word-win-16.0.19127/page-0001.png",
          "imageSha256": "...64 lowercase hexadecimal characters..."
        }
      ]
    }
  ]
}
```

`imagePath` is a normalized POSIX path relative to the viewer manifest. The
corpus ID, revision, and `sourceManifestSha256` must exactly match the
materialized reference report. Browser user agent, binary version, browser and
host platform, viewport, device scale factor, locale, and timezone are measured
from the launched browser/runtime. The font fingerprint is computed from the
installed files. None of those manifest fields is copied from the requested
CLI strings.

Thresholds are explicit and tracked separately from the corpus:

```json
{
  "schemaVersion": 1,
  "metricScriptSha256": "...SHA-256 of scripts/measure_png_visual_diff.py...",
  "expectedRendererEnvironment": {
    "browser": "Mozilla/5.0 (...) HeadlessChrome/138.0.7204.0 Safari/537.36",
    "browserVersion": "138.0.7204.0",
    "platform": "Linux x86_64",
    "hostPlatform": "linux",
    "hostArchitecture": "x64",
    "viewport": "816x1056",
    "deviceScaleFactor": 1,
    "fontSetFingerprintSha256": "...64 lowercase hexadecimal characters...",
    "locale": "en-US",
    "timezone": "UTC"
  },
  "comparisonWidth": 816,
  "comparisonHeight": 1056,
  "tolerance": 18,
  "inkThreshold": 24,
  "verticalBands": 12,
  "horizontalBands": 8,
  "gridColumns": 6,
  "gridRows": 8,
  "maxPageDimensionRelativeDiff": 0.01,
  "maxPageAspectRatioRelativeDiff": 0.01,
  "maxMeanAbsoluteDiff": 0.04,
  "maxRootMeanSquareDiff": 0.09,
  "maxMismatchRatio": 0.12,
  "maxLayoutStructureDiff": 0.025
}
```

Those numbers are examples, not universal recommendations. Establish them from
a reviewed baseline using a pinned Word capture, rasterizer, browser, and font
set. Pin the comparator bytes with
`shasum -a 256 scripts/measure_png_visual_diff.py`; the gate rejects any other
script and rejects metric output whose resolution, tolerance, or structural
settings differ from this file. A page must satisfy both physical-geometry
limits before it reaches the four raster/structure measurements.
`expectedRendererEnvironment` is also a hard gate: every field must exactly
match the measured viewer manifest before page geometry or pixels are compared.
After intentionally updating the browser image or fonts, capture once, review
the measured renderer block, and update this pin together with the baseline.

Run the complete gate with the existing Python measurement script:

```sh
node scripts/word-oracle/compare-pages.mjs \
  --references .artifacts/word-reference-pages/reference-pages.json \
  --viewer .artifacts/viewer-pages/viewer-pages.json \
  --thresholds tests/fidelity/word-thresholds.json \
  --out .artifacts/word-comparison/report.json
```

The gate always checks the exact page-number set first. A mismatch writes a
report, skips later work, and exits nonzero. With a matching page set it verifies
viewer width, height, and aspect ratio against the Word PDF, verifies every PNG
path and digest, then invokes the hash-pinned metric script live.
Hash failures, missing or duplicate metric results, and hard-threshold failures
also write a report and exit nonzero. The report contains input-manifest hashes,
renderer and rasterizer metadata, thresholds, measurement settings, per-page
digests and metrics, and deterministic failure records.

Contract-only validation does not require Pillow, NumPy, Poppler, a browser, or
binary image fixtures:

```sh
node --test tests/fidelity/word-oracle-comparison.test.mjs
```

## Capture react-docx viewer pages

Start the playground separately, then run the offline viewer capture against
that supplied URL. The command validates the Word oracle manifest, verifies
every selected DOCX digest, and does not contact Microsoft or another oracle
provider:

```sh
pnpm --filter @extend-ai/react-docx-playground dev \
  --host 127.0.0.1 --port 4173

node scripts/word-oracle/capture-viewer-pages.mjs \
  --manifest tests/fidelity-corpus/word-oracle.json \
  --references .artifacts/word-reference-pages/reference-pages.json \
  --base-url http://127.0.0.1:4173 \
  --out-dir .artifacts/viewer-pages \
  --font-fingerprint 0123456789abcdef0123456789abcdef0123456789abcdef0123456789abcdef
```

`--references` is optional. With it, the capture emits exactly the
case/reference identities present in the materialized report and verifies that
the report came from the same source-manifest bytes. Without it, each source is
captured once and associated with all references declared for that case.
`--case` and `--reference` can narrow either form further.

The font fingerprint is an expected pin, not reported metadata. The capture
recomputes the installed inventory through Fontconfig, rejects a mismatch, and
emits only the computed value. Generate or verify it explicitly with:

```sh
pnpm word-oracle:font-fingerprint
pnpm word-oracle:font-fingerprint -- --expected <sha256>
```

The capture also verifies that the launched browser honored the requested
viewport, device scale factor, locale, and timezone. The comparison gate then
compares all measured environment fields to `expectedRendererEnvironment`, so
arbitrary but structurally valid capture settings cannot pass.

The base URL is loopback-only by default, and a cross-origin redirect is always
rejected before the DOCX is attached to the page. `--allow-remote` is an
explicit escape hatch for a trusted non-loopback development deployment; it
does not permit that deployment to redirect to another origin.

Readiness is condition-based rather than a fixed delay. For each local DOCX,
the capture waits until the playground reports the matching file as loaded,
the pagination overlay is gone, `document.fonts` is loaded, every viewer image
has decoded dimensions, animations have stopped, all zero-based page surfaces
are mounted, and page geometry plus rendered DOM content remain unchanged for
eight consecutive animation frames. A per-document timeout is only a failure
deadline. Playwright runs with WebDriver enabled, which makes the viewer mount
the complete page stack instead of virtualizing offscreen pages.

Each `[data-docx-page-surface]` is captured as an individual PNG. Its stable CSS
geometry is converted from 96-DPI CSS pixels to points and recorded beside the
hash so page-size and aspect-ratio errors cannot be hidden by image
normalization. The output is written through a temporary directory and moved
into place only after all
sources succeed:

```text
.artifacts/viewer-pages/
  viewer-pages.json
  sections/
    word-win-16.0.19127/
      page-0001.png
```

Every PNG is SHA-256 hashed and `viewer-pages.json` is validated against the
same contract consumed by `compare-pages.mjs`. Existing output is preserved
unless `--force` is supplied. Contract tests do not launch a browser:

```sh
node --test tests/fidelity/viewer-capture-contract.test.mjs
```

## Deterministic edit, export, and reopen scenarios

`run-edit-roundtrip.mjs` turns model-level editing behavior into a reproducible
artifact pipeline. It uses the development-only `window.__DOCX_TEST_HOOKS__`
exposed by the playground, so the supplied URL must be a running development
build rather than a production deployment.

Edit actions live in a separate, strict, versioned manifest. The root corpus
identity and `sourceManifestSha256` bind the scenarios to the exact Word oracle
manifest bytes. Each scenario references a source case; its own `id` becomes
the proposed case ID for registering the exported DOCX as a new Word oracle:

```json
{
  "$schema": "../../scripts/word-oracle/edit-actions.schema.json",
  "schemaVersion": 1,
  "corpus": {
    "id": "word-fidelity",
    "revision": "2026-07-09"
  },
  "sourceManifestSha256": "...64 lowercase hexadecimal characters...",
  "scenarios": [
    {
      "id": "sections-edited-heading",
      "sourceCaseId": "sections",
      "description": "Restyle and replace the first heading",
      "expected": {
        "semanticModel": "changed-from-source",
        "paragraphTexts": [{ "nodeIndex": 0, "text": "Updated heading" }]
      },
      "actions": [
        {
          "type": "select-paragraph",
          "nodeIndex": 0,
          "expect": {
            "semanticModel": "unchanged",
            "effect": {
              "kind": "selection",
              "selection": { "kind": "paragraph", "nodeIndex": 0 }
            }
          }
        },
        {
          "type": "set-text-range",
          "range": {
            "start": {
              "location": { "kind": "paragraph", "nodeIndex": 0 },
              "offset": 0
            },
            "end": {
              "location": { "kind": "paragraph", "nodeIndex": 0 },
              "offset": 7
            }
          },
          "expect": {
            "semanticModel": "unchanged",
            "effect": {
              "kind": "active-text-range",
              "range": {
                "start": {
                  "location": { "kind": "paragraph", "nodeIndex": 0 },
                  "offset": 0
                },
                "end": {
                  "location": { "kind": "paragraph", "nodeIndex": 0 },
                  "offset": 7
                }
              }
            }
          }
        },
        {
          "type": "toggle-bold",
          "expect": {
            "semanticModel": "changed",
            "effect": {
              "kind": "text-style",
              "range": {
                "start": {
                  "location": { "kind": "paragraph", "nodeIndex": 0 },
                  "offset": 0
                },
                "end": {
                  "location": { "kind": "paragraph", "nodeIndex": 0 },
                  "offset": 7
                }
              },
              "property": "bold",
              "value": true
            }
          }
        },
        {
          "type": "set-font-family",
          "fontFamily": "Aptos Display",
          "expect": {
            "semanticModel": "changed",
            "effect": {
              "kind": "text-style",
              "range": {
                "start": {
                  "location": { "kind": "paragraph", "nodeIndex": 0 },
                  "offset": 0
                },
                "end": {
                  "location": { "kind": "paragraph", "nodeIndex": 0 },
                  "offset": 7
                }
              },
              "property": "fontFamily",
              "value": "Aptos Display"
            }
          }
        },
        {
          "type": "set-font-size",
          "fontSizePt": 18,
          "expect": {
            "semanticModel": "changed",
            "effect": {
              "kind": "text-style",
              "range": {
                "start": {
                  "location": { "kind": "paragraph", "nodeIndex": 0 },
                  "offset": 0
                },
                "end": {
                  "location": { "kind": "paragraph", "nodeIndex": 0 },
                  "offset": 7
                }
              },
              "property": "fontSizePt",
              "value": 18
            }
          }
        },
        {
          "type": "commit-paragraph-text",
          "nodeIndex": 0,
          "text": "Updated heading",
          "expect": {
            "semanticModel": "changed",
            "effect": {
              "kind": "paragraph-text",
              "nodeIndex": 0,
              "text": "Updated heading"
            }
          }
        }
      ]
    }
  ]
}
```

The contract intentionally exposes only an allowlisted hook subset. Supported
action types are:

- Selection and text: `select-paragraph`, `select-table-cell`,
  `set-text-range`, `commit-paragraph-text`, and `commit-table-cell-text`.
- Character formatting: `toggle-bold`, `toggle-italic`, `toggle-underline`,
  `toggle-strike`, `set-text-color`, `set-highlight`, `set-font-family`, and
  `set-font-size`.
- Paragraph formatting: `set-alignment`, `toggle-list`, and
  `set-line-spacing`.
- Table structure: `insert-table-row`, `insert-table-column`,
  `delete-table-row`, and `delete-table-column`.
- Annotations: `accept-tracked-change`, `reject-tracked-change`,
  `create-comment`, and `set-comment-resolved`.
- History: `undo` and `redo`.

Colors are lowercase `#rrggbb` values or `null` to clear. Alignment can also be
`null`. A `set-text-range` action accepts `null` to clear the active range.
Body locations use `{ "kind": "paragraph", "nodeIndex": ... }`; table-cell
locations additionally require table, row, cell, and paragraph indexes. Extra
properties, unknown actions, unsafe indexes, reversed ranges, control
characters in font names, and unbound source cases are rejected before a
browser starts.

Every action has both a semantic transition and a mandatory action-specific
`expect.effect`. Selection and active-range effects repeat the exact requested
selection/range. Text and paragraph formatting effects name the exact target
range/location, property, and resulting value. Text commits repeat their exact
paragraph or table-cell coordinates and result text. Table edits provide the
exact resulting row/column shape. Undo/redo provide the exact restored semantic
digest and history capabilities. Tracked-change actions identify the revision
by change ID, revision ID, kind, location, and text, then assert its removal,
remaining count, and exact result text. Comment actions identify the exact
anchor, stable ID, comment ID, body, metadata, resolution state, and counts.
`create-comment` therefore requires explicit `author`, `initials`, and `date`
options; runner-generated dates are not accepted in deterministic manifests.

The runner reads each target before and after its action. The target itself must
change for mutating range, paragraph, and table operations, while the exact
declared revision must disappear after tracked-change acceptance/rejection.
The declared result value must be present afterward. A mutation elsewhere in
the model cannot satisfy an action merely because the global semantic digest
changed. Selection/range actions must leave the semantic model unchanged; all
other actions must change it. Every scenario also declares whether the final
model differs from its source. Optional exact final digest, paragraph text,
table-cell text, and table-shape assertions make deterministic scenarios more
specific. Export/reopen must reproduce the same semantic digest.

Run the development playground and replay the manifest:

```sh
pnpm --filter @extend-ai/react-docx-playground dev \
  --host 127.0.0.1 --port 4173

pnpm run word-oracle:edit-roundtrip \
  --manifest tests/fidelity-corpus/word-oracle.json \
  --actions tests/fidelity-corpus/edit-actions.json \
  --base-url http://127.0.0.1:4173 \
  --out-dir .artifacts/word-edit-results
```

`--scenario` can select one or more scenario IDs. Browser, viewport, DPR,
locale, timezone, renderer build, and optional font fingerprint overrides are
recorded in the result. Existing output is preserved unless `--force` is used.

For each scenario the runner:

1. Verifies the source DOCX digest before upload.
2. Waits for the matching loaded status, hooks, fonts, images, pages, and stable
   summary/render state without a fixed sleep.
3. Validates current paragraph/table targets and replays one allowlisted action
   at a time, waiting for consecutive stable hook/render frames between them
   and enforcing both its semantic transition and exact target/effect
   postcondition.
4. Captures the actual browser download and hashes the exported DOCX.
5. Opens a fresh playground page, imports that exported DOCX, fails on any
   import error, and compares its full canonical editable-body fingerprint plus
   structural summary with the pre-export model. Internal source XML,
   generated block IDs, object URLs, and mutation-provenance fields are omitted
   from the fingerprint; text, formatting, tables, fields, and image semantics
   are retained.
6. Writes all artifacts atomically after every selected scenario succeeds.

The result tree is directly traceable:

```text
.artifacts/word-edit-results/
  edit-results.json
  sections-edited-heading/
    sections-edited-heading.docx
```

Every scenario in `edit-results.json` records the source/action digests,
per-action outcomes, source/edited/reopened summaries, export size and SHA-256,
and this registration fragment:

```json
{
  "wordOracleRegistration": {
    "caseId": "sections-edited-heading",
    "source": {
      "path": "sections-edited-heading/sections-edited-heading.docx",
      "sha256": "...exported DOCX digest..."
    }
  }
}
```

Use that local source with `record-reference.mjs` after Word or Graph produces
its PDF reference. The action contract tests need no browser or binary fixture:

```sh
node --test tests/fidelity/edit-action-contract.test.mjs
```

The generated annotation browser smoke imports one safe insertion and one
comment, accepts the change, verifies undo/redo, resolves the comment, exports,
and freshly reopens the DOCX while requiring annotation counts, resolution
state, and the semantic digest to survive:

```sh
pnpm exec playwright test tests/visual/annotation-roundtrip.spec.ts
```
