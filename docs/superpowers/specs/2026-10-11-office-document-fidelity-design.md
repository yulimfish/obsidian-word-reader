# Office document fidelity: approved direction and product specification

Date: 2026-10-11  
Work item: W2, dependent on W1  
Stage: specification draft ready for user review; implementation and plan not authorized by this artifact  
Design approval: user said “批准，暂无基准” (direction approved; no reference baseline available).

## 1. Purpose and original requirements

Original request: “持续优化直到我的文档 pptWord 可以在 obsidian 里被打开。必要时看看其他开源项目怎么实现的，要求 100% 复原。”

The user wants to read the two identified course documents inside Obsidian without losing their equations, layout, or presentation behavior. Opening successfully is necessary but is not the same as full fidelity. The approved direction is to augment the existing local, read-only renderer, never execute OLE, compare against Microsoft Office page and animation states, and reassess a local Office engine if this route cannot meet the strict target.

| ID | Requirement | Boundary |
| --- | --- | --- |
| R1 | Both identified original files open and remain usable inside Obsidian. | Opening in an external application alone does not satisfy this. |
| R2 | Reproduce all visible text, equations, images, geometry, pagination, and layout with 100% fidelity. | No missing objects, substituted text equations, approximate layout, or unapproved pixel tolerance. |
| R3 | Preserve the PPT's animation states, triggers, ordering, timing, and transitions. | A static final slide or PDF is not a substitute for its presentation behavior. |
| R4 | Preserve offline/local-only, read-only, secure operation and Obsidian compatibility. | No OLE/macros execution, telemetry, remote code, dynamic script injection, or undisclosed network requests. |
| R5 | Establish evidence using real Microsoft Office references. | DOM hashes and passing build/tests do not establish visual parity. |

The approved purpose does not imply approval to install software, launch applications, automate another process, export course material externally, or edit originals.

## 2. Scope and measured evidence

Only the following files in the supplied `2026-10` directory are fidelity targets. This specification does not generalize acceptance to all DOCX/PPTX files.

### D1: Word document

- File: `第4-5讲 习题四 极限存在准则和两个重要极限.docx`.
- Size: 125,738 bytes.
- SHA-256: `0c25bd8fbf8834212b53a9e617bd2b66c0a6db1c75f8027ccf1dd3abebf43e43`.
- Main document has 55 paragraph elements and 50 independent OLE object references; 50 embedding parts.
- W1 classified these as 10 Equation.3 and 40 Equation.DSMT4 objects, with no native `oMath` in the main document.
- Media inventory: 45 WMF and one PNG. W1 found 50 VML `imagedata` references. Counts are not one-image-per-equation guarantees: previews may be reused.
- Cached document properties say four pages. Actual Office pagination has not been observed; four pages is not an acceptance baseline.

### D2: PowerPoint presentation

- File: `第6讲 2.7连续函数+2.8闭区间上连续函数的性质.pptx`.
- Size: 2,466,540 bytes.
- SHA-256: `613347c276dda1f2c3e86798cbaa68b7e28df19a5d06aa59937110abd59811e3`.
- 29 slide parts; 146 independent OLE references identified by slide part plus relationship ID; 146 embedding parts.
- W1 observed 292 raw `oleObj` occurrences across alternative representations. They must not be interpreted as 292 independent equations.
- Media inventory: 96 WMF, 49 EMF, 10 PNG, three JPEG.
- 27 slide parts contain timing structures. W1 observed 44 `anim` and 152 `animEffect` elements in slide XML and 189 `AlternateContent` elements across the package. These are structural counts, not a count of actual playback steps or supported effects.
- All 29 slides, including slides without timing, remain in scope.

W2 re-read both files and verified their sizes, SHA-256 values, paragraph/slide counts, independent OLE reference counts, media extensions, and the count of slides containing timing. W1 found no External relationships in either package. XML and media presence do not prove the preview bytes are correct, decodable, or visually faithful.

Relevant fonts observed by W1 include Times New Roman and 宋体; D2 additionally uses names such as 楷体_GB2312 and 黑体. Installed font availability, font versions, and actual substitution remain unverified.

## 3. Current implementation and evidence limits

| Area | Actual artifact | Consequence |
| --- | --- | --- |
| File entry | `src/main.ts:69–92` registers reader views and extensions. | Format registration already exists; it is not the primary identified failure. |
| DOCX package gate | `src/docx/DocxAdapter.ts:41–57`, `validateDocxPackage`. | Reads through the vault, validates ZIP, then invokes the shared policy. |
| OLE rejection | `src/ooxml/packagePolicy.ts:55–63,92–99`, `enforceOoxmlPackagePolicy`. | Any embedding/oleObjects part or OLE content type rejects the package; both samples match this condition. |
| DOCX rendering | `src/docx/docxRenderer.ts:30–41`, `renderDocx`. | Uses `docx-preview` `parseAsync` and `renderDocument`; options retain dimensions/fonts and enable page breaks. This is not proof of Word pagination equivalence. |
| PPTX package gate | `src/pptx/pptxPackage.ts:168–223`, `PptxPackage.load`. | Applies shared policy before creating slide contexts. |
| PPTX media | `src/pptx/pptxPackage.ts:398–410`, `getImageMimeType`. | Recognizes JPEG/BMP/PNG only, not WMF/EMF. |
| PPTX alternatives | `src/pptx/pptxRenderer.ts:271–365`, `renderLayer`/`renderShape`. | Shape traversal only accepts five element types, not `AlternateContent`; both top-level and grouped traversal need a consistent alternative-selection boundary. |
| PPTX visual tests | `tests/pptxVisualSnapshot.test.ts:19–70,107–133`. | Simulated DOM/style serialization is hashed, not painted pixels compared to Office. |
| Current security tests | `tests/docxPackageSafety.test.ts:14–27`. | Explicitly expects rejection of OLE parts; safe static-preview handling requires new positive and negative security criteria rather than merely removing this test. |

Other existing coverage includes `tests/pptxPackage.integration.test.ts`, `tests/pptxRenderer.integration.test.ts`, and `tests/ooxmlPackageSafety.test.ts`. Their existing presence does not constitute sample-specific visual acceptance.

W1's attempted in-memory invocation failed because `esbuild` was not installed. Neither W1 nor W2 has reproduced the samples in the running Obsidian app or run the required source-change check suite. The rejection conclusion is grounded in package structure and the explicit source branch, not an observed live error screen.

## 4. Divergence, comparison, and selected direction

### Alternatives

1. **Augment the current renderer with safe static previews and measured fidelity correction.** Best fit for R1/R4 and existing integration. Requires legacy vector decoding, Word layout corrections, and real PPT timing support; R2/R3 remain high-risk until measured.
2. **Integrate a mature local Office engine.** Potentially stronger layout and presentation support for R2/R3. Adds packaging, isolation, licensing, fonts, runtime and installation dependencies for R4; still needs R5 and cannot be assumed pixel-identical to Microsoft Office.
3. **Display Office-exported fixed pages/slides locally.** Can preserve a particular static reference output, but a static export alone fails R3 and may weaken search/copy/navigation. Needs Office-produced artifacts, currently absent. Not selected as a silent replacement.

### Decision

Use alternative 1 first, as approved. Keep the reader shell, local file entry, lifecycle ownership, and existing safe package limits. Separate static visual data from active embedded payloads, rather than treating all OLE as executable or all local bytes as trusted. Require evidence for each correction. If strict fidelity is demonstrably unattainable on this route, report the gap and reassess alternative 2 with the user before introducing an engine or dependency.

Upstream references reviewed in W1:

- [docxjs / docx-preview](https://github.com/VolodymyrBaydalka/docxjs): Apache-2.0; upstream describes HTML limitations, no real-time automatic pagination, and only `renderAsync` as stable. Internal parse/render integration requires compatibility coverage. This library is not an assurance of Word pixel parity.
- [ONLYOFFICE DocumentServer](https://github.com/ONLYOFFICE/DocumentServer): AGPL-3.0; components include conversion core, fonts and sdkjs, with presentation animation/transition support advertised. Useful as an architectural reference; no sample-specific validation or licensing/integration approval exists.

No WMF/EMF dependency has been chosen. Selection must later establish actual required opcode support, malformed-input handling, offline packaging, compatible licensing, and rendered results. Upstream compatibility claims alone do not meet R5.

## 5. Technical behavior and component boundaries

### 5.1 Safe package admission and preview association (R1/R2/R4)

- Retain ZIP size/count/compression limits, path validation, encryption/damage detection, and denial of macro/ActiveX/script-capable content.
- Permit only a validated, inert static-preview path for the legacy equation representations required by D1/D2. An OLE class name or `.bin` extension is not a trust decision.
- Locate preview relationships structurally and validate that targets remain inside the package. Record document part, relationship, selected branch, shape identity and preview association for each logical equation occurrence.
- Treat embedded OLE payload bytes as opaque and non-executable. Never instantiate COM, launch OLE handlers, evaluate code, activate links, or use the payload as a script/image source. No equation activation or editing is provided.
- Unsupported, missing, external, ambiguous or malformed previews produce an explicit object-level diagnostic and make content/fidelity acceptance fail. They must not trigger hidden network requests or substitution with guessed formula text.
- Do not broadly weaken the shared XLSX policy. Scope any changed admission behavior to the validated DOCX/PPTX preview path; unrelated formats retain their prior safety guarantees.
- Original ZIP bytes and files remain unchanged. Any renderer-compatible transformed package/model exists only in bounded local memory; no transformed course file overwrites the source.

### 5.2 Legacy image decoding (R2/R4)

- Decode the referenced WMF/EMF records needed by these documents into safe local visual output with correct dimensions, coordinate mapping, clipping, pen/brush/text semantics and font dependencies.
- Reject invalid records, unsupported escape/executable behavior, out-of-bounds reads, recursive/oversized allocations and excessive work; cancel through reader lifecycle ownership.
- Do not claim browser MIME support is sufficient or rename WMF/EMF as PNG. Exported SVG, if used internally, must be constructed from permitted drawing primitives without scripts, external URLs, event handlers or arbitrary embedded markup. Do not change the package's general script-media admission policy to accommodate it.
- Preserve raster inputs and equation preview quality. Caches and blob URLs must be scoped, bounded, cancelled and revoked consistently with existing sessions.

### 5.3 DOCX layout and accessible reading (R1/R2/R4)

- Resolve each of the 50 logical OLE occurrences to its static visual representation, including VML inline/anchored placement and surrounding text alignment.
- Preserve visible glyphs, formula structure, superscripts/subscripts as painted in the preview, line breaks, spacing, page geometry and pagination. Do not rebuild legacy equations using guessed MathML/LaTeX as a fidelity shortcut.
- Preserve text search/copy and existing zoom/fit/navigation where supported. Provide existing accessible names or neutral equation descriptions without inventing formula semantics. Bitmap/vector previews do not automatically provide accessible mathematical semantics.
- A missing font must be disclosed as an unmet fidelity dependency, not silently accepted through fallback.

### 5.4 PPTX normalization and playback (R1/R2/R3/R4)

- Select exactly one supported `AlternateContent` representation using namespace/capability rules, with fallback where appropriate. Preserve placement, group transforms, crop, ordering and identity; do not render Choice and Fallback together.
- Render static equation previews at their original locations and retain shape identifiers needed by timing targets. Deduplication must not remove legitimately repeated visible occurrences sharing an image.
- Preserve master/layout/theme inheritance and all visible objects actually present in these samples. Feature discovery is evidence-driven, not an unlimited Office feature implementation.
- Map timing targets, click/automatic triggers, entrance/exit/emphasis/motion behavior and slide transitions to deterministic, cancellable local playback. Handle initial, intermediate and final states rather than showing every element immediately.
- Leaving a slide, reopening a document or closing the view must cancel playback and restore a defined initial state; no leaked timers or cross-slide effects.
- Any unsupported actual effect/target/trigger remains an explicit failed R3 criterion. Never equate parsing a timing tree with executing it correctly.
- Preserve existing keyboard navigation and readable controls. Any later new playback controls or DOM hierarchy changes require the UI wireframe/confirmation gate; this specification does not approve an unpresented visual layout.

## 6. Definition of strict 100% and acceptance matrix

Strict acceptance is for D1/D2 under a recorded reference environment and an exhaustive inventory of their pages, objects, playback steps and sampled frames. It is not a promise about all Office documents, devices or the mathematical continuum of animation time.

| Criterion | Trace | Required result and evidence |
| --- | --- | --- |
| A1: Original integrity | R4 | Both hashes match section 2 before and after verification; no course-file writes. |
| A2: Opening and lifecycle | R1/R4 | Both open inside Obsidian; all D1 reference pages and 29/29 D2 slides are navigable; close/reopen, reload, zoom/fit and existing search/copy paths remain usable; no crashes, retained playback or broken preview resources. Capture real-app evidence. |
| A3: D1 completeness | R2 | Map all 55 paragraph elements (including empty/layout paragraphs) and 50 logical OLE occurrences to the rendered document; all visible text and previews match the reference, with no omission/duplication/error placeholders. Actual page count and every page break equal the Office baseline; do not hard-code cached four-page metadata as truth. |
| A4: D2 completeness | R2 | All 29 slides and 146 logical OLE references are accounted for, including legitimate repeated occurrences; selected alternatives render once; all reference-visible text/geometry/media/layers/crops match. Unreferenced media inventory need not appear onscreen, but every reference-visible object must. |
| A5: Static pixel equality | R2/R5 | For every D1 page and D2 reference state, registered content-area RGBA images have identical dimensions and zero differing pixels; maximum channel difference is zero. Report a diff image and counts for every failure. |
| A6: Animation behavior | R3/R5 | Enumerate actual playback sequences for all 29 slides, explicitly including the 27 with timing; initial state, every trigger, stable state, intermediate effect frames and transition match reference ordering, targets and scheduled timings. At a fixed recorded 60 fps reference sampling grid, compare every frame of every effect/transition segment using A5. Repeated deterministically timestamped playback must produce the same results. No unsupported effect passes. |
| A7: Offline and safe handling | R4 | Verification with outbound networking disabled succeeds; no external resource attempts or payload execution. Negative fixtures retain rejection/containment of macros, ActiveX, unsafe paths/relationships, invalid previews, vector escapes, excessive work and script media. |
| A8: Quality and compatibility | R1/R4 | All project-required checks pass with zero warnings/errors; supported APIs match manifest minimum, desktop access is guarded, popout owner-document/window behavior and listener/timer disposal are verified. |
| A9: Evidence provenance | R5 | Baseline identifies exact original hashes, reference software/environment, fonts, rendering/capture parameters, pages/state/frame inventory and capture provenance. DOM-only snapshots and CI green status cannot satisfy A5/A6. |

No SSIM threshold, perceptual approximation, antialiasing exemption, ignored equation region, missing-font fallback, rescaling, or blur may convert a failure into a pass. Crop only the predeclared document/slide content rectangle, excluding application chrome; no content-dependent masks. Compare lossless captures with fixed color/alpha handling. Antialiasing differences count as differences. Any requested tolerance is a separate scope change requiring explicit approval.

Animation sampling is a stated measurement boundary, not proof of equality between sampled instants. If references cannot provide deterministic timestamps/frames or documented timing, A6 is blocked rather than waived. Stable states alone do not establish animation fidelity. When reduced-motion behavior is active, report that separately and do not use it to bypass the normal-playback reference test.

Results must distinguish **opening verified**, **structural/content coverage verified**, **static parity verified**, **animation parity verified**, and **full acceptance**. The first two may progress before references exist; they must never be reported as 100% fidelity.

## 7. Baseline availability and safe acquisition protocol

### Current limitation

The user has no baseline. W2 inspected application names and bundle metadata only, without launching applications or opening documents:

- `/Applications`: WPS Office 12.1.28496, Pages Creator Studio 15.4, Keynote Creator Studio 15.4 were found.
- No Microsoft Word, Microsoft PowerPoint, LibreOffice, OpenOffice or ONLYOFFICE application was found by the name-filtered, top-level inventory of `/Applications` and `~/Applications`.
- This bounded inventory does not prove Microsoft Office is absent from all disks or that installed alternatives can safely/readily render these documents. Licensing, installed fonts, unlocked UI, permissions and export capabilities were not tested.

WPS/Pages/Keynote output may help diagnostics only after separate authorization; it is not a Microsoft Office baseline. No substitute becomes authoritative without user-approved change to R5.

### Minimum baseline acquisition protocol (dependency, not current authorization)

1. Obtain explicit authorization for the chosen reference workflow: user-supplied Office captures/exports, or an authorized local Microsoft Office environment. Do not install or launch software merely because this specification exists. No upload to cloud converters.
2. Use read-only working copies in an approved local location; record source and copy hashes. Disable macros, external-link updates and embedded-object activation. Never activate/edit equations or open their OLE handlers. If safe rendering requires activation or conversion that would execute embedded content, stop and report the blocker.
3. Record exact Word/PowerPoint version/build, OS, installed/used font files and versions, locale, page setup, renderer/Obsidian version, viewport, content rectangle, physical output size, DPR, zoom, color profile and capture method. Missing fields make the strict baseline incomplete; no invented values.
4. For D1, acquire lossless captures of every Office-rendered page at fixed output dimensions and record real pagination. PDF export can preserve page inventory, but any PDF-to-image conversion must be pinned and used consistently; do not conflate PDF rasterizer pixels with an unrecorded live Word capture.
5. For D2, acquire all 29 slide reference images plus a trigger/state manifest and lossless, timestamped frame sequences for every effect and transition at the fixed 60 fps grid. Include playback duration, target IDs, trigger order and initial/final states. Static PDF/export alone is insufficient.
6. Confirm repeatable reference outputs before evaluating the plugin. Preserve captures and a manifest locally outside the source documents; no sample binaries or private course reference images are committed or published without explicit consent. Recheck original hashes.

**Baseline blocker B1:** No authoritative Office output or confirmed authorized Office reference environment currently exists. Therefore A5/A6/A9 and full 100% acceptance are blocked. This does not block writing this specification or later authorized opening/content work, but no implementation success can remove B1 by assertion.

## 8. Dependencies, assumptions, risks, and reassessment

### Dependencies and assumptions

- D1/D2 remain byte-identical to section 2. Changed inputs require a new inventory and reference manifest, not silent reuse of old acceptance.
- Existing desktop Obsidian integration is retained; W1 read `manifest.json` as minimum version 1.12.7 and desktop-only. Any API upgrade must be evaluated explicitly.
- Static previews are present but their complete visual validity is not yet established. No assumption that every formula is recoverable merely because a preview relationship exists.
- Reference font availability/rights and baseline provenance are unresolved dependencies, not permission to redistribute fonts.
- A bounded decoder with required record coverage and compatible license must be demonstrated before dependency selection. No dependency install is part of W2.

### Principal risks

- HTML text layout and automatic pagination may not match Word; preview text may depend on missing fonts.
- WMF/EMF support may omit required records or have unsafe complexity; no decoder is presumed trustworthy.
- Choice/Fallback normalization may alter identity or object order and break animation targets.
- PPT timing semantics, automatic triggers, motion and transitions may exceed current renderer capabilities.
- Browser and Office antialiasing/color pipelines may make exact pixel equality unattainable even when geometry is correct. This remains a failed strict criterion, not a hidden tolerance.
- Screenshots without deterministic playback capture cannot prove timing parity.
- Replacing a blanket OLE rejection with overly broad admission would weaken security, including unrelated XLSX behavior.

### Triggers for local-engine reassessment

Reassess when any of the following is demonstrated and recorded: a required preview is missing/unrecoverable without payload execution; a decoder cannot safely support a required record; required Word pagination/layout cannot match the fixed reference; a required PPT effect/trigger cannot be faithfully represented; or systematic pixel differences persist due to the rendering architecture rather than an isolated defect.

Do not invent a promise of endless optimization or silently switch engines. Produce the failing document/page/object/state, expected versus actual evidence, root cause and attempted corrections. Present a local-engine feasibility recommendation covering offline behavior, isolation/no OLE execution, fonts, size, licensing and distribution. A new engine requires a revised specification and explicit user approval; it must pass the same fidelity criteria and is not automatically pixel-perfect.

## 9. Verification obligations and non-goals

Future source changes require `npm run lint`, `npm run typecheck`, `npm test`, `npm run build`, and `npm run security:scan`, preferably the project's `npm run check`. Warnings are blockers. These checks supplement, never replace, A1–A9. Sample-specific security, package association, decoder, actual browser paint, Obsidian lifecycle and playback evidence are needed. Fresh separate-context reviews must assess the written spec/plan and later implementation/fixes; the author cannot claim independent review.

Non-goals: Office editing; OLE activation or equation editing; macro support; online conversion; telemetry; mobile support; XLSX fidelity expansion; all-document compatibility promises; font redistribution; replacing current reader features with static screenshots alone; unrelated refactoring; dependency installation or app/process automation during specification authoring.

## 10. Approval gate and author self-review

This is a product/technical specification, not an implementation plan. No task sequence, dependency install, source changes or UI changes are authorized by its existence. The next gate is user review and approval of this written specification. Only after that approval may the implementation plan be authored. Controller must also arrange fresh-context specification review; this author does not launch another agent.

Author self-review checklist:

- R1–R5 retained and traced to A1–A9; opening and strict fidelity are distinct.
- Alternatives compared and the approved existing-renderer direction selected; engine adoption is gated.
- Two-document scope, no OLE execution, no original modification and no silent tolerance preserved.
- Measured counts distinguished from cached properties and actual runtime/visual evidence.
- Baseline absence explicitly blocks visual/animation/full acceptance; installed alternative apps do not remove it.
- Static preview/security boundary, vector decoding, alternative selection, equation accessibility limits and animation states covered.
- No placeholder decisions masquerade as approved implementation choices; unresolved dependencies are named explicitly.
- No tests, live application verification, pixel parity, font availability or independent review claimed as passed.

W2 completion evidence is the written artifact, read-back/scope verification and unchanged source-document hashes. Independent review and user approval remain pending.
