# JSZip writer-route exclusion and false-positive guard correction

- Date: 2026-10-18
- Work ID: W140; owner: tower-high-capability
Stage: design Spec ready for user review; not an executable Plan or implementation authorization.

## 1. Purpose, authority, and scope

Enable offline, local, read-only DOCX preview in Obsidian while removing actual JSZip archive-generation and compression-writer capability from the shipped runtime closure. Preserve ZIP loading, STORE entry reading, DEFLATE decompression, CRC-related reading behavior, and the existing safe preview pipeline. A filename shared by reading and writing is not proof of a writer route.

The user's “是” approved the architectural direction **to write this Spec only**. The parent specification remains `docs/superpowers/specs/2026-10-11-office-document-fidelity-design.md`; the governing existing Plan is `docs/superpowers/plans/2026-10-11-office-document-fidelity.md`, especially W75, W95 and W139. This focused Spec neither relaxes their acceptance criteria nor marks any checkbox complete.

### Original requirements and stable trace IDs

| ID | Required outcome / constraint |
| --- | --- |
| WR1 | Remove JSZip generation/write routes and compression writers from both authoritative production JavaScript closures; no usable JSZip writer API remains. |
| WR2 | Preserve read-only loading of stored and deflated DOCX entries, CRC semantics, real docx-preview integration, cancellation, errors and resource boundaries. |
| WR3 | Preserve original Office files and caller buffers byte-for-byte; remain offline/local-only, with no OLE payload reading or execution, telemetry, network access, remote code or dynamic script injection. |
| WR4 | Preserve isolated-Worker renderer verification, engine/main ownership, paired build binding, Obsidian compatibility and all existing byte/security/license gates. |
| WR5 | Distinguish historical observations, current graph proof, evidence-only capture, S5.3 acceptance and strict Office pixel parity. Fail closed on unresolved writer ownership. |
| WR6 | Written-Spec approval precedes the separate executable Plan; fresh independent Plan review and checkpoint precede any implementation. |

**Terminology boundary:** WR1 concerns JSZip/general archive generation and compression writer implementations, including STORE writer routes inside JSZip. It does not prohibit the already scoped project-owned, bounded in-memory STORE serializer used to construct a sanitized preview package (`src/ooxml/storedZipWriter.ts`). That existing component is explicitly engine-owned in the current distribution contract. It is not an Office save/export facility and must not become one. Claiming “no ZIP-producing code anywhere” would contradict that existing architecture; this Spec makes no such claim. If the user intends to prohibit even that internal serializer, Tower must obtain a scope decision before planning, rather than silently removing it or broadening the exception.

## 2. Evidence and uncertainty

### Confirmed historical facts, not a new production-graph measurement

- Recovered project memory `mem_1792225330555_on7esmb8d` (2026-10-17) reports JSZip 3.10.1's read path reaching `lib/crc32.js`, `lib/flate.js` and `lib/compressions.js`. CRC metadata and DEFLATE reading need shared code; banning these filenames as generation/deflate code falsely rejected that build. Its measured closure lacked `lib/generate/**` and `pako/lib/deflate.js`, and engine externals were empty. These absences do **not** prove that shared modules contain no reachable writers.
- The same historical report says a narrowly targeted alias for JSZip's optional Node-readable-stream probe/adapter allowed the browser dual build to proceed. This is not a reason to stub arbitrary stream imports.
- W139 records W137's attempt-3 capture exiting successfully with the expected unchanged `crc32.js` ownership rejection, and W138's fresh evidence audit returning PASS_HIGH_RISK with zero findings. W140 accepts these as supplied reports and has not accessed evidence directories. Their result is **guard failed / evidence only**, with `s5_3Pass: false`, not distribution acceptance.
- W69/S5.3 remains blocked/unpassed. A successful evidence capture or its independent audit is not a successful ownership guard, S5.3, S5.4/S5.5, real Obsidian opening, or 100% Office fidelity. The original Office-reference blocker and strict zero-differing-pixel requirements remain unchanged.

### Source-only observations in W140

`scripts/docx-build-config.mjs` builds browser-platform, minified production CJS with two entries, metafiles, in-memory outputs and no engine externals. Its `assertOutputOwnership` currently rejects JSZip input names matching generation, flate, compress or CRC prefixes; this conflates shared reader modules with writers. Its normalized input inventory alone cannot prove function-level absence in mixed modules.

`scripts/read-only-zip-plugin.mjs` aliases `jszip` to the local read host and `pako` to inflate. Its readable-stream alias is importer-limited to JSZip support and Node output-adapter modules. `src/ooxml/jszipReader.ts` imports JSZip load/ZipObject internals, uses a multi-argument `file` method to populate the in-memory read model, and explicitly loads with `checkCRC32: false`. Internal entry registration is not by itself archive generation; blindly removing that method would break loading. CRC presence is not proof that every read currently verifies CRC.

The manifest is desktop-only, version 3.5.0, minimum Obsidian 1.12.7; root lock declarations pin JSZip 3.10.1 and docx-preview 0.3.7. Exact installed dependency bytes, all resolved pako instances, current emitted function reachability and a complete writer inventory are **not established here**. No current production build or evidence read was performed.

## 3. Alternatives, comparison and decision

| Approach | Requirements / trade-off | Disposition |
| --- | --- | --- |
| Proven writer-edge pruning with source-origin-preserving decode-only boundaries | Best preserves WR2–WR4 and existing dependencies. Requires exact graph, mixed-module ownership and production-byte proof for WR1/WR5; cannot be a regex-only guard fix. | Selected direction. |
| Merely narrow the filename regex or provide broad empty/no-op replacements | Cheap but may admit mixed-module writers or silently break readers; fails WR1/WR2/WR5. Missing pako Deflate alone does not exclude JSZip STORE generation. | Rejected. |
| Replace the ZIP reader dependency or maintain a fork | Could provide an explicit decode-only surface, but adds parser/security compatibility work, maintenance, licensing and dependency/distribution costs. | Deferred; reconsider only if safe narrow separation cannot be proven, under a revised approved scope. |

Select narrow build-graph reachability pruning / route exclusion at **proven writer-only imports**. Retain shared CRC/inflate/compression-dispatch modules only with proof that their surviving behavior is decode-only. Do not reopen a general engine/dependency selection exercise.

## 4. Intended architecture and invariants

### Graph and source-origin boundary (WR1/WR2/WR5)

The later Plan must identify each actual production entry, import edge, physical/virtual source identity and mixed read/write component using the authoritative graph bound to dependency bytes and output hashes. Classify each relevant module/route as read-required, writer-only or unresolved; record why and which callers reach it. Enumerate writers from this graph, not from guessed filenames.

Historical candidates requiring explicit investigation include generation entrypoints, `flate.compressWorker`, STORE `compressWorker`, `zipObject._compressWorker`, `compressedObject.createWorkerFrom`, and the `_pakoAction` Inflate/Deflate dispatcher. These are investigation obligations from the existing Plan, not a complete or approved current exclusion list. Aliased/computed access, prototype methods and constructor dispatch must not hide retained write capability.

Remove proven writer-only edges while preserving actual read callers. If mixed modules require a decode-only boundary, its exact replacement content, dependency version/source binding, import scope and absence proof must be specified in the executable Plan before Build. Unknown source versions, changed imports, ambiguous ownership or unproven dispatcher behavior block acceptance. Do not simply admit a mixed module because it contains Inflate or because no dedicated generation file appears.

**Build/runtime invariant:** neither shipped output contains a reachable JSZip archive generator, STORE/DEFLATE writer implementation, writer dispatcher route, or usable JSZip generation/export API. Removal must be structural, not merely a caller promise or a runtime disabled flag. A forbidden call, if a compatibility surface must exist, must reject explicitly before work/output; no successful empty archive, silent no-op or hidden fallback. Such a rejecting surface cannot retain the writer implementation behind it.

### Read and isolation boundary (WR2–WR4)

Retain the existing admission → sanitized in-memory preview package → real docx-preview path, without OLE payload reads or execution and without mutating original bytes. Preserve worker isolation and parent-global restoration guarantees in verification; do not treat the existing Node-worker parse test as proof of live Obsidian rendering or complete runtime isolation. Any discovered mismatch between required isolation and current integration is reported, not silently redesigned in this work.

Preserve the existing optional Node-stream alias unchanged unless graph evidence demonstrates a necessary focused change. Separately test exact intended importers, unrelated importer noninterception, unsupported stream behavior and normal buffer-based reads. Never generalize it into a global Node shim or a successful no-op writer.

Retain ZIP admission limits, path/content policy, cancellation/error propagation, owner-document/popout behavior, desktop guards and lifecycle disposal. CRC metadata reading and explicit CRC-check behavior must remain testable: this work must not silently change the current `checkCRC32: false` policy to true, nor remove CRC support needed for an explicitly check-enabled loader test. The later Plan must name the real check-enabled seam and corruption oracle; any product-policy change needs separate approval.

### Distribution boundary (WR4)

Keep the four-file distribution: `main.js`, `docx-engine.cjs`, `manifest.json`, `styles.css`. Preserve main ≤ 512,000 B; archive ≤ 8,388,608 B; payload ≤ 8,388,192 B; four-member STORE overhead 416 B. Engine allowance remains measured archive cap minus actual main/manifest/styles bytes and overhead, not an enlarged arbitrary cap. Keep paired build identity/digest binding, no engine-owned code in main, zero engine externals, no extra runtime chunks/maps, and existing security/license checks. No version/minimum-API change is selected by this Spec.

## 5. Measurable acceptance and failure behavior

These are future acceptance obligations, not verification performed in W140.

| Criterion | Trace | Required evidence / expected result |
| --- | --- | --- |
| A1: Authoritative writer exclusion | WR1/WR5 | Both actual production outputs and metafiles, exact options, input/dependency hashes, entry edges and output hashes bound to one snapshot; a graph-derived enumerated writer inventory with each route absent/unreachable. Mixed-module emitted-code ownership proof is required where input names cannot settle absence. No unexplained origin, dynamic route or unresolved writer is accepted. |
| A2: Writer-negative sensitivity | WR1/WR5 | Binding/distribution tests reject deliberately retained live writer routes through the exact production minification pipeline in each output: direct generation, STORE writers, DEFLATE dispatcher, property aliases/computed access and unknown/over-limit analysis cases. Demonstrate that each negative fixture survives compilation and fails for the writer invariant, not missing exports or setup failure. Any analyzer requires a finite supported subset, work limits and fail-closed diagnostics in the Plan. |
| A3: Reader-positive coverage | WR2 | Real read-only host loads deterministic STORE and DEFLATE DOCX fixtures with exact entry bytes/text/media. CRC-only and inflate-only positives are admitted; an explicit check-enabled CRC fixture accepts correct CRC and rejects corruption, while the existing default policy is unchanged. No blanket CRC/flate/compressions filename rejection. |
| A4: Parser/render lifecycle | WR2/WR4 | Existing real docx-preview isolated-worker integration remains passing, plus worker render-result coverage appropriate to the actual integration. Cancellation before/during work, parser errors, malformed archives and boundary/over-boundary size cases preserve defined rejection, disposal and no stale output. Parent globals/descriptors remain unchanged. A parse-only result is labelled parse-only. |
| A5: API and alias boundary | WR1/WR2 | Runtime probes of the production-bound reader/engine surface cannot generate an archive through JSZip APIs or retained writer methods; attempted forbidden access is absent or explicitly rejects with no output. Reader entry registration still works. Exact optional stream alias is separately covered, unrelated imports are not intercepted. Runtime probes supplement A1, never replace structural proof. |
| A6: Integrity and local-only behavior | WR3 | Synthetic original buffers compared byte-for-byte before/after success, rejection and cancellation; fixture hashes unchanged. Spies/traps show zero OLE payload reads, filesystem writes to sources, network attempts, telemetry or execution/injection. Any real-document verification remains separately authorized. |
| A7: Binding/distribution/regressions | WR4/WR5 | Existing capture-binding and distribution assertions plus new graph/exclusion regressions pass without weakening input/control/virtual binding or ownership. Actual measured four-file payload/archive and engine/main digest bindings meet section 4; engine externals empty. Future source completion requires `npm run check` with zero warnings/errors. S5.4/S5.5 and release gates remain separate and mandatory when reached. |
| A8: Authorization and provenance | WR5/WR6 | User approval of this written Spec, separate executable Plan, fresh independent Plan PASS and recorded pre-Build checkpoint exist. Reports separately state capture completion, guard result, S5.3 status, implementation review and Office parity; none is inferred from another. |

Unknown graph edges, changed dependency bytes, unclassified virtual modules, unsupported proof constructs, surviving writers, reader regressions, budget overruns or missing binding invalidate acceptance and stop publication. Diagnostics identify the entry, import/source/route and reason where known; errors are retained, not swallowed. Cancellation produces no adopted stale preview; malformed/unsafe input rejects locally without fallback network or original modification. A guard correction alone cannot satisfy A1–A7.

## 6. Decisions reserved for the executable Plan

After written-Spec approval, the high-capability Plan owner must resolve: exact graph/input versions and source hashes; precise import aliases/replacements and source-origin mapping; the permitted module/route allowlist and forbidden inventory; any finite proof algorithm and its bounds; exact target files/interfaces and fixtures; CRC-check seam; per-test command/invocation budget and failure fingerprints; production build/capture gates and evidence lifecycle; ordered dependencies, expected results and independent review/checkpoint evidence. These are deliberately unapproved implementation decisions, not placeholders implementers may fill opportunistically.

No Plan is authored in W140. No guard correction, source/test/config change, dependency installation, build/test/capture run, full release, evidence access/reuse/repair/cleanup, Git staging/commit/push, Office/app automation or document operation is authorized by this Spec. Existing dirty/untracked work and all retained evidence must be preserved. New capture or reuse requires its own explicit scope and reviewed lifecycle; historical capture permission is not recycled.

Open questions for the later Plan are whether the exact current mixed modules admit a bounded decode-only separation and which production-byte proof establishes it. If that cannot be established within the selected narrow approach, report blocked to Tower; do not weaken WR1, invent a module list, add a fork or silently widen scope.

## 7. Approval gate and author self-review

The next gate is **user review/approval of this written Spec**. Approval of the earlier direction is not approval of this artifact. Only afterward may the separate executable writer-route-exclusion Plan be written. Implementation cannot start until that Plan also receives the required user review/execution authorization, a fresh independent-context review PASS and the recorded checkpoint required by W95. Neither Spec approval nor independent review alone authorizes code or evidence operations.

W140 author self-review covers placeholders, internal consistency, single-workstream scope, ambiguity and testability: WR1–WR6 map to A1–A8; shared read modules are not banned by name; JSZip writer exclusion is distinct from the existing bounded sanitizer serializer; current CRC policy and historical graph evidence are not overstated; uncertainty blocks rather than broadens acceptance. Self-review is not independent review. W140 validation is full Spec read-back and `git diff --check` only; no runtime checks, build, capture, evidence inspection or Office parity are claimed.
