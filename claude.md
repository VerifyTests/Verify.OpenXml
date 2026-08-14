# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Verify.OpenXml is a [Verify](https://github.com/VerifyTests/Verify) plugin that converts Excel (xlsx) and Word (docx) documents into human-readable, deterministic formats for snapshot testing. It registers stream and file converters with Verify so that test assertions can be made against Office documents.

## Build & Test Commands

**Important:** Commands must be run from the `src/` directory (where `global.json` lives), not from the repo root.

```bash
# Build
cd src
dotnet build --configuration Release

# Run all tests — all three test projects. `Tests/Tests.csproj` alone is not the full suite.
cd src
dotnet test Verify.OpenXml.slnx

# Run a single test
cd src
dotnet test Tests/Tests.csproj --filter "FullyQualifiedName~Samples.VerifyExcel"
```

Requires .NET SDK 10.0 (preview). See `src/global.json` for exact version.

Tests use NUnit and target `net10.0` only. The library targets `net472;net48;net8.0;net9.0;net10.0`.

## Architecture

All source lives under `src/`. Build from the `src` directory; the solution is `src/Verify.OpenXml.slnx`.

### Library (`src/Verify.OpenXml/`)

Entry point is `VerifyOpenXml.Initialize()` which registers six converters with Verify — a stream converter per extension (`xlsx`, `docx`, `pptx`) and a file converter per document type (`SpreadsheetDocument`, `WordprocessingDocument`, `PresentationDocument`).

Only the stream converters extract anything. Verify runs the stream converter registered for an extension over the targets a file converter returns, so each file converter just clones the package and hands it over (`ToPackage`). Doing the extraction in both places produced every info, text and PNG target twice. `ToPackage` is generic because `Clone` resolves the package factory from the static type and throws when called through the `OpenXmlPackage` base.

Key files:
- **VerifyOpenXml.cs** — Initialization and `ToPackage`
- **VerifyOpenXml_Excel.cs** — Excel conversion (stream→CSV, metadata extraction, deterministic binary output)
- **VerifyOpenXml_Word.cs** — Word conversion (text/font/property extraction, deterministic binary output)
- **VerifyOpenXml_Powerpoint.cs** — Powerpoint conversion
- **MorphRenderer.cs** — Backend probing and the shared PNG page targets (`net10.0` only)
- **Info.cs / WordInfo.cs** — Data models for extracted document metadata

Each stream converter returns a `ConversionResult` containing:
1. An info object (metadata serialized to JSON)
2. Text targets (CSV for Excel sheets, TXT for Word/Powerpoint text)
3. A deterministic binary copy of the original document (via `DeterministicIoPackaging`)
4. On `net10.0` with a Morph backend present, one PNG per rendered page

All three document types render. Morph exposes a separate converter per type (`DocumentConverter`, `ExcelConverter`, `PowerPointConverter`) with no common base, so `MorphRenderer` captures each one's `ConvertToImageData` as a delegate and shares a single `AddPages`. Rendering reads the deterministic package when one was built and the raw clone otherwise — `DeterministicPackage` only normalizes zip container metadata, so the pixels are the same either way. This is why each converter clones unconditionally rather than only inside the `IsTargetExcluded` branch.

A page is not a sheet or a slide by definition: Word paginates by layout, Powerpoint emits one page per slide, and Excel paginates by *print* layout — a long sheet spills onto several pages, and each visible sheet starts a new one.

Document text is a target only — never also a property on the info object, or it lands in two snapshot files.

Word text extraction walks the body in document order and treats content controls (`w:sdt` — `SdtBlock`, `SdtRun`, `SdtRow`, `SdtCell`) as transparent, emitting their content as Word displays it.

The deterministic binary output is critical — `DeterministicIoPackaging` ensures identical binary output across .NET runtimes so that `.verified.xlsx`/`.verified.docx` files are stable.

### Tests (`src/Tests/`)

- **ModuleInitializer.cs** — Calls `VerifyOpenXml.Initialize()` via `[ModuleInitializer]`
- **Samples.cs** — Core tests verifying Excel/Word files, streams, and document objects
- Verified snapshot files (`.verified.txt`, `.verified.csv`, `.verified.xlsx`, `.verified.docx`) live alongside tests

`sample.pptx` is hand-built rather than authored in PowerPoint, and two things about it are load-bearing:

- The title placeholder carries an explicit `<a:xfrm>`. Its layout and master are stubs with no placeholder geometry, so without it Morph has nothing to position the text with and renders a blank slide — the text targets still pass, so the blank PNG is easy to accept by mistake.
- `ppt/slideLayouts/_rels/slideLayout1.xml.rels` exists. A slideLayout part must relate to its slideMaster; without it PowerPoint offers to repair both the sample and every `.verified.pptx` derived from it.
- The master's `<p:bgRef idx="1001">` names `bg1`, not `phClr`. `phClr` is the placeholder the theme's `bgFillStyleLst` substitutes *into*, so supplying it as the substitution colour is circular and PowerPoint renders the slide solid black.

Nothing in the test suite notices any of this — the OpenXML SDK, DeterministicIoPackaging and Morph all read the deck fine, and Morph renders the background white either way. Only opening a `.verified.pptx` in PowerPoint surfaces it.

When tests fail, Verify produces `.received.*` files showing actual output. Compare these against `.verified.*` files. To accept new output, replace the verified file with the received file.

### Rendering-backend tests (`src/Tests.ImageSharp/`, `src/Tests.Skia/`)

These exist to exercise the two Morph PNG-rendering backends. Each references one backend package and links the test sources from `Tests` (`<Compile Include="..\Tests\**\*.cs">`, excluding the unit-test files and `ModuleInitializer.cs`), so they run the same Samples/Excel/Word tests a third time.

**They keep their own copies of every snapshot** — their `ModuleInitializer` calls `DerivePathInfo` with the project directory. A change to converter output therefore has to be re-verified in all three directories, and `dotnet test Tests/Tests.csproj` going green means nothing about the other two. Regenerate with `dotnet test Verify.OpenXml.slnx`.

### Regenerating snapshots

Two Verify behaviours make binary (`.docx`/`.xlsx`/`.pptx`) snapshots easy to get wrong:

- They are compared through a comparer, not by raw bytes, so a stale verified package that is merely *equivalent* keeps passing and can sit in the repo for a long time.
- The binary targets set `BypassComparersForSubsequentOnDifference`, so as soon as an earlier text target differs, the binary targets skip comparison and are written out as `.received.*` regardless. A blanket `mv *.received.* *.verified.*` will silently rewrite binary snapshots that were never actually compared.

So after changing a converter, delete the affected `.verified.*` in **all three** test directories, regenerate, and then run the suite twice — the second run is what proves the accepted output is reproducible rather than an artifact of the transitional state.

## Key Dependencies

- **DocumentFormat.OpenXml** — OpenXML SDK for reading Office documents
- **DeterministicIoPackaging** — Makes ZIP-based package output byte-identical across runtimes
- **Verify** — Snapshot testing framework

Package versions are centrally managed in `src/Directory.Packages.props`.

## Build Configuration

- `TreatWarningsAsErrors` is enabled
- `EnforceCodeStyleInBuild` is enabled
- Namespace: `VerifyTests`
- Global type alias: `CharSpan` = `System.ReadOnlySpan<char>` (defined in `src/Directory.Build.props`)
