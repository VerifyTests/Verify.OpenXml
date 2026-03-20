# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Verify.OpenXml is a [Verify](https://github.com/VerifyTests/Verify) plugin that converts Excel (xlsx) and Word (docx) documents into human-readable, deterministic formats for snapshot testing. It registers stream and file converters with Verify so that test assertions can be made against Office documents.

## Build & Test Commands

```bash
# Build
dotnet build src --configuration Release

# Run all tests
dotnet test src --configuration Release

# Run a single test
dotnet test src --configuration Release --filter "FullyQualifiedName~Samples.VerifyExcel"
```

Requires .NET SDK 10.0 (preview). See `src/global.json` for exact version.

Tests use NUnit and target `net10.0` only. The library targets `net472;net48;net8.0;net9.0;net10.0`.

## Architecture

All source lives under `src/`. There is no solution file; build from the `src` directory.

### Library (`src/Verify.OpenXml/`)

Entry point is `VerifyOpenXml.Initialize()` which registers four converters with Verify:
- Stream converter for `xlsx` → `ConvertExcel`
- Stream converter for `docx` → `ConvertWord`
- File converter for `SpreadsheetDocument` (Excel)
- File converter for `WordprocessingDocument` (Word)

Key files:
- **VerifyOpenXml.cs** — Initialization, Excel conversion (stream→CSV, metadata extraction, deterministic binary output)
- **VerifyOpenXml_Word.cs** — Word conversion (text/font/property extraction, deterministic binary output)
- **Info.cs / WordInfo.cs** — Data models for extracted document metadata

Each converter returns a `ConversionResult` containing:
1. An info object (metadata serialized to JSON)
2. Text targets (CSV for Excel sheets, TXT for Word text)
3. A deterministic binary copy of the original document (via `DeterministicIoPackaging`)

The deterministic binary output is critical — `DeterministicIoPackaging` ensures identical binary output across .NET runtimes so that `.verified.xlsx`/`.verified.docx` files are stable.

### Tests (`src/Tests/`)

- **ModuleInitializer.cs** — Calls `VerifyOpenXml.Initialize()` via `[ModuleInitializer]`
- **Samples.cs** — Core tests verifying Excel/Word files, streams, and document objects
- Verified snapshot files (`.verified.txt`, `.verified.csv`, `.verified.xlsx`, `.verified.docx`) live alongside tests

When tests fail, Verify produces `.received.*` files showing actual output. Compare these against `.verified.*` files. To accept new output, replace the verified file with the received file.

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
