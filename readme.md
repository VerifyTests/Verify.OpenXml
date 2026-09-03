# <img src="/src/icon.png" height="30px"> Verify.OpenXML

[![Discussions](https://img.shields.io/badge/Verify-Discussions-yellow?svg=true&label=)](https://github.com/orgs/VerifyTests/discussions)
[![Build status](https://img.shields.io/appveyor/build/SimonCropp/verify-openxml)](https://ci.appveyor.com/project/SimonCropp/verify-openxml)
[![NuGet Status](https://img.shields.io/nuget/v/Verify.OpenXML.svg)](https://www.nuget.org/packages/Verify.OpenXML/)

Extends [Verify](https://github.com/VerifyTests/Verify) to allow verification of Word, Excel, and PowerPoint documents via [OpenXML](https://github.com/dotnet/Open-XML-SDK/).<!-- singleLineInclude: intro. path: /docs/intro.include.md -->


## Features


### Excel (xlsx)

 * Converts workbooks to CSV format for each worksheet
 * Extracts formulas and displays them alongside cell values
 * Captures document properties (title, subject, keywords, description, category, status, company, manager)
 * Captures custom document properties
 * Supports date scrubbing and GUID scrubbing for deterministic tests
 * Generates deterministic XLSX output using DeterministicIoPackaging
 * Optionally renders each page to PNG via [Morph](https://github.com/SimonCropp/Morph) (opt-in)


### Word (docx)

 * Extracts document text content from paragraphs and tables
 * Captures document properties (title, subject, keywords, description, category, status, revision)
 * Captures custom document properties
 * Extracts font information
 * Generates deterministic DOCX output using DeterministicIoPackaging
 * Optionally renders each page to PNG via [Morph](https://github.com/SimonCropp/Morph) (opt-in)


### PowerPoint (pptx)

 * Extracts slide text from every slide, separated by `---`
 * Captures document properties (title, subject, keywords, description, category, status, revision)
 * Reports slide count
 * Generates deterministic PPTX output using DeterministicIoPackaging
 * Optionally renders each slide to PNG via [Morph](https://github.com/SimonCropp/Morph) (opt-in)


**See [Milestones](../../milestones?state=closed) for release notes.**


## Sponsors


### Entity Framework Extensions<!-- include: sponsors. path: /docs/sponsors.include.md -->

[Entity Framework Extensions](https://entityframework-extensions.net/?utm_source=simoncropp&utm_medium=Verify.OpenXML) is a major sponsor and is proud to contribute to the development this project.

[![Entity Framework Extensions](https://raw.githubusercontent.com/VerifyTests/Verify.OpenXML/refs/heads/main/docs/zzz.png)](https://entityframework-extensions.net/?utm_source=simoncropp&utm_medium=Verify.OpenXML)

### Developed using JetBrains IDEs

[![JetBrains logo.](https://raw.githubusercontent.com/VerifyTests/Verify.OpenXml/main/docs/jetbrains.png)](https://jb.gg/OpenSourceSupport)<!-- endInclude -->


## NuGet

 * https://nuget.org/packages/Verify.OpenXML


## Usage


### Enable Verify.OpenXml

<!-- snippet: enable -->
<a id='snippet-enable'></a>
```cs
[ModuleInitializer]
public static void Initialize() =>
    VerifyOpenXml.Initialize();
```
<sup><a href='/src/Verify.OpenXml.Tests/ModuleInitializer.cs#L3-L9' title='Snippet source file'>snippet source</a> | <a href='#snippet-enable' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### Excel


#### Verify a file

<!-- snippet: VerifyExcel -->
<a id='snippet-VerifyExcel'></a>
```cs
[Test]
public Task VerifyExcel() =>
    VerifyFile("sample.xlsx");
```
<sup><a href='/src/Verify.OpenXml.Tests/Samples.cs#L4-L10' title='Snippet source file'>snippet source</a> | <a href='#snippet-VerifyExcel' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Verify a Stream

<!-- snippet: VerifyExcelStream -->
<a id='snippet-VerifyExcelStream'></a>
```cs
[Test]
public Task VerifyExcelStream()
{
    var stream = new MemoryStream(File.ReadAllBytes("sample.xlsx"));
    return Verify(stream, "xlsx");
}
```
<sup><a href='/src/Verify.OpenXml.Tests/Samples.cs#L37-L46' title='Snippet source file'>snippet source</a> | <a href='#snippet-VerifyExcelStream' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Verify a SpreadsheetDocument

<!-- snippet: SpreadsheetDocument -->
<a id='snippet-SpreadsheetDocument'></a>
```cs
[Test]
public async Task VerifySpreadsheetDocument()
{
    await using var stream = File.OpenRead("sample.xlsx");
    using var reader = SpreadsheetDocument.Open(stream, false);
    await Verify(reader);
}
```
<sup><a href='/src/Verify.OpenXml.Tests/Samples.cs#L25-L35' title='Snippet source file'>snippet source</a> | <a href='#snippet-SpreadsheetDocument' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Example snapshot

<!-- snippet: Samples.VerifyExcel.verified.csv -->
<a id='snippet-Samples.VerifyExcel.verified.csv'></a>
```csv
0,First Name,Last Name,Gender,Country,Date,Age,Id,Formula
1,Dulce,Abril,Female,United States,2017-10-15,32,1562,G2+H21594 (G2+H2)
2,Mara,Hashimoto,Female,Great Britain,2016-08-16,25,1582,1607
3,Philip,Gent,Male,France,2015-05-21,36,2587,2623
4,Kathleen,Hanner,Female,United States,2017-10-15,25,3549,3574
5,Nereida,Magwood,Female,United States,2016-08-16,58,2468,2526
6,Gaston,Brumm,Male,United States,2015-05-21,24,2554,2578
```
<sup><a href='/src/Verify.OpenXml.Tests/Samples.VerifyExcel.verified.csv#L1-L7' title='Snippet source file'>snippet source</a> | <a href='#snippet-Samples.VerifyExcel.verified.csv' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### Word


#### Verify a file

<!-- snippet: VerifyWord -->
<a id='snippet-VerifyWord'></a>
```cs
[Test]
public Task VerifyWord() =>
    VerifyFile("sample.docx")
        .Snapshot(
            """
            {
              Properties: {
                Subject: Test Subject,
                Title: Sample Document
              }
            }
            """);
```
<sup><a href='/src/Verify.OpenXml.Tests/Samples.cs#L48-L63' title='Snippet source file'>snippet source</a> | <a href='#snippet-VerifyWord' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Verify a Stream

<!-- snippet: VerifyWordStream -->
<a id='snippet-VerifyWordStream'></a>
```cs
[Test]
public Task VerifyWordStream()
{
    var stream = new MemoryStream(File.ReadAllBytes("sample.docx"));
    return Verify(stream, "docx")
        .Snapshot(
            """
            {
              Properties: {
                Subject: Test Subject,
                Title: Sample Document
              }
            }
            """);
}
```
<sup><a href='/src/Verify.OpenXml.Tests/Samples.cs#L86-L104' title='Snippet source file'>snippet source</a> | <a href='#snippet-VerifyWordStream' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### Binary output across .NET frameworks

When verifying binary package output (xlsx, docx, nupkg, etc.) across multiple target frameworks (e.g. net48 and net10.0), the binary output may differ due to Deflate compression implementation differences. The XML content within entries is identical — only the compressed bytes differ. Use `UniqueForRuntime` to generate framework-specific verified files:

```cs
await Verify(stream, extension: "xlsx")
    .UniqueForRuntime();
```

See [Verify Naming docs](https://github.com/VerifyTests/Verify/blob/main/docs/naming.md) for more details.


#### Verify a WordprocessingDocument

<!-- snippet: WordprocessingDocument -->
<a id='snippet-WordprocessingDocument'></a>
```cs
[Test]
public async Task VerifyWordprocessingDocument()
{
    await using var stream = File.OpenRead("sample.docx");
    using var reader = WordprocessingDocument.Open(stream, false);
    await Verify(reader)
        .Snapshot(
            """
            {
              Properties: {
                Subject: Test Subject,
                Title: Sample Document
              }
            }
            """);
}
```
<sup><a href='/src/Verify.OpenXml.Tests/Samples.cs#L65-L84' title='Snippet source file'>snippet source</a> | <a href='#snippet-WordprocessingDocument' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### PowerPoint


#### Verify a file

<!-- snippet: VerifyPowerpoint -->
<a id='snippet-VerifyPowerpoint'></a>
```cs
[Test]
public Task VerifyPowerpoint() =>
    VerifyFile("sample.pptx")
        .Snapshot(
            """
            {
              Properties: {
                Title: Sample Presentation
              },
              SlideCount: 1
            }
            """);
```
<sup><a href='/src/Verify.OpenXml.Tests/Samples.cs#L106-L121' title='Snippet source file'>snippet source</a> | <a href='#snippet-VerifyPowerpoint' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Verify a Stream

<!-- snippet: VerifyPowerpointStream -->
<a id='snippet-VerifyPowerpointStream'></a>
```cs
[Test]
public Task VerifyPowerpointStream()
{
    var stream = new MemoryStream(File.ReadAllBytes("sample.pptx"));
    return Verify(stream, "pptx")
        .Snapshot(
            """
            {
              Properties: {
                Title: Sample Presentation
              },
              SlideCount: 1
            }
            """);
}
```
<sup><a href='/src/Verify.OpenXml.Tests/Samples.cs#L190-L208' title='Snippet source file'>snippet source</a> | <a href='#snippet-VerifyPowerpointStream' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Verify a PresentationDocument

<!-- snippet: PresentationDocument -->
<a id='snippet-PresentationDocument'></a>
```cs
[Test]
public async Task VerifyPresentationDocument()
{
    await using var stream = File.OpenRead("sample.pptx");
    using var reader = PresentationDocument.Open(stream, false);
    await Verify(reader)
        .Snapshot(
            """
            {
              Properties: {
                Title: Sample Presentation
              },
              SlideCount: 1
            }
            """);
}
```
<sup><a href='/src/Verify.OpenXml.Tests/Samples.cs#L169-L188' title='Snippet source file'>snippet source</a> | <a href='#snippet-PresentationDocument' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


## Render pages to PNG (opt-in)

Verify.OpenXml can additionally snapshot a rendered PNG of every page of a `.docx`, `.xlsx`, or `.pptx` using the [Morph](https://github.com/SimonCropp/Morph) renderer. This catches visual regressions (layout, fonts, images, tables) that the text-based snapshot would miss.


### Enabling rendering

The base [`Morph`](https://nuget.org/packages/Morph) package is referenced automatically by Verify.OpenXml on `net10.0`. To turn on rendering, add **exactly one** backend package to the test project:

[`Morph.Skia`](https://nuget.org/packages/Morph.Skia) — uses [SkiaSharp](https://github.com/mono/SkiaSharp):

```xml
<PackageReference Include="Morph.Skia" />
```

or [`Morph.ImageSharp`](https://nuget.org/packages/Morph.ImageSharp) — uses [ImageSharp](https://github.com/SixLabors/ImageSharp), fully managed:

```xml
<PackageReference Include="Morph.ImageSharp" />
```

The backend is detected at runtime by probing for the assembly. No code changes are needed in `ModuleInitializer.cs` — the existing `VerifyOpenXml.Initialize()` call picks it up automatically.


### Output

When a backend is present, every verification (file, stream, or document object) produces additional PNG targets — one per rendered page — alongside the existing binary and text targets. What counts as a page differs per document type:

 * **Word** - one page per laid-out page of the document.
 * **PowerPoint** - one page per slide, in `p:sldIdLst` order.
 * **Excel** - pages come from the print layout rather than the sheet: a long sheet paginates downward, and each visible sheet starts a new page with its own paper size and orientation.

A single rendered page is written without an index:

```
Samples.VerifyWord.verified.docx
Samples.VerifyWord#00.verified.txt
Samples.VerifyWord#01.verified.txt
Samples.VerifyWord.verified.png
```

Multiple pages are indexed in page order. For example a two-sheet workbook:

```
Samples.MultipleSheets.verified.xlsx
Samples.MultipleSheets.verified.txt
Samples.MultipleSheets#Sheet1.verified.csv
Samples.MultipleSheets#Sheet2.verified.csv
Samples.MultipleSheets#00.verified.png
Samples.MultipleSheets#01.verified.png
```


### Backend selection rules

 * **Neither backend referenced** — rendering is silently skipped. The text and binary targets are still produced. This is the default for consumers who do not opt in.
 * **One backend referenced** — that backend is used for all verifications.
 * **Both backends referenced** — an exception is thrown on the first verification with a clear message. Pick one.


### Target framework support

Rendering is only available on `net10.0` because Morph targets `net10.0` only. On `net472`, `net48`, `net8.0`, and `net9.0`, verification continues to produce only the existing text and binary targets — the rendering code is conditionally compiled out.


### Cross-platform PNG stability

PNG output from Skia and ImageSharp depends on installed fonts and platform-specific rasterization. A `.verified.png` generated on one machine may not be byte-identical on another OS or with different fonts installed. Recommendations:

 * Generate and commit `.verified.png` files from a single canonical machine (often a CI agent).
 * For cross-platform CI, combine with `UniqueForOSPlatform()` so each OS gets its own `.verified.png`:

```cs
await Verify(stream, "docx")
    .UniqueForOSPlatform();
```

 * Consider [PNG SSIM comparer](https://github.com/VerifyTests/Verify/blob/main/docs/comparer.md#png-ssim-comparer) for tolerance-based image diffing.

See [Verify Naming docs](https://github.com/VerifyTests/Verify/blob/main/docs/naming.md) for the full list of `UniqueFor*` modifiers.


### Sharing one test suite across both backends

For a worked example of running the same test suite against both backends side-by-side, see the [`Tests.Skia`](/src/Tests.Skia) and [`Tests.ImageSharp`](/src/Tests.ImageSharp) projects in this repository. Both projects link the source files from [`Tests`](/src/Tests) and use `DerivePathInfo` in their `ModuleInitializer` to redirect snapshots into the per-backend project directory:

```cs
[ModuleInitializer]
public static void Initialize()
{
    VerifyOpenXml.Initialize();

    var projectDir = ProjectDir();
    Verifier.DerivePathInfo(
        (sourceFile, projectDirectory, type, method) =>
            new(directory: projectDir, typeName: type.Name, methodName: method.Name));
}

static string ProjectDir([CallerFilePath] string here = "") =>
    Path.GetDirectoryName(here)!;
```

This pattern lets a single set of tests produce two parallel sets of `.verified.*` snapshots — one per rendering backend.


## Exclude the document

The source document is included in the snapshot as a `.verified.xlsx`, `.verified.docx`, or `.verified.pptx`. Building the deterministic package is expensive, and committing it is not always wanted. [`ExcludeTargets`](https://github.com/VerifyTests/Verify/blob/main/docs/converter.md#excluding-targets) drops it from a verification and skips the build, while the info, text, csv, and rendered pages still verify:

<!-- snippet: ExcludeExcel -->
<a id='snippet-ExcludeExcel'></a>
```cs
// Skips the .verified.xlsx (and building it), keeping the info and csv sheets.
[Test]
public Task ExcludeExcel() =>
    VerifyFile("sample.xlsx")
        .ExcludeTargets("xlsx");
```
<sup><a href='/src/Verify.OpenXml.Tests/Samples.cs#L123-L131' title='Snippet source file'>snippet source</a> | <a href='#snippet-ExcludeExcel' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

The same applies to `docx` and `pptx`. To exclude for every test, call `VerifierSettings.ExcludeTargets("xlsx")` at initialization.
