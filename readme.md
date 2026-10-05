# <img src="/src/icon.png" height="30px"> Verify.OpenXML

[![Discussions](https://img.shields.io/badge/Verify-Discussions-yellow?svg=true&label=)](https://github.com/orgs/VerifyTests/discussions)
[![Build status](https://github.com/VerifyTests/Verify.OpenXml/actions/workflows/build.yml/badge.svg)](https://github.com/VerifyTests/Verify.OpenXml/actions/workflows/build.yml)
[![NuGet Status](https://img.shields.io/nuget/v/Verify.OpenXML.svg)](https://www.nuget.org/packages/Verify.OpenXML/)

Extends [Verify](https://github.com/VerifyTests/Verify) to allow verification of Word, Excel, and PowerPoint documents via [OpenXML](https://github.com/dotnet/Open-XML-SDK/).<!-- singleLineInclude: intro. path: /docs/intro.include.md -->


## Features


### Excel (xlsx)

 * Converts each worksheet to CSV, in a file named for the worksheet
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

 * Extracts the text of each slide, as the text of its page
 * Captures document properties (title, subject, keywords, description, category, status, revision)
 * Reports the slide count, as the page count
 * Generates deterministic PPTX output using DeterministicIoPackaging
 * Optionally renders each slide to PNG via [Morph](https://github.com/SimonCropp/Morph) (opt-in)


### Paged documents

How the files of a document are named, where its text goes, and which of them are verified, is decided by Verify's [paged documents](https://github.com/VerifyTests/Verify/blob/main/docs/paged-documents.md) support, which every Verify plugin that splits a document into pages shares. See [The files](#the-files) and [Choosing what is verified](#choosing-what-is-verified).


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


### The files

For a test `Samples.VerifyWord` verifying a docx:

| File | Holds |
| --- | --- |
| `Samples.VerifyWord.verified.docx` | The document, made deterministic |
| `Samples.VerifyWord.verified.txt` | The info file: the properties and fonts of the document, its page count, and its text |
| `Samples.VerifyWord#page_0001.verified.png` | The first page, drawn. Requires a [rendering backend](#render-pages-to-png-opt-in) |

A pptx is verified the same way, with a page for each slide. An xlsx has its sheets in place of text: each is a csv named for the sheet, `Samples.VerifyExcel#Sheet1.verified.csv`, and its info file holds the properties of the workbook, its sheets and their columns.

The info file has the shape every paged document has, with what is read from the document under `Document`. Where a page of a Word document ends is only known once it is laid out. So with a [rendering backend](#render-pages-to-png-opt-in) its text is under the page it is on, as that of a presentation is, and without one it is read as one text:

<!-- snippet: Samples.VerifyWord.verified.txt -->
<a id='snippet-Samples.VerifyWord.verified.txt'></a>
```txt
{
  Document: {
    Properties: {
      Subject: Test Subject,
      Title: Sample Document
    }
  },
  Text:
Hello World! This is a sample Word document.
This is the second paragraph with some more text.
}
```
<sup><a href='/src/Verify.OpenXml.Tests/Samples.VerifyWord.verified.txt#L1-L11' title='Snippet source file'>snippet source</a> | <a href='#snippet-Samples.VerifyWord.verified.txt' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

A presentation is read slide by slide, so the text of a slide is under its page:

<!-- snippet: Samples.VerifyPowerpoint.verified.txt -->
<a id='snippet-Samples.VerifyPowerpoint.verified.txt'></a>
```txt
{
  Document: {
    Properties: {
      Title: Sample Presentation
    }
  },
  PageCount: 1,
  Pages: [
    {
      Number: 1,
      Text: Hello, PowerPoint!
    }
  ]
}
```
<sup><a href='/src/Verify.OpenXml.Tests/Samples.VerifyPowerpoint.verified.txt#L1-L14' title='Snippet source file'>snippet source</a> | <a href='#snippet-Samples.VerifyPowerpoint.verified.txt' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

`PageCount` is the number of slides of a presentation. The pages of a Word document and of a workbook only exist once they are laid out, so their `PageCount` is written when the pages are [rendered](#render-pages-to-png-opt-in).


### Choosing what is verified

What a document is split into is controlled by Verify's settings for [paged documents](https://github.com/VerifyTests/Verify/blob/main/docs/paged-documents.md).

`ExcludeDerivedTargets("png")` leaves out the rendered pages, keeping the document and its text:

<!-- snippet: ExcludeRenderedPages -->
<a id='snippet-ExcludeRenderedPages'></a>
```cs
[Test]
public Task ExcludeRenderedPages() =>
    VerifyFile("sample.docx")
        .ExcludeDerivedTargets("png");
```
<sup><a href='/src/Verify.OpenXml.Tests/Samples.cs#L147-L154' title='Snippet source file'>snippet source</a> | <a href='#snippet-ExcludeRenderedPages' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

`ExcludeDerivedTargets("csv")` leaves out the sheets of a workbook the same way, and [`ExcludeTargets`](#exclude-the-document) the document itself.

The text is in the info file by default. `PageText` moves it to a file of its own, or leaves it out with `PageTextPlacement.None`. The file is `#page_0001.verified.txt` for each slide of a presentation and each page of a Word document, or a single `#text.verified.txt` for a Word document that is read as one text. Here with the rendered pages left out as well:

<!-- snippet: PageTextPerPage -->
<a id='snippet-PageTextPerPage'></a>
```cs
[Test]
public Task PageTextPerPage() =>
    VerifyFile("sample.pptx")
        .PageText(PageTextPlacement.PerPage)
        .ExcludeDerivedTargets("png");
```
<sup><a href='/src/Verify.OpenXml.Tests/Samples.cs#L137-L145' title='Snippet source file'>snippet source</a> | <a href='#snippet-PageTextPerPage' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

What these leave out is not produced at all (pages are not rendered, text is not read, sheets are not converted), so they also save work.

`PagesToInclude` limits the pages that are verified, to the first pages of a document or to those a delegate accepts. The document itself is still verified whole, and `PageCount` is still the number of pages it has:

<!-- snippet: PagesToInclude -->
<a id='snippet-PagesToInclude'></a>
```cs
[Test]
public async Task PagesToInclude()
{
    using var presentation = ThreeSlides();
    await Verify(presentation)
        .PagesToInclude(2);
}
```
<sup><a href='/src/Verify.OpenXml.Tests/PowerpointPagesTests.cs#L9-L19' title='Snippet source file'>snippet source</a> | <a href='#snippet-PagesToInclude' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

It applies to pages only: their images, and their text. The sheets of a workbook belong to no page, so they are verified whole, and so is the text of a Word document that is read as one text, for want of a rendering backend. A paragraph that runs over the end of a page is whole in the text of the page it starts on, and a row of a table likewise. A document is also rendered whole, so the pages left out are still drawn before they are dropped.

Each can also be set for every test, on `VerifierSettings`:

<!-- snippet: InitializeOutputs -->
<a id='snippet-InitializeOutputs'></a>
```cs
[ModuleInitializer]
public static void Initialize()
{
    VerifyOpenXml.Initialize();

    // For every test: no text and no rendered pages
    VerifierSettings.PageText(PageTextPlacement.None);
    VerifierSettings.ExcludeDerivedTargets("png");
}
```
<sup><a href='/src/StaticSettingsTests/ModuleInitializer.cs#L3-L15' title='Snippet source file'>snippet source</a> | <a href='#snippet-InitializeOutputs' title='Start of snippet'>anchor</a></sup>
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
<sup><a href='/src/Verify.OpenXml.Tests/Samples.cs#L3-L9' title='Snippet source file'>snippet source</a> | <a href='#snippet-VerifyExcel' title='Start of snippet'>anchor</a></sup>
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
<sup><a href='/src/Verify.OpenXml.Tests/Samples.cs#L36-L45' title='Snippet source file'>snippet source</a> | <a href='#snippet-VerifyExcelStream' title='Start of snippet'>anchor</a></sup>
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
<sup><a href='/src/Verify.OpenXml.Tests/Samples.cs#L24-L34' title='Snippet source file'>snippet source</a> | <a href='#snippet-SpreadsheetDocument' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Example snapshot

The sheet `Sheet1`, as `Samples.VerifyExcel#Sheet1.verified.csv`:

<!-- snippet: Samples.VerifyExcel#Sheet1.verified.csv -->
<a id='snippet-Samples.VerifyExcel#Sheet1.verified.csv'></a>
```csv
0,First Name,Last Name,Gender,Country,Date,Age,Id,Formula
1,Dulce,Abril,Female,United States,2017-10-15,32,1562,G2+H21594 (G2+H2)
2,Mara,Hashimoto,Female,Great Britain,2016-08-16,25,1582,1607
3,Philip,Gent,Male,France,2015-05-21,36,2587,2623
4,Kathleen,Hanner,Female,United States,2017-10-15,25,3549,3574
5,Nereida,Magwood,Female,United States,2016-08-16,58,2468,2526
6,Gaston,Brumm,Male,United States,2015-05-21,24,2554,2578
```
<sup><a href='/src/Verify.OpenXml.Tests/Samples.VerifyExcel%23Sheet1.verified.csv#L1-L7' title='Snippet source file'>snippet source</a> | <a href='#snippet-Samples.VerifyExcel#Sheet1.verified.csv' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### Word


#### Verify a file

<!-- snippet: VerifyWord -->
<a id='snippet-VerifyWord'></a>
```cs
[Test]
public Task VerifyWord() =>
    VerifyFile("sample.docx");
```
<sup><a href='/src/Verify.OpenXml.Tests/Samples.cs#L47-L53' title='Snippet source file'>snippet source</a> | <a href='#snippet-VerifyWord' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Verify a Stream

<!-- snippet: VerifyWordStream -->
<a id='snippet-VerifyWordStream'></a>
```cs
[Test]
public Task VerifyWordStream()
{
    var stream = new MemoryStream(File.ReadAllBytes("sample.docx"));
    return Verify(stream, "docx");
}
```
<sup><a href='/src/Verify.OpenXml.Tests/Samples.cs#L67-L76' title='Snippet source file'>snippet source</a> | <a href='#snippet-VerifyWordStream' title='Start of snippet'>anchor</a></sup>
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
    await Verify(reader);
}
```
<sup><a href='/src/Verify.OpenXml.Tests/Samples.cs#L55-L65' title='Snippet source file'>snippet source</a> | <a href='#snippet-WordprocessingDocument' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### PowerPoint


#### Verify a file

<!-- snippet: VerifyPowerpoint -->
<a id='snippet-VerifyPowerpoint'></a>
```cs
[Test]
public Task VerifyPowerpoint() =>
    VerifyFile("sample.pptx");
```
<sup><a href='/src/Verify.OpenXml.Tests/Samples.cs#L78-L84' title='Snippet source file'>snippet source</a> | <a href='#snippet-VerifyPowerpoint' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Verify a Stream

<!-- snippet: VerifyPowerpointStream -->
<a id='snippet-VerifyPowerpointStream'></a>
```cs
[Test]
public Task VerifyPowerpointStream()
{
    var stream = new MemoryStream(File.ReadAllBytes("sample.pptx"));
    return Verify(stream, "pptx");
}
```
<sup><a href='/src/Verify.OpenXml.Tests/Samples.cs#L126-L135' title='Snippet source file'>snippet source</a> | <a href='#snippet-VerifyPowerpointStream' title='Start of snippet'>anchor</a></sup>
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
    await Verify(reader);
}
```
<sup><a href='/src/Verify.OpenXml.Tests/Samples.cs#L114-L124' title='Snippet source file'>snippet source</a> | <a href='#snippet-PresentationDocument' title='Start of snippet'>anchor</a></sup>
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
 * **Excel** - one page per visible sheet, drawn whole as the one image, however long the sheet is and whatever paper its page setup names. A hidden sheet has no page.

A page is named for its number, counted from 1, whether the document has one page or several:

```
Samples.VerifyWord.verified.docx
Samples.VerifyWord.verified.txt
Samples.VerifyWord#page_0001.verified.png
```

For example a two-sheet workbook, where each sheet is a page:

```
Samples.MultipleSheets.verified.xlsx
Samples.MultipleSheets.verified.txt
Samples.MultipleSheets#Sheet1.verified.csv
Samples.MultipleSheets#Sheet2.verified.csv
Samples.MultipleSheets#page_0001.verified.png
Samples.MultipleSheets#page_0002.verified.png
```

Rendering also puts the `PageCount` of a Word document or a workbook in its info file.

`ExcludeDerivedTargets("png")` [leaves the pages out](#choosing-what-is-verified), for one verification or for every test, and nothing is rendered.


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
<sup><a href='/src/Verify.OpenXml.Tests/Samples.cs#L86-L94' title='Snippet source file'>snippet source</a> | <a href='#snippet-ExcludeExcel' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

The same applies to `docx` and `pptx`. To exclude for every test, call `VerifierSettings.ExcludeTargets("xlsx")` at initialization.


## Reviewing changes

A change to a document is a change to several files: the document, its info file, and every page and sheet. Verify tells the diff tool that the pages, the sheets and the info file were derived from the document, and [DiffEngineViewer](https://github.com/VerifyTests/DiffEngine/blob/main/docs/viewer.md#files-derived-from-a-document), which draws the pages of a Word, Excel or PowerPoint document itself, shows them as one row and accepts them together. Other diff tools are given each file, as before.

When the document has changed, its pages, sheets and info file are compared exactly, skipping any [comparer](https://github.com/VerifyTests/Verify/blob/main/docs/comparer.md) registered for them.


## Migrating from 1.x

Version 2 moves to the paged document support in Verify 33.3. Pages and single sheets are renamed, the text moves into the info file, and `OpenXmlOutputs` gives way to Verify's own settings.


### Settings

`Initialize` no longer takes an `OpenXmlOutputs`. Each output that could be left out of it is now left out by a setting of Verify, for every test on `VerifierSettings` or for a single verification:

| 1.x: not in `OpenXmlOutputs` | 2.x |
| --- | --- |
| `Png` | `ExcludeDerivedTargets("png")` |
| `Text` | `PageText(PageTextPlacement.None)` |
| `Csv` | `ExcludeDerivedTargets("csv")` |

So `Initialize(OpenXmlOutputs.Csv)` becomes:

```cs
VerifyOpenXml.Initialize();
VerifierSettings.PageText(PageTextPlacement.None);
VerifierSettings.ExcludeDerivedTargets("png");
```


### Files

For a test `Tests.Report`:

| 1.x | 2.x |
| --- | --- |
| `Tests.Report.verified.png`, the page of a document with one | `Tests.Report#page_0001.verified.png` |
| `Tests.Report#00.verified.png`, `#01`, the pages of a document with several | `Tests.Report#page_0001.verified.png`, `#page_0002` |
| `Tests.Report.verified.csv`, the sheet of a workbook with one | `Tests.Report#Sheet1.verified.csv`, by the name of the sheet |
| `Tests.Report#Sheet1.verified.csv`, a sheet of a workbook with several | The same |
| `Tests.Report#00.verified.txt`, the info of a docx or pptx | `Tests.Report.verified.txt` |
| `Tests.Report#01.verified.txt`, the text of a docx or pptx | In the info file, or with `PageText(PageTextPlacement.PerPage)` in `#page_0001.verified.txt` for each slide of a pptx and each page of a docx, or `#text.verified.txt` for a docx read without a rendering backend |
| `Tests.Report.verified.txt`, the info of an xlsx | The same |
| `Tests.Report.verified.docx`, `.xlsx`, `.pptx` | The same |

A renamed snapshot shows as a new file and a pending delete. Accepting both, or running once with [AutoVerify](https://github.com/VerifyTests/Verify/blob/main/docs/autoverify.md), moves a test over. The content of a page and of a sheet is unchanged, so source control shows each as a rename.


### The info file

What was at the top of the info file is now under `Document`, with the page count and the text beside it. For a docx:

```
{                                    {
  Properties: {                        Document: {
    Title: Sample Document               Properties: {
  },                                       Title: Sample Document
  Fonts: [                               },
    Aptos                                Fonts: [
  ]                                        Aptos
}                                        ]
                                       },
                                       PageCount: 1,
                                       Text: The text of the document
                                     }
```

For an xlsx the properties of the workbook, with its `Sheets`, move under `Document` the same way.

For a pptx `SlideCount` is now `PageCount`, and the text is the `Text` of each page, where it was one text with `---` between the slides. Slides are read in the order they are shown in, that of `p:sldIdLst`, so that the text of a slide is on the page it is drawn on. They were read in the order of the slide parts, which is not the order of a deck whose slides have been moved.
