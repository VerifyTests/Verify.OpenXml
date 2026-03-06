# <img src="/src/icon.png" height="30px"> Verify.OpenXML

[![Discussions](https://img.shields.io/badge/Verify-Discussions-yellow?svg=true&label=)](https://github.com/orgs/VerifyTests/discussions)
[![Build status](https://img.shields.io/appveyor/build/SimonCropp/verify-openxml)](https://ci.appveyor.com/project/SimonCropp/verify-openxml)
[![NuGet Status](https://img.shields.io/nuget/v/Verify.OpenXML.svg)](https://www.nuget.org/packages/Verify.OpenXML/)

Extends [Verify](https://github.com/VerifyTests/Verify) to allow verification of Word and Excel documents via [OpenXML](https://github.com/dotnet/Open-XML-SDK/).<!-- singleLineInclude: intro. path: /docs/intro.include.md -->

Supports Excel (xlsx) and Word (docx) documents.

## Features

### Excel (xlsx)

 * Converts workbooks to CSV format for each worksheet
 * Extracts formulas and displays them alongside cell values
 * Captures document metadata (title, subject, creator, keywords, category, etc.)
 * Supports date scrubbing and GUID scrubbing for deterministic tests
 * Generates deterministic XLSX output using DeterministicIoPackaging

### Word (docx)

 * Extracts document text content from paragraphs and tables
 * Captures document properties (title, subject, creator, keywords, etc.)
 * Captures custom document properties
 * Extracts font information
 * Generates deterministic DOCX output using DeterministicIoPackaging

**See [Milestones](../../milestones?state=closed) for release notes.**

 

## Sponsors


### Entity Framework Extensions<!-- include: sponsors. path: /docs/sponsors.include.md -->

[Entity Framework Extensions](https://entityframework-extensions.net/?utm_source=simoncropp&utm_medium=Verify.OpenXML) is a major sponsor and is proud to contribute to the development this project.

[![Entity Framework Extensions](https://raw.githubusercontent.com/VerifyTests/Verify.OpenXML/refs/heads/main/docs/zzz.png)](https://entityframework-extensions.net/?utm_source=simoncropp&utm_medium=Verify.OpenXML)

### Developed using JetBrains IDEs

[![JetBrains logo.](https://resources.jetbrains.com/storage/products/company/brand/logos/jetbrains.svg)](https://jb.gg/OpenSourceSupport)<!-- endInclude -->


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
<sup><a href='/src/Tests/ModuleInitializer.cs#L3-L9' title='Snippet source file'>snippet source</a> | <a href='#snippet-enable' title='Start of snippet'>anchor</a></sup>
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
<sup><a href='/src/Tests/Samples.cs#L4-L10' title='Snippet source file'>snippet source</a> | <a href='#snippet-VerifyExcel' title='Start of snippet'>anchor</a></sup>
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
<sup><a href='/src/Tests/Samples.cs#L33-L42' title='Snippet source file'>snippet source</a> | <a href='#snippet-VerifyExcelStream' title='Start of snippet'>anchor</a></sup>
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
<sup><a href='/src/Tests/Samples.cs#L21-L31' title='Snippet source file'>snippet source</a> | <a href='#snippet-SpreadsheetDocument' title='Start of snippet'>anchor</a></sup>
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
<sup><a href='/src/Tests/Samples.VerifyExcel.verified.csv#L1-L7' title='Snippet source file'>snippet source</a> | <a href='#snippet-Samples.VerifyExcel.verified.csv' title='Start of snippet'>anchor</a></sup>
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
<sup><a href='/src/Tests/Samples.cs#L44-L50' title='Snippet source file'>snippet source</a> | <a href='#snippet-VerifyWord' title='Start of snippet'>anchor</a></sup>
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
<sup><a href='/src/Tests/Samples.cs#L64-L73' title='Snippet source file'>snippet source</a> | <a href='#snippet-VerifyWordStream' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


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
<sup><a href='/src/Tests/Samples.cs#L52-L62' title='Snippet source file'>snippet source</a> | <a href='#snippet-WordprocessingDocument' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->
