public class ExcelUnitTests
{
    [Test]
    public async Task EscapeCsvValue_NoSpecial() =>
        await Assert.That(VerifyOpenXml.EscapeCsvValue("plain")).IsEqualTo("plain");

    [Test]
    public async Task EscapeCsvValue_Comma() =>
        await Assert.That(VerifyOpenXml.EscapeCsvValue("a,b")).IsEqualTo("\"a,b\"");

    [Test]
    public async Task EscapeCsvValue_Quote() =>
        await Assert.That(VerifyOpenXml.EscapeCsvValue("say \"hi\"")).IsEqualTo("\"say \"\"hi\"\"\"");

    [Test]
    public async Task EscapeCsvValue_Newline() =>
        await Assert.That(VerifyOpenXml.EscapeCsvValue("a\nb")).IsEqualTo("\"a\nb\"");

    [Test]
    public async Task EscapeCsvValue_CarriageReturn() =>
        await Assert.That(VerifyOpenXml.EscapeCsvValue("a\rb")).IsEqualTo("\"a\rb\"");

    [Test]
    public async Task GetHeaderCellValue_SharedString()
    {
        var shared = new List<SharedStringItem>
        {
            new(new Text("First")),
            new(new Text("Second"))
        };
        var cell = new Cell
        {
            DataType = CellValues.SharedString,
            CellValue = new("1")
        };
        await Assert.That(VerifyOpenXml.GetHeaderCellValue(cell, shared)).IsEqualTo("Second");
    }

    [Test]
    public async Task GetHeaderCellValue_InlineString()
    {
        var cell = new Cell
        {
            DataType = CellValues.InlineString,
            InlineString = new(new Text("Inline"))
        };
        await Assert.That(VerifyOpenXml.GetHeaderCellValue(cell, null)).IsEqualTo("Inline");
    }

    [Test]
    public async Task GetHeaderCellValue_Plain()
    {
        var cell = new Cell
        {
            CellValue = new("42")
        };
        await Assert.That(VerifyOpenXml.GetHeaderCellValue(cell, null)).IsEqualTo("42");
    }

    [Test]
    public async Task IsCellDateFormatted_NoStyleIndex_False()
    {
        using var doc = CreateWorkbook(addStyles: false);
        var cell = new Cell();
        await Assert.That(VerifyOpenXml.IsCellDateFormatted(cell, doc.WorkbookPart!)).IsFalse();
    }

    [Test]
    public async Task IsCellDateFormatted_NoStylesPart_False()
    {
        using var doc = CreateWorkbook(addStyles: false);
        var cell = new Cell
        {
            StyleIndex = 0
        };
        await Assert.That(VerifyOpenXml.IsCellDateFormatted(cell, doc.WorkbookPart!)).IsFalse();
    }

    [Test]
    public async Task IsCellDateFormatted_BuiltInRange1()
    {
        using var doc = CreateWorkbookWithFormats(14);
        var cell = new Cell
        {
            StyleIndex = 0
        };
        await Assert.That(VerifyOpenXml.IsCellDateFormatted(cell, doc.WorkbookPart!)).IsTrue();
    }

    [Test]
    public async Task IsCellDateFormatted_BuiltInRange2()
    {
        using var doc = CreateWorkbookWithFormats(177);
        var cell = new Cell
        {
            StyleIndex = 0
        };
        await Assert.That(VerifyOpenXml.IsCellDateFormatted(cell, doc.WorkbookPart!)).IsTrue();
    }

    [Test]
    public async Task IsCellDateFormatted_BuiltInRange3()
    {
        using var doc = CreateWorkbookWithFormats(182);
        var cell = new Cell
        {
            StyleIndex = 0
        };
        await Assert.That(VerifyOpenXml.IsCellDateFormatted(cell, doc.WorkbookPart!)).IsTrue();
    }

    [Test]
    public async Task IsCellDateFormatted_CustomDateFormat()
    {
        using var doc = CreateWorkbookWithFormats(200, customFormatCode: "yyyy-mm-dd");
        var cell = new Cell
        {
            StyleIndex = 0
        };
        await Assert.That(VerifyOpenXml.IsCellDateFormatted(cell, doc.WorkbookPart!)).IsTrue();
    }

    [Test]
    public async Task IsCellDateFormatted_CustomNonDateFormat()
    {
        using var doc = CreateWorkbookWithFormats(201, customFormatCode: "0.00");
        var cell = new Cell
        {
            StyleIndex = 0
        };
        await Assert.That(VerifyOpenXml.IsCellDateFormatted(cell, doc.WorkbookPart!)).IsFalse();
    }

    [Test]
    public async Task IsCellDateFormatted_UnknownFormatId_False()
    {
        using var doc = CreateWorkbookWithFormats(500); // not built-in, not in numberingFormats
        var cell = new Cell
        {
            StyleIndex = 0
        };
        await Assert.That(VerifyOpenXml.IsCellDateFormatted(cell, doc.WorkbookPart!)).IsFalse();
    }

    [Test]
    public async Task GetColumnInfos_NoRows_ReturnsNull()
    {
        using var doc = SpreadsheetDocument.Create(new MemoryStream(), SpreadsheetDocumentType.Workbook);
        var wbPart = doc.AddWorkbookPart();
        wbPart.Workbook = new(new Sheets());
        var wsPart = wbPart.AddNewPart<WorksheetPart>();
        wsPart.Worksheet = new(new SheetData());
        await Assert.That(VerifyOpenXml.GetColumnInfos(wsPart, wbPart)).IsNull();
    }

    [Test]
    public async Task GetColumnInfos_WithRowsAndCustomWidths()
    {
        using var doc = SpreadsheetDocument.Create(new MemoryStream(), SpreadsheetDocumentType.Workbook);
        var wbPart = doc.AddWorkbookPart();
        wbPart.Workbook = new(new Sheets());
        var wsPart = wbPart.AddNewPart<WorksheetPart>();

        var sheetData = new SheetData(
            new Row(
                new Cell
                {
                    DataType = CellValues.InlineString,
                    InlineString = new(new Text("Name"))
                },
                new Cell
                {
                    DataType = CellValues.InlineString,
                    InlineString = new(new Text("Age"))
                })
            {
                RowIndex = 1u
            });

        var columns = new Columns(
            new Column
            {
                Min = 1,
                Max = 1,
                Width = 20.5,
                CustomWidth = true
            },
            new Column
            {
                Min = 2,
                Max = 2,
                Width = 10.123,
                CustomWidth = false
            });

        wsPart.Worksheet = new(columns, sheetData);

        var result = VerifyOpenXml.GetColumnInfos(wsPart, wbPart)!;
        await Assert.That(result).Count().IsEqualTo(2);
        await Assert.That(result[0].Name).IsEqualTo("Name");
        await Assert.That(result[0].Width).IsEqualTo(20.5);
        await Assert.That(result[1].Name).IsEqualTo("Age");
        await Assert.That(result[1].Width).IsNull();
    }

    [Test]
    public async Task GetColumnInfos_SkipsLeadingEmptyRow()
    {
        using var doc = SpreadsheetDocument.Create(new MemoryStream(), SpreadsheetDocumentType.Workbook);
        var wbPart = doc.AddWorkbookPart();
        wbPart.Workbook = new(new Sheets());
        var wsPart = wbPart.AddNewPart<WorksheetPart>();

        var sheetData = new SheetData(
            // A leading hidden row emitted as a cell-less <row/> must not be picked as the header.
            new Row
            {
                RowIndex = 1u,
                Hidden = true
            },
            new Row(
                new Cell
                {
                    DataType = CellValues.InlineString,
                    InlineString = new(new Text("Name"))
                },
                new Cell
                {
                    DataType = CellValues.InlineString,
                    InlineString = new(new Text("Age"))
                })
            {
                RowIndex = 2u
            });
        wsPart.Worksheet = new(sheetData);

        var result = VerifyOpenXml.GetColumnInfos(wsPart, wbPart)!;
        await Assert.That(result).Count().IsEqualTo(2);
        await Assert.That(result[0].Name).IsEqualTo("Name");
        await Assert.That(result[1].Name).IsEqualTo("Age");
    }

    [Test]
    public async Task GetColumnInfos_RichText_SharedString()
    {
        using var doc = SpreadsheetDocument.Create(new MemoryStream(), SpreadsheetDocumentType.Workbook);
        var wbPart = doc.AddWorkbookPart();
        wbPart.Workbook = new(new Sheets());

        var sharedStringPart = wbPart.AddNewPart<SharedStringTablePart>();
        sharedStringPart.SharedStringTable = new(
            new SharedStringItem(new Text("Plain")),
            new SharedStringItem(
                new Run(new Text("Rich")),
                new Run(new Text("Text"))));

        var wsPart = wbPart.AddNewPart<WorksheetPart>();
        var sheetData = new SheetData(
            new Row(
                new Cell
                {
                    DataType = CellValues.InlineString,
                    InlineString = new(new Text("ColA")),
                    CellReference = "A1"
                },
                new Cell
                {
                    DataType = CellValues.InlineString,
                    InlineString = new(new Text("ColB")),
                    CellReference = "B1"
                })
            {
                RowIndex = 1u
            },
            new Row(
                new Cell
                {
                    DataType = CellValues.SharedString,
                    CellValue = new("0"),
                    CellReference = "A2"
                },
                new Cell
                {
                    DataType = CellValues.SharedString,
                    CellValue = new("1"),
                    CellReference = "B2"
                })
            {
                RowIndex = 2u
            });
        wsPart.Worksheet = new(sheetData);

        var result = VerifyOpenXml.GetColumnInfos(wsPart, wbPart)!;
        await Assert.That(result[0].ContainsRichText).IsFalse();
        await Assert.That(result[1].ContainsRichText).IsTrue();
    }

    [Test]
    public async Task GetColumnInfos_RichText_InlineString()
    {
        using var doc = SpreadsheetDocument.Create(new MemoryStream(), SpreadsheetDocumentType.Workbook);
        var wbPart = doc.AddWorkbookPart();
        wbPart.Workbook = new(new Sheets());

        var wsPart = wbPart.AddNewPart<WorksheetPart>();
        var sheetData = new SheetData(
            new Row(
                new Cell
                {
                    DataType = CellValues.InlineString,
                    InlineString = new(new Text("Col")),
                    CellReference = "A1"
                })
            {
                RowIndex = 1u
            },
            new Row(
                new Cell
                {
                    DataType = CellValues.InlineString,
                    InlineString = new(new Run(new Text("Styled"))),
                    CellReference = "A2"
                })
            {
                RowIndex = 2u
            });
        wsPart.Worksheet = new(sheetData);

        var result = VerifyOpenXml.GetColumnInfos(wsPart, wbPart)!;
        await Assert.That(result[0].ContainsRichText).IsTrue();
    }

    [Test]
    public async Task GetColumnInfos_RichText_HeaderRowIgnored()
    {
        using var doc = SpreadsheetDocument.Create(new MemoryStream(), SpreadsheetDocumentType.Workbook);
        var wbPart = doc.AddWorkbookPart();
        wbPart.Workbook = new(new Sheets());

        var wsPart = wbPart.AddNewPart<WorksheetPart>();
        var sheetData = new SheetData(
            new Row(
                new Cell
                {
                    DataType = CellValues.InlineString,
                    InlineString = new(new Run(new Text("HeaderRich"))),
                    CellReference = "A1"
                })
            {
                RowIndex = 1u
            });
        wsPart.Worksheet = new(sheetData);

        var result = VerifyOpenXml.GetColumnInfos(wsPart, wbPart)!;
        await Assert.That(result[0].ContainsRichText).IsFalse();
    }

    [Test]
    public async Task BuildSheetInfos_MultipleSheets()
    {
        using var doc = SpreadsheetDocument.Create(new MemoryStream(), SpreadsheetDocumentType.Workbook);
        var wbPart = doc.AddWorkbookPart();
        wbPart.Workbook = new(new Sheets());
        var sheets = wbPart.Workbook.GetFirstChild<Sheets>()!;

        AddSheet(wbPart, sheets, "Alpha", 1);
        AddSheet(wbPart, sheets, "Beta", 2);

        var infos = VerifyOpenXml.BuildSheetInfos(wbPart);
        await Assert.That(infos.Select(_ => _.Name)).IsEquivalentTo(["Alpha", "Beta"], CollectionOrdering.Matching);
    }

    static SpreadsheetDocument CreateWorkbook(bool addStyles)
    {
        var doc = SpreadsheetDocument.Create(new MemoryStream(), SpreadsheetDocumentType.Workbook);
        var wbPart = doc.AddWorkbookPart();
        wbPart.Workbook = new(new Sheets());
        if (addStyles)
        {
            var stylesPart = wbPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = new();
        }

        return doc;
    }

    static SpreadsheetDocument CreateWorkbookWithFormats(uint numFormatId, string? customFormatCode = null)
    {
        var doc = CreateWorkbook(addStyles: true);
        var stylesPart = doc.WorkbookPart!.WorkbookStylesPart!;
        stylesPart.Stylesheet!.Append(new CellFormats(
            new CellFormat
            {
                NumberFormatId = numFormatId
            }));

        if (customFormatCode != null)
        {
            stylesPart.Stylesheet.Append(
                new NumberingFormats(
                    new NumberingFormat
                    {
                        NumberFormatId = numFormatId,
                        FormatCode = customFormatCode
                    }));
        }

        return doc;
    }

    static void AddSheet(WorkbookPart wbPart, Sheets sheets, string name, uint id)
    {
        var wsPart = wbPart.AddNewPart<WorksheetPart>();
        wsPart.Worksheet = new(new SheetData());
        sheets.Append(new Sheet
        {
            Id = wbPart.GetIdOfPart(wsPart),
            SheetId = id,
            Name = name
        });
    }
}
