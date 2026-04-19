[TestFixture]
public class ExcelUnitTests
{

    [Test]
    public void EscapeCsvValue_NoSpecial() =>
        Assert.That(VerifyOpenXml.EscapeCsvValue("plain"), Is.EqualTo("plain"));

    [Test]
    public void EscapeCsvValue_Comma() =>
        Assert.That(VerifyOpenXml.EscapeCsvValue("a,b"), Is.EqualTo("\"a,b\""));

    [Test]
    public void EscapeCsvValue_Quote() =>
        Assert.That(VerifyOpenXml.EscapeCsvValue("say \"hi\""), Is.EqualTo("\"say \"\"hi\"\"\""));

    [Test]
    public void EscapeCsvValue_Newline() =>
        Assert.That(VerifyOpenXml.EscapeCsvValue("a\nb"), Is.EqualTo("\"a\nb\""));

    [Test]
    public void EscapeCsvValue_CarriageReturn() =>
        Assert.That(VerifyOpenXml.EscapeCsvValue("a\rb"), Is.EqualTo("\"a\rb\""));

    [Test]
    public void GetHeaderCellValue_SharedString()
    {
        var shared = new List<SharedStringItem>
        {
            new(new Text("First")),
            new(new Text("Second"))
        };
        var cell = new Cell { DataType = CellValues.SharedString, CellValue = new("1") };
        Assert.That(VerifyOpenXml.GetHeaderCellValue(cell, shared), Is.EqualTo("Second"));
    }

    [Test]
    public void GetHeaderCellValue_InlineString()
    {
        var cell = new Cell
        {
            DataType = CellValues.InlineString,
            InlineString = new(new Text("Inline"))
        };
        Assert.That(VerifyOpenXml.GetHeaderCellValue(cell, null), Is.EqualTo("Inline"));
    }

    [Test]
    public void GetHeaderCellValue_Plain()
    {
        var cell = new Cell { CellValue = new("42") };
        Assert.That(VerifyOpenXml.GetHeaderCellValue(cell, null), Is.EqualTo("42"));
    }

    [Test]
    public void IsCellDateFormatted_NoStyleIndex_False()
    {
        using var doc = CreateWorkbook(addStyles: false);
        var cell = new Cell();
        Assert.That(VerifyOpenXml.IsCellDateFormatted(cell, doc.WorkbookPart!), Is.False);
    }

    [Test]
    public void IsCellDateFormatted_NoStylesPart_False()
    {
        using var doc = CreateWorkbook(addStyles: false);
        var cell = new Cell { StyleIndex = 0 };
        Assert.That(VerifyOpenXml.IsCellDateFormatted(cell, doc.WorkbookPart!), Is.False);
    }

    [Test]
    public void IsCellDateFormatted_BuiltInRange1()
    {
        using var doc = CreateWorkbookWithFormats(14);
        var cell = new Cell { StyleIndex = 0 };
        Assert.That(VerifyOpenXml.IsCellDateFormatted(cell, doc.WorkbookPart!), Is.True);
    }

    [Test]
    public void IsCellDateFormatted_BuiltInRange2()
    {
        using var doc = CreateWorkbookWithFormats(177);
        var cell = new Cell { StyleIndex = 0 };
        Assert.That(VerifyOpenXml.IsCellDateFormatted(cell, doc.WorkbookPart!), Is.True);
    }

    [Test]
    public void IsCellDateFormatted_BuiltInRange3()
    {
        using var doc = CreateWorkbookWithFormats(182);
        var cell = new Cell { StyleIndex = 0 };
        Assert.That(VerifyOpenXml.IsCellDateFormatted(cell, doc.WorkbookPart!), Is.True);
    }

    [Test]
    public void IsCellDateFormatted_CustomDateFormat()
    {
        using var doc = CreateWorkbookWithFormats(200, customFormatCode: "yyyy-mm-dd");
        var cell = new Cell { StyleIndex = 0 };
        Assert.That(VerifyOpenXml.IsCellDateFormatted(cell, doc.WorkbookPart!), Is.True);
    }

    [Test]
    public void IsCellDateFormatted_CustomNonDateFormat()
    {
        using var doc = CreateWorkbookWithFormats(201, customFormatCode: "0.00");
        var cell = new Cell { StyleIndex = 0 };
        Assert.That(VerifyOpenXml.IsCellDateFormatted(cell, doc.WorkbookPart!), Is.False);
    }

    [Test]
    public void IsCellDateFormatted_UnknownFormatId_False()
    {
        using var doc = CreateWorkbookWithFormats(500); // not built-in, not in numberingFormats
        var cell = new Cell { StyleIndex = 0 };
        Assert.That(VerifyOpenXml.IsCellDateFormatted(cell, doc.WorkbookPart!), Is.False);
    }

    [Test]
    public void GetColumnInfos_NoRows_ReturnsNull()
    {
        using var doc = SpreadsheetDocument.Create(new MemoryStream(), SpreadsheetDocumentType.Workbook);
        var wbPart = doc.AddWorkbookPart();
        wbPart.Workbook = new(new Sheets());
        var wsPart = wbPart.AddNewPart<WorksheetPart>();
        wsPart.Worksheet = new(new SheetData());
        Assert.That(VerifyOpenXml.GetColumnInfos(wsPart, wbPart), Is.Null);
    }

    [Test]
    public void GetColumnInfos_WithRowsAndCustomWidths()
    {
        using var doc = SpreadsheetDocument.Create(new MemoryStream(), SpreadsheetDocumentType.Workbook);
        var wbPart = doc.AddWorkbookPart();
        wbPart.Workbook = new(new Sheets());
        var wsPart = wbPart.AddNewPart<WorksheetPart>();

        var sheetData = new SheetData(
            new Row(
                new Cell { DataType = CellValues.InlineString, InlineString = new(new Text("Name")) },
                new Cell { DataType = CellValues.InlineString, InlineString = new(new Text("Age")) })
            {
                RowIndex = 1u
            });

        var columns = new Columns(
            new Column { Min = 1, Max = 1, Width = 20.5, CustomWidth = true },
            new Column { Min = 2, Max = 2, Width = 10.123, CustomWidth = false });

        wsPart.Worksheet = new(columns, sheetData);

        var result = VerifyOpenXml.GetColumnInfos(wsPart, wbPart)!;
        Assert.That(result, Has.Count.EqualTo(2));
        Assert.That(result[0].Name, Is.EqualTo("Name"));
        Assert.That(result[0].Width, Is.EqualTo(20.5));
        Assert.That(result[1].Name, Is.EqualTo("Age"));
        Assert.That(result[1].Width, Is.Null);
    }

    [Test]
    public void GetColumnInfos_RichText_SharedString()
    {
        using var doc = SpreadsheetDocument.Create(new MemoryStream(), SpreadsheetDocumentType.Workbook);
        var wbPart = doc.AddWorkbookPart();
        wbPart.Workbook = new(new Sheets());

        var sharedStringPart = wbPart.AddNewPart<SharedStringTablePart>();
        sharedStringPart.SharedStringTable = new(
            new SharedStringItem(new Text("Plain")),
            new SharedStringItem(
                new DocumentFormat.OpenXml.Spreadsheet.Run(new Text("Rich")),
                new DocumentFormat.OpenXml.Spreadsheet.Run(new Text("Text"))));

        var wsPart = wbPart.AddNewPart<WorksheetPart>();
        var sheetData = new SheetData(
            new Row(
                new Cell { DataType = CellValues.InlineString, InlineString = new(new Text("ColA")), CellReference = "A1" },
                new Cell { DataType = CellValues.InlineString, InlineString = new(new Text("ColB")), CellReference = "B1" })
            {
                RowIndex = 1u
            },
            new Row(
                new Cell { DataType = CellValues.SharedString, CellValue = new("0"), CellReference = "A2" },
                new Cell { DataType = CellValues.SharedString, CellValue = new("1"), CellReference = "B2" })
            {
                RowIndex = 2u
            });
        wsPart.Worksheet = new(sheetData);

        var result = VerifyOpenXml.GetColumnInfos(wsPart, wbPart)!;
        Assert.That(result[0].ContainsRichText, Is.False);
        Assert.That(result[1].ContainsRichText, Is.True);
    }

    [Test]
    public void GetColumnInfos_RichText_InlineString()
    {
        using var doc = SpreadsheetDocument.Create(new MemoryStream(), SpreadsheetDocumentType.Workbook);
        var wbPart = doc.AddWorkbookPart();
        wbPart.Workbook = new(new Sheets());

        var wsPart = wbPart.AddNewPart<WorksheetPart>();
        var sheetData = new SheetData(
            new Row(
                new Cell { DataType = CellValues.InlineString, InlineString = new(new Text("Col")), CellReference = "A1" })
            {
                RowIndex = 1u
            },
            new Row(
                new Cell
                {
                    DataType = CellValues.InlineString,
                    InlineString = new(new DocumentFormat.OpenXml.Spreadsheet.Run(new Text("Styled"))),
                    CellReference = "A2"
                })
            {
                RowIndex = 2u
            });
        wsPart.Worksheet = new(sheetData);

        var result = VerifyOpenXml.GetColumnInfos(wsPart, wbPart)!;
        Assert.That(result[0].ContainsRichText, Is.True);
    }

    [Test]
    public void GetColumnInfos_RichText_HeaderRowIgnored()
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
                    InlineString = new(new DocumentFormat.OpenXml.Spreadsheet.Run(new Text("HeaderRich"))),
                    CellReference = "A1"
                })
            {
                RowIndex = 1u
            });
        wsPart.Worksheet = new(sheetData);

        var result = VerifyOpenXml.GetColumnInfos(wsPart, wbPart)!;
        Assert.That(result[0].ContainsRichText, Is.False);
    }

    [Test]
    public void BuildSheetInfos_MultipleSheets()
    {
        using var doc = SpreadsheetDocument.Create(new MemoryStream(), SpreadsheetDocumentType.Workbook);
        var wbPart = doc.AddWorkbookPart();
        wbPart.Workbook = new(new Sheets());
        var sheets = wbPart.Workbook.GetFirstChild<Sheets>()!;

        AddSheet(wbPart, sheets, "Alpha", 1);
        AddSheet(wbPart, sheets, "Beta", 2);

        var infos = VerifyOpenXml.BuildSheetInfos(wbPart);
        Assert.That(infos.Select(_ => _.Name), Is.EqualTo(["Alpha", "Beta"]));
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
            new CellFormat { NumberFormatId = numFormatId }));

        if (customFormatCode != null)
        {
            stylesPart.Stylesheet.Append(new NumberingFormats(
                new NumberingFormat { NumberFormatId = numFormatId, FormatCode = customFormatCode }));
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
