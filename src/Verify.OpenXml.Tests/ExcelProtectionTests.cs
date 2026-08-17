[TestFixture]
public class ExcelProtectionTests
{
    [Test]
    public Task Protection()
    {
        using var document = CreateDocument();
        return Verify(document);
    }

    static SpreadsheetDocument CreateDocument()
    {
        var document = SpreadsheetDocument.Create(new MemoryStream(), SpreadsheetDocumentType.Workbook);

        var wbPart = document.AddWorkbookPart();
        wbPart.Workbook = new(
            new WorkbookProtection
            {
                WorkbookPassword = new("DAA7"),
                LockStructure = true,
                LockWindows = false
            },
            new Sheets());

        var wsPart = wbPart.AddNewPart<WorksheetPart>();
        var sheetData = new SheetData(
            new Row(
                new Cell
                {
                    DataType = CellValues.InlineString,
                    InlineString = new(new Text("Header"))
                })
            {
                RowIndex = 1u
            });

        var sheetProtection = new SheetProtection
        {
            Sheet = true,
            Password = new("DAA7"),
            Objects = true,
            Scenarios = true,
            FormatCells = true,
            FormatColumns = true,
            FormatRows = true,
            InsertColumns = true,
            InsertRows = true,
            InsertHyperlinks = true,
            DeleteColumns = true,
            DeleteRows = true,
            SelectLockedCells = false,
            SelectUnlockedCells = false,
            Sort = false,
            AutoFilter = false,
            PivotTables = true
        };
        wsPart.Worksheet = new(
            sheetData,
            sheetProtection,
            // A4. A sheet stating no paper size takes the renderer's region default — Letter in
            // North America, A4 elsewhere — so the rendered page snapshot would depend on where
            // the test ran rather than on the workbook.
            new PageSetup
            {
                PaperSize = 9
            });

        var sheets = wbPart.Workbook.GetFirstChild<Sheets>()!;
        sheets.Append(new Sheet
        {
            Id = wbPart.GetIdOfPart(wsPart),
            SheetId = 1,
            Name = "Sheet1"
        });

        return document;
    }
}
