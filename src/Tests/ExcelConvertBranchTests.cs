// Full-pipeline Verify tests that build a SpreadsheetDocument in memory to
// exercise branches of VerifyOpenXml.Convert / GetCellValue that aren't reachable
// from sample.xlsx (Boolean cells, Date-typed cells, formula formatting).
[TestFixture]
public class ExcelConvertBranchTests
{
    [Test]
    public Task BooleanAndDateAndFormulaCells()
    {
        using var document = CreateDocument();
        return Verify(document);
    }

    static SpreadsheetDocument CreateDocument()
    {
        var document = SpreadsheetDocument.Create(new MemoryStream(), SpreadsheetDocumentType.Workbook);

        var wbPart = document.AddWorkbookPart();
        wbPart.Workbook = new Workbook(new Sheets());

        var sst = wbPart.AddNewPart<SharedStringTablePart>();
        sst.SharedStringTable = new SharedStringTable(new SharedStringItem(new Text("Header")));

        var stylesPart = wbPart.AddNewPart<WorkbookStylesPart>();
        stylesPart.Stylesheet = new Stylesheet(
            new CellFormats(
                new CellFormat(),
                new CellFormat { NumberFormatId = 14, ApplyNumberFormat = true }));

        var wsPart = wbPart.AddNewPart<WorksheetPart>();
        var sheetData = new SheetData(
            new Row(
                new Cell { DataType = CellValues.SharedString, CellValue = new("0") })
            {
                RowIndex = 1u
            },
            new Row(
                new Cell { DataType = CellValues.Boolean, CellValue = new("1") },
                new Cell { DataType = CellValues.Boolean, CellValue = new("0") },
                new Cell { DataType = CellValues.Date, CellValue = new("45000") },
                new Cell { CellValue = new("45000"), StyleIndex = 1u },
                new Cell
                {
                    CellValue = new("99"),
                    CellFormula = new CellFormula("1+98")
                })
            {
                RowIndex = 2u
            });
        wsPart.Worksheet = new Worksheet(sheetData);

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
