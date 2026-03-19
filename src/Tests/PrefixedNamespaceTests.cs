using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

// Verifies that xlsx files created by the OpenXml SDK (which uses prefixed
// default namespaces like <x:worksheet xmlns:x="...">) are handled correctly.
[TestFixture]
public class PrefixedNamespaceTests
{
    [Test]
    public async Task OpenXmlSdkSpreadsheet()
    {
        using var document = CreateSpreadsheet();
        await Verify(document);
    }

    static SpreadsheetDocument CreateSpreadsheet()
    {
        var stream = new MemoryStream();
        var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);

        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new(new Sheets());

        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new(new SheetData());

        var sheets = workbookPart.Workbook.GetFirstChild<Sheets>()!;
        sheets.Append(new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = 1,
            Name = "Sheet1"
        });

        var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>()!;

        var row = new Row { RowIndex = 1 };
        row.Append(new Cell
        {
            CellReference = "A1",
            DataType = CellValues.InlineString,
            InlineString = new(new Text("Hello"))
        });
        row.Append(new Cell
        {
            CellReference = "B1",
            DataType = CellValues.InlineString,
            InlineString = new(new Text("World"))
        });
        sheetData.Append(row);

        return document;
    }
}
