// Exercises ColumnInfo.Note sourced from cell notes (legacy comments) on the header row.
// Pairs the notes with a column-metadata custom XML part so the ApplyColumnMetadata
// carry-through (which rebuilds each ColumnInfo) is covered too.
[TestFixture]
public class ExcelCommentTests
{
    [Test]
    public Task HeaderNotes()
    {
        using var document = CreateDocument();
        return Verify(document);
    }

    static SpreadsheetDocument CreateDocument()
    {
        var document = SpreadsheetDocument.Create(new MemoryStream(), SpreadsheetDocumentType.Workbook);
        var wbPart = document.AddWorkbookPart();
        wbPart.Workbook = new(new Sheets());

        var wsPart = wbPart.AddNewPart<WorksheetPart>();
        wsPart.Worksheet = new(
            new SheetData(
                new Row(
                    Header("Id"),
                    Header("Name"),
                    Header("Salary"))
                {
                    RowIndex = 1u
                }));

        // Notes on the Id (A1) and Salary (C1) headers; Name (B1) has none.
        // A note on a data cell (B2) is ignored — the model only attributes header notes.
        var commentsPart = wsPart.AddNewPart<WorksheetCommentsPart>();
        commentsPart.Comments = new(
            new Authors(new Author("")),
            new CommentList(
                Note("A1", "Read-only — assigned by payroll."),
                Note("C1", "Gross annual salary in USD."),
                Note("B2", "Ignored: not a header cell.")));

        var customPart = wbPart.AddCustomXmlPart(CustomXmlPartType.CustomXml);
        using (var stream = customPart.GetStream(FileMode.Create))
        using (var writer = new StreamWriter(stream))
        {
            writer.Write(
                """
                <columnMetadata xmlns="urn:test:column-metadata">
                  <sheet name="Sheet1">
                    <column index="1" property="Id" />
                    <column index="3" property="Salary" />
                  </sheet>
                </columnMetadata>
                """);
        }

        var sheets = wbPart.Workbook.GetFirstChild<Sheets>()!;
        sheets.Append(new Sheet
        {
            Id = wbPart.GetIdOfPart(wsPart),
            SheetId = 1,
            Name = "Sheet1"
        });

        return document;

        static Cell Header(string value) =>
            new()
            {
                DataType = CellValues.InlineString,
                InlineString = new(new Text(value))
            };

        static Comment Note(string reference, string text) =>
            new(
                new CommentText(
                    new Run(
                        new Text(text)
                        {
                            Space = SpaceProcessingModeValues.Preserve
                        })))
            {
                Reference = reference,
                AuthorId = 0U
            };
    }
}
