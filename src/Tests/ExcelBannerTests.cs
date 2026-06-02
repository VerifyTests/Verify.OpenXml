// Verifies that a merged "banner" row above the header — a single horizontal merge from column A,
// as emitted for an instruction row above the data — is skipped when locating the header, so the
// real columns (Id/Name/Salary) are surfaced rather than the banner text. Also covers that the
// banner's rich text is not mis-attributed to the column beneath it.
[TestFixture]
public class ExcelBannerTests
{
    [Test]
    public Task MergedBannerRowIsSkipped()
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
                // Row 1: a merged banner spanning all columns, carrying rich text.
                new Row(BannerCell())
                {
                    RowIndex = 1u
                },
                // Row 2: the real header.
                new Row(
                    Header("Id"),
                    Header("Name"),
                    Header("Salary"))
                {
                    RowIndex = 2u
                },
                // Row 3: a data row.
                new Row(
                    Inline("1"),
                    Inline("John"),
                    Inline("75000"))
                {
                    RowIndex = 3u
                }),
            new MergeCells(
                new MergeCell
                {
                    Reference = "A1:C1"
                })
            {
                Count = 1
            });

        // Metadata is keyed by the real column index (1/2/3); skipping the banner is what lets it
        // line up with the Id/Name/Salary header instead of the lone banner cell.
        var customPart = wbPart.AddCustomXmlPart(CustomXmlPartType.CustomXml);
        using (var stream = customPart.GetStream(FileMode.Create))
        using (var writer = new StreamWriter(stream))
        {
            writer.Write(
                """
                <columnMetadata xmlns="urn:test:column-metadata">
                  <sheet name="Sheet1">
                    <column index="1" property="Id" />
                    <column index="2" property="Name" />
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

        static Cell Inline(string value) =>
            new()
            {
                DataType = CellValues.InlineString,
                InlineString = new(new Text(value))
            };

        static Cell BannerCell() =>
            new()
            {
                CellReference = "A1",
                DataType = CellValues.InlineString,
                InlineString = new(
                    new Run(
                        new RunProperties(new Bold()),
                        new Text("Instructions: ")),
                    new Run(
                        new Text("fill in every field.")))
            };
    }
}
