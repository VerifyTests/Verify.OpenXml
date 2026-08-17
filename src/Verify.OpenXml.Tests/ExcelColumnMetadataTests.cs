// Exercises ColumnInfo.Metadata sourced from a custom XML part shaped like
// <sheet name="..."><column index="N" {anyAttr}/></sheet>. The reader is namespace-agnostic
// and surfaces every attribute (other than index) verbatim.
[TestFixture]
public class ExcelColumnMetadataTests
{
    [Test]
    public Task ColumnMetadataFromCustomXml()
    {
        using var document = CreateDocument(
            """
            <columnMetadata xmlns="urn:test:column-metadata">
              <sheet name="Sheet1">
                <column index="1" property="Id" nullable="false" />
                <column index="2" property="Name" />
                <column index="3" property="Notes" nullable="true" />
              </sheet>
            </columnMetadata>
            """);
        return Verify(document);
    }

    [Test]
    public Task ColumnMetadataIgnoresNamespace()
    {
        // No xmlns at all — reader still picks it up by structure.
        using var document = CreateDocument(
            """
            <whatever>
              <sheet name="Sheet1">
                <column index="1" foo="bar" />
                <column index="2" baz="qux" />
              </sheet>
            </whatever>
            """);
        return Verify(document);
    }

    [Test]
    public Task SheetMetadataFromCustomXml()
    {
        // Sheet-level attributes (anything other than `name`) flow to SheetInfo.Metadata. The
        // mechanism mirrors ColumnInfo.Metadata so producers like Excelsior can attach per-sheet
        // annotations such as `bannerRows="1"` without coordinating schema changes here.
        using var document = CreateDocument(
            """
            <columnMetadata xmlns="urn:test:column-metadata">
              <sheet name="Sheet1" bannerRows="1" origin="import">
                <column index="1" property="Id" />
              </sheet>
            </columnMetadata>
            """);
        return Verify(document);
    }

    [Test]
    public Task ColumnMetadataMissingPartIsBenign()
    {
        // No custom XML part at all — Metadata stays null on each column.
        using var document = CreateDocument(customXml: null);
        return Verify(document);
    }

    static SpreadsheetDocument CreateDocument(string? customXml)
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
                    Header("Notes"))
                {
                    RowIndex = 1u
                }),
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

        if (customXml != null)
        {
            var customPart = wbPart.AddCustomXmlPart(CustomXmlPartType.CustomXml);
            using var stream = customPart.GetStream(FileMode.Create);
            using var writer = new StreamWriter(stream);
            writer.Write(customXml);
        }

        return document;

        static Cell Header(string value) =>
            new()
            {
                DataType = CellValues.InlineString,
                InlineString = new(new Text(value))
            };
    }
}
