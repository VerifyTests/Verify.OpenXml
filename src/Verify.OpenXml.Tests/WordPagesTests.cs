using W = DocumentFormat.OpenXml.Wordprocessing;

// Verifies that the text of a document is split by the page it is on, and limited by PagesToInclude
// as the images are. Where a page ends is only known to a renderer, so in a project without one the
// text is that of the whole document.
public class WordPagesTests
{
    [Test]
    public async Task TextOfEachPage()
    {
        using var document = ThreePages();
        await Verify(document)
            .ExcludeDerivedTargets("png");
    }

    [Test]
    public async Task PagesToInclude()
    {
        using var document = ThreePages();
        await Verify(document)
            .PagesToInclude(2)
            .ExcludeDerivedTargets("png");
    }

    // sample.docx, so that the styles and fonts are there, with a second page that starts with a
    // paragraph and has a table, and a third that starts with a table.
    static WordprocessingDocument ThreePages()
    {
        // Copied, since a stream over the bytes could not grow to hold what is added
        var stream = new MemoryStream();
        using (var file = File.OpenRead("sample.docx"))
        {
            file.CopyTo(stream);
        }

        var document = WordprocessingDocument.Open(stream, true);
        var body = document.MainDocumentPart!.Document!.Body!;
        var last = body.Elements<W.Paragraph>().Last();

        last.InsertAfterSelf(Table("Third page, first row", "Third page, second row"));
        last.InsertAfterSelf(PageBreak());
        last.InsertAfterSelf(Table("Second page, row"));
        last.InsertAfterSelf(Text("Second page, paragraph"));
        last.InsertAfterSelf(PageBreak());

        return document;

        static W.Paragraph Text(string text) =>
            new(new W.Run(new W.Text(text)));

        static W.Paragraph PageBreak() =>
            new(
                new W.Run(
                    new W.Break
                    {
                        Type = W.BreakValues.Page
                    }));

        static W.Table Table(params string[] rows)
        {
            var table = new W.Table();
            foreach (var row in rows)
            {
                table.Append(
                    new W.TableRow(
                        new W.TableCell(Text(row)),
                        new W.TableCell(Text("cell"))));
            }

            return table;
        }
    }
}
