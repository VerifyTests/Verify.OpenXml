using S = DocumentFormat.OpenXml.Spreadsheet;

// Verifies that a page of a workbook is a sheet, hidden or not, and that PagesToInclude limits the
// csv of a sheet as it does its image. Which sheet a page is does not depend on a renderer, so the
// csv is limited the same way in a project without one.
public class ExcelPagesTests
{
    [Test]
    public Task PagesToInclude() =>
        VerifyFile(ProjectFiles.sample_multiple_sheets_xlsx.Path)
            .PagesToInclude(_ => _ == 2);

    // A hidden sheet is verified as any other: it has a csv, and where there is a renderer a page.
    // What says it is hidden is HiddenSheets in the info file.
    [Test]
    public async Task AHiddenSheetIsAPage()
    {
        using var workbook = FirstSheetHidden();
        await Verify(workbook);
    }

    // Hidden sheets are counted, so the first page is the hidden sheet and the second the one that
    // is not.
    [Test]
    public async Task AHiddenSheetIsCountedByPagesToInclude()
    {
        using var workbook = FirstSheetHidden();
        await Verify(workbook)
            .PagesToInclude(_ => _ == 2);
    }

    // And the hidden sheet is the one that is verified when the first page is asked for
    [Test]
    public async Task AHiddenSheetIsTheFirstPage()
    {
        using var workbook = FirstSheetHidden();
        await Verify(workbook)
            .PagesToInclude(1);
    }

    // A sheet that only code can unhide is verified as one Excel can
    [Test]
    public async Task AVeryHiddenSheetIsAPage()
    {
        using var workbook = FirstSheetHidden(S.SheetStateValues.VeryHidden);
        await Verify(workbook);
    }

    static SpreadsheetDocument FirstSheetHidden() =>
        FirstSheetHidden(S.SheetStateValues.Hidden);

    static SpreadsheetDocument FirstSheetHidden(S.SheetStateValues state)
    {
        // Copied, since a stream over the bytes could not grow
        var stream = new MemoryStream();
        using (var file = File.OpenRead(ProjectFiles.sample_multiple_sheets_xlsx.Path))
        {
            file.CopyTo(stream);
        }

        var workbook = SpreadsheetDocument.Open(stream, true);
        var sheet = workbook.WorkbookPart!.Workbook!.Sheets!.Elements<S.Sheet>().First();
        sheet.State = state;
        return workbook;
    }
}
