[TestFixture]
public class Samples
{
    #region VerifyExcel

    [Test]
    public Task VerifyExcel() =>
        VerifyFile("sample.xlsx");

    #endregion

    [Test]
    public Task MultipleSheets() =>
        VerifyFile("sample_multiple_sheets.xlsx");

    [Test]
    public Task HiddenRow() =>
        VerifyFile("sample_hidden_row.xlsx");

    [Test]
    public Task DontScrub() =>
        VerifyFile("sample.xlsx")
            .DontScrubGuids().DontScrubDateTimes();

    #region SpreadsheetDocument

    [Test]
    public async Task VerifySpreadsheetDocument()
    {
        await using var stream = File.OpenRead("sample.xlsx");
        using var reader = SpreadsheetDocument.Open(stream, false);
        await Verify(reader);
    }

    #endregion

    #region VerifyExcelStream

    [Test]
    public Task VerifyExcelStream()
    {
        var stream = new MemoryStream(File.ReadAllBytes("sample.xlsx"));
        return Verify(stream, "xlsx");
    }

    #endregion

    #region VerifyWord

    [Test]
    public Task VerifyWord() =>
        VerifyFile("sample.docx");

    #endregion

    #region WordprocessingDocument

    [Test]
    public async Task VerifyWordprocessingDocument()
    {
        await using var stream = File.OpenRead("sample.docx");
        using var reader = WordprocessingDocument.Open(stream, false);
        await Verify(reader);
    }

    #endregion

    #region VerifyWordStream

    [Test]
    public Task VerifyWordStream()
    {
        var stream = new MemoryStream(File.ReadAllBytes("sample.docx"));
        return Verify(stream, "docx");
    }

    #endregion

    #region VerifyPowerpoint

    [Test]
    public Task VerifyPowerpoint() =>
        VerifyFile("sample.pptx");

    #endregion

    #region PresentationDocument

    [Test]
    public async Task VerifyPresentationDocument()
    {
        await using var stream = File.OpenRead("sample.pptx");
        using var reader = PresentationDocument.Open(stream, false);
        await Verify(reader);
    }

    #endregion

    #region VerifyPowerpointStream

    [Test]
    public Task VerifyPowerpointStream()
    {
        var stream = new MemoryStream(File.ReadAllBytes("sample.pptx"));
        return Verify(stream, "pptx");
    }

    #endregion
}