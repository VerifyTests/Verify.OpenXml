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
    public Task DontScrub() =>
        VerifyFile("sample.xlsx")
            .DontScrubGuids().DontScrubDateTimes();

    #region SpreadsheetDocument

    [Test]
    public Task VerifySpreadsheetDocument()
    {
        using var stream = File.OpenRead("sample.xlsx");
        using var reader = SpreadsheetDocument.Open(stream, false);
        return Verify(reader);
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
}