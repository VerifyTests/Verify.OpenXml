// Initialized with only OpenXmlOutputs.Csv, so the docx has no txt target and no png pages.
public class Tests
{
    [Test]
    public Task Excel() =>
        VerifyFile("sample.xlsx");

    [Test]
    public Task Word() =>
        VerifyFile("sample.docx");
}
