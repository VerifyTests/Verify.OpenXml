// Initialized with only OpenXmlOutputs.Csv, so the docx has no txt target and no png pages.
public class Tests
{
    [Test]
    public Task Excel() =>
        VerifyFile(ProjectFiles.sample_xlsx.Path);

    [Test]
    public Task Word() =>
        VerifyFile(ProjectFiles.sample_docx.Path);
}
