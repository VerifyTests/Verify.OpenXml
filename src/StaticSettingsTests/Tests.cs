// Initialized with PageText(None) and ExcludeDerivedTargets("png"), so the docx has no text and no png pages.
public class Tests
{
    [Test]
    public Task Excel() =>
        VerifyFile(ProjectFiles.sample_xlsx.Path);

    [Test]
    public Task Word() =>
        VerifyFile(ProjectFiles.sample_docx.Path);
}
