public static class ModuleInitializer
{
    [ModuleInitializer]
    public static void Initialize()
    {
        VerifyOpenXml.Initialize();

        var projectDir = ProjectDir();
        // Render from the committed fonts only, never the machine's. Font metrics decide column
        // widths and so where text wraps, so a face that resolves differently here and on CI moves
        // the rendered pixels; pinning this also turns a fixture asking for a font the directory
        // does not carry into an outright failure rather than a snapshot mismatch elsewhere.
        VerifyOpenXml.FontDirectory = Path.Combine(projectDir, "..", "Fonts");
        // A4 whatever the machine's region says. A sheet stating no paperSize otherwise renders
        // Letter on a North American agent and A4 here, which reads as a rendering regression.
        VerifyOpenXml.UseLetterPageSize = false;
        DerivePathInfo(
            (_, _, type, method) =>
                new(directory: projectDir, typeName: type.Name, methodName: method.Name));
    }

    [ModuleInitializer]
    public static void InitializeOther()
    {
        VerifierSettings.Inline(maxLines: 10, applyMaxLinesToExisting: true);
        VerifierSettings.InitializePlugins();
    }

    static string ProjectDir([CallerFilePath] string here = "") =>
        Path.GetDirectoryName(here)!;
}
