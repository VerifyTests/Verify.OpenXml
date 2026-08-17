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
        DerivePathInfo(
            (_, _, type, method) =>
                new(directory: projectDir, typeName: type.Name, methodName: method.Name));
    }

    [ModuleInitializer]
    public static void InitializeOther() =>
        VerifierSettings.InitializePlugins();

    static string ProjectDir([CallerFilePath] string here = "") =>
        Path.GetDirectoryName(here)!;
}
