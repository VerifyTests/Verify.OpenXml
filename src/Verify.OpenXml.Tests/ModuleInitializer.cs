public static class ModuleInitializer
{
    #region enable

    [ModuleInitializer]
    public static void Initialize() =>
        VerifyOpenXml.Initialize();

    #endregion

    // No rendering backend is referenced here, so these pin nothing today. They are set anyway so
    // that referencing one later cannot quietly start rendering from the machine's fonts, or on the
    // machine's paper.
    [ModuleInitializer]
    public static void InitializeRendering()
    {
        VerifyOpenXml.FontDirectory = Path.Combine(ProjectDir(), "..", "Fonts");
        VerifyOpenXml.UseLetterPageSize = false;
    }

    [ModuleInitializer]
    public static void InitializeOther() =>
        VerifierSettings.InitializePlugins();

    static string ProjectDir([CallerFilePath] string here = "") =>
        Path.GetDirectoryName(here)!;
}