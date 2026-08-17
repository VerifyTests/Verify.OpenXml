public static class ModuleInitializer
{
    #region enable

    [ModuleInitializer]
    public static void Initialize() =>
        VerifyOpenXml.Initialize();

    #endregion

    // No rendering backend is referenced here, so this pins nothing today. It is set anyway so that
    // referencing one later cannot quietly start rendering from the machine's fonts.
    [ModuleInitializer]
    public static void InitializeFonts() =>
        VerifyOpenXml.FontDirectory = Path.Combine(ProjectDir(), "..", "Fonts");

    [ModuleInitializer]
    public static void InitializeOther() =>
        VerifierSettings.InitializePlugins();

    static string ProjectDir([CallerFilePath] string here = "") =>
        Path.GetDirectoryName(here)!;
}