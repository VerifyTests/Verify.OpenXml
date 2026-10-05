public static class ModuleInitializer
{
    #region InitializeOutputs

    [ModuleInitializer]
    public static void Initialize()
    {
        VerifyOpenXml.Initialize();

        // For every test: no text and no rendered pages
        VerifierSettings.PageText(PageTextPlacement.None);
        VerifierSettings.ExcludeDerivedTargets("png");
    }

    #endregion
}
