public static class ModuleInitializer
{
    #region InitializeOutputs

    [ModuleInitializer]
    public static void Initialize() =>
        VerifyOpenXml.Initialize(OpenXmlOutputs.Csv);

    #endregion
}
