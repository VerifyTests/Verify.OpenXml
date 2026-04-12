using System.Runtime.CompilerServices;

public static class ModuleInitializer
{
    [ModuleInitializer]
    public static void Initialize()
    {
        VerifyOpenXml.Initialize();

        var projectDir = ProjectDir();
        Verifier.DerivePathInfo(
            (sourceFile, projectDirectory, type, method) =>
                new(directory: projectDir, typeName: type.Name, methodName: method.Name));
    }

    [ModuleInitializer]
    public static void InitializeOther() =>
        VerifierSettings.InitializePlugins();

    static string ProjectDir([CallerFilePath] string here = "") =>
        Path.GetDirectoryName(here)!;
}
