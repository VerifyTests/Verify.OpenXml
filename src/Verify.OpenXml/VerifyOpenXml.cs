namespace VerifyTests;

public static partial class VerifyOpenXml
{
    public static bool Initialized { get; private set; }

    public static void Initialize()
    {
        if (Initialized)
        {
            throw new("Already Initialized");
        }

        Initialized = true;

        VerifierSettings.RegisterStreamConverter("xlsx", (_, target, settings) => ConvertExcel(target, settings));
        VerifierSettings.RegisterFileConverter<SpreadsheetDocument>(ConvertExcel);

        VerifierSettings.RegisterStreamConverter("docx", (_, target, settings) => ConvertWord(target, settings));
        VerifierSettings.RegisterFileConverter<WordprocessingDocument>(ConvertWord);

        VerifierSettings.RegisterStreamConverter("pptx", (_, target, settings) => ConvertPowerpoint(target, settings));
        VerifierSettings.RegisterFileConverter<PresentationDocument>(ConvertPowerpoint);
    }
}
