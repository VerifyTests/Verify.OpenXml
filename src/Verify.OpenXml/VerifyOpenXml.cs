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
        VerifierSettings.RegisterFileConverter<SpreadsheetDocument>((document, _) => ToPackage(document, "xlsx"));

        VerifierSettings.RegisterStreamConverter("docx", (_, target, settings) => ConvertWord(target, settings));
        VerifierSettings.RegisterFileConverter<WordprocessingDocument>((document, _) => ToPackage(document, "docx"));

        VerifierSettings.RegisterStreamConverter("pptx", (_, target, settings) => ConvertPowerpoint(target, settings));
        VerifierSettings.RegisterFileConverter<PresentationDocument>((document, _) => ToPackage(document, "pptx"));
    }

    /// <summary>
    /// Verify runs the stream converter registered for <paramref name="extension" /> over the target returned
    /// here, and that converter is what produces the info, the text and the rendered pages. So a file converter
    /// only has to hand over the package: building the other targets here too wrote every snapshot twice.
    /// </summary>
    /// <remarks>
    /// Generic because <c>Clone</c> resolves the package factory from the static type: called through the
    /// <see cref="OpenXmlPackage" /> base it throws.
    /// </remarks>
    static ConversionResult ToPackage<TPackage>(TPackage package, string extension)
        where TPackage : OpenXmlPackage
    {
        var stream = new MemoryStream();
        package.Clone(stream);
        stream.Position = 0;
        return new(null, [new(extension, stream)]);
    }
}
