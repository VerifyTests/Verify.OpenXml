namespace VerifyTests;

public static partial class VerifyOpenXml
{
    public static bool Initialized { get; private set; }

    /// <summary>
    /// Directory holding the fonts to render pages with. When set, only fonts found here are used —
    /// the machine's installed fonts are ignored entirely and a face the directory does not carry
    /// throws rather than silently resolving to something else.
    /// </summary>
    /// <remarks>
    /// Leaving this null renders with whatever the machine has installed, which is right for a
    /// one-off render and wrong for a snapshot: the same document measures differently where a face
    /// resolves differently, and in a spreadsheet that shifts column widths and so the wrapping. Set
    /// it to a directory committed alongside the tests to make the rendered snapshots reproducible,
    /// and to turn a fixture that reaches for an unavailable font into a failure at the point of the
    /// mistake rather than a snapshot mismatch on someone else's machine.
    /// </remarks>
    public static string? FontDirectory { get; set; }

    /// <summary>
    /// The paper to render on when the document states none — a worksheet with no
    /// <c>pageSetup/@paperSize</c>, or a docx with no <c>w:pgSz</c>. <c>true</c> is US Letter,
    /// <c>false</c> is A4. A document that does state its paper size is unaffected.
    /// </summary>
    /// <remarks>
    /// Leaving this null takes the machine's region — Letter in North America, A4 elsewhere — which
    /// is what Word and Excel do, and which makes the rendered page size depend on where the tests
    /// run. A workbook stating no paper size is the common case rather than the rare one, so set
    /// this for snapshots that have to survive a move between machines.
    /// </remarks>
    public static bool? UseLetterPageSize { get; set; }

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
    /// Whether rendering will produce PNG targets. Always false below <c>net10.0</c>, where the Morph
    /// integration is compiled out. Cloning the source package is only worth it for the deterministic
    /// binary target or for rendering, so each converter checks this before cloning.
    /// </summary>
    static bool RenderingEnabled =>
#if NET10_0_OR_GREATER
        MorphRenderer.Enabled;
#else
        false;
#endif

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
