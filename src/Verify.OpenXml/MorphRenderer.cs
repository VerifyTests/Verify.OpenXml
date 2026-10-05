#if NET10_0_OR_GREATER
using Morph;

namespace VerifyTests;

static class MorphRenderer
{
    // Morph has no common base across its three converters, so each is captured as the one method
    // this needs. Null when no backend is referenced.
    static Renderer? word;
    static Renderer? excel;
    static Renderer? powerpoint;

    delegate IReadOnlyList<byte[]> Renderer(Stream package, ImageExportOptions options);

    /// <summary>
    /// Whether a backend was found. All three renderers come from the same assembly, so this holds for
    /// every document type or none.
    /// </summary>
    public static bool Enabled => word != null;

    static MorphRenderer()
    {
        var directory = Path.GetDirectoryName(typeof(MorphRenderer).Assembly.Location)!;

        var skiaPath = Path.Combine(directory, "Morph.Skia.dll");
        var imageSharpPath = Path.Combine(directory, "Morph.ImageSharp.dll");

        var hasSkia = File.Exists(skiaPath);
        var hasImageSharp = File.Exists(imageSharpPath);

        if (hasSkia && hasImageSharp)
        {
            throw new("Cannot reference both Morph.Skia and Morph.ImageSharp. Pick one rendering backend.");
        }

        string assemblyPath;
        string prefix;
        if (hasSkia)
        {
            assemblyPath = skiaPath;
            prefix = "Skia";
        }
        else if (hasImageSharp)
        {
            assemblyPath = imageSharpPath;
            prefix = "ImageSharp";
        }
        else
        {
            return;
        }

        var assembly = Assembly.LoadFrom(assemblyPath);
        word = Load<DocumentConverter>(assembly, $"Morph.{prefix}DocumentConverter").ConvertToImageData;
        excel = Load<ExcelConverter>(assembly, $"Morph.{prefix}ExcelConverter").ConvertToImageData;
        powerpoint = Load<PowerPointConverter>(assembly, $"Morph.{prefix}PowerPointConverter").ConvertToImageData;
    }

    static T Load<T>(Assembly assembly, string typeName)
    {
        var type = assembly.GetType(typeName, throwOnError: true)!;
        return (T) Activator.CreateInstance(type)!;
    }

    // Each renders every page, in page order. Only called once Enabled has been checked, which is
    // what says the renderer is there.
    public static IReadOnlyList<byte[]> RenderWord(Stream docx) =>
        Render(word!, docx);

    // A sheet is drawn whole, as the one image, rather than as it prints. So a page is a sheet,
    // however long it is and whatever paper its page setup names.
    public static IReadOnlyList<byte[]> RenderExcel(Stream xlsx) =>
        Render(
            excel!,
            xlsx,
            Options() with
            {
                SheetPagination = SheetPagination.OnePagePerSheet
            });

    public static IReadOnlyList<byte[]> RenderPowerpoint(Stream pptx) =>
        Render(powerpoint!, pptx);

    // The 1 based page each bookmark of a docx is on, by the name of the bookmark. Laid out with the
    // options the pages are drawn with, so a bookmark is on the page its text is drawn on.
    public static IReadOnlyDictionary<string, int> WordBookmarkPages(Stream docx)
    {
        docx.Position = 0;
        return DocumentConverter.GetBookmarkPages(docx, Options());
    }

    static IReadOnlyList<byte[]> Render(Renderer render, Stream package) =>
        Render(render, package, Options());

    static IReadOnlyList<byte[]> Render(Renderer render, Stream package, ImageExportOptions options)
    {
        package.Position = 0;
        using var copy = new MemoryStream();
        package.CopyTo(copy);
        package.Position = 0;
        copy.Position = 0;

        return render(copy, options);
    }

    static ImageExportOptions Options() =>
        new()
        {
            DeterministicRendering = true,
            FontDirectory = VerifyOpenXml.FontDirectory,
            DefaultFont = VerifyOpenXml.DefaultFont,
            UseLetterPageSize = VerifyOpenXml.UseLetterPageSize
        };
}
#endif
