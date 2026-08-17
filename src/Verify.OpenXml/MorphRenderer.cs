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

    public static void AddWordPages(Stream docx, List<Target> targets) =>
        AddPages(word, docx, targets);

    public static void AddExcelPages(Stream xlsx, List<Target> targets) =>
        AddPages(excel, xlsx, targets);

    public static void AddPowerpointPages(Stream pptx, List<Target> targets) =>
        AddPages(powerpoint, pptx, targets);

    static void AddPages(Renderer? render, Stream package, List<Target> targets)
    {
        if (render == null)
        {
            return;
        }

        package.Position = 0;
        using var copy = new MemoryStream();
        package.CopyTo(copy);
        package.Position = 0;
        copy.Position = 0;

        var pages = render(
            copy,
            new()
            {
                DeterministicRendering = true,
                FontDirectory = VerifyOpenXml.FontDirectory,
                UseLetterPageSize = VerifyOpenXml.UseLetterPageSize
            });
        foreach (var page in pages)
        {
            targets.Add(new("png", new MemoryStream(page)));
        }
    }
}
#endif
