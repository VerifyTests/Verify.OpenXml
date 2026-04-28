#if NET10_0_OR_GREATER
using Morph;

static class MorphRenderer
{
    public static DocumentConverter? Instance { get; }

    static MorphRenderer()
    {
        var directory = Path.GetDirectoryName(typeof(MorphRenderer).Assembly.Location)!;

        var skiaPath = Path.Combine(directory, "Morph.OpenXml.Skia.dll");
        var imageSharpPath = Path.Combine(directory, "Morph.OpenXml.ImageSharp.dll");

        var hasSkia = File.Exists(skiaPath);
        var hasImageSharp = File.Exists(imageSharpPath);

        if (hasSkia && hasImageSharp)
        {
            throw new("Cannot reference both Morph.OpenXml.Skia and Morph.OpenXml.ImageSharp. Pick one rendering backend.");
        }

        if (hasSkia)
        {
            Instance = Load(skiaPath, "Morph.SkiaDocumentConverter");
        }
        else if (hasImageSharp)
        {
            Instance = Load(imageSharpPath, "Morph.ImageSharpDocumentConverter");
        }
    }

    static DocumentConverter Load(string assemblyPath, string typeName)
    {
        var assembly = Assembly.LoadFrom(assemblyPath);
        var type = assembly.GetType(typeName, throwOnError: true)!;
        return (DocumentConverter) Activator.CreateInstance(type)!;
    }
}
#endif
