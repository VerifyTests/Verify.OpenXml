#if NET10_0_OR_GREATER
using System.Reflection;
using WordRender;

namespace VerifyTests;

static class MorphRenderer
{
    static readonly Lazy<DocumentConverter?> converter = new(Resolve);

    public static DocumentConverter? Instance => converter.Value;

    static DocumentConverter? Resolve()
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
            return Load(skiaPath, "WordRender.Skia.DocumentConverter");
        }

        if (hasImageSharp)
        {
            return Load(imageSharpPath, "WordRender.ImageSharp.DocumentConverter");
        }

        return null;
    }

    static DocumentConverter Load(string assemblyPath, string typeName)
    {
        var assembly = Assembly.LoadFrom(assemblyPath);
        var type = assembly.GetType(typeName, throwOnError: true)!;
        return (DocumentConverter) Activator.CreateInstance(type)!;
    }
}
#endif
