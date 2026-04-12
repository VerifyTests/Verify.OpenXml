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
        var skia = TryLoad("Morph.OpenXml.Skia");
        var imageSharp = TryLoad("Morph.OpenXml.ImageSharp");

        if (skia != null && imageSharp != null)
        {
            throw new("Cannot reference both Morph.OpenXml.Skia and Morph.OpenXml.ImageSharp. Pick one rendering backend.");
        }

        if (skia != null)
        {
            return Create(skia, "WordRender.Skia.DocumentConverter");
        }

        if (imageSharp != null)
        {
            return Create(imageSharp, "WordRender.ImageSharp.DocumentConverter");
        }

        return null;
    }

    static Assembly? TryLoad(string name)
    {
        try
        {
            return Assembly.Load(new AssemblyName(name));
        }
        catch (FileNotFoundException)
        {
            return null;
        }
        catch (FileLoadException)
        {
            return null;
        }
    }

    static DocumentConverter Create(Assembly assembly, string typeName)
    {
        var type = assembly.GetType(typeName, throwOnError: true)!;
        return (DocumentConverter) Activator.CreateInstance(type)!;
    }
}
#endif
