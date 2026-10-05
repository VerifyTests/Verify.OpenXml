using PText = DocumentFormat.OpenXml.Drawing.Text;
using PParagraph = DocumentFormat.OpenXml.Drawing.Paragraph;
using SlideId = DocumentFormat.OpenXml.Presentation.SlideId;

namespace VerifyTests;

public static partial class VerifyOpenXml
{
    static ConversionResult ConvertPowerpoint(Stream stream, IReadOnlyDictionary<string, object> settings)
    {
        using var document = PresentationDocument.Open(stream, false, new()
        {
            AutoSave = false
        });
        return ConvertPowerpoint(document, settings);
    }

    static ConversionResult ConvertPowerpoint(PresentationDocument document, IReadOnlyDictionary<string, object> settings)
    {
        // Names the pages, places their text, and says which pages and which of their outputs the
        // verification wants
        var conversion = new PagedConversion(settings)
        {
            Info = GetPowerpointInfo(document)
        };
        var slides = GetSlides(document);

        // Building the deterministic pptx is expensive, so skip it when the pptx target is excluded.
        // The text and info are extracted from the document, so they are unaffected.
        var buildDeterministic = !settings.IsTargetExcluded("pptx");
        var render = RenderingEnabled(conversion);

        using var sourceStream = new MemoryStream();
        if (buildDeterministic ||
            render)
        {
            document.Clone(sourceStream);
            sourceStream.Position = 0;
        }

        // ReSharper disable once TooWideLocalVariableScope
        // ReSharper disable once RedundantAssignment
        Stream? deterministic = null;
        if (buildDeterministic)
        {
            deterministic = DeterministicPackage.Convert(sourceStream);
            conversion.Source(new("pptx", deterministic));
        }

        IReadOnlyList<byte[]>? images = null;
#if NET10_0_OR_GREATER
        // Rendering needs a package stream. Reuse the deterministic pptx when built; otherwise render
        // from the raw clone (DeterministicPackage only normalizes zip container metadata, not content,
        // so the rendered pixels are the same either way).
        if (render)
        {
            images = MorphRenderer.RenderPowerpoint(deterministic ?? sourceStream);
        }
#endif

        // A page is a slide. The renderer draws every slide, in the same order as GetSlides, so the
        // image and the text of a slide are paired by its number.
        var includeText = conversion.IncludeText;
        foreach (var number in conversion.Pages(slides.Count))
        {
            Stream? image = null;
            if (images != null)
            {
                image = new MemoryStream(images[number - 1]);
            }

            string? text = null;
            if (includeText)
            {
                text = GetSlideText(slides[number - 1]);
            }

            conversion.AddPage(number, image, text);
        }

        return conversion.Build();
    }

    /// <summary>
    /// Document metadata, or null when the presentation carries none — so no empty <c>Document</c> is written.
    /// </summary>
    internal static PowerpointInfo? GetPowerpointInfo(PresentationDocument document)
    {
        var properties = GetPowerpointProperties(document);
        if (properties == null)
        {
            return null;
        }

        return new()
        {
            Properties = properties
        };
    }

    /// <summary>
    /// Slides in presentation order. <c>p:sldIdLst</c> is authoritative: it is the order the slides
    /// are shown and rendered in, whereas <c>PresentationPart.SlideParts</c> is the order the parts
    /// were related in, which reordering a deck does not change.
    /// </summary>
    internal static List<SlidePart> GetSlides(PresentationDocument document)
    {
        var slides = new List<SlidePart>();
        var presentationPart = document.PresentationPart;
        var slideIds = presentationPart?.Presentation?.SlideIdList;
        if (presentationPart == null ||
            slideIds == null)
        {
            return slides;
        }

        foreach (var slideId in slideIds.Elements<SlideId>())
        {
            var relationshipId = slideId.RelationshipId?.Value;
            if (relationshipId != null &&
                presentationPart.TryGetPartById(relationshipId, out var part) &&
                part is SlidePart slidePart)
            {
                slides.Add(slidePart);
            }
        }

        return slides;
    }

    internal static string? GetSlideText(SlidePart slidePart)
    {
        var builder = new StringBuilder();
        if (AppendSlideText(builder, slidePart))
        {
            return builder.ToString();
        }

        return null;
    }

    internal static Dictionary<string, object?>? GetPowerpointProperties(PresentationDocument document) =>
        GetCoreProperties(document);

    internal static bool AppendSlideText(StringBuilder builder, SlidePart slidePart)
    {
        var slide = slidePart.Slide;
        if (slide == null)
        {
            return false;
        }

        var startLength = builder.Length;

        foreach (var paragraph in slide.Descendants<PParagraph>())
        {
            var paragraphStart = builder.Length;
            foreach (var text in paragraph.Descendants<PText>())
            {
                builder.Append(text.Text);
            }

            if (builder.Length > paragraphStart)
            {
                builder.AppendLine();
            }
        }

        if (builder.Length == startLength)
        {
            return false;
        }

        builder.TrimEnd();
        return builder.Length > startLength;
    }
}

class PowerpointInfo
{
    public required Dictionary<string, object?> Properties { get; init; }
}
