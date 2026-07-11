using PText = DocumentFormat.OpenXml.Drawing.Text;
using PParagraph = DocumentFormat.OpenXml.Drawing.Paragraph;

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
        var info = GetPowerpointInfo(document);
        var text = GetPowerpointText(document);

        List<Target> targets = [];
        // Building the deterministic pptx is expensive, so skip it when the pptx target is excluded.
        // The text and info are extracted from the document, so they are unaffected.
        if (!settings.IsTargetExcluded("pptx"))
        {
            using var sourceStream = new MemoryStream();
            document.Clone(sourceStream);
            sourceStream.Position = 0;
            var resultStream = DeterministicPackage.Convert(sourceStream);
            targets.Add(
                new("pptx", resultStream)
                {
                    BypassComparersForSubsequentOnDifference = true
                });
        }

        // The text is its own target, so it is deliberately absent from the info. Carrying it in both
        // wrote the slide text to two snapshot files.
        if (!string.IsNullOrWhiteSpace(text))
        {
            targets.Add(new("txt", text!));
        }

        return new(info, targets);
    }

    internal static PowerpointInfo GetPowerpointInfo(PresentationDocument document) =>
        new()
        {
            Properties = GetPowerpointProperties(document),
            SlideCount = document.PresentationPart?.SlideParts.Count() ?? 0
        };

    internal static string? GetPowerpointText(PresentationDocument document)
    {
        var presentationPart = document.PresentationPart;
        if (presentationPart == null)
        {
            return null;
        }

        var builder = new StringBuilder();
        foreach (var slidePart in presentationPart.SlideParts)
        {
            var before = builder.Length;
            if (before > 0)
            {
                builder.Append("\n---\n");
            }

            if (!AppendSlideText(builder, slidePart))
            {
                builder.Length = before;
            }
        }

        return builder.Length > 0 ? builder.ToString() : null;
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
    public Dictionary<string, object?>? Properties { get; init; }
    public required int SlideCount { get; init; }
}
