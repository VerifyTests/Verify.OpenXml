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

        using var sourceStream = new MemoryStream();
        document.Clone(sourceStream);
        sourceStream.Position = 0;
        var resultStream = DeterministicPackage.Convert(sourceStream);

        List<Target> targets =
        [
            new("pptx", resultStream)
            {
                BypassComparersForSubsequentOnDifference = true
            }
        ];

        if (!string.IsNullOrWhiteSpace(info.Text))
        {
            targets.Add(new("txt", info.Text!));
        }

        return new(info, targets);
    }

    internal static PowerpointInfo GetPowerpointInfo(PresentationDocument document)
    {
        var presentationPart = document.PresentationPart;
        var builder = new StringBuilder();
        var slideCount = 0;

        if (presentationPart?.SlideParts != null)
        {
            foreach (var slidePart in presentationPart.SlideParts)
            {
                slideCount++;
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
        }

        return new()
        {
            Properties = GetPowerpointProperties(document),
            SlideCount = slideCount,
            Text = builder.Length > 0 ? builder.ToString() : null
        };
    }

    internal static Dictionary<string, object?>? GetPowerpointProperties(PresentationDocument document)
    {
        var packageProperties = document.PackageProperties;
        var properties = new Dictionary<string, object?>();

        AddPropertyIfNotEmpty(properties, "Title", packageProperties.Title);
        AddPropertyIfNotEmpty(properties, "Subject", packageProperties.Subject);
        AddPropertyIfNotEmpty(properties, "Creator", packageProperties.Creator);
        AddPropertyIfNotEmpty(properties, "Keywords", packageProperties.Keywords);
        AddPropertyIfNotEmpty(properties, "Description", packageProperties.Description);
        AddPropertyIfNotEmpty(properties, "Category", packageProperties.Category);
        AddPropertyIfNotEmpty(properties, "LastModifiedBy", packageProperties.LastModifiedBy);
        AddPropertyIfNotEmpty(properties, "ContentStatus", packageProperties.ContentStatus);
        AddPropertyIfNotEmpty(properties, "Revision", packageProperties.Revision);

        return properties.Count > 0 ? properties : null;
    }

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
    public string? Text { get; init; }
}
