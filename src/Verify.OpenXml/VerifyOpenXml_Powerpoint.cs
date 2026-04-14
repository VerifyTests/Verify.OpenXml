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

        List<Target> targets = [new("pptx", resultStream)];

        if (!string.IsNullOrWhiteSpace(info.Text))
        {
            targets.Add(new("txt", info.Text!));
        }

        return new(info, targets);
    }

    static PowerpointInfo GetPowerpointInfo(PresentationDocument document)
    {
        var presentationPart = document.PresentationPart;
        var slideTexts = new List<string>();

        if (presentationPart?.SlideParts != null)
        {
            foreach (var slidePart in presentationPart.SlideParts)
            {
                var text = GetSlideText(slidePart);
                if (!string.IsNullOrEmpty(text))
                {
                    slideTexts.Add(text);
                }
            }
        }

        return new()
        {
            Properties = GetPowerpointProperties(document),
            SlideCount = presentationPart?.SlideParts.Count() ?? 0,
            Text = slideTexts.Count > 0 ? string.Join("\n---\n", slideTexts) : null
        };
    }

    static Dictionary<string, object?>? GetPowerpointProperties(PresentationDocument document)
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

    static string GetSlideText(SlidePart slidePart)
    {
        var builder = new StringBuilder();

        var slide = slidePart.Slide;
        if (slide == null)
        {
            return string.Empty;
        }

        foreach (var paragraph in slide.Descendants<PParagraph>())
        {
            var paragraphText = new StringBuilder();
            foreach (var text in paragraph.Descendants<PText>())
            {
                paragraphText.Append(text.Text);
            }

            if (paragraphText.Length > 0)
            {
                builder.AppendLine(paragraphText.ToString());
            }
        }

        return builder.ToString().TrimEnd();
    }
}

class PowerpointInfo
{
    public Dictionary<string, object?>? Properties { get; init; }
    public required int SlideCount { get; init; }
    public string? Text { get; init; }
}
