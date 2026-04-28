using WordFont = DocumentFormat.OpenXml.Wordprocessing.Font;
using WordTable = DocumentFormat.OpenXml.Wordprocessing.Table;
using WordRun = DocumentFormat.OpenXml.Wordprocessing.Run;
using WordText = DocumentFormat.OpenXml.Wordprocessing.Text;
using WordBreak = DocumentFormat.OpenXml.Wordprocessing.Break;

namespace VerifyTests;

public static partial class VerifyOpenXml
{
    static ConversionResult ConvertWord(Stream stream, IReadOnlyDictionary<string, object> settings)
    {
        using var document = WordprocessingDocument.Open(stream, false, new()
        {
            AutoSave = false
        });
        return ConvertWord(document, settings);
    }

    static ConversionResult ConvertWord(WordprocessingDocument document, IReadOnlyDictionary<string, object> settings)
    {
        var info = GetWordInfo(document);

        // Create deterministic DOCX output
        using var sourceStream = new MemoryStream();
        document.Clone(sourceStream);
        sourceStream.Position = 0;
        var resultStream = DeterministicPackage.Convert(sourceStream);

        List<Target> targets = [new("docx", resultStream)];

        // Add text content as txt target
        if (!string.IsNullOrWhiteSpace(info.Text))
        {
            targets.Add(new("txt", info.Text!));
        }

#if NET10_0_OR_GREATER
        AddRenderedPages(resultStream, targets);
#endif

        return new(info, targets);
    }

#if NET10_0_OR_GREATER
    static void AddRenderedPages(Stream docxStream, List<Target> targets)
    {
        var renderer = MorphRenderer.Instance;
        if (renderer == null)
        {
            return;
        }

        docxStream.Position = 0;
        using var copy = new MemoryStream();
        docxStream.CopyTo(copy);
        docxStream.Position = 0;
        copy.Position = 0;

        var pages = renderer.ConvertToImageData(
            copy,
            new()
            {
                DeterministicRendering = true
            });
        for (var i = 0; i < pages.Count; i++)
        {
            targets.Add(new("png", new MemoryStream(pages[i]), $"page{i + 1:D2}"));
        }
    }
#endif

    static WordInfo GetWordInfo(WordprocessingDocument document)
    {
        var (fonts, embeddedFonts) = GetWordDocumentFonts(document);

        return new()
        {
            Properties = GetWordProperties(document),
            CustomProperties = GetWordCustomProperties(document),
            Fonts = fonts,
            EmbeddedFonts = embeddedFonts,
            Text = GetWordDocumentText(document)
        };
    }

    internal static (List<string>? fonts, List<string>? embeddedFonts) GetWordDocumentFonts(WordprocessingDocument document)
    {
        var fontTablePart = document.MainDocumentPart?.FontTablePart;
        if (fontTablePart?.Fonts == null)
        {
            return (null, null);
        }

        var fonts = new List<string>();
        var embeddedFonts = new List<string>();

        foreach (var font in fontTablePart.Fonts.Elements<WordFont>())
        {
            var fontName = font.Name?.Value;
            if (fontName == null)
            {
                continue;
            }

            fonts.Add(fontName);

            // Check if font has embedded data by looking for EmbedRegularFont, EmbedBoldFont, etc. child elements
            if (font.GetFirstChild<EmbedRegularFont>() != null ||
                font.GetFirstChild<EmbedBoldFont>() != null ||
                font.GetFirstChild<EmbedItalicFont>() != null ||
                font.GetFirstChild<EmbedBoldItalicFont>() != null)
            {
                embeddedFonts.Add(fontName);
            }
        }

        return (
            fonts.Count > 0 ? fonts.OrderBy(_ => _).ToList() : null,
            embeddedFonts.Count > 0 ? embeddedFonts.OrderBy(_ => _).ToList() : null
        );
    }

    internal static Dictionary<string, object?>? GetWordProperties(WordprocessingDocument document)
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

    static void AddPropertyIfNotEmpty(Dictionary<string, object?> properties, string key, string? value)
    {
        if (!string.IsNullOrWhiteSpace(value))
        {
            properties[key] = value;
        }
    }

    internal static Dictionary<string, object?>? GetWordCustomProperties(WordprocessingDocument document)
    {
        var customFilePropertiesPart = document.CustomFilePropertiesPart;
        if (customFilePropertiesPart?.Properties == null)
        {
            return null;
        }

        var properties = new Dictionary<string, object?>();

        foreach (var property in customFilePropertiesPart.Properties.Elements<DocumentFormat.OpenXml.CustomProperties.CustomDocumentProperty>())
        {
            var name = property.Name?.Value;
            if (name == null)
            {
                continue;
            }

            object? value = property.InnerText;

            // Try to get typed value
            if (property.VTBool != null)
            {
                value = property.VTBool.Text == "true";
            }
            else if (property.VTInt32 != null)
            {
                value = int.Parse(property.VTInt32.Text);
            }
            else if (property.VTFloat != null)
            {
                value = float.Parse(property.VTFloat.Text);
            }
            else if (property.VTDouble != null)
            {
                value = double.Parse(property.VTDouble.Text);
            }
            else if (property.VTDate != null)
            {
                value = property.VTDate.Text;
            }
            else if (property.VTLPWSTR != null)
            {
                value = property.VTLPWSTR.Text;
            }

            properties[name] = value;
        }

        return properties.Count > 0 ? properties : null;
    }

    internal static string? GetWordDocumentText(WordprocessingDocument document)
    {
        var body = document.MainDocumentPart?.Document?.Body;

        if (body == null)
        {
            return null;
        }

        var builder = new StringBuilder();
        foreach (var paragraph in body.Elements<Paragraph>())
        {
            if (AppendWordParagraphText(builder, paragraph))
            {
                builder.AppendLine();
            }
        }

        foreach (var table in body.Elements<WordTable>())
        {
            foreach (var row in table.Elements<TableRow>())
            {
                var firstCell = true;
                var anyCell = false;
                foreach (var cell in row.Elements<TableCell>())
                {
                    if (!firstCell)
                    {
                        builder.Append('\t');
                    }

                    firstCell = false;
                    anyCell = true;

                    foreach (var paragraph in cell.Elements<Paragraph>())
                    {
                        AppendWordParagraphText(builder, paragraph);
                    }
                }

                if (anyCell)
                {
                    builder.AppendLine();
                }
            }
        }

        builder.TrimEnd();
        if (builder.Length == 0)
        {
            return null;
        }

        var result = builder.ToString();
        return string.IsNullOrEmpty(result) ? null : result;
    }

    internal static bool AppendWordParagraphText(StringBuilder builder, Paragraph paragraph)
    {
        var startLength = builder.Length;

        foreach (var run in paragraph.Elements<WordRun>())
        {
            foreach (var text in run.Elements<WordText>())
            {
                builder.Append(text.Text);
            }

            foreach (var _ in run.Elements<TabChar>())
            {
                builder.Append('\t');
            }

            foreach (var wordBreak in run.Elements<WordBreak>())
            {
                if (wordBreak.Type?.Value == BreakValues.Page)
                {
                    builder.AppendLine();
                    builder.AppendLine("--- Page Break ---");
                }
                else
                {
                    builder.AppendLine();
                }
            }
        }

        return builder.Length > startLength;
    }
}
