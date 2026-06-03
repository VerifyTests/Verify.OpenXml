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

        List<Target> targets =
        [
            new("docx", resultStream)
            {
                BypassComparersForSubsequentOnDifference = true
            }
        ];

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
        foreach (var page in pages)
        {
            targets.Add(new("png", new MemoryStream(page)));
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

    internal static Dictionary<string, object?>? GetWordProperties(WordprocessingDocument document) =>
        GetCoreProperties(document);

    internal static Dictionary<string, object?>? GetWordCustomProperties(WordprocessingDocument document) =>
        ReadCustomProperties(document.CustomFilePropertiesPart);

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
