using DocumentFormat.OpenXml;
using WordFont = DocumentFormat.OpenXml.Wordprocessing.Font;
using WordTable = DocumentFormat.OpenXml.Wordprocessing.Table;
using WordRun = DocumentFormat.OpenXml.Wordprocessing.Run;
using WordText = DocumentFormat.OpenXml.Wordprocessing.Text;
using WordBreak = DocumentFormat.OpenXml.Wordprocessing.Break;
using WordHyperlink = DocumentFormat.OpenXml.Wordprocessing.Hyperlink;

namespace VerifyTests;

public static partial class VerifyOpenXml
{
    static ConversionResult ConvertWord(Stream stream, IReadOnlyDictionary<string, object> settings)
    {
        using var document = WordprocessingDocument.Open(
            stream,
            false,
            new()
            {
                AutoSave = false
            });
        return ConvertWord(document, settings);
    }

    static ConversionResult ConvertWord(WordprocessingDocument document, IReadOnlyDictionary<string, object> settings)
    {
        var info = GetWordInfo(document);
        var text = GetWordDocumentText(document);

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

        // The text is its own target, so it is deliberately absent from the info. Carrying it in both
        // wrote the document text to two snapshot files.
        if (!string.IsNullOrWhiteSpace(text))
        {
            targets.Add(new("txt", text!));
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

    /// <summary>
    /// Document metadata, or null when the document carries none — so no empty info snapshot is written.
    /// </summary>
    static WordInfo? GetWordInfo(WordprocessingDocument document)
    {
        var (fonts, embeddedFonts) = GetWordDocumentFonts(document);
        var properties = GetWordProperties(document);
        var customProperties = GetWordCustomProperties(document);

        if (properties == null &&
            customProperties == null &&
            fonts == null &&
            embeddedFonts == null)
        {
            return null;
        }

        return new()
        {
            Properties = properties,
            CustomProperties = customProperties,
            Fonts = fonts,
            EmbeddedFonts = embeddedFonts
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
        AppendBlocks(builder, body);

        builder.TrimEnd();
        if (builder.Length == 0)
        {
            return null;
        }

        var result = builder.ToString();
        return string.IsNullOrEmpty(result) ? null : result;
    }

    /// <summary>
    /// Block content in document order — a table preceding a paragraph must be emitted first. Content
    /// controls (<c>w:sdt</c>) are transparent: their content is emitted as though the control were not
    /// there, which is what Word displays.
    /// </summary>
    static void AppendBlocks(StringBuilder builder, OpenXmlElement parent)
    {
        foreach (var child in parent.ChildElements)
        {
            switch (child)
            {
                case Paragraph paragraph:
                    if (AppendWordParagraphText(builder, paragraph))
                    {
                        builder.AppendLine();
                    }

                    break;
                case WordTable table:
                    AppendRows(builder, table);
                    break;
                case SdtBlock sdt when Content(sdt) is { } content:
                    AppendBlocks(builder, content);
                    break;
            }
        }
    }

    static void AppendRows(StringBuilder builder, OpenXmlElement parent)
    {
        foreach (var child in parent.ChildElements)
        {
            switch (child)
            {
                case TableRow row:
                    AppendRow(builder, row);
                    break;
                case SdtRow sdt when Content(sdt) is { } content:
                    AppendRows(builder, content);
                    break;
            }
        }
    }

    static void AppendRow(StringBuilder builder, TableRow row)
    {
        var firstCell = true;
        var anyCell = false;

        foreach (var cell in Cells(row))
        {
            if (!firstCell)
            {
                builder.Append('\t');
            }

            firstCell = false;
            anyCell = true;

            AppendCell(builder, cell);
        }

        if (anyCell)
        {
            builder.AppendLine();
        }
    }

    static IEnumerable<TableCell> Cells(OpenXmlElement row)
    {
        foreach (var child in row.ChildElements)
        {
            switch (child)
            {
                case TableCell cell:
                    yield return cell;

                    break;
                case SdtCell sdt when Content(sdt) is { } content:
                    foreach (var nested in Cells(content))
                    {
                        yield return nested;
                    }

                    break;
            }
        }
    }

    // Cell content stays on the row's line, so paragraphs within a cell are concatenated.
    static void AppendCell(StringBuilder builder, OpenXmlElement parent)
    {
        foreach (var child in parent.ChildElements)
        {
            switch (child)
            {
                case Paragraph paragraph:
                    AppendWordParagraphText(builder, paragraph);
                    break;
                case WordTable table:
                    AppendRows(builder, table);
                    break;
                case SdtBlock sdt when Content(sdt) is { } content:
                    AppendCell(builder, content);
                    break;
            }
        }
    }

    internal static bool AppendWordParagraphText(StringBuilder builder, Paragraph paragraph)
    {
        var startLength = builder.Length;

        AppendInline(builder, paragraph);

        return builder.Length > startLength;
    }

    static void AppendInline(StringBuilder builder, OpenXmlElement parent)
    {
        foreach (var child in parent.ChildElements)
        {
            switch (child)
            {
                case WordRun run:
                    AppendRun(builder, run);
                    break;
                case WordHyperlink hyperlink:
                    AppendInline(builder, hyperlink);
                    break;
                case SdtRun sdt when Content(sdt) is { } content:
                    AppendInline(builder, content);
                    break;
            }
        }
    }

    // Run children in document order: text, tabs and breaks can interleave.
    static void AppendRun(StringBuilder builder, WordRun run)
    {
        foreach (var child in run.ChildElements)
        {
            switch (child)
            {
                case WordText text:
                    builder.Append(text.Text);
                    break;
                case TabChar:
                    builder.Append('\t');
                    break;
                case WordBreak wordBreak:
                    if (wordBreak.Type?.Value == BreakValues.Page)
                    {
                        builder.AppendLine();
                        builder.AppendLine("--- Page Break ---");
                    }
                    else
                    {
                        builder.AppendLine();
                    }

                    break;
            }
        }
    }

    static OpenXmlElement? Content(SdtElement sdt) =>
        sdt.ChildElements.FirstOrDefault(_ => _.LocalName == "sdtContent");
}
