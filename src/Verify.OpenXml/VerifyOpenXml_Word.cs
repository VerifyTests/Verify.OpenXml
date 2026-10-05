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
        // Names the pages, places the text, and says which of them the verification wants
        var conversion = new PagedConversion(settings)
        {
            Info = GetWordInfo(document)
        };

        // The text is read from the body, which knows nothing of where a page ends. Only a renderer
        // does, so the text is that of each page where there is one, and of the document where not.
        var pageText = TextByPage(conversion);
        if (conversion.IncludeText &&
            !pageText)
        {
            conversion.Text(GetWordDocumentText(document));
        }

        // Building the deterministic docx is expensive, so skip it when the docx target is excluded.
        var buildDeterministic = !settings.IsTargetExcluded("docx");
        var render = RenderingEnabled(conversion);

        using var sourceStream = new MemoryStream();
        if (buildDeterministic ||
            render ||
            pageText)
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
            conversion.Source(new("docx", deterministic));
        }

#if NET10_0_OR_GREATER
        // Rendering needs a package stream. Reuse the deterministic docx when built; otherwise render
        // from the raw clone (DeterministicPackage only normalizes zip container metadata, not content,
        // so the rendered pixels are the same either way).
        if (render)
        {
            conversion.AddImages(MorphRenderer.RenderWord(deterministic ?? sourceStream));
        }

        if (pageText)
        {
            // An entry for every page, so the count is there when no page is drawn. A page with
            // both an image and text is added twice, once for each. They are the one page to
            // PagedConversion, which goes by the number.
            var texts = MorphRenderer.WordPageTexts(deterministic ?? sourceStream);
            foreach (var number in conversion.Pages(texts.Count))
            {
                conversion.AddPage(number, text: texts[number - 1]);
            }
        }
#endif

        return conversion.Build();
    }

    /// <summary>
    /// Whether the text is read page by page, so that <c>PagesToInclude</c> limits it as it does
    /// the images. That takes laying the document out, so it is false where there is no renderer:
    /// below <c>net10.0</c>, and when no Morph backend is referenced.
    /// </summary>
    static bool TextByPage(PagedConversion conversion) =>
#if NET10_0_OR_GREATER
        MorphRenderer.Enabled &&
        conversion.IncludeText;
#else
        false;
#endif

    /// <summary>
    /// Document metadata, or null when the document carries none — so no empty <c>Document</c> is written.
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
