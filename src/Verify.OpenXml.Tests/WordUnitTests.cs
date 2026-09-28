using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.VariantTypes;
using WordFont = DocumentFormat.OpenXml.Wordprocessing.Font;
using WordText = DocumentFormat.OpenXml.Wordprocessing.Text;
using Body = DocumentFormat.OpenXml.Wordprocessing.Body;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using TabChar = DocumentFormat.OpenXml.Wordprocessing.TabChar;
using Break = DocumentFormat.OpenXml.Wordprocessing.Break;
using BreakValues = DocumentFormat.OpenXml.Wordprocessing.BreakValues;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;
using EmbedRegularFont = DocumentFormat.OpenXml.Wordprocessing.EmbedRegularFont;
using EmbedBoldFont = DocumentFormat.OpenXml.Wordprocessing.EmbedBoldFont;
using EmbedItalicFont = DocumentFormat.OpenXml.Wordprocessing.EmbedItalicFont;
using EmbedBoldItalicFont = DocumentFormat.OpenXml.Wordprocessing.EmbedBoldItalicFont;
using CustomProps = DocumentFormat.OpenXml.CustomProperties.Properties;
public class WordUnitTests
{
    static WordprocessingDocument CreateDoc(Body? body = null)
    {
        var stream = new MemoryStream();
        var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var main = doc.AddMainDocumentPart();
        main.Document = new(body ?? new Body());
        return doc;
    }

    [Test]
    public async Task GetWordDocumentText_NoMainPart_ReturnsNull()
    {
        var stream = new MemoryStream();
        using var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        await Assert.That(VerifyOpenXml.GetWordDocumentText(doc)).IsNull();
    }

    [Test]
    public async Task GetWordDocumentText_EmptyBody_ReturnsNull()
    {
        using var doc = CreateDoc();
        await Assert.That(VerifyOpenXml.GetWordDocumentText(doc)).IsNull();
    }

    [Test]
    public async Task GetWordDocumentText_ParagraphsOnly()
    {
        var body = new Body(MakeParagraph("Hello"), MakeParagraph("World"));
        using var doc = CreateDoc(body);
        var text = VerifyOpenXml.GetWordDocumentText(doc);
        await Assert.That(text).Contains("Hello");
        await Assert.That(text).Contains("World");
    }

    [Test]
    public async Task GetWordDocumentText_EmptyParagraphsSkipped()
    {
        var body = new Body(new Paragraph(), MakeParagraph("Only"));
        using var doc = CreateDoc(body);
        var text = VerifyOpenXml.GetWordDocumentText(doc)!;
        await Assert.That(text.TrimEnd()).IsEqualTo("Only");
    }

    [Test]
    public async Task GetWordDocumentText_AllEmpty_ReturnsNull()
    {
        var body = new Body(new Paragraph());
        using var doc = CreateDoc(body);
        await Assert.That(VerifyOpenXml.GetWordDocumentText(doc)).IsNull();
    }

    [Test]
    public async Task GetWordDocumentText_PreservesDocumentOrder()
    {
        var row = new DocumentFormat.OpenXml.Wordprocessing.TableRow(
            new DocumentFormat.OpenXml.Wordprocessing.TableCell(MakeParagraph("cell")));
        var body = new Body(
            MakeParagraph("before"),
            new Table(row),
            MakeParagraph("after"));
        using var doc = CreateDoc(body);

        var text = VerifyOpenXml.GetWordDocumentText(doc)!;
        var lines = text.Split('\n').Select(_ => _.Trim()).Where(_ => _.Length > 0).ToList();
        await Assert.That(lines).IsEquivalentTo(["before", "cell", "after"], CollectionOrdering.Matching);
    }

    [Test]
    public async Task GetWordDocumentText_InlineContentControl()
    {
        var sdt = new DocumentFormat.OpenXml.Wordprocessing.SdtRun(
            new DocumentFormat.OpenXml.Wordprocessing.SdtContentRun(
                new DocumentFormat.OpenXml.Wordprocessing.Run(new WordText("controlled"))));
        var paragraph = new Paragraph(
            new DocumentFormat.OpenXml.Wordprocessing.Run(new WordText("before ")),
            sdt);
        using var doc = CreateDoc(new(paragraph));

        await Assert.That(VerifyOpenXml.GetWordDocumentText(doc)!.TrimEnd()).IsEqualTo("before controlled");
    }

    [Test]
    public async Task GetWordDocumentText_BlockContentControl()
    {
        var sdt = new DocumentFormat.OpenXml.Wordprocessing.SdtBlock(
            new DocumentFormat.OpenXml.Wordprocessing.SdtContentBlock(MakeParagraph("inside")));
        using var doc = CreateDoc(new(sdt));

        await Assert.That(VerifyOpenXml.GetWordDocumentText(doc)!.TrimEnd()).IsEqualTo("inside");
    }

    [Test]
    public async Task GetWordDocumentText_ContentControlInTableCell()
    {
        var sdt = new DocumentFormat.OpenXml.Wordprocessing.SdtBlock(
            new DocumentFormat.OpenXml.Wordprocessing.SdtContentBlock(MakeParagraph("value")));
        var row = new DocumentFormat.OpenXml.Wordprocessing.TableRow(
            new DocumentFormat.OpenXml.Wordprocessing.TableCell(MakeParagraph("label")),
            new DocumentFormat.OpenXml.Wordprocessing.TableCell(sdt));
        using var doc = CreateDoc(new(new Table(row)));

        await Assert.That(VerifyOpenXml.GetWordDocumentText(doc)!.TrimEnd()).IsEqualTo("label\tvalue");
    }

    [Test]
    public async Task AppendWordParagraphText_Hyperlink()
    {
        var hyperlink = new DocumentFormat.OpenXml.Wordprocessing.Hyperlink(
            new DocumentFormat.OpenXml.Wordprocessing.Run(new WordText("linked")));
        await Assert.That(Render(new(hyperlink))).IsEqualTo("linked");
    }

    [Test]
    public async Task AppendWordParagraphText_InterleavedTextAndTab()
    {
        var run = new DocumentFormat.OpenXml.Wordprocessing.Run(
            new WordText("a"),
            new TabChar(),
            new WordText("b"));
        await Assert.That(Render(new(run))).IsEqualTo("a\tb");
    }

    [Test]
    public async Task GetWordDocumentText_WithTable()
    {
        var row1 = new DocumentFormat.OpenXml.Wordprocessing.TableRow(
            new DocumentFormat.OpenXml.Wordprocessing.TableCell(MakeParagraph("a1")),
            new DocumentFormat.OpenXml.Wordprocessing.TableCell(MakeParagraph("b1")));
        var row2 = new DocumentFormat.OpenXml.Wordprocessing.TableRow(
            new DocumentFormat.OpenXml.Wordprocessing.TableCell(MakeParagraph("a2")),
            new DocumentFormat.OpenXml.Wordprocessing.TableCell(MakeParagraph("b2")));
        var body = new Body(new Table(row1, row2));
        using var doc = CreateDoc(body);
        var text = VerifyOpenXml.GetWordDocumentText(doc)!;
        await Assert.That(text).Contains("a1\tb1");
        await Assert.That(text).Contains("a2\tb2");
    }

    [Test]
    public async Task AppendWordParagraphText_TextAndTab()
    {
        var run = new DocumentFormat.OpenXml.Wordprocessing.Run(new WordText("Hi"), new TabChar());
        await Assert.That(Render(new(run))).IsEqualTo("Hi\t");
    }

    [Test]
    public async Task AppendWordParagraphText_PageBreak()
    {
        var run = new DocumentFormat.OpenXml.Wordprocessing.Run(
            new WordText("Before"),
            new Break
            {
                Type = BreakValues.Page
            });
        await Assert.That(Render(new(run))).Contains("--- Page Break ---");
    }

    [Test]
    public async Task AppendWordParagraphText_LineBreak()
    {
        var run = new DocumentFormat.OpenXml.Wordprocessing.Run(new WordText("A"), new Break());
        var result = Render(new(run));
        await Assert.That(result).StartsWith("A");
        await Assert.That(result).DoesNotContain("Page Break");
    }

    [Test]
    public async Task AppendWordParagraphText_Empty_ReturnsFalse()
    {
        var builder = new StringBuilder();
        await Assert.That(VerifyOpenXml.AppendWordParagraphText(builder, new())).IsFalse();
        await Assert.That(builder.Length).IsZero();
    }

    static string Render(Paragraph paragraph)
    {
        var builder = new StringBuilder();
        VerifyOpenXml.AppendWordParagraphText(builder, paragraph);
        return builder.ToString();
    }

    [Test]
    public async Task GetWordDocumentFonts_NoFontTablePart_ReturnsNulls()
    {
        using var doc = CreateDoc();
        var (fonts, embedded) = VerifyOpenXml.GetWordDocumentFonts(doc);
        await Assert.That(fonts).IsNull();
        await Assert.That(embedded).IsNull();
    }

    [Test]
    public async Task GetWordDocumentFonts_AllFourEmbedTypes()
    {
        using var doc = CreateDoc();
        var fontPart = doc.MainDocumentPart!.AddNewPart<FontTablePart>();
        fontPart.Fonts = new(
            MakeFont(
                "Regular",
                new EmbedRegularFont
                {
                    FontKey = "{x}"
                }),
            MakeFont(
                "Bold",
                new EmbedBoldFont
                {
                    FontKey = "{x}"
                }),
            MakeFont(
                "Italic",
                new EmbedItalicFont
                {
                    FontKey = "{x}"
                }),
            MakeFont(
                "BoldItalic",
                new EmbedBoldItalicFont
                {
                    FontKey = "{x}"
                }),
            MakeFont("NoEmbed"),
            new WordFont());

        var (fonts, embedded) = VerifyOpenXml.GetWordDocumentFonts(doc);
        await Assert.That(fonts).IsEquivalentTo(["Bold", "BoldItalic", "Italic", "NoEmbed", "Regular"]);
        await Assert.That(embedded).IsEquivalentTo(["Bold", "BoldItalic", "Italic", "Regular"]);
    }

    [Test]
    public async Task GetWordDocumentFonts_OnlyNullNames_ReturnsNulls()
    {
        using var doc = CreateDoc();
        var fontPart = doc.MainDocumentPart!.AddNewPart<FontTablePart>();
        fontPart.Fonts = new(new WordFont());
        var (fonts, embedded) = VerifyOpenXml.GetWordDocumentFonts(doc);
        await Assert.That(fonts).IsNull();
        await Assert.That(embedded).IsNull();
    }

    [Test]
    public async Task GetWordProperties_AllEmpty_ReturnsNull()
    {
        using var doc = CreateDoc();
        await Assert.That(VerifyOpenXml.GetWordProperties(doc)).IsNull();
    }

    [Test]
    public async Task GetWordProperties_Populated()
    {
        using var doc = CreateDoc();
        var props = doc.PackageProperties;
        props.Title = "T";
        props.Subject = "S";
        props.Creator = "C";
        props.Keywords = "K";
        props.Description = "D";
        props.Category = "Cat";
        props.LastModifiedBy = "L";
        props.ContentStatus = "Draft";
        props.Revision = "1";

        var result = VerifyOpenXml.GetWordProperties(doc)!;
        await Assert.That(result["Title"]).IsEqualTo("T");
        await Assert.That(result["Subject"]).IsEqualTo("S");
        await Assert.That(result["Keywords"]).IsEqualTo("K");
        await Assert.That(result["Description"]).IsEqualTo("D");
        await Assert.That(result["Category"]).IsEqualTo("Cat");
        await Assert.That(result["ContentStatus"]).IsEqualTo("Draft");
        await Assert.That(result["Revision"]).IsEqualTo("1");
        // Creator and LastModifiedBy are intentionally omitted (DeterministicIoPackaging strips them).
        await Assert.That(result.ContainsKey("Creator")).IsFalse();
        await Assert.That(result.ContainsKey("LastModifiedBy")).IsFalse();
    }

    [Test]
    public async Task GetWordCustomProperties_NoPart_ReturnsNull()
    {
        using var doc = CreateDoc();
        await Assert.That(VerifyOpenXml.GetWordCustomProperties(doc)).IsNull();
    }

    [Test]
    public async Task GetWordCustomProperties_AllVariantTypes()
    {
        using var doc = CreateDoc();
        var part = doc.AddCustomFilePropertiesPart();
        var props = new CustomProps();
        part.Properties = props;
        var pid = 2;
        props.Append(MakeCustomProp("BoolProp", pid++, new VTBool("true")));
        props.Append(MakeCustomProp("IntProp", pid++, new VTInt32("42")));
        props.Append(MakeCustomProp("FloatProp", pid++, new VTFloat("1.5")));
        props.Append(MakeCustomProp("DoubleProp", pid++, new VTDouble("2.5")));
        props.Append(MakeCustomProp("DateProp", pid++, new VTDate("2025-01-01T00:00:00Z")));
        props.Append(MakeCustomProp("StringProp", pid++, new VTLPWSTR("hello")));
        props.Append(MakeCustomProp("UnknownProp", pid++, new VTLPSTR("raw")));
        props.Append(new CustomDocumentProperty
        {
            FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}",
            PropertyId = pid
        });

        var result = VerifyOpenXml.GetWordCustomProperties(doc)!;
        await Assert.That((bool)result["BoolProp"]!).IsTrue();
        await Assert.That(result["IntProp"]).IsEqualTo(42);
        await Assert.That(result["FloatProp"]).IsEqualTo(1.5f);
        await Assert.That(result["DoubleProp"]).IsEqualTo(2.5d);
        await Assert.That(result["DateProp"]).IsEqualTo("2025-01-01T00:00:00Z");
        await Assert.That(result["StringProp"]).IsEqualTo("hello");
        await Assert.That(result["UnknownProp"]).IsEqualTo("raw");
    }

    [Test]
    public async Task GetWordCustomProperties_Empty_ReturnsNull()
    {
        using var doc = CreateDoc();
        var part = doc.AddCustomFilePropertiesPart();
        part.Properties = new();
        await Assert.That(VerifyOpenXml.GetWordCustomProperties(doc)).IsNull();
    }

    static Paragraph MakeParagraph(string text) =>
        new(new DocumentFormat.OpenXml.Wordprocessing.Run(new WordText(text)));

    static WordFont MakeFont(string name, params OpenXmlElement[] children)
    {
        var font = new WordFont
        {
            Name = name
        };
        foreach (var child in children)
        {
            font.Append(child);
        }

        return font;
    }

    static CustomDocumentProperty MakeCustomProp(string name, int pid, OpenXmlElement value)
    {
        var prop = new CustomDocumentProperty
        {
            FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}",
            PropertyId = pid,
            Name = name
        };
        prop.Append(value);
        return prop;
    }
}
