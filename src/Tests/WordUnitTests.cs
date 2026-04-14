using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.VariantTypes;
using WordFont = DocumentFormat.OpenXml.Wordprocessing.Font;
using WordText = DocumentFormat.OpenXml.Wordprocessing.Text;
using Body = DocumentFormat.OpenXml.Wordprocessing.Body;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using TabChar = DocumentFormat.OpenXml.Wordprocessing.TabChar;
using Break = DocumentFormat.OpenXml.Wordprocessing.Break;
using BreakValues = DocumentFormat.OpenXml.Wordprocessing.BreakValues;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;
using TableRow = DocumentFormat.OpenXml.Wordprocessing.TableRow;
using TableCell = DocumentFormat.OpenXml.Wordprocessing.TableCell;
using EmbedRegularFont = DocumentFormat.OpenXml.Wordprocessing.EmbedRegularFont;
using EmbedBoldFont = DocumentFormat.OpenXml.Wordprocessing.EmbedBoldFont;
using EmbedItalicFont = DocumentFormat.OpenXml.Wordprocessing.EmbedItalicFont;
using EmbedBoldItalicFont = DocumentFormat.OpenXml.Wordprocessing.EmbedBoldItalicFont;
using CustomProps = DocumentFormat.OpenXml.CustomProperties.Properties;

[TestFixture]
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
    public void GetWordDocumentText_NoMainPart_ReturnsNull()
    {
        var stream = new MemoryStream();
        using var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        Assert.That(VerifyOpenXml.GetWordDocumentText(doc), Is.Null);
    }

    [Test]
    public void GetWordDocumentText_EmptyBody_ReturnsNull()
    {
        using var doc = CreateDoc();
        Assert.That(VerifyOpenXml.GetWordDocumentText(doc), Is.Null);
    }

    [Test]
    public void GetWordDocumentText_ParagraphsOnly()
    {
        var body = new Body(MakeParagraph("Hello"), MakeParagraph("World"));
        using var doc = CreateDoc(body);
        var text = VerifyOpenXml.GetWordDocumentText(doc);
        Assert.That(text, Does.Contain("Hello"));
        Assert.That(text, Does.Contain("World"));
    }

    [Test]
    public void GetWordDocumentText_EmptyParagraphsSkipped()
    {
        var body = new Body(new Paragraph(), MakeParagraph("Only"));
        using var doc = CreateDoc(body);
        var text = VerifyOpenXml.GetWordDocumentText(doc)!;
        Assert.That(text.TrimEnd(), Is.EqualTo("Only"));
    }

    [Test]
    public void GetWordDocumentText_AllEmpty_ReturnsNull()
    {
        var body = new Body(new Paragraph());
        using var doc = CreateDoc(body);
        Assert.That(VerifyOpenXml.GetWordDocumentText(doc), Is.Null);
    }

    [Test]
    public void GetWordDocumentText_WithTable()
    {
        var row1 = new TableRow(
            new TableCell(MakeParagraph("a1")),
            new TableCell(MakeParagraph("b1")));
        var row2 = new TableRow(
            new TableCell(MakeParagraph("a2")),
            new TableCell(MakeParagraph("b2")));
        var body = new Body(new Table(row1, row2));
        using var doc = CreateDoc(body);
        var text = VerifyOpenXml.GetWordDocumentText(doc)!;
        Assert.That(text, Does.Contain("a1\tb1"));
        Assert.That(text, Does.Contain("a2\tb2"));
    }

    [Test]
    public void GetWordParagraphText_TextAndTab()
    {
        var run = new Run(new WordText("Hi"), new TabChar());
        Assert.That(VerifyOpenXml.GetWordParagraphText(new Paragraph(run)), Is.EqualTo("Hi\t"));
    }

    [Test]
    public void GetWordParagraphText_PageBreak()
    {
        var run = new Run(new WordText("Before"), new Break { Type = BreakValues.Page });
        var result = VerifyOpenXml.GetWordParagraphText(new Paragraph(run));
        Assert.That(result, Does.Contain("--- Page Break ---"));
    }

    [Test]
    public void GetWordParagraphText_LineBreak()
    {
        var run = new Run(new WordText("A"), new Break());
        var result = VerifyOpenXml.GetWordParagraphText(new Paragraph(run));
        Assert.That(result, Does.StartWith("A"));
        Assert.That(result, Does.Not.Contain("Page Break"));
    }

    [Test]
    public void GetWordDocumentFonts_NoFontTablePart_ReturnsNulls()
    {
        using var doc = CreateDoc();
        var (fonts, embedded) = VerifyOpenXml.GetWordDocumentFonts(doc);
        Assert.That(fonts, Is.Null);
        Assert.That(embedded, Is.Null);
    }

    [Test]
    public void GetWordDocumentFonts_AllFourEmbedTypes()
    {
        using var doc = CreateDoc();
        var fontPart = doc.MainDocumentPart!.AddNewPart<FontTablePart>();
        fontPart.Fonts = new(
            MakeFont("Regular", new EmbedRegularFont { FontKey = "{x}" }),
            MakeFont("Bold", new EmbedBoldFont { FontKey = "{x}" }),
            MakeFont("Italic", new EmbedItalicFont { FontKey = "{x}" }),
            MakeFont("BoldItalic", new EmbedBoldItalicFont { FontKey = "{x}" }),
            MakeFont("NoEmbed"),
            new WordFont());

        var (fonts, embedded) = VerifyOpenXml.GetWordDocumentFonts(doc);
        Assert.That(fonts, Is.EquivalentTo(["Bold", "BoldItalic", "Italic", "NoEmbed", "Regular"]));
        Assert.That(embedded, Is.EquivalentTo(["Bold", "BoldItalic", "Italic", "Regular"]));
    }

    [Test]
    public void GetWordDocumentFonts_OnlyNullNames_ReturnsNulls()
    {
        using var doc = CreateDoc();
        var fontPart = doc.MainDocumentPart!.AddNewPart<FontTablePart>();
        fontPart.Fonts = new(new WordFont());
        var (fonts, embedded) = VerifyOpenXml.GetWordDocumentFonts(doc);
        Assert.That(fonts, Is.Null);
        Assert.That(embedded, Is.Null);
    }

    [Test]
    public void GetWordProperties_AllEmpty_ReturnsNull()
    {
        using var doc = CreateDoc();
        Assert.That(VerifyOpenXml.GetWordProperties(doc), Is.Null);
    }

    [Test]
    public void GetWordProperties_Populated()
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
        Assert.That(result["Title"], Is.EqualTo("T"));
        Assert.That(result["Subject"], Is.EqualTo("S"));
        Assert.That(result["Creator"], Is.EqualTo("C"));
        Assert.That(result["Keywords"], Is.EqualTo("K"));
        Assert.That(result["Description"], Is.EqualTo("D"));
        Assert.That(result["Category"], Is.EqualTo("Cat"));
        Assert.That(result["LastModifiedBy"], Is.EqualTo("L"));
        Assert.That(result["ContentStatus"], Is.EqualTo("Draft"));
        Assert.That(result["Revision"], Is.EqualTo("1"));
    }

    [Test]
    public void GetWordCustomProperties_NoPart_ReturnsNull()
    {
        using var doc = CreateDoc();
        Assert.That(VerifyOpenXml.GetWordCustomProperties(doc), Is.Null);
    }

    [Test]
    public void GetWordCustomProperties_AllVariantTypes()
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
        Assert.That(result["BoolProp"], Is.EqualTo(true));
        Assert.That(result["IntProp"], Is.EqualTo(42));
        Assert.That(result["FloatProp"], Is.EqualTo(1.5f));
        Assert.That(result["DoubleProp"], Is.EqualTo(2.5d));
        Assert.That(result["DateProp"], Is.EqualTo("2025-01-01T00:00:00Z"));
        Assert.That(result["StringProp"], Is.EqualTo("hello"));
        Assert.That(result["UnknownProp"], Is.EqualTo("raw"));
    }

    [Test]
    public void GetWordCustomProperties_Empty_ReturnsNull()
    {
        using var doc = CreateDoc();
        var part = doc.AddCustomFilePropertiesPart();
        part.Properties = new();
        Assert.That(VerifyOpenXml.GetWordCustomProperties(doc), Is.Null);
    }

    static Paragraph MakeParagraph(string text) =>
        new(new Run(new WordText(text)));

    static WordFont MakeFont(string name, params OpenXmlElement[] children)
    {
        var font = new WordFont { Name = name };
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
