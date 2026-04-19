using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

[TestFixture]
public class PowerpointUnitTests
{
    [Test]
    public void GetPowerpointProperties_AllEmpty_ReturnsNull()
    {
        using var doc = CreateEmptyDoc();
        Assert.That(VerifyOpenXml.GetPowerpointProperties(doc), Is.Null);
    }

    [Test]
    public void GetPowerpointProperties_Populated()
    {
        using var doc = CreateEmptyDoc();
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

        var result = VerifyOpenXml.GetPowerpointProperties(doc)!;
        Assert.That(result["Title"], Is.EqualTo("T"));
        Assert.That(result["Creator"], Is.EqualTo("C"));
        Assert.That(result["Revision"], Is.EqualTo("1"));
    }

    [Test]
    public void GetPowerpointInfo_NoSlides()
    {
        using var doc = CreateEmptyDoc();
        var info = VerifyOpenXml.GetPowerpointInfo(doc);
        Assert.That(info.SlideCount, Is.EqualTo(0));
        Assert.That(info.Text, Is.Null);
    }

    [Test]
    public void GetPowerpointInfo_WithSlidesAndText()
    {
        using var doc = CreateEmptyDoc();
        var presPart = doc.PresentationPart!;
        AddSlide(presPart, "First");
        AddSlide(presPart, "Second");

        var info = VerifyOpenXml.GetPowerpointInfo(doc);
        Assert.That(info.SlideCount, Is.EqualTo(2));
        Assert.That(info.Text, Does.Contain("First"));
        Assert.That(info.Text, Does.Contain("Second"));
        Assert.That(info.Text, Does.Contain("---"));
    }

    [Test]
    public void GetPowerpointInfo_SlideWithNoText_TextIsNull()
    {
        using var doc = CreateEmptyDoc();
        var presPart = doc.PresentationPart!;
        AddEmptySlide(presPart);

        var info = VerifyOpenXml.GetPowerpointInfo(doc);
        Assert.That(info.SlideCount, Is.EqualTo(1));
        Assert.That(info.Text, Is.Null);
    }

    [Test]
    public void AppendSlideText_EmptySlide_ReturnsFalse()
    {
        using var doc = CreateEmptyDoc();
        var slidePart = doc.PresentationPart!.AddNewPart<SlidePart>();
        slidePart.Slide = new(
            new CommonSlideData(new ShapeTree(
                new NonVisualGroupShapeProperties(
                    new NonVisualDrawingProperties { Id = 1, Name = "" },
                    new NonVisualGroupShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()),
                new GroupShapeProperties(new A.TransformGroup()))));

        var builder = new StringBuilder();
        Assert.That(VerifyOpenXml.AppendSlideText(builder, slidePart), Is.False);
        Assert.That(builder.Length, Is.Zero);
    }

    [Test]
    public void AppendSlideText_WithParagraphs()
    {
        using var doc = CreateEmptyDoc();
        var presPart = doc.PresentationPart!;
        var slidePart = AddSlide(presPart, "Line1", "Line2");
        var builder = new StringBuilder();
        Assert.That(VerifyOpenXml.AppendSlideText(builder, slidePart), Is.True);
        var text = builder.ToString();
        Assert.That(text, Does.Contain("Line1"));
        Assert.That(text, Does.Contain("Line2"));
    }

    [Test]
    public void AppendSlideText_ParagraphWithNoText_SkippedFromOutput()
    {
        using var doc = CreateEmptyDoc();
        var presPart = doc.PresentationPart!;
        var slidePart = presPart.AddNewPart<SlidePart>();
        slidePart.Slide = BuildSlide(
            new A.Paragraph(),
            new A.Paragraph(new A.Run(new A.RunProperties(), new A.Text("Only"))));
        var builder = new StringBuilder();
        VerifyOpenXml.AppendSlideText(builder, slidePart);
        Assert.That(builder.ToString(), Is.EqualTo("Only"));
    }

    static PresentationDocument CreateEmptyDoc()
    {
        var doc = PresentationDocument.Create(new MemoryStream(), PresentationDocumentType.Presentation);
        var presPart = doc.AddPresentationPart();
        presPart.Presentation = new(new SlideIdList());
        return doc;
    }

    static SlidePart AddSlide(PresentationPart presPart, params string[] lines)
    {
        var slidePart = presPart.AddNewPart<SlidePart>();
        var paragraphs = lines.Select(_ =>
            (OpenXmlElement) new A.Paragraph(
                new A.Run(new A.RunProperties(), new A.Text(_)))).ToArray();
        slidePart.Slide = BuildSlide(paragraphs);
        return slidePart;
    }

    static SlidePart AddEmptySlide(PresentationPart presPart)
    {
        var slidePart = presPart.AddNewPart<SlidePart>();
        slidePart.Slide = BuildSlide();
        return slidePart;
    }

    static Slide BuildSlide(params OpenXmlElement[] paragraphs)
    {
        var textBody = new TextBody(new A.BodyProperties(), new A.ListStyle());
        foreach (var p in paragraphs)
        {
            textBody.Append(p);
        }

        var shape = new Shape(
            new NonVisualShapeProperties(
                new NonVisualDrawingProperties { Id = 2, Name = "Text" },
                new NonVisualShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new ShapeProperties(),
            textBody);

        return new(
            new CommonSlideData(
                new ShapeTree(
                    new NonVisualGroupShapeProperties(
                        new NonVisualDrawingProperties { Id = 1, Name = "" },
                        new NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(new A.TransformGroup()),
                    shape)));
    }
}
