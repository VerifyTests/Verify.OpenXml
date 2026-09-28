using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
public class PowerpointUnitTests
{
    [Test]
    public async Task GetPowerpointProperties_AllEmpty_ReturnsNull()
    {
        using var doc = CreateEmptyDoc();
        await Assert.That(VerifyOpenXml.GetPowerpointProperties(doc)).IsNull();
    }

    [Test]
    public async Task GetPowerpointProperties_Populated()
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
        await Assert.That(result["Title"]).IsEqualTo("T");
        await Assert.That(result["Revision"]).IsEqualTo("1");
        // Creator and LastModifiedBy are intentionally omitted (DeterministicIoPackaging strips them).
        await Assert.That(result.ContainsKey("Creator")).IsFalse();
        await Assert.That(result.ContainsKey("LastModifiedBy")).IsFalse();
    }

    [Test]
    public async Task GetPowerpointInfo_NoSlides()
    {
        using var doc = CreateEmptyDoc();
        var info = VerifyOpenXml.GetPowerpointInfo(doc);
        await Assert.That(info.SlideCount).IsEqualTo(0);
        await Assert.That(VerifyOpenXml.GetPowerpointText(doc)).IsNull();
    }

    [Test]
    public async Task GetPowerpointInfo_WithSlidesAndText()
    {
        using var doc = CreateEmptyDoc();
        var presPart = doc.PresentationPart!;
        AddSlide(presPart, "First");
        AddSlide(presPart, "Second");

        var info = VerifyOpenXml.GetPowerpointInfo(doc);
        await Assert.That(info.SlideCount).IsEqualTo(2);

        var text = VerifyOpenXml.GetPowerpointText(doc);
        await Assert.That(text).Contains("First");
        await Assert.That(text).Contains("Second");
        await Assert.That(text).Contains("---");
    }

    [Test]
    public async Task GetPowerpointInfo_SlideWithNoText_TextIsNull()
    {
        using var doc = CreateEmptyDoc();
        var presPart = doc.PresentationPart!;
        AddEmptySlide(presPart);

        var info = VerifyOpenXml.GetPowerpointInfo(doc);
        await Assert.That(info.SlideCount).IsEqualTo(1);
        await Assert.That(VerifyOpenXml.GetPowerpointText(doc)).IsNull();
    }

    [Test]
    public async Task AppendSlideText_EmptySlide_ReturnsFalse()
    {
        using var doc = CreateEmptyDoc();
        var slidePart = doc.PresentationPart!.AddNewPart<SlidePart>();
        slidePart.Slide = new(
            new CommonSlideData(new ShapeTree(
                new NonVisualGroupShapeProperties(
                    new NonVisualDrawingProperties
                    {
                        Id = 1,
                        Name = ""
                    },
                    new NonVisualGroupShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()),
                new GroupShapeProperties(new A.TransformGroup()))));

        var builder = new StringBuilder();
        await Assert.That(VerifyOpenXml.AppendSlideText(builder, slidePart)).IsFalse();
        await Assert.That(builder.Length).IsZero();
    }

    [Test]
    public async Task AppendSlideText_WithParagraphs()
    {
        using var doc = CreateEmptyDoc();
        var presPart = doc.PresentationPart!;
        var slidePart = AddSlide(presPart, "Line1", "Line2");
        var builder = new StringBuilder();
        await Assert.That(VerifyOpenXml.AppendSlideText(builder, slidePart)).IsTrue();
        var text = builder.ToString();
        await Assert.That(text).Contains("Line1");
        await Assert.That(text).Contains("Line2");
    }

    [Test]
    public async Task AppendSlideText_ParagraphWithNoText_SkippedFromOutput()
    {
        using var doc = CreateEmptyDoc();
        var presPart = doc.PresentationPart!;
        var slidePart = presPart.AddNewPart<SlidePart>();
        slidePart.Slide = BuildSlide(
            new A.Paragraph(),
            new A.Paragraph(new A.Run(new A.RunProperties(), new A.Text("Only"))));
        var builder = new StringBuilder();
        VerifyOpenXml.AppendSlideText(builder, slidePart);
        await Assert.That(builder.ToString()).IsEqualTo("Only");
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
        var paragraphs = lines.Select(OpenXmlElement (_) =>
            new A.Paragraph(
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
                new NonVisualDrawingProperties
                {
                    Id = 2,
                    Name = "Text"
                },
                new NonVisualShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new ShapeProperties(),
            textBody);

        return new(
            new CommonSlideData(
                new ShapeTree(
                    new NonVisualGroupShapeProperties(
                        new NonVisualDrawingProperties
                        {
                            Id = 1,
                            Name = ""
                        },
                        new NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(new A.TransformGroup()),
                    shape)));
    }
}
